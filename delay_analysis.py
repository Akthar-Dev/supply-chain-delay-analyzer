
import json
import urllib.request
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                      
import matplotlib.pyplot as plt
import pandas as pd

INPUT_DIR = Path("input_files")
OUTPUT_DIR = Path("delay_output")
MODEL = "qwen2.5-coder:3b"
OLLAMA_URL = "http://localhost:11434"
USE_AI = True                              

SUPPORTED = (".xlsx", ".xlsm", ".xls", ".csv")

COLUMN_ALIASES = {
    "order_id": ["order id", "orderid", "po", "po number", "po no", "order", "shipment id"],
    "promised": ["promised date", "promised", "due date", "expected date", "planned date"],
    "actual": ["actual date", "actual", "delivery date", "delivered date", "received date"],
    "ordered_qty": ["ordered qty", "ordered quantity", "order qty", "qty ordered"],
    "delivered_qty": ["delivered qty", "delivered quantity", "received qty", "qty delivered"],
}
CORE_COLUMNS = set(COLUMN_ALIASES)
OUTCOME_VALUES = {"delayed", "delivered", "partially delivered", "in transit"}
NO_REASON = {"no delay", "none", "nan", "n/a", "na", "no reason", "on time", "not delayed"}
COLORS = {"Late": "#d9534f", "On time": "#5cb85c", "Early": "#5bc0de"}



def load_file(path: Path) -> pd.DataFrame:
    """Read a CSV, or every sheet of an Excel file stacked together."""
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    sheets = pd.read_excel(path, sheet_name=None)
    return pd.concat([s for s in sheets.values() if not s.empty], ignore_index=True)


def standardise(df: pd.DataFrame):
    """Rename the five core columns, clean their types, and detect stage columns."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lowered = {c.lower().replace("_", " ").replace("*", "").strip(): c for c in df.columns}

    rename = {}
    for std, aliases in COLUMN_ALIASES.items():
        match = next((lowered[a] for a in aliases if a in lowered), None)
        if match is None:
            raise ValueError(f"no '{std}' column found (looked for: {', '.join(aliases)})")
        rename[match] = std
    df = df.rename(columns=rename)

    df["order_id"] = df["order_id"].astype(str).str.replace("*", "", regex=False).str.strip()
    df["promised"] = pd.to_datetime(df["promised"])
    df["actual"] = pd.to_datetime(df["actual"])
    for col in ("ordered_qty", "delivered_qty"):
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ""))

  
    stages = {}
    for col in df.columns:
        for kind in ("planned", "actual"):
            if col.lower().endswith(" " + kind):
                stages.setdefault(col[: -len(kind)].strip(), {})[kind] = col
    stages = {name: cols for name, cols in stages.items() if len(cols) == 2}
    return df, stages

def compute_metrics(df: pd.DataFrame):
    """Add delay columns to every shipment and return the summary KPIs."""
    df = df.copy()
    df["delay_days"] = (df["actual"] - df["promised"]).dt.days        
    df["on_time"] = df["delay_days"] <= 0
    df["in_full"] = df["delivered_qty"] >= df["ordered_qty"]
    df["otif"] = df["on_time"] & df["in_full"]
    df["status"] = df["delay_days"].apply(
        lambda d: "Late" if d > 0 else ("Early" if d < 0 else "On time"))

    late = df[df["delay_days"] > 0]
    summary = {
        "total_orders": len(df),
        "late_orders": len(late),
        "on_time_orders": int(df["on_time"].sum()),
        "avg_delay_late_orders_days": round(late["delay_days"].mean(), 2) if len(late) else 0.0,
        "max_delay_days": int(df["delay_days"].max()),
        "otd_rate_pct": round(df["on_time"].mean() * 100, 1),
        "otif_rate_pct": round(df["otif"].mean() * 100, 1),
        "fill_rate_pct": round(df["delivered_qty"].sum() / df["ordered_qty"].sum() * 100, 1),
        "short_shipped_orders": int((~df["in_full"]).sum()),
        "on_time_but_short_orders": int((df["on_time"] & ~df["in_full"]).sum()),
    }
    return df, summary


def stage_root_cause(df: pd.DataFrame, stages: dict):
    """Extra days (actual - planned) per stage, for late orders only."""
    if not stages:
        return None
    late = df[df["delay_days"] > 0]
    rows = []
    for stage, cols in stages.items():
        planned = pd.to_numeric(late[cols["planned"]], errors="coerce")
        actual = pd.to_numeric(late[cols["actual"]], errors="coerce")
        variance = actual - planned
        rows.append({
            "stage": stage,
            "avg_planned_days": round(planned.mean(), 2),
            "avg_actual_days": round(actual.mean(), 2),
            "avg_variance_days": round(variance.mean(), 2),
            "total_variance_days": round(variance.sum(), 2),
        })
    rc = pd.DataFrame(rows).fillna(0)
    extra = rc["total_variance_days"].clip(lower=0)
    rc["share_of_delay_pct"] = (extra / extra.sum() * 100).round(1) if extra.sum() else 0.0
    return rc.sort_values("total_variance_days", ascending=False).reset_index(drop=True)


def find_factor_columns(df: pd.DataFrame, stages: dict) -> list:
    """Text columns worth grouping by (reason, carrier, warehouse ...).
    Skips IDs, outcome columns and columns that just repeat another one."""
    skip = CORE_COLUMNS | {c for cols in stages.values() for c in cols.values()}
    found = []
    for c in df.columns:
        is_text = df[c].dtype == object or str(df[c].dtype).startswith(("str", "category"))
        if c in skip or not is_text:
            continue
        name = c.lower()
        if "on-time" in name or "on time" in name:           
            continue
        if set(df[c].astype(str).str.lower()) & OUTCOME_VALUES:
            continue
        if 2 <= df[c].nunique() <= 20:                       
            found.append(c)

    keep = []
    for c in found:
        redundant = False
        for k in list(keep):
            n_c, n_k = df[c].nunique(), df[k].nunique()
            if len(df[[c, k]].dropna().drop_duplicates()) == max(n_c, n_k):
                if n_c >= n_k:
                    redundant = True
                else:
                    keep.remove(k)
                break
        if not redundant:
            keep.append(c)
    return keep


def factor_table(df: pd.DataFrame, factors: list):
    """For every value of every factor: orders, late orders, average delay,
    gap vs the overall average, and share of all delay days."""
    overall_avg = df["delay_days"].mean()
    overall_total = df["delay_days"].clip(lower=0).sum()
    rows = []

    def add(factor, value, group):
        total = group["delay_days"].clip(lower=0).sum()
        rows.append({
            "factor": factor, "value": str(value), "orders": len(group),
            "late_orders": int((group["delay_days"] > 0).sum()),
            "avg_delay_days": round(group["delay_days"].mean(), 2),
            "gap_vs_avg_days": round(group["delay_days"].mean() - overall_avg, 2),
            "total_delay_days": int(total),
            "share_of_delay_pct": round(total / overall_total * 100, 1) if overall_total else 0.0,
        })

    for col in factors:
        for value, group in df.groupby(col):
            add(col, value, group)
    for label, group in (("Short-shipped", df[~df["in_full"]]), ("Full quantity", df[df["in_full"]])):
        if len(group):
            add("Quantity", label, group)
    if not rows:
        return None
    return pd.DataFrame(rows).sort_values("avg_delay_days", ascending=False).reset_index(drop=True)


def split_factors(table):
    """(delay reasons ranked by total delay days, other factors ranked by gap vs average)."""
    if table is None:
        return None, None
    is_reason = table["factor"].str.lower().str.contains("reason")
    reasons = table[is_reason & ~table["value"].str.lower().isin(NO_REASON)]
    reasons = reasons.sort_values("total_delay_days", ascending=False)
    others = table[~is_reason & (table["orders"] >= 30)]
    others = others.reindex(others["gap_vs_avg_days"].abs().sort_values(ascending=False).index)
    return (reasons if len(reasons) else None), (others if len(others) else None)


def make_chart(df, summary, stage_rc, table, out_png: Path):
    has_cause_panel = stage_rc is not None or table is not None
    fig, axes = plt.subplots(1, 3 if has_cause_panel else 2,
                             figsize=(18 if has_cause_panel else 12, 5.2))

   
    ax = axes[0]
    if len(df) > 40:
        counts = df["delay_days"].value_counts().sort_index()
        ax.bar(counts.index, counts.values, width=0.8,
               color=[COLORS["Late"] if d > 0 else COLORS["On time"] if d == 0 else COLORS["Early"]
                      for d in counts.index])
        ax.set_title(f"Delay Distribution - {len(df)} shipments\n(days vs promised date; + = late)")
        ax.set_xlabel("Days vs promised date")
        ax.set_ylabel("Number of shipments")
        if len(counts) <= 25:
            for x, v in zip(counts.index, counts.values):
                ax.text(x, v, str(v), ha="center", va="bottom", fontsize=8)
    else:
        bars = ax.bar(df["order_id"], df["delay_days"], color=[COLORS[s] for s in df["status"]])
        ax.axhline(0, color="black", linewidth=0.8)
        ax.set_title("Delay per Shipment (days)\n+ = late, - = early")
        ax.set_ylabel("Days vs promised date")
        ax.tick_params(axis="x", rotation=45)
        for bar, v in zip(bars, df["delay_days"]):
            ax.annotate(f"{v:+d}", (bar.get_x() + bar.get_width() / 2, v),
                        ha="center", va="bottom" if v >= 0 else "top", fontsize=9)
    ax.legend([plt.Rectangle((0, 0), 1, 1, color=c) for c in COLORS.values()],
              COLORS.keys(), loc="best", fontsize=8)

  
    ax = axes[1]
    rates = [summary["otd_rate_pct"], summary["otif_rate_pct"], summary["fill_rate_pct"]]
    bars = ax.bar(["On-Time\nDelivery", "OTIF", "Fill Rate"], rates,
                  color=["#4c72b0", "#dd8452", "#55a868"])
    ax.set_ylim(0, 110)
    ax.set_ylabel("%")
    ax.set_title(f"Delivery Performance\n(Avg delay of late orders: "
                 f"{summary['avg_delay_late_orders_days']} days)")
    for bar, v in zip(bars, rates):
        ax.text(bar.get_x() + bar.get_width() / 2, v + 1.5, f"{v}%", ha="center")

   
    if stage_rc is not None:
        ax = axes[2]
        r = stage_rc.sort_values("total_variance_days")
        ax.barh(r["stage"], r["total_variance_days"],
                color=["#d9534f" if v > 0 else "#999999" for v in r["total_variance_days"]])
        ax.set_title("Root Cause: Stage Delay Contribution\n(late orders, total extra days)")
        ax.set_xlabel("Extra days vs plan")
        for y, v in enumerate(r["total_variance_days"]):
            ax.text(v + 0.05, y, f"{v:+.1f}", va="center")
    elif table is not None:
        ax = axes[2]
        reasons, _ = split_factors(table)
        if reasons is not None:
            r = reasons.head(8).iloc[::-1]
            biggest = r["total_delay_days"].max()
            ax.barh(r["value"], r["total_delay_days"], color="#d9534f")
            ax.set_xlim(0, biggest * 1.6)
            for y, (v, share) in enumerate(zip(r["total_delay_days"], r["share_of_delay_pct"])):
                ax.text(v + biggest * 0.01, y, f"{v:,} days ({share}%)", va="center", fontsize=9)
            ax.set_title("Root Cause: Total Delay Days by Reason")
            ax.set_xlabel("Total days of delay")
        else:
            r = table.head(8).iloc[::-1]
            ax.barh(r["factor"] + ": " + r["value"], r["avg_delay_days"],
                    color=["#d9534f" if v > 0 else "#5cb85c" for v in r["avg_delay_days"]])
            ax.set_title("Root Cause: Avg Delay by Factor (days)")
            ax.set_xlabel("Average delay (days)")
            for y, v in enumerate(r["avg_delay_days"]):
                ax.text(v + 0.05 if v >= 0 else v - 0.05, y, f"{v:+.1f}", va="center",
                        ha="left" if v >= 0 else "right")

    fig.suptitle("Supply Chain Delay Analysis", fontsize=15, fontweight="bold")
    fig.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)



def ensure_model() -> bool:
    """Check that Ollama is running and the model is installed (pull it if not)."""
    try:
        with urllib.request.urlopen(OLLAMA_URL + "/api/tags", timeout=10) as r:
            installed = [m["name"] for m in json.loads(r.read()).get("models", [])]
    except Exception as e:
        print(f"[Ollama not reachable at {OLLAMA_URL}. Start it with `ollama serve`. ({e})]")
        return False

    if (MODEL if ":" in MODEL else MODEL + ":latest") in installed:
        print(f"Model '{MODEL}' found in local Ollama.")
        return True

    print(f"Model '{MODEL}' not installed - pulling it now (one-time download)...")
    try:
        req = urllib.request.Request(OLLAMA_URL + "/api/pull",
                                     data=json.dumps({"name": MODEL}).encode(),
                                     headers={"Content-Type": "application/json"})
        last = ""
        with urllib.request.urlopen(req, timeout=3600) as r:
            for line in r:
                status = json.loads(line).get("status", "")
                if status != last:
                    print("   ", status)
                    last = status
        return True
    except Exception as e:
        print(f"[Could not pull model: {e}]")
        return False


def ask_ollama(prompt: str) -> str:
    payload = json.dumps({
        "model": MODEL,
        "stream": False,
        "options": {"temperature": 0.2, "num_ctx": 4096, "num_predict": 700},
        "messages": [
            {"role": "system", "content":
             "You are a supply chain analyst. Use ONLY the numbers provided. "
             "Do not invent data or causes that are not in the data. "
             "Be concise and practical. Follow the requested format exactly."},
            {"role": "user", "content": prompt},
        ],
    }).encode()
    req = urllib.request.Request(OLLAMA_URL + "/api/chat", data=payload,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            return json.loads(resp.read())["message"]["content"].strip()
    except Exception as e:
        return (f"[AI unavailable: the request to Ollama failed ({e}). "
                f"Check that `ollama serve` is running and that MODEL '{MODEL}' appears in `ollama list`.]")


def build_prompt(summary, stage_rc, table) -> str:
    """Give the model exact, pre-calculated facts so it only has to explain them."""
    facts = [
        f"- Orders analysed: {summary['total_orders']}. Late orders: {summary['late_orders']}. "
        f"On-time orders: {summary['on_time_orders']}.",
        f"- Average delay of late orders: {summary['avg_delay_late_orders_days']} days "
        f"(longest delay {summary['max_delay_days']} days).",
        f"- On-time delivery rate: {summary['otd_rate_pct']}%. "
        f"On-time in-full rate: {summary['otif_rate_pct']}%. Quantity fill rate: {summary['fill_rate_pct']}%.",
        f"- Short-shipped orders: {summary['short_shipped_orders']}; of these, "
        f"{summary['on_time_but_short_orders']} arrived on time but still failed the in-full test.",
    ]
    parts = ["FACTS (copy numbers exactly; do not calculate new numbers or percentages):\n"
             + "\n".join(facts)]
    if stage_rc is not None:
        parts.append("Stage-wise delay (late orders, days):\n" + stage_rc.to_string(index=False))

    reasons, others = split_factors(table)
    if reasons is not None:
        cols = ["value", "orders", "avg_delay_days", "total_delay_days", "share_of_delay_pct"]
        parts.append("Delay reasons ranked by share of total delay days "
                     "(orders = number of orders with that reason):\n"
                     + reasons[cols].rename(columns={"value": "reason"}).to_string(index=False))
    if others is not None:
        cols = ["factor", "value", "orders", "avg_delay_days", "gap_vs_avg_days"]
        parts.append("Other factors, biggest gap vs overall average delay "
                     "(a gap under 0.5 day is NOT meaningful):\n"
                     + others.head(6)[cols].to_string(index=False))
    parts.append(
        "Write a short report for a manager in plain English with exactly these 4 headings:\n"
        "1. OVERVIEW - 3 sentences using the FACTS.\n"
        "2. MAIN DELAY CAUSES - the top 3 reasons by share of total delay days, with their numbers.\n"
        "3. OTHER ISSUES - short shipments, and say whether other factors show any meaningful difference.\n"
        "4. RECOMMENDATIONS - 3 specific actions, each tied to one of the top causes.")
    return "\n\n".join(parts)


def print_shipments(df):
    print("\n1. SHIPMENT-BY-SHIPMENT DELAY")
    print("-" * 60)
    show = df[["order_id", "promised", "actual", "delay_days", "status",
               "ordered_qty", "delivered_qty", "otif"]].copy()
    show["promised"] = show["promised"].dt.strftime("%b %d, %Y")
    show["actual"] = show["actual"].dt.strftime("%b %d, %Y")
    if len(show) > 20:
        print(show.head(20).to_string(index=False))
        print(f"... {len(show) - 20} more shipments not shown")
    else:
        print(show.to_string(index=False))


def print_metrics(df, summary):
    print("\n2. KEY METRICS")
    print("-" * 60)
    ids = list(df.loc[df["delay_days"] > 0, "order_id"])
    late_ids = ", ".join(ids[:5]) + (f" ... +{len(ids) - 5} more" if len(ids) > 5 else "")
    n = summary["total_orders"]
    print(f"Average Delay Time (late orders only: {late_ids or 'none'}): "
          f"{summary['avg_delay_late_orders_days']} days")
    print(f"On-Time Delivery rate (OTD):  {summary['on_time_orders']}/{n} = {summary['otd_rate_pct']}%")
    print(f"On-Time In-Full rate (OTIF):  {int(df['otif'].sum())}/{n} = {summary['otif_rate_pct']}%")
    print(f"Quantity fill rate:           {summary['fill_rate_pct']}%")
    print(f"Short-shipped orders:         {summary['short_shipped_orders']}")


def print_root_cause(stage_rc, table):
    print("\n3. ROOT CAUSE ANALYSIS")
    print("-" * 60)
    if stage_rc is not None:
        print("Stage-wise lead time variance (late orders):")
        print(stage_rc.to_string(index=False))
        top = stage_rc.iloc[0]
        print(f"\nBiggest bottleneck: {top['stage']} ({top['share_of_delay_pct']}% of stage delay)")
    else:
        print("No '<Stage> Planned' / '<Stage> Actual' columns found - stage analysis skipped.")
    if table is None:
        return

    reasons, others = split_factors(table)
    if reasons is not None:
        print("\nDelay reasons ranked by total delay days:")
        cols = ["value", "orders", "avg_delay_days", "total_delay_days", "share_of_delay_pct"]
        print(reasons[cols].rename(columns={"value": "reason"}).to_string(index=False))
        missing = table[table["factor"].str.lower().str.contains("reason")
                        & table["value"].str.lower().isin(NO_REASON) & (table["late_orders"] > 0)]
        if len(missing):
            print(f"({int(missing['late_orders'].sum())} late orders have no real reason recorded "
                  f"= {missing['share_of_delay_pct'].sum():.1f}% of all delay days)")
    if others is not None:
        print("\nOther factors with the biggest gap vs the overall average "
              "(a gap under ~0.5 day is not meaningful):")
        print(others.head(10)[["factor", "value", "orders", "avg_delay_days",
                               "gap_vs_avg_days"]].to_string(index=False))
    print(f"\n(Showing the top 10 of {len(table)} factors.)")



def analyse(raw: pd.DataFrame, title: str, out_dir: Path, ai_ready: bool) -> dict:
    df, stages = standardise(raw)
    factors = find_factor_columns(df, stages)      
    df, summary = compute_metrics(df)
    stage_rc = stage_root_cause(df, stages)
    table = factor_table(df, factors)

    print("\n" + "#" * 70)
    print(f"#  {title}")
    print("#" * 70)
    print_shipments(df)
    print_metrics(df, summary)
    print_root_cause(stage_rc, table)

    out_dir.mkdir(parents=True, exist_ok=True)
    png = out_dir / "delay_analysis.png"
    make_chart(df, summary, stage_rc, table, png)

    if ai_ready:
        print(f"\n4. AI OVERVIEW OF DELAY REASONS  (model: {MODEL})")
        print("-" * 60)
        print("Asking local model, please wait...\n")
        print(ask_ollama(build_prompt(summary, stage_rc, table)))
        print("\n[Note: written by a small local model - check numbers against sections 2 and 3.]")

    print(f"\nSaved chart  -> {png}")
    return {"name": title, **summary}


def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    files = [f for f in sorted(INPUT_DIR.rglob("*"))
             if f.is_file() and f.suffix.lower() in SUPPORTED and not f.name.startswith(("~$", "."))]
    if not files:
        print(f"No spreadsheet files found in '{INPUT_DIR.resolve()}'.")
        print("Put your .xlsx / .xls / .csv files in that folder and run the script again.")
        return

    print(f"Scanning '{INPUT_DIR.resolve()}' ... found {len(files)} file(s):")
    for f in files:
        print(f"   - {f.relative_to(INPUT_DIR)}")

    ai_ready = ensure_model() if USE_AI else False

    results = []
    for f in files:
        try:
            folder = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in f.stem)
            results.append(analyse(load_file(f), f"FILE: {f.name}", OUTPUT_DIR / folder, ai_ready))
        except Exception as e:                  
            print(f"\n  ! Skipping {f.name}: {e}")

    if len(results) > 1:
        cols = ["name", "total_orders", "late_orders", "avg_delay_late_orders_days",
                "otd_rate_pct", "otif_rate_pct", "fill_rate_pct"]
        print("\n" + "=" * 70)
        print("OVERALL COMPARISON ACROSS FILES")
        print("=" * 70)
        print(pd.DataFrame(results)[cols].to_string(index=False))


if __name__ == "__main__":
    main()