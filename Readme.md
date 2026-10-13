# Supply Chain Delay Analyzer

A Python tool that reads shipment spreadsheets, measures delivery delays, finds the root causes, draws a summary chart, and (optionally) asks a local AI model to write a plain-English report for managers.

Pandas does all the calculations. The AI only explains numbers that were already computed, so the figures stay reliable even with a small model.

## Features

- Reads `.xlsx`, `.xlsm`, `.xls` and `.csv` files (every sheet of an Excel file is combined)
- Flexible column matching, so headers like `PO Number`, `Order ID` or `Shipment ID` all work
- Per-shipment delay in days (positive = late, negative = early)
- Key delivery metrics:
  - **OTD**: On-Time Delivery rate
  - **OTIF**: On-Time In-Full rate
  - **Fill rate**: delivered quantity vs ordered quantity
  - Average delay of late orders, maximum delay, short-shipped orders
- Root cause analysis:
  - **By stage**, if the file has `<Stage> Planned` / `<Stage> Actual` column pairs
  - **By reason, carrier, warehouse, etc.**, detected automatically from text columns
- 3-panel PNG chart: delay distribution, delivery performance, root cause
- Optional AI-written summary using a local [Ollama](https://ollama.com) model (no data leaves your machine)
- Processes multiple files in one run and prints a comparison table

## Requirements

- Python 3.9+
- Packages: `pandas`, `matplotlib`, `openpyxl`, `xlrd`
- (Optional) [Ollama](https://ollama.com) for the AI summary

## Installation

```bash
git clone https://github.com/YOUR-USERNAME/YOUR-REPO.git
cd YOUR-REPO
pip install -r requirements.txt
```

`requirements.txt`:

```
pandas
matplotlib
openpyxl
xlrd
```

## Usage

1. Put your spreadsheet files in the `input_files/` folder (it is created on first run).
2. Run the script:

   ```bash
   python delay_analysis.py
   ```

3. Read the console report and open the chart at `delay_output/<file name>/delay_analysis.png`.

### Using the AI summary

1. Install Ollama and start it:

   ```bash
   ollama serve
   ```

2. The script checks for the model set in `MODEL` (default `qwen2.5-coder:3b`) and downloads it automatically the first time.

To skip the AI part, set `USE_AI = False` at the top of the script.

## Required columns

Your file needs these five columns. Header names are not case sensitive, and underscores and `*` are ignored.

| Meaning | Accepted header names |
|---|---|
| Order ID | order id, orderid, po, po number, po no, order, shipment id |
| Promised date | promised date, promised, due date, expected date, planned date |
| Actual date | actual date, actual, delivery date, delivered date, received date |
| Ordered quantity | ordered qty, ordered quantity, order qty, qty ordered |
| Delivered quantity | delivered qty, delivered quantity, received qty, qty delivered |

If one is missing, that file is skipped and the others still run. You can add more names in `COLUMN_ALIASES`.

### Optional columns

- **Stage columns**: pairs such as `Picking Planned` and `Picking Actual` (in days) enable the stage bottleneck analysis.
- **Factor columns**: any text column with 2 to 20 distinct values (carrier, warehouse, region...) is analysed automatically.
- **Delay reason column**: a column whose header contains the word "reason" is ranked by total delay days.

## Output

For each input file:

- A console report with four sections:
  1. Shipment-by-shipment delay
  2. Key metrics
  3. Root cause analysis
  4. AI overview (if enabled)
- `delay_output/<file name>/delay_analysis.png`

## Configuration

Edit these values at the top of the script:

| Setting | Purpose | Default |
|---|---|---|
| `INPUT_DIR` | Folder to scan for files | `input_files` |
| `OUTPUT_DIR` | Folder for charts | `delay_output` |
| `MODEL` | Ollama model name | `qwen2.5-coder:3b` |
| `OLLAMA_URL` | Ollama server address | `http://localhost:11434` |
| `USE_AI` | Turn the AI summary on or off | `True` |

## Notes and limitations

- Delays are calculated in whole days, so time of day is ignored.
- Rows with missing or invalid dates can cause errors, so clean the data first.
- Other factors need at least 30 orders to appear in rankings, to avoid misleading small samples.
- A small local model can still make mistakes. Always check the AI text against the metrics in sections 2 and 3.

## Privacy

Everything runs locally. Do not commit your input data. Keep `input_files/` and `delay_output/` in `.gitignore`.

## License

Add a license of your choice (for example MIT).