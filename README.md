## 📌 Project Overview
This project is a **Guild Hunt Report ETL Pipeline** built to eliminate manual Excel cleanup for guild/operations tracking.

Instead of someone manually opening a raw hunt-tracker export, deleting junk rows, fixing percentage formatting, and color-coding performance by hand every week, this **Python-based ETL (Extract, Transform, Load) pipeline** does it automatically. Point it at a raw `.xlsx` export and it produces a clean, filtered, formatted Excel report ready to share.

---

## 💼 What It Solves
* **No More Manual Cleanup:** Removes stray "Total" rows, inconsistent percentage formats, and messy columns automatically.
* **Fair Reporting:** Automatically filters out players who haven't been in the guild long enough (or have no hunt-date data at all) so they don't skew the report.
* **Consistent Formatting:** Every report comes out with the same header styling, rank highlighting, and color-coded performance tiers — no manual formatting required.

---

## ⚙️ How It Works (The ETL Logic)

### 🔹 Phase 1: Extract (Data Ingestion)
* Reads a raw `.xlsx` hunt export directly into a Pandas DataFrame via `pandas.read_excel`.
* Validates that the required columns (`Name`, `Points (Hunt)`, `Goal Percentage (Hunt)`) are present before continuing.

### 🔹 Phase 2: Transform (Data Processing)
Once the file is loaded, **Pandas** takes over:
* **Cleaning:** Drops any `Total` summary row left over from the source export.
* **Guild-Presence Filtering:** Compares `First Hunt Time` and `Last Hunt Time` to calculate how many days each player has actually been active, and filters out:
  * Players with fewer than 4 days in the guild.
  * "0% hunters" with no valid date data at all.
* **Normalization:** Parses `Points (Hunt)` into clean integers and `Goal Percentage (Hunt)` into a consistent, rounded percentage value (handling values given as either `0–1` or `0–100`).
* **Sorting:** Orders the final list by goal percentage, highest performers first.

### 🔹 Phase 3: Load (Formatted Report Generation)
The cleaned data is written straight back out to Excel using `openpyxl`:
* Bold, colored header row.
* Top 3 performers get rank labels and distinct highlight colors.
* Remaining rows are color-coded by performance tier (≥100%, ≥90%, ≥70%, below 70%).
* A special "Kidnape me" row (used for guild bank tracking) is detected and formatted separately.

---

## 🚀 Key Features

| Feature | Description | Tech Stack |
| :--- | :--- | :--- |
| **Guild-Presence Filter** | Drops players under 4 days tenure or with no hunt-date data. | Pandas |
| **Percentage Normalization** | Handles messy percent formats (`"85%"`, `0.85`, `85`) and rounds consistently. | Pandas |
| **Color-Coded Report** | Auto-highlights top 3 ranks and performance tiers in the output `.xlsx`. | `openpyxl` |
| **Column Auto-Formatting** | Sets consistent column widths and header styling on every run. | `openpyxl` |

---

## 🛠️ Usage

### Requirements
```
pip install pandas openpyxl
```

### Run it
```
python etl_pipeline.py <input_file.xlsx> [output_dir]
```
* `input_file.xlsx` — the raw hunt export to process.
* `output_dir` — optional, defaults to `downloads/`.

The pipeline prints a short summary (rows removed, players filtered) and writes a timestamped report, e.g. `downloads/Processed_Report_20261026_143000.xlsx`.

### Use it as a library
```python
from etl_pipeline import run_pipeline

output_path = run_pipeline("raw_export.xlsx", output_dir="reports")
```

---

## 📁 Project Structure
```
etl_pipeline.py   # Extract, transform, and load logic (this project)
downloads/        # Default output directory for generated reports
```

> Note: this pipeline is purely a data-processing script — it has no bot, webhook, or messaging integration. It reads a file in and writes a formatted file out.
