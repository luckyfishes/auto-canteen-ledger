# AutoCanteenLedger

A Python desktop tool that automates canteen ledger generation by merging a template workbook with supplier delivery bills. Built with **Tkinter** and **openpyxl**.

## Overview

Managing canteen ledgers by hand is tedious and error-prone. This tool lets you:

1. Select a **ledger template** workbook with pre-built day sheets and a summary sheet.
2. Select a **delivery bill** workbook from your supplier (with a `送货单` sheet).
3. Click **Run** — it fills the dates, auto-adjusts row counts, updates totals, and cleans up unused sheets.

## Features

- **One-click workflow** — just pick a template, a bill, and an output path.
- **Automatic `.xls` → `.xlsx` conversion** — works with legacy Excel files too (via `win32com` or `xlrd` fallback).
- **Cross-workbook sheet copy** — duplicates the `送货单` sheet into the output with full formatting preserved (fonts, borders, merges, column widths, row heights).
- **Dynamic row adjustment** — each day sheet expands or shrinks to match the number of delivered items for that date.
- **Summary sheet auto-update** — keeps only the days with data, updates totals, and rewrites date labels.
- **Clean output** — removes empty date sheets so only the days with actual data remain.
- **Progress bar + live log** — real-time feedback during processing.

## Installation

```bash
git clone https://github.com/neo/auto-canteen-ledger.git
cd auto-canteen-ledger
pip install -r requirements.txt
```

### Requirements

- Python 3.8+
- `pandas`
- `openpyxl`
- `xlrd` (optional — legacy `.xls` fallback)
- `pywin32` (optional — for better `.xls` → `.xlsx` conversion on Windows via Excel COM)

> **Note:** On Windows, `pywin32` is recommended for `.xls` conversion because it preserves all formatting. Without it, the tool falls back to `xlrd`, which preserves cell values only.

## Usage

```bash
python fill_ledger_gui.py
```

1. **Ledger Template** — select an `.xlsx` template with daily sheets (`1` through `31`) and a `汇总` sheet.
2. **Bill File** — select the supplier's delivery bill (`.xlsx`, `.xlsm`, or `.xls`).
3. **Output File** — choose where to save the generated ledger. Defaults to `{bill_name}_输出结果.xlsx` next to the bill.
4. Click **▶ Start Fill**.

## How It Works

1. **Copy Template** — creates a fresh copy of the template so the original stays untouched.
2. **Convert Bill** — if the bill is `.xls`, converts it to `.xlsx` (using Excel COM or `xlrd`).
3. **Copy Delivery Sheet** — copies the `送货单` sheet into the output workbook with full formatting.
4. **Parse Orders** — scans the `送货单` sheet, groups items by delivery date, extracts item name, unit, quantity, price, amount, and notes.
5. **Fill Daily Sheets** — for each date with data, finds the matching sheet (e.g., `1` for the 1st), dynamically adjusts rows to fit the number of items, and fills the data. Auto-formats with `SUM` formula for the total row.
6. **Update Summary** — removes rows for dates without data, rewrites date labels, updates the `SUM` total.
7. **Remove Empty Sheets** — deletes all daily sheets that have no data.
8. **Save Result** — writes the final workbook to the chosen path.

## Template Structure

Your ledger template should be an `.xlsx` file with:

- **31 daily sheets** named `1` through `31` — each has a fixed data region starting at row 4, with a `合计` (Total) row at the bottom.
- **A `汇总` (Summary) sheet** — lists all dates in the month and auto-sums the totals.

### Example Daily Sheet (`1`):

| Row | Column A | Column B | Column C | Column D | Column E | Column F | Column G |
|-----|----------|----------|----------|----------|----------|----------|----------|
| 1   | Header area... | | | | | | |
| ... | | | | | | | |
| 4   | `日期` | `品名` | `单位` | `数量` | `单价` | `金额` | `备注` |
| 5   | *(filled by tool)* | *(filled)* | *(filled)* | *(filled)* | *(filled)* | *(filled)* | *(filled)* |
| ... | | | | | | | |
| N   | `合计` | | | | | `=SUM(F4:F{N-1})` | |

### Example Summary Sheet (`汇总`):

| Row | Column A | Column B |
|-----|----------|----------|
| 1   | `日期` | `金额` |
| 2   | `2026.01.01` | *(filled by tool)* |
| 3   | `2026.01.02` | *(filled)* |
| ... | | |
| 32  | `合计` | `=SUM(B2:B31)` |

## Bill File Structure

The supplier's delivery bill must contain a sheet named **exactly** `送货单` with the following layout:

- **Row 9 (column J)** contains delivery date in the format `送货日期：YYYY-MM-DD`.
- **Rows below** contain items with:
  - Column B (index 1): `序号` (sequence number)
  - Column D (index 3): `品名` (item name)
  - Column G (index 6): `数量` (quantity)
  - Column H (index 7): `单位` (unit)
  - Column I (index 8): `单价` (price)
  - Column J (index 9): `金额` (amount)
  - Column K (index 10): `备注` (notes)

## Screenshot

```
┌──────────────────────────────────────────────┐
│  🍖  AutoCanteenLedger                        │
│  Template: [____________________] [Browse]   │
│  Bill:    [____________________] [Browse]   │
│  Output:  [____________________] [Browse]   │
│                                              │
│  [  ▶  Start Fill  ]                         │
│  [=====================>    ] 80%            │
│  ------------------------------------------- │
│  ✅ Copied template → output.xlsx            │
│  ✅ Delivery sheet copied                    │
│  ✅ Parsed 5 delivery dates: [1, 3, 5, 7, 9] │
│  📅 Sheet 1: 12 items ...                    │
│  ...                                         │
└──────────────────────────────────────────────┘
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Contributing

Issues and PRs welcome. Please open an issue first to discuss major changes.

## Author

- **neo** — created for automating canteen ledger workflows at a school/institution cafeteria.
