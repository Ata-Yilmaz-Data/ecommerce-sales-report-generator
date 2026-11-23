# E-commerce Sales & Profit Report Generator

A Python CLI pipeline built with **Pandas** and **OpenPyXL** that cleans raw e-commerce CSV exports, joins optional product unit costs, deducts returns from net figures, and generates a styled multi-sheet Excel workbook plus a plain-text executive summary.

## Key Features
- **Gross Profit (after COGS) via `--costs`:** Real store exports rarely include profit. Join a separate unit-cost CSV by SKU or product name to compute COGS, Gross Profit, and Gross Margin.
- **Unbiased margin calculation:** Products missing unit costs are listed on a dedicated `Missing_Costs` sheet and excluded from the margin denominator (`Costed_Net_Sales`), so they do not artificially lower margins. A `Cost_Coverage` metric shows what share of net sales has cost data.
- **Honest profit labels:** If the source CSV already contains a `Profit` column (e.g. Superstore), it is used as provided and labeled *"Profit (source column)"* - never as "Gross Profit after COGS".
- **Returns deducted from net figures:** Separates Gross Sales, Returned Value, and Net Sales across flag-based return columns, external return CSVs, and negative-row accounting ledgers.
- **Configurable column mapping (`--preset` / `--config`):** Ships with `global_superstore`, `superstore`, and an illustrative Polish-schema example (`example_pl`). Adapt to any store CSV via JSON without editing code.
- **Multi-currency formatting (`--currency`):** Formats Excel cells and text summaries for `PLN`, `EUR`, `USD`, `GBP`, or any custom ISO code.
- **Data-integrity safeguards:** Strict date parsing (`--date-format` / `--dayfirst`) with partial-failure warnings, US and European number parsing (`1,234.50` and `1.234,50`), safe `row_id`-based deduplication, >5% data-loss warnings, and multi-encoding CSV reading (`utf-8-sig`, `windows-1250`, `windows-1252`, `latin1`).

## Requirements
- **Python 3.10+** (uses PEP 604 `X | Y` union type syntax)
- `pandas >= 2.0`, `openpyxl >= 3.1`

## Installation

Download or clone this repository, then install the two dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### 1. Run on Global Superstore (51,290 order rows)

```bash
python sales_report_generator.py \
  --input "Global_Superstore2.csv" \
  --preset global_superstore \
  --currency USD
```

> **Headers, dates & returns:** The `global_superstore` preset accepts both header styles seen in public mirrors (`Order ID` and `Order.ID`) and both ISO timestamps (`2011-01-07 00:00:00.000`) and `DD-MM-YYYY` dates. If your copy uses another date format, the report prints a date warning - override it with `--date-format`. The standard single-file export has no returns column, so the return rate is reported as `0.00%` (`None detected`) unless a separate returns CSV is passed via `--returns`.

### 2. Run with synthetic demo data (Polish schema + unit-cost join)
Generates synthetic sales, unit-cost, and config files (1 uncosted SKU, 2 duplicate row IDs, 1 invalid date) in `output/demo_data/`:

```bash
python sales_report_generator.py --demo --currency PLN
```

> The `--demo` dataset is 100% synthetic and exists only to demonstrate cleaning, cost joining, and margin-coverage behavior.

### 3. Run on a custom store export + supplier unit costs
Copy `configs/example_pl.json`, adjust the column names to match your export, and pass it with `--config`:

```bash
python sales_report_generator.py \
  --input store_sales.csv \
  --costs supplier_unit_costs.csv \
  --config configs/my_store.json \
  --currency PLN \
  --date-format "%d.%m.%Y"
```

## Outputs (`output/`)

1. **`sales_report.xlsx`**
   - `Executive_KPI`: Gross vs. Net Sales, sales with cost/profit data, coverage %, profit, margin, order and value return rates.
   - `Product_Profitability`: product-level net sales, profit, margin, and return rate.
   - `Monthly_Trend`: monthly net sales, MoM growth, margin, coverage, and return rate.
   - `Category_Returns`: category-level return exposure and profitability.
   - `Missing_Costs` *(conditional)*: products without unit costs.
2. **`executive_summary.txt`**: plain-text summary with top performers, negative-margin products, return hotspots, and data-quality warnings.

## Example Output

### Global Superstore (51,290 order rows)
Real run: `python sales_report_generator.py --input superstore.csv --preset global_superstore --currency USD`. The dataset itself is not included in this repository.

```text
==============================================================
               E-COMMERCE SALES & PROFIT REPORT
==============================================================
Period          : 2011-01-01 to 2014-12-31
Data Cleaning   : 51290 raw rows -> 0 duplicates, 0 invalid rows removed (51290 valid)
Dedup Strategy  : By unique row_id
Date Parsing    : ISO-8601
Return Source   : None detected (0% assumed)
Profit Basis    : Profit column as provided by the source CSV (returned orders excluded)

--- OVERALL PERFORMANCE (NET OF RETURNS) ---
Gross Sales     : $12,642,905.00
Returned Value  : $0.00 (0.00% of gross)
Net Sales       : $12,642,905.00
Net Sales w/ Profit Data: $12,642,905.00 (100.00% coverage)
Profit (source) : $1,467,457.29 (as defined by the source dataset)
Profit Margin   : 11.61%
Order Returns   : 0.00% (0 / 25035 orders)

--- MONTHLY TREND (NET SALES) ---
Best Month      : 2014-11 ($555,312.00)
Lowest Month    : 2011-02 ($91,152.00)

--- TOP 3 PRODUCTS BY PROFIT (SOURCE COLUMN) ---
- Canon imageCLASS 2200 Advanced Copier: $25,199.93 (Margin: 40.9% | Return Rate: 0.0%)
- Cisco Smart Phone, Full Size: $17,238.52 (Margin: 22.6% | Return Rate: 0.0%)
- Motorola Smart Phone, Full Size: $17,027.11 (Margin: 23.3% | Return Rate: 0.0%)

--- ATTENTION: NEGATIVE-PROFIT PRODUCTS ---
- Cubify CubeX 3D Printer Double Head Print: -$8,879.97 (Margin: -80.0%)
- Lexmark MX611dhe Monochrome Laser Printer: -$4,589.97 (Margin: -27.3%)
- Motorola Smart Phone, Cordless: -$4,447.04 (Margin: -11.4%)

--- RETURN RISK ---
Highest Return Category: Furniture (0.00% of orders)

--- ACCOUNTING NOTES ---
* Profit is taken as provided in the source CSV; which costs it includes is
  defined by the dataset, not by this script.
==============================================================
```

<details>
<summary>Synthetic demo run (Polish schema + unit-cost join)</summary>

```text
==============================================================
            E-COMMERCE SALES & GROSS PROFIT REPORT            
==============================================================
Period          : 2025-01-01 to 2025-12-31
Data Cleaning   : 502 raw rows -> 2 duplicates, 1 invalid rows removed (499 valid)
Dedup Strategy  : By unique row_id
Date Parsing    : Strict format (%d.%m.%Y)
Return Source   : Inline returned/status column
Profit Basis    : Gross Profit (after COGS) via --costs joined on 'sku'

--- OVERALL PERFORMANCE (NET OF RETURNS) ---
Gross Sales     : 588,739.45 PLN
Returned Value  : 64,581.10 PLN (10.97% of gross)
Net Sales       : 524,158.35 PLN
Costed Net Sales: 472,907.70 PLN (90.22% coverage)
Gross Profit    : 106,357.70 PLN (after COGS only)
Gross Margin    : 22.49% (calculated on costed sales only)
Order Returns   : 13.55% (34 / 251 orders)

--- MONTHLY TREND (NET SALES) ---
Best Month      : 2025-10 (76,771.55 PLN)
Lowest Month    : 2025-08 (24,962.25 PLN)

--- TOP 3 PRODUCTS BY GROSS PROFIT (AFTER COGS) ---
- Ergonomic Desk Chair: 35,076.00 PLN (Margin: 25.4% | Return Rate: 19.7%)
- 24-inch IPS Monitor: 26,292.65 PLN (Margin: 20.9% | Return Rate: 3.3%)
- USB-C Docking Station: 13,716.20 PLN (Margin: 41.2% | Return Rate: 5.4%)

--- RETURN RISK ---
Highest Return Category: Home Office (15.23% of orders)

--- ACCOUNTING NOTES ---
* Gross Profit deducts product unit costs (COGS) only; marketplace commissions,
  shipping, ads, payment fees, and overhead are NOT deducted.
* VAT consistency: Ensure sales and unit costs use the same tax basis
  (both net ex-VAT or both gross incl-VAT) to avoid distorted margins.

--- DATA QUALITY WARNINGS ---
! 1 row(s) (0.2%) failed date parsing with Strict format (%d.%m.%Y). Sample values: ['invalid-date'].
! 1 product(s) missing unit costs in --costs CSV. Excluded from Gross Profit and the margin denominator (see 'Missing_Costs' sheet).
==============================================================
```

</details>

## Tests

```bash
python tests/test_all.py
```

40 dependency-free checks, passing on pandas 2.2 and pandas 3.0. They cover number parsing, the demo pipeline (gross profit and margin are cross-checked against an independent recomputation from the raw CSVs), Superstore-style files (latin-1 encoding with `DD-MM-YYYY` dates, and dot-style headers with ISO timestamps, both using the source `Profit` column), ledger-style returns, and failure modes such as a wrong date format.

## Limitations & Accounting Assumptions
1. **Gross Profit (after COGS) only - not net operating profit.** With `--costs`, profit deducts product unit costs only. Marketplace commissions (e.g. Allegro fees), shipping, advertising, payment fees, packaging, and overhead are **not** included.
2. **Source Profit column.** When the input already has a `Profit` column, it is used as provided; what costs it includes is defined by the dataset, not by this script.
3. **VAT / tax basis.** The script performs no VAT conversion. Sales values and unit costs **must use the same tax basis** (both net or both gross). Mixing VAT-inclusive prices with net supplier costs overstates profit and margin.
4. **Return accounting.** Returned orders are excluded from Net Sales and profit, assuming recovered inventory with no restocking or return-shipping loss.
5. **No partial returns via external returns CSV.** With `--returns`, matching is by `Order ID`, so the **entire order** is treated as returned. Partial returns are reflected only if the main CSV flags individual line items or contains negative return rows.
6. **Ambiguous number formats.** A lone value like `1.234` is read as a decimal unless the same column also contains comma decimals (`149,90`), in which case it is read as a thousands separator.
7. **Illustrative Polish mapping only.** The `example_pl` preset is a customizable example, not an official Allegro (or any marketplace) export format. Check your real export's column names and edit the JSON config accordingly.
