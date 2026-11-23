"""
Dependency-free test suite. Run from the project root:

    python tests/test_all.py

Covers: number parsing, demo pipeline (with an independent recomputation of gross profit),
a Superstore-style file (latin-1 encoding, DD-MM-YYYY dates, source Profit column),
ledger-style returns, and failure modes (wrong date format).
"""
import math
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import sales_report_generator as srg  # noqa: E402

PASSED = 0


def check(condition: bool, message: str) -> None:
    global PASSED
    if not condition:
        raise AssertionError(message)
    PASSED += 1
    print(f"  ok - {message}")


def kpi_dict(reports: dict[str, pd.DataFrame]) -> dict:
    k = reports["Executive_KPI"]
    return dict(zip(k["Metric"], k["Value"]))


# ------------------------------------------------------------------
def test_numeric_parser(tmp: Path) -> None:
    print("[1] Numeric parser")
    csv = tmp / "numbers.csv"
    csv.write_text('id;raw_sales\n1;"1,234.50"\n2;"1.234,50"\n3;"149,90"\n4;\n5;abc\n', encoding="utf-8")
    df = srg.load_csv_safely(csv, sep=";")
    parsed, failed, samples = srg.parse_numeric_series(df["raw_sales"])
    check(parsed.iloc[0] == 1234.50, "1,234.50 -> 1234.50")
    check(parsed.iloc[1] == 1234.50, "1.234,50 -> 1234.50")
    check(parsed.iloc[2] == 149.90, "149,90 -> 149.90")
    check(math.isnan(parsed.iloc[3]), "empty cell -> NaN")
    check(math.isnan(parsed.iloc[4]), "'abc' -> NaN")
    check(failed == 1 and samples == ["abc"], "only 'abc' counted as a parse failure")

    p, f, _ = srg.parse_numeric_series(pd.Series(["1.234.567", "149,90", "-17,5"], dtype="string"))
    check(p.tolist() == [1234567.0, 149.9, -17.5], "multi-dot thousands and negative decimal comma")
    p, f, _ = srg.parse_numeric_series(pd.Series(["2309.65", "-17.5", "1234"], dtype="string"))
    check(p.tolist() == [2309.65, -17.5, 1234.0] and f == 0, "plain US-style numbers untouched")


# ------------------------------------------------------------------
def test_demo_pipeline(tmp: Path) -> None:
    print("[2] Demo pipeline + independent cross-check")
    sales_path, costs_path, config_path = srg.create_demo_files(tmp)
    config = srg.load_mapping_config("example_pl", config_path)
    raw = srg.load_csv_safely(sales_path)
    costs = srg.load_csv_safely(costs_path)

    clean, missing, stats = srg.clean_and_enrich(raw, config, costs_df=costs)
    reports = srg.build_reports(clean, missing, stats, "PLN")
    kpi = kpi_dict(reports)

    check(stats["raw_rows"] == 502, "502 raw rows")
    check(stats["duplicates_removed"] == 2, "2 duplicate row IDs removed")
    check(stats["invalid_rows_removed"] == 1, "1 invalid-date row removed")
    check(stats["clean_rows"] == 499, "499 valid rows")
    check(stats["profit_kind"] == "cogs", "profit computed from costs (COGS)")
    check(list(missing["sku"]) == ["SKU-302"], "Missing_Costs lists exactly SKU-302")
    check(0 < kpi["Cost Coverage (% of Net Sales)"] < 1, "cost coverage is below 100%")

    # Independent recomputation straight from the raw CSVs (no pipeline code involved)
    r = raw.drop_duplicates(subset=["ID pozycji"]).copy()
    r["d"] = pd.to_datetime(r["Data zamówienia"], format="%d.%m.%Y", errors="coerce")
    r = r.dropna(subset=["d"])
    r = r[r["Zwrot"] != "Tak"]
    r = r.merge(costs[["SKU", "Koszt jednostkowy"]], on="SKU", how="inner")  # costed rows only
    expected_profit = float((r["Wartość sprzedaży"] - r["Koszt jednostkowy"] * r["Sztuki"]).sum())
    expected_costed = float(r["Wartość sprzedaży"].sum())
    got_profit = kpi["Gross Profit after COGS (PLN)"]
    got_costed = kpi["Costed Net Sales (PLN)"]
    check(abs(got_profit - expected_profit) < 0.01, f"gross profit matches independent calc ({got_profit:,.2f})")
    check(abs(got_costed - expected_costed) < 0.01, f"costed net sales match independent calc ({got_costed:,.2f})")
    check(abs(kpi["Gross Margin after COGS (on Costed Sales)"] - expected_profit / expected_costed) < 1e-4,
          "margin = profit / costed sales (not diluted by uncosted SKU)")

    # Summary + Excel generation must not crash
    text = srg.generate_summary_text(reports, stats, "PLN")
    check("Gross Profit" in text and "Missing_Costs" in text, "summary shows gross profit and the missing-cost warning")
    out = tmp / "out"
    srg.export_to_excel(reports, out / "r.xlsx", "PLN")
    with pd.ExcelFile(out / "r.xlsx") as xf:  # close the file handle (required on Windows)
        sheets = xf.sheet_names
    check(sheets == ["Executive_KPI", "Product_Profitability", "Monthly_Trend", "Category_Returns", "Missing_Costs"],
          "Excel has the 5 expected sheets")


# ------------------------------------------------------------------
def test_superstore_style(tmp: Path) -> None:
    print("[3] Superstore-style file (latin-1, DD-MM-YYYY, source Profit column)")
    rows = [
        # Row ID, Order ID, Order Date, Product ID, Product Name, Category, Sales, Quantity, Profit
        (1, "A-1", "31-07-2012", "P1", "Caf\u00e9 Table", "Furniture", "1,000.50", 2, 100.25),
        (2, "A-1", "31-07-2012", "P2", "Chair, Wooden", "Furniture", "500.00", 1, -20.00),
        (3, "A-2", "01-08-2012", "P3", "Pen", "Office", "10.00", 5, 3.00),
        (4, "A-3", "15-09-2012", "P3", "Pen", "Office", "20.00", 10, 6.00),
    ]
    df = pd.DataFrame(rows, columns=["Row ID", "Order ID", "Order Date", "Product ID", "Product Name",
                                     "Category", "Sales", "Quantity", "Profit"])
    path = tmp / "gs.csv"
    df.to_csv(path, index=False, encoding="latin1")

    summary, xlsx, _ = srg.run_pipeline(path, "global_superstore", None, None, None, "USD", tmp / "gs_out")
    check("Profit (source)" in summary and "Gross Profit" not in summary.split("--- OVERALL")[1].split("Order Returns")[0],
          "source Profit is NOT labeled 'Gross Profit after COGS'")
    check("None detected" in summary, "no returns column -> 'None detected (0% assumed)'")

    config = srg.load_mapping_config("global_superstore", None)
    clean, _, stats = srg.clean_and_enrich(srg.load_csv_safely(path), config)
    check(stats["profit_kind"] == "source", "profit_kind == 'source'")
    check(abs(float(clean["sales"].sum()) - (1000.50 + 500 + 10 + 20)) < 1e-9, "sales parsed incl. '1,000.50'")
    check(clean["order_date"].min() == pd.Timestamp(2012, 7, 31), "DD-MM-YYYY parsed correctly (31 Jul 2012)")
    check("Caf\u00e9 Table" in clean["product_name"].tolist(), "latin-1 accented product name survives")
    check("Chair, Wooden" in clean["product_name"].tolist(), "quoted comma inside a product name survives")



# ------------------------------------------------------------------
def test_dot_headers_iso_dates(tmp: Path) -> None:
    print("[3b] Dot-style headers + ISO timestamps (layout of the Kaggle Global Superstore mirror)")
    df = pd.DataFrame({
        "Category": ["Office Supplies", "Furniture", "Technology"],
        "Customer.Name": ["A B", "C D", "E F"],
        "Order.Date": ["2011-01-07 00:00:00.000", "2011-01-08 00:00:00.000", "2011-02-01 00:00:00.000"],
        "Order.ID": ["CA-2011-1", "CA-2011-2", "CA-2011-3"],
        "Product.ID": ["OFF-1", "FUR-1", "TEC-1"],
        "Product.Name": ["Xerox 225", "Table", "Phone"],
        "Profit": [9.3312, -4.5, 50.0],
        "Quantity": [3, 1, 2],
        "Row.ID": [36624, 36625, 36626],
        "Sales": [19, 100, 400],
        "Market2": ["North America"] * 3,
    })
    path = tmp / "dots.csv"
    df.to_csv(path, index=False)
    config = srg.load_mapping_config("global_superstore", None)
    clean, _, stats = srg.clean_and_enrich(srg.load_csv_safely(path), config)
    check(stats["clean_rows"] == 3 and stats["invalid_rows_removed"] == 0, "all 3 rows mapped and kept")
    check(stats["date_mode"] == "ISO-8601", "ISO timestamps with milliseconds parsed as ISO-8601")
    check(clean["order_date"].min() == pd.Timestamp(2011, 1, 7), "first date is 2011-01-07")
    check(stats["profit_kind"] == "source" and stats["duplicates_removed"] == 0, "source Profit used, no duplicates")
    summary, _, _ = srg.run_pipeline(path, "global_superstore", None, None, None, "USD", tmp / "dots_out")
    check("ISO-8601" in summary, "summary reports ISO-8601 date parsing")

# ------------------------------------------------------------------
def test_ledger_returns(tmp: Path) -> None:
    print("[4] Negative-row (ledger) returns")
    df = pd.DataFrame({
        "Order ID": ["1", "2", "3"],
        "Order Date": ["2025-01-05", "2025-01-06", "2025-01-07"],
        "Product Name": ["A", "A", "A"],
        "Sales": [100.0, 50.0, -50.0],
        "Quantity": [2, 1, -1],
    })
    config = srg.load_mapping_config("superstore", None)
    clean, missing, stats = srg.clean_and_enrich(df, config, cli_date_format="%Y-%m-%d")
    rep = srg.build_reports(clean, missing, stats, "USD")
    kpi = kpi_dict(rep)
    check(stats["return_source"].startswith("Negative"), "ledger mode detected")
    check(kpi["Gross Sales (USD)"] == 150.0, "gross sales counts only positive rows (150)")
    check(kpi["Returned Sales Value (USD)"] == 50.0, "returned value = 50")
    check(kpi["Net Sales (USD)"] == 100.0, "net sales = 100 (not double-deducted)")
    check(kpi["Gross Profit after COGS (USD)"] == "N/A", "no profit data -> N/A, no crash")
    srg.generate_summary_text(rep, stats, "USD")
    check(True, "summary works without any profit data")


# ------------------------------------------------------------------
def test_failure_modes(tmp: Path) -> None:
    print("[5] Failure modes")
    df = pd.DataFrame({
        "Order ID": ["1", "2"], "Order Date": ["31-07-2012", "01-08-2012"],
        "Product Name": ["A", "B"], "Sales": [1.0, 2.0],
    })
    config = srg.load_mapping_config("superstore", None)  # expects %m/%d/%Y -> every date fails
    try:
        srg.clean_and_enrich(df, config)
        check(False, "wrong date format should raise")
    except ValueError as e:
        check("All dates failed to parse" in str(e), "wrong date format -> clear error message")

    df2 = pd.DataFrame({"Order ID": ["1"], "Order Date": ["01/02/2025"], "Sales": [1.0]})
    try:
        srg.clean_and_enrich(df2, config)
        check(False, "missing column should raise")
    except ValueError as e:
        check("Missing required mapped columns" in str(e), "missing column -> clear error message")


# ------------------------------------------------------------------
if __name__ == "__main__":
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
        tmp = Path(d)
        test_numeric_parser(tmp)
        test_demo_pipeline(tmp)
        test_superstore_style(tmp)
        test_dot_headers_iso_dates(tmp)
        test_ledger_returns(tmp)
        test_failure_modes(tmp)
    print(f"\nALL TESTS PASSED ({PASSED} checks)")
