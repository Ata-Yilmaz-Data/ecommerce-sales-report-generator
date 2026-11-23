import argparse
import json
import random
import re
import sys
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ==========================================
# BUILT-IN PRESETS & CURRENCY FORMATS
# ==========================================
# Column values may be a single header name or a list of accepted alternatives
# (Kaggle mirrors differ: "Order ID" vs "Order.ID").
_SUPERSTORE_COLUMNS = {
    "row_id": ["Row ID", "Row.ID"],
    "order_id": ["Order ID", "Order.ID"],
    "order_date": ["Order Date", "Order.Date"],
    "sku": ["Product ID", "Product.ID"],
    "product_name": ["Product Name", "Product.Name"],
    "category": "Category",
    "sales": "Sales",
    "quantity": "Quantity",
    "profit": "Profit",
    "returned": "Returned",
}
_SUPERSTORE_COSTS = {
    "sku": ["Product ID", "Product.ID"],
    "product_name": ["Product Name", "Product.Name"],
    "unit_cost": "Unit Cost",
}

PRESETS: dict[str, dict[str, Any]] = {
    # Kaggle "Sample - Superstore" (US dates, ~10k rows)
    "superstore": {
        "date_format": "%m/%d/%Y",
        "dayfirst": False,
        "columns": dict(_SUPERSTORE_COLUMNS),
        "cost_columns": dict(_SUPERSTORE_COSTS),
    },
    # "Global Superstore" (~51k rows). Mirrors differ: dates are either ISO
    # ("2011-01-07 00:00:00.000") or DD-MM-YYYY. ISO is tried first, then explicit
    # day-first parsing; force a format with --date-format if needed.
    "global_superstore": {
        "date_format": None,
        "dayfirst": True,
        "columns": dict(_SUPERSTORE_COLUMNS),
        "cost_columns": dict(_SUPERSTORE_COSTS),
    },
    # Illustrative Polish-language column mapping (NOT an official Allegro export spec)
    "example_pl": {
        "date_format": "%d.%m.%Y",
        "dayfirst": True,
        "columns": {
            "row_id": "ID pozycji",
            "order_id": "Numer zamówienia",
            "order_date": "Data zamówienia",
            "sku": "SKU",
            "product_name": "Nazwa oferty",
            "category": "Kategoria",
            "sales": "Wartość sprzedaży",
            "quantity": "Sztuki",
            "returned": "Zwrot",
        },
        "cost_columns": {
            "sku": "SKU",
            "product_name": "Nazwa oferty",
            "unit_cost": "Koszt jednostkowy",
        },
    },
}

CURRENCY_EXCEL_FORMATS: dict[str, str] = {
    "PLN": '#,##0.00 "PLN"',
    "EUR": "€#,##0.00",
    "USD": "$#,##0.00",
    "GBP": "£#,##0.00",
}

RETURN_TRUE_TOKENS = {"yes", "true", "1", "returned", "refunded", "tak", "zwrot"}
INVALID_ROW_WARN_THRESHOLD = 0.05


# ==========================================
# HELPERS
# ==========================================
def is_num(x: Any) -> bool:
    """True for real numbers (Python or NumPy), False for strings like 'N/A' and for NaN."""
    return isinstance(x, (int, float, np.integer, np.floating)) and not pd.isna(x)


def format_currency_text(value: float, currency: str) -> str:
    """Formats numeric currency consistently for plain-text summaries."""
    symbols = {"USD": "$", "EUR": "€", "GBP": "£"}
    if currency in symbols:
        sign = "-" if value < 0 else ""
        return f"{sign}{symbols[currency]}{abs(value):,.2f}"
    return f"{value:,.2f} {currency}"


def safe_divide(
    numerator: pd.Series,
    denominator: pd.Series,
    fill_zero: bool = False,
) -> pd.Series:
    """Divides two numeric Series safely using float('nan') to keep float64 dtype."""
    num = pd.to_numeric(numerator, errors="coerce").astype("float64")
    den = pd.to_numeric(denominator, errors="coerce").astype("float64").replace(0.0, float("nan"))
    result = num / den
    return result.fillna(0.0) if fill_zero else result


def parse_numeric_series(series: pd.Series) -> tuple[pd.Series, int, list[str]]:
    """
    Parses numeric columns supporting US (1,234.50) and European (1.234,50 / 1 234,50 / 149,90)
    formats without silently corrupting thousand separators.
    Returns (float64_series, failed_count, failed_raw_samples).
    Empty cells become NaN and are NOT counted as failures.
    """
    if pd.api.types.is_numeric_dtype(series):
        return series.astype("float64"), 0, []

    raw_str = (
        series.astype("string")
        .str.strip()
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    raw_non_null = raw_str.notna() & (raw_str != "")

    uses_dot_decimal = bool(raw_str.dropna().str.contains(r"\.\d{1,2}$", regex=True).any())
    uses_comma_decimal = bool(raw_str.dropna().str.contains(r",\d{1,2}$", regex=True).any())

    def _normalize(val: Any) -> Any:
        if pd.isna(val) or val == "":
            return None
        s = str(val)
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                return s.replace(".", "").replace(",", ".")  # 1.234,50
            return s.replace(",", "")  # 1,234.50
        if "," in s:
            if s.count(",") > 1:
                return s.replace(",", "")  # 1,234,567
            if uses_dot_decimal and re.match(r"^-?\d{1,3},\d{3}$", s):
                return s.replace(",", "")  # 1,234 in a dot-decimal column
            return s.replace(",", ".")  # 149,90
        if "." in s:
            if s.count(".") > 1:
                return s.replace(".", "")  # 1.234.567
            if uses_comma_decimal and not uses_dot_decimal and re.match(r"^-?\d{1,3}\.\d{3}$", s):
                return s.replace(".", "")  # 1.234 in a comma-decimal column
        return s

    normalized = raw_str.map(_normalize)
    parsed = pd.to_numeric(normalized, errors="coerce").astype("float64")
    failed_mask = raw_non_null & parsed.isna()
    failed_count = int(failed_mask.sum())
    failed_samples = raw_str[failed_mask].astype(str).head(3).tolist()
    return parsed, failed_count, failed_samples


def profit_labels(kind: str, currency: str) -> dict[str, str]:
    """
    Single source of truth for profit-related names, so a profit value taken from the source CSV
    is never mislabeled as 'Gross Profit after COGS'.
    kind: 'cogs' (computed from --costs) or 'source' (Profit column supplied by the dataset).
    """
    if kind == "source":
        return {
            "profit_metric": f"Profit, source column ({currency})",
            "margin_metric": "Profit Margin, source column (on rows with profit data)",
            "coverage_metric": "Profit Data Coverage (% of Net Sales)",
            "costed_metric": f"Net Sales with Profit Data ({currency})",
            "profit_col": "Profit",
            "margin_col": "Margin",
            "coverage_col": "Profit_Coverage",
            "costed_col": "Net_Sales_With_Profit",
            "profit_title": "PROFIT (SOURCE COLUMN)",
            "negative_title": "NEGATIVE-PROFIT PRODUCTS",
        }
    return {
        "profit_metric": f"Gross Profit after COGS ({currency})",
        "margin_metric": "Gross Margin after COGS (on Costed Sales)",
        "coverage_metric": "Cost Coverage (% of Net Sales)",
        "costed_metric": f"Costed Net Sales ({currency})",
        "profit_col": "Gross_Profit",
        "margin_col": "Gross_Margin",
        "coverage_col": "Cost_Coverage",
        "costed_col": "Costed_Net_Sales",
        "profit_title": "GROSS PROFIT (AFTER COGS)",
        "negative_title": "NEGATIVE GROSS PROFIT PRODUCTS",
    }


# ==========================================
# 1. SYNTHETIC DEMO DATA GENERATOR
# ==========================================
def create_demo_files(out_dir: Path) -> tuple[Path, Path, Path]:
    """
    Generates synthetic Polish-schema e-commerce CSVs for testing:
    - sales CSV WITHOUT a Profit column
    - separate unit costs CSV (1 SKU intentionally omitted to verify cost-coverage logic)
    - mapping JSON config (loaded during --demo execution)
    """
    random.seed(42)
    catalog = [
        ("SKU-101", "Electronics", "Logitech Wireless Mouse", 149.0, 82.0, 0.05),
        ("SKU-102", "Electronics", "USB-C Docking Station", 249.0, 140.0, 0.07),
        ("SKU-103", "Electronics", "24-inch IPS Monitor", 799.0, 610.0, 0.08),
        ("SKU-201", "Home Office", "Ergonomic Desk Chair", 980.0, 690.0, 0.16),
        ("SKU-202", "Home Office", "Standing Desk 120cm", 1150.0, 1080.0, 0.14),
        ("SKU-203", "Home Office", "LED Desk Lamp", 129.0, 58.0, 0.04),
        ("SKU-301", "Accessories", "Laptop Stand Aluminum", 119.0, 46.0, 0.03),
        ("SKU-302", "Accessories", "Mechanical Keyboard", 349.0, None, 0.06),  # Missing cost test
    ]

    dates = pd.date_range("2025-01-01", "2025-12-31", freq="D")
    sales_rows = []
    for row_num in range(1, 501):
        sku, cat, name, unit_price, _, ret_prob = random.choice(catalog)
        qty = random.randint(1, 4)
        discount = random.choice([0.0, 0.0, 0.0, 0.10, 0.15])
        gross_sales = round(unit_price * qty * (1 - discount), 2)
        is_returned = "Tak" if random.random() < ret_prob else "Nie"

        sales_rows.append({
            "ID pozycji": f"LINE-{row_num:05d}",
            "Numer zamówienia": f"ORD-2025-{1000 + (row_num // 2)}",
            "Data zamówienia": random.choice(dates).strftime("%d.%m.%Y"),
            "SKU": sku,
            "Nazwa oferty": f" {name} " if row_num % 17 == 0 else name,
            "Kategoria": cat,
            "Wartość sprzedaży": gross_sales,
            "Sztuki": qty,
            "Zwrot": is_returned,
        })

    sales_df = pd.DataFrame(sales_rows)
    # Inject 2 duplicate row IDs and 1 invalid date to test data cleaning
    sales_df = pd.concat([sales_df, sales_df.iloc[:2]], ignore_index=True)
    sales_df.loc[15, "Data zamówienia"] = "invalid-date"

    costs_rows = [
        {"SKU": sku, "Nazwa oferty": name, "Koszt jednostkowy": cost}
        for sku, _, name, _, cost, _ in catalog
        if cost is not None
    ]
    costs_df = pd.DataFrame(costs_rows)

    demo_dir = out_dir / "demo_data"
    demo_dir.mkdir(parents=True, exist_ok=True)
    sales_path = demo_dir / "demo_sales_pl.csv"
    costs_path = demo_dir / "demo_product_costs_pl.csv"
    config_path = demo_dir / "demo_config_pl.json"

    sales_df.to_csv(sales_path, index=False, encoding="utf-8-sig")
    costs_df.to_csv(costs_path, index=False, encoding="utf-8-sig")
    config_path.write_text(json.dumps(PRESETS["example_pl"], indent=2, ensure_ascii=False), encoding="utf-8")

    return sales_path, costs_path, config_path


# ==========================================
# 2. IO, CONFIG & COLUMN MAPPING
# ==========================================
def load_csv_safely(path: Path, sep: str | None = None) -> pd.DataFrame:
    """Handles common European/US encodings and auto-detects comma vs semicolon delimiters."""
    for enc in ("utf-8-sig", "utf-8", "windows-1250", "windows-1252", "latin1"):
        try:
            if sep:
                return pd.read_csv(path, encoding=enc, sep=sep)
            return pd.read_csv(path, encoding=enc, sep=None, engine="python")
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Unable to decode CSV file with supported encodings: {path}")


def load_mapping_config(preset_name: str, config_path: Path | None) -> dict[str, Any]:
    """Loads a built-in preset and merges optional user JSON overrides."""
    base = json.loads(json.dumps(PRESETS[preset_name]))
    if config_path:
        user_cfg = json.loads(config_path.read_text(encoding="utf-8"))
        for key in ("date_format", "dayfirst"):
            if key in user_cfg:
                base[key] = user_cfg[key]
        if "columns" in user_cfg:
            base["columns"].update(user_cfg["columns"])
        if "cost_columns" in user_cfg:
            base["cost_columns"].update(user_cfg["cost_columns"])
    return base


def apply_column_mapping(df: pd.DataFrame, col_map: dict[str, Any]) -> pd.DataFrame:
    """Renames source columns to canonical internal names using config mapping.
    Each mapping value is a header name or a list of accepted alternatives (first match wins)."""
    df = df.copy()
    df.columns = df.columns.str.strip()
    rename_dict: dict[str, str] = {}
    for canonical, src in col_map.items():
        candidates = [src] if isinstance(src, str) else list(src or [])
        for cand in candidates:
            cand = str(cand).strip()
            if cand and cand in df.columns and cand not in rename_dict:
                rename_dict[cand] = canonical
                break
    return df.rename(columns=rename_dict)


# ==========================================
# 3. DATA CLEANING, COST MERGE & NET LOGIC
# ==========================================
def parse_dates_strictly(
    series: pd.Series,
    date_format: str | None,
    dayfirst: bool,
) -> tuple[pd.Series, str]:
    """Parses dates strictly without ambiguous format='mixed' guessing."""
    cleaned = series.astype("string").str.strip()
    if date_format:
        parsed = pd.to_datetime(cleaned, format=date_format, errors="coerce")
        return parsed, f"Strict format ({date_format})"

    iso_parsed = pd.to_datetime(cleaned, format="ISO8601", errors="coerce")
    if iso_parsed.notna().mean() >= 0.9:
        return iso_parsed, "ISO-8601"

    parsed = pd.to_datetime(cleaned, dayfirst=dayfirst, errors="coerce")
    order_label = "DD/MM/YYYY (dayfirst=True)" if dayfirst else "MM/DD/YYYY (dayfirst=False)"
    return parsed, f"Explicit {order_label}"


def clean_and_enrich(
    raw_df: pd.DataFrame,
    config: dict[str, Any],
    costs_df: pd.DataFrame | None = None,
    returns_df: pd.DataFrame | None = None,
    cli_date_format: str | None = None,
    cli_dayfirst: bool | None = None,
    drop_exact_duplicates: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """
    Cleans sales data, joins unit costs (if provided) or uses the source Profit column,
    excludes returns from net sales/profit, and tracks which sales have profit data so
    margins are never distorted by missing costs.
    """
    stats: dict[str, Any] = {"raw_rows": len(raw_df), "warnings": []}
    df = apply_column_mapping(raw_df, config["columns"])

    required = {"order_id", "order_date", "product_name", "sales"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing required mapped columns: {', '.join(sorted(missing))}. "
            "Check --preset or provide a custom --config JSON."
        )

    # Normalize text columns safely across Pandas 2.x and 3.x
    text_cols = df.select_dtypes(include=["object", "string"]).columns
    for col in text_cols:
        df[col] = (
            df[col]
            .astype("string")
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )

    # Safe deduplication: only deduplicate by unique row_id unless opted in
    if "row_id" in df.columns and df["row_id"].notna().any():
        before = len(df)
        df = df.drop_duplicates(subset=["row_id"], keep="first")
        stats["duplicates_removed"] = before - len(df)
        stats["dedup_mode"] = "By unique row_id"
    else:
        exact_dupes = int(df.duplicated().sum())
        if drop_exact_duplicates and exact_dupes > 0:
            df = df.drop_duplicates(keep="first")
            stats["duplicates_removed"] = exact_dupes
            stats["dedup_mode"] = "Exact row match (explicit --drop-exact-duplicates)"
        else:
            stats["duplicates_removed"] = 0
            stats["dedup_mode"] = "Skipped (no row_id column)"
            if exact_dupes > 0:
                stats["warnings"].append(
                    f"Detected {exact_dupes} identical rows without a row_id column; "
                    "kept to avoid deleting legitimate split line items. "
                    "Pass --drop-exact-duplicates to remove them."
                )

    # Strict date parsing with partial AND full failure diagnostics
    raw_dates = df["order_date"].copy()
    date_format = cli_date_format or config.get("date_format")
    dayfirst = config.get("dayfirst", True) if cli_dayfirst is None else cli_dayfirst
    df["order_date"], stats["date_mode"] = parse_dates_strictly(raw_dates, date_format, dayfirst)

    bad_date_mask = df["order_date"].isna()
    bad_date_count = int(bad_date_mask.sum())
    if bad_date_count == len(df):
        sample_raw = raw_dates.dropna().astype(str).head(3).tolist()
        raise ValueError(
            f"All dates failed to parse using {stats['date_mode']}. "
            f"Sample raw values: {sample_raw}. Pass the correct --date-format."
        )
    if bad_date_count > 0:
        bad_date_pct = bad_date_count / len(df)
        bad_samples = raw_dates[bad_date_mask].astype(str).head(3).tolist()
        prefix = "HIGH INVALID DATE RATE: " if bad_date_pct > INVALID_ROW_WARN_THRESHOLD else ""
        stats["warnings"].append(
            f"{prefix}{bad_date_count} row(s) ({bad_date_pct * 100:.1f}%) failed date parsing "
            f"with {stats['date_mode']}. Sample values: {bad_samples}."
        )

    # Numeric conversions (US & European formats) with parse-failure diagnostics
    for num_col in ("sales", "quantity", "profit"):
        if num_col in df.columns:
            parsed_col, failed_cnt, failed_samples = parse_numeric_series(df[num_col])
            df[num_col] = parsed_col
            if failed_cnt > 0:
                fail_pct = failed_cnt / len(df)
                prefix = "HIGH NUMERIC PARSE FAILURE: " if fail_pct > INVALID_ROW_WARN_THRESHOLD else ""
                stats["warnings"].append(
                    f"{prefix}{failed_cnt} value(s) ({fail_pct * 100:.1f}%) in '{num_col}' "
                    f"could not be parsed as numbers. Sample values: {failed_samples}."
                )

    if "quantity" not in df.columns:
        df["quantity"] = 1.0
    if "category" not in df.columns:
        df["category"] = "Uncategorized"
    if "sku" not in df.columns:
        df["sku"] = pd.NA

    # Drop rows missing critical fields and warn if data loss exceeds 5%
    before_drop = len(df)
    df = df.dropna(subset=["order_id", "order_date", "product_name", "sales"]).copy()
    stats["invalid_rows_removed"] = before_drop - len(df)

    if df.empty:
        raise ValueError("No valid rows remaining after cleaning. Verify column mapping and date format.")

    drop_rate = stats["invalid_rows_removed"] / before_drop if before_drop else 0.0
    if drop_rate > INVALID_ROW_WARN_THRESHOLD:
        stats["warnings"].append(
            f"HIGH DATA LOSS: {drop_rate * 100:.1f}% of rows ({stats['invalid_rows_removed']} / {before_drop}) "
            "were dropped due to missing or unparseable required fields. Verify --date-format and numeric formats."
        )

    # Return detection (flag-based vs negative ledger rows)
    has_negative_returns = bool((df["sales"] < 0).any() or (df["quantity"] < 0).any())

    if returns_df is not None:
        ret_mapped = apply_column_mapping(returns_df, config["columns"])
        if "order_id" not in ret_mapped.columns:
            raise ValueError("Returns CSV must contain a mapped order_id column.")
        ret_ids = set(ret_mapped["order_id"].astype("string").str.strip().dropna())
        df["is_returned"] = df["order_id"].isin(ret_ids)
        stats["return_source"] = "External returns CSV (full-order match by order_id)"
        return_mode = "flag"
    elif "returned" in df.columns:
        df["is_returned"] = (
            df["returned"].astype("string").str.lower().str.strip().isin(RETURN_TRUE_TOKENS).fillna(False)
        )
        stats["return_source"] = "Inline returned/status column"
        return_mode = "flag"
    elif has_negative_returns:
        df["is_returned"] = (df["sales"] < 0) | (df["quantity"] < 0)
        stats["return_source"] = "Negative sales/quantity ledger rows"
        return_mode = "ledger"
    else:
        df["is_returned"] = False
        stats["return_source"] = "None detected (0% assumed)"
        return_mode = "flag"

    # Compute Gross, Returned, and Net Sales/Quantity
    if return_mode == "ledger":
        df["gross_sales"] = df["sales"].where(df["sales"] > 0, 0.0)
        df["returned_sales"] = (-df["sales"]).where(df["sales"] < 0, 0.0)
        df["net_sales"] = df["sales"]
        df["net_quantity"] = df["quantity"]
    else:
        df["gross_sales"] = df["sales"]
        df["returned_sales"] = df["sales"].where(df["is_returned"], 0.0)
        df["net_sales"] = df["sales"].where(~df["is_returned"], 0.0)
        df["net_quantity"] = df["quantity"].where(~df["is_returned"], 0.0)

    # Profit: computed from --costs (COGS only) or taken from the source Profit column
    missing_costs_df = pd.DataFrame(columns=["sku", "product_name", "net_sales"])
    stats["profit_kind"] = "none"

    if costs_df is not None:
        cost_map = config.get("cost_columns", {})
        c_df = apply_column_mapping(costs_df, cost_map)
        if "unit_cost" not in c_df.columns:
            raise ValueError("Costs CSV must contain a mapped 'unit_cost' column.")

        for col in c_df.select_dtypes(include=["object", "string"]).columns:
            c_df[col] = c_df[col].astype("string").str.strip()
        c_df["unit_cost"], cost_fail_cnt, cost_fail_samples = parse_numeric_series(c_df["unit_cost"])
        if cost_fail_cnt > 0:
            stats["warnings"].append(
                f"{cost_fail_cnt} value(s) in costs CSV 'unit_cost' failed numeric parsing: {cost_fail_samples}."
            )

        join_key = "sku" if ("sku" in c_df.columns and df["sku"].notna().any()) else "product_name"
        c_lookup = c_df.dropna(subset=[join_key, "unit_cost"]).drop_duplicates(subset=[join_key], keep="last")
        df = df.merge(c_lookup[[join_key, "unit_cost"]], on=join_key, how="left")

        missing_mask = df["unit_cost"].isna()
        if missing_mask.any():
            missing_costs_df = (
                df.loc[missing_mask]
                .groupby(["sku", "product_name"], dropna=False, as_index=False)["net_sales"]
                .sum()
                .sort_values("net_sales", ascending=False)
            )
            stats["warnings"].append(
                f"{len(missing_costs_df)} product(s) missing unit costs in --costs CSV. "
                "Excluded from Gross Profit and the margin denominator (see 'Missing_Costs' sheet)."
            )

        df["cogs"] = df["unit_cost"] * df["net_quantity"]
        df["gross_profit"] = df["net_sales"] - df["cogs"]
        stats["profit_source"] = f"Gross Profit (after COGS) via --costs joined on '{join_key}'"
        stats["profit_kind"] = "cogs"

    elif "profit" in df.columns and df["profit"].notna().any():
        if return_mode == "ledger":
            df["gross_profit"] = df["profit"]
        else:
            df["gross_profit"] = df["profit"].where(~df["is_returned"], 0.0)
        df["cogs"] = df["net_sales"] - df["gross_profit"]
        stats["profit_source"] = "Profit column as provided by the source CSV (returned orders excluded)"
        stats["profit_kind"] = "source"

    else:
        df["cogs"] = float("nan")
        df["gross_profit"] = float("nan")
        stats["profit_source"] = "Unavailable (provide --costs CSV for Gross Profit & Margin)"
        stats["warnings"].append(
            "No Profit column or --costs CSV provided. Generated Revenue & Return report only."
        )

    stats["has_profit"] = stats["profit_kind"] != "none"

    # Net sales ONLY for rows where profit is known (accurate margin denominator)
    df["costed_net_sales"] = df["net_sales"].where(df["gross_profit"].notna(), 0.0)
    df["returned_order_id"] = df["order_id"].where(df["is_returned"], pd.NA)
    df["year_month"] = df["order_date"].dt.to_period("M").astype(str)
    stats["clean_rows"] = len(df)

    return df, missing_costs_df, stats


# ==========================================
# 4. REPORT AGGREGATIONS
# ==========================================
def build_reports(
    df: pd.DataFrame,
    missing_costs_df: pd.DataFrame,
    stats: dict[str, Any],
    currency: str,
) -> dict[str, pd.DataFrame]:
    """Builds Executive KPI, Product Profitability, Monthly Trend, and Category sheets."""
    has_profit = stats["has_profit"]
    lb = profit_labels(stats["profit_kind"], currency)

    gross_sales = float(df["gross_sales"].sum())
    returned_sales = float(df["returned_sales"].sum())
    net_sales = float(df["net_sales"].sum())
    costed_net_sales = float(df["costed_net_sales"].sum()) if has_profit else 0.0
    coverage = (costed_net_sales / net_sales) if (has_profit and net_sales != 0) else float("nan")

    profit = float(df["gross_profit"].sum(min_count=1)) if has_profit else float("nan")
    margin = (
        (profit / costed_net_sales)
        if (has_profit and pd.notna(profit) and costed_net_sales != 0)
        else float("nan")
    )

    total_orders = int(df["order_id"].nunique())
    returned_orders = int(df["returned_order_id"].nunique())
    order_return_rate = (returned_orders / total_orders) if total_orders else 0.0
    value_return_rate = (returned_sales / gross_sales) if gross_sales else 0.0

    def _v(x: float, fmt: str) -> tuple[Any, str]:
        return (round(x, 4 if fmt == "percent" else 2), fmt) if pd.notna(x) else ("N/A", "text")

    cn_val, cn_fmt = _v(costed_net_sales, "currency") if has_profit else ("N/A", "text")
    cov_val, cov_fmt = _v(coverage, "percent")
    pr_val, pr_fmt = _v(profit, "currency")
    mg_val, mg_fmt = _v(margin, "percent")

    kpi_rows = [
        {"Metric": "Reporting Period", "Value": f"{df['order_date'].min():%Y-%m-%d} to {df['order_date'].max():%Y-%m-%d}", "Format": "text"},
        {"Metric": f"Gross Sales ({currency})", "Value": round(gross_sales, 2), "Format": "currency"},
        {"Metric": f"Returned Sales Value ({currency})", "Value": round(returned_sales, 2), "Format": "currency"},
        {"Metric": f"Net Sales ({currency})", "Value": round(net_sales, 2), "Format": "currency"},
        {"Metric": lb["costed_metric"], "Value": cn_val, "Format": cn_fmt},
        {"Metric": lb["coverage_metric"], "Value": cov_val, "Format": cov_fmt},
        {"Metric": lb["profit_metric"], "Value": pr_val, "Format": pr_fmt},
        {"Metric": lb["margin_metric"], "Value": mg_val, "Format": mg_fmt},
        {"Metric": "Total Orders", "Value": total_orders, "Format": "int"},
        {"Metric": "Returned Orders", "Value": returned_orders, "Format": "int"},
        {"Metric": "Order Return Rate", "Value": round(order_return_rate, 4), "Format": "percent"},
        {"Metric": "Sales Value Return Rate", "Value": round(value_return_rate, 4), "Format": "percent"},
        {"Metric": "Net Units Sold", "Value": int(df["net_quantity"].sum()), "Format": "int"},
    ]
    kpi_df = pd.DataFrame(kpi_rows)

    def _aggregate_group(group_cols: list[str], rename_map: dict[str, str], sort_col: str, ascending: bool) -> pd.DataFrame:
        agg_spec: dict[str, tuple[str, Any]] = {
            "Gross_Sales": ("gross_sales", "sum"),
            "Returned_Sales": ("returned_sales", "sum"),
            "Net_Sales": ("net_sales", "sum"),
            "Net_Units": ("net_quantity", "sum"),
            "Orders": ("order_id", "nunique"),
            "Returned_Orders": ("returned_order_id", "nunique"),
        }
        if has_profit:
            agg_spec[lb["costed_col"]] = ("costed_net_sales", "sum")
            agg_spec[lb["profit_col"]] = ("gross_profit", lambda s: s.sum(min_count=1))

        grouped = df.groupby(group_cols, dropna=False, as_index=False).agg(**agg_spec).rename(columns=rename_map)

        if has_profit:
            grouped[lb["margin_col"]] = safe_divide(grouped[lb["profit_col"]], grouped[lb["costed_col"]], fill_zero=False)
            grouped[lb["coverage_col"]] = safe_divide(grouped[lb["costed_col"]], grouped["Net_Sales"], fill_zero=True)

        grouped["Return_Rate"] = safe_divide(grouped["Returned_Orders"], grouped["Orders"], fill_zero=True)
        return grouped.sort_values(sort_col, ascending=ascending).reset_index(drop=True)

    prod_df = _aggregate_group(
        ["product_name", "category"],
        {"product_name": "Product Name", "category": "Category"},
        sort_col="Net_Sales",
        ascending=False,
    )

    monthly_df = _aggregate_group(
        ["year_month"],
        {"year_month": "Month"},
        sort_col="Month",
        ascending=True,
    )
    monthly_df["MoM_Net_Sales_Growth"] = safe_divide(
        monthly_df["Net_Sales"].diff(),
        monthly_df["Net_Sales"].shift(1),
        fill_zero=False,
    )

    cat_df = _aggregate_group(
        ["category"],
        {"category": "Category"},
        sort_col="Return_Rate",
        ascending=False,
    )

    reports: dict[str, pd.DataFrame] = {
        "Executive_KPI": kpi_df,
        "Product_Profitability": prod_df,
        "Monthly_Trend": monthly_df,
        "Category_Returns": cat_df,
    }
    if not missing_costs_df.empty:
        reports["Missing_Costs"] = missing_costs_df.rename(
            columns={"sku": "SKU", "product_name": "Product Name", "net_sales": "Net_Sales"}
        )

    return reports


# ==========================================
# 5. FORMATTED EXCEL EXPORT
# ==========================================
def export_to_excel(reports: dict[str, pd.DataFrame], output_path: Path, currency: str) -> None:
    """Exports styled multi-sheet Excel workbook with dynamic currency formatting."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    curr_fmt = CURRENCY_EXCEL_FORMATS.get(currency, f'#,##0.00 "{currency}"')

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    currency_cols = {
        "Gross_Sales", "Returned_Sales", "Net_Sales",
        "Costed_Net_Sales", "Net_Sales_With_Profit",
        "Gross_Profit", "Profit",
    }
    percent_cols = {
        "Gross_Margin", "Margin", "Cost_Coverage", "Profit_Coverage",
        "Return_Rate", "MoM_Net_Sales_Growth",
    }
    int_cols = {"Net_Units", "Orders", "Returned_Orders"}

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in reports.items():
            export_df = df.drop(columns=["Format"]) if sheet_name == "Executive_KPI" else df
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for col_idx, col_name in enumerate(export_df.columns, start=1):
                for row_idx in range(2, len(export_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value is None:
                        continue
                    if sheet_name == "Executive_KPI" and col_idx == 2:
                        fmt_type = df.iloc[row_idx - 2]["Format"]
                        if fmt_type == "currency":
                            cell.number_format = curr_fmt
                        elif fmt_type == "percent":
                            cell.number_format = "0.00%"
                        elif fmt_type == "int":
                            cell.number_format = "#,##0"
                    elif col_name in currency_cols:
                        cell.number_format = curr_fmt
                    elif col_name in percent_cols:
                        cell.number_format = "0.00%"
                    elif col_name in int_cols:
                        cell.number_format = "#,##0"

                max_len = max(
                    [len(str(col_name))]
                    + [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, len(export_df) + 2)]
                )
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 44)


# ==========================================
# 6. ENGLISH EXECUTIVE SUMMARY GENERATOR
# ==========================================
def generate_summary_text(
    reports: dict[str, pd.DataFrame],
    stats: dict[str, Any],
    currency: str,
) -> str:
    """Produces a concise, crash-safe English executive summary."""
    kind = stats["profit_kind"]
    lb = profit_labels(kind, currency)
    kpi = dict(zip(reports["Executive_KPI"]["Metric"], reports["Executive_KPI"]["Value"]))
    products = reports["Product_Profitability"]
    monthly = reports["Monthly_Trend"]
    categories = reports["Category_Returns"]

    best_month = monthly.loc[monthly["Net_Sales"].idxmax()]
    worst_month = monthly.loc[monthly["Net_Sales"].idxmin()]
    highest_ret_cat = categories.iloc[0]

    title = "E-COMMERCE SALES & PROFIT REPORT" if kind == "source" else "E-COMMERCE SALES & GROSS PROFIT REPORT"
    lines = [
        "==============================================================",
        f"{title:^62}",
        "==============================================================",
        f"Period          : {kpi['Reporting Period']}",
        f"Data Cleaning   : {stats['raw_rows']} raw rows -> {stats['duplicates_removed']} duplicates, "
        f"{stats['invalid_rows_removed']} invalid rows removed ({stats['clean_rows']} valid)",
        f"Dedup Strategy  : {stats['dedup_mode']}",
        f"Date Parsing    : {stats['date_mode']}",
        f"Return Source   : {stats['return_source']}",
        f"Profit Basis    : {stats['profit_source']}",
        "",
        "--- OVERALL PERFORMANCE (NET OF RETURNS) ---",
        f"Gross Sales     : {format_currency_text(kpi[f'Gross Sales ({currency})'], currency)}",
        f"Returned Value  : {format_currency_text(kpi[f'Returned Sales Value ({currency})'], currency)} "
        f"({kpi['Sales Value Return Rate'] * 100:.2f}% of gross)",
        f"Net Sales       : {format_currency_text(kpi[f'Net Sales ({currency})'], currency)}",
    ]

    costed_val = kpi[lb["costed_metric"]]
    cov_val = kpi[lb["coverage_metric"]]
    pr_val = kpi[lb["profit_metric"]]
    mg_val = kpi[lb["margin_metric"]]

    if stats["has_profit"] and is_num(pr_val):
        costed_str = format_currency_text(costed_val, currency) if is_num(costed_val) else "N/A"
        cov_str = f"{cov_val * 100:.2f}% coverage" if is_num(cov_val) else "N/A coverage"
        mg_str = f"{mg_val * 100:.2f}%" if is_num(mg_val) else "N/A (no non-zero net sales with profit data)"
        if kind == "source":
            lines.extend([
                f"Net Sales w/ Profit Data: {costed_str} ({cov_str})",
                f"Profit (source) : {format_currency_text(pr_val, currency)} (as defined by the source dataset)",
                f"Profit Margin   : {mg_str}",
            ])
        else:
            lines.extend([
                f"Costed Net Sales: {costed_str} ({cov_str})",
                f"Gross Profit    : {format_currency_text(pr_val, currency)} (after COGS only)",
                f"Gross Margin    : {mg_str} (calculated on costed sales only)",
            ])
    else:
        lines.append("Gross Profit    : N/A (provide --costs CSV or a Profit column to enable margin analysis)")

    lines.extend([
        f"Order Returns   : {kpi['Order Return Rate'] * 100:.2f}% "
        f"({kpi['Returned Orders']} / {kpi['Total Orders']} orders)",
        "",
        "--- MONTHLY TREND (NET SALES) ---",
        f"Best Month      : {best_month['Month']} ({format_currency_text(best_month['Net_Sales'], currency)})",
        f"Lowest Month    : {worst_month['Month']} ({format_currency_text(worst_month['Net_Sales'], currency)})",
    ])

    pcol, mcol = lb["profit_col"], lb["margin_col"]
    if stats["has_profit"]:
        valid_profit = products.dropna(subset=[pcol])
        top_profit = valid_profit.sort_values(pcol, ascending=False).head(3)
        loss_makers = valid_profit[valid_profit[pcol] < 0].sort_values(pcol).head(3)

        lines.append(f"\n--- TOP 3 PRODUCTS BY {lb['profit_title']} ---")
        for _, row in top_profit.iterrows():
            margin_str = f"{row[mcol] * 100:.1f}%" if pd.notna(row[mcol]) else "N/A"
            lines.append(
                f"- {row['Product Name']}: {format_currency_text(row[pcol], currency)} "
                f"(Margin: {margin_str} | Return Rate: {row['Return_Rate'] * 100:.1f}%)"
            )

        if not loss_makers.empty:
            lines.append(f"\n--- ATTENTION: {lb['negative_title']} ---")
            for _, row in loss_makers.iterrows():
                margin_str = f"{row[mcol] * 100:.1f}%" if pd.notna(row[mcol]) else "N/A"
                lines.append(
                    f"- {row['Product Name']}: {format_currency_text(row[pcol], currency)} "
                    f"(Margin: {margin_str})"
                )
    else:
        lines.append("\n--- TOP 3 PRODUCTS BY NET SALES ---")
        for _, row in products.head(3).iterrows():
            lines.append(
                f"- {row['Product Name']}: {format_currency_text(row['Net_Sales'], currency)} "
                f"(Return Rate: {row['Return_Rate'] * 100:.1f}%)"
            )

    lines.extend([
        "",
        "--- RETURN RISK ---",
        f"Highest Return Category: {highest_ret_cat['Category']} "
        f"({highest_ret_cat['Return_Rate'] * 100:.2f}% of orders)",
    ])

    if kind == "cogs":
        lines.extend([
            "",
            "--- ACCOUNTING NOTES ---",
            "* Gross Profit deducts product unit costs (COGS) only; marketplace commissions,",
            "  shipping, ads, payment fees, and overhead are NOT deducted.",
            "* VAT consistency: Ensure sales and unit costs use the same tax basis",
            "  (both net ex-VAT or both gross incl-VAT) to avoid distorted margins.",
        ])
    elif kind == "source":
        lines.extend([
            "",
            "--- ACCOUNTING NOTES ---",
            "* Profit is taken as provided in the source CSV; which costs it includes is",
            "  defined by the dataset, not by this script.",
        ])

    if stats["warnings"]:
        lines.append("\n--- DATA QUALITY WARNINGS ---")
        for w in stats["warnings"]:
            lines.append(f"! {w}")

    lines.append("==============================================================")
    return "\n".join(lines)


# ==========================================
# 7. CLI ENTRYPOINT
# ==========================================
def run_pipeline(
    sales_path: Path,
    preset_name: str,
    config_path: Path | None,
    costs_path: Path | None,
    returns_path: Path | None,
    currency: str,
    outdir: Path,
    sep: str | None = None,
    date_format: str | None = None,
    dayfirst: bool | None = None,
    drop_exact_duplicates: bool = False,
) -> tuple[str, Path, Path]:
    """Runs the full pipeline and returns (summary_text, excel_path, summary_path)."""
    config = load_mapping_config(preset_name, config_path)
    raw_df = load_csv_safely(sales_path, sep=sep)
    costs_df = load_csv_safely(costs_path, sep=sep) if costs_path else None
    returns_df = load_csv_safely(returns_path, sep=sep) if returns_path else None

    clean_df, missing_costs_df, stats = clean_and_enrich(
        raw_df=raw_df,
        config=config,
        costs_df=costs_df,
        returns_df=returns_df,
        cli_date_format=date_format,
        cli_dayfirst=dayfirst,
        drop_exact_duplicates=drop_exact_duplicates,
    )

    reports = build_reports(clean_df, missing_costs_df, stats, currency)
    excel_path = outdir / "sales_report.xlsx"
    summary_path = outdir / "executive_summary.txt"

    export_to_excel(reports, excel_path, currency)
    summary_text = generate_summary_text(reports, stats, currency)
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(summary_text, encoding="utf-8")
    return summary_text, excel_path, summary_path


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(
        description="E-commerce Sales & Profit Report Generator"
    )
    parser.add_argument("--input", type=Path, help="Path to sales CSV export")
    parser.add_argument("--costs", type=Path, default=None, help="Optional product unit costs CSV (SKU/Product + Unit Cost)")
    parser.add_argument("--returns", type=Path, default=None, help="Optional separate returns CSV")
    parser.add_argument("--preset", choices=list(PRESETS.keys()), default="global_superstore", help="Column mapping preset (default: global_superstore)")
    parser.add_argument("--config", type=Path, default=None, help="Custom JSON column mapping config")
    parser.add_argument("--currency", type=str, default="PLN", help="ISO currency code, e.g., PLN, EUR, USD (default: PLN)")
    parser.add_argument("--date-format", type=str, default=None, help="Strict strftime date format, e.g., %%d.%%m.%%Y")
    parser.add_argument("--dayfirst", action=argparse.BooleanOptionalAction, default=None, help="Parse DD/MM/YYYY when no date format is set")
    parser.add_argument("--drop-exact-duplicates", action="store_true", help="Allow full-row dedup when no row_id column exists")
    parser.add_argument("--sep", type=str, default=None, help="Force CSV delimiter (e.g., ';' or ',')")
    parser.add_argument("--outdir", type=Path, default=Path("output"), help="Output directory (default: output)")
    parser.add_argument("--demo", action="store_true", help="Generate synthetic Polish-schema demo CSVs and run report")
    args = parser.parse_args()

    currency = args.currency.upper()

    if args.demo:
        sales_path, costs_path, config_path = create_demo_files(args.outdir)
        print(f"[Info] Synthetic demo files generated in: {sales_path.parent}")
        print(f"[Info] Loaded generated demo config: {config_path}")
        preset_name = "example_pl"
        returns_path = None
    elif args.input:
        sales_path = args.input
        costs_path = args.costs
        config_path = args.config
        preset_name = args.preset
        returns_path = args.returns
    else:
        parser.error("Provide --input <sales.csv> or run with --demo.")

    summary_text, excel_path, summary_path = run_pipeline(
        sales_path=sales_path,
        preset_name=preset_name,
        config_path=config_path,
        costs_path=costs_path,
        returns_path=returns_path,
        currency=currency,
        outdir=args.outdir,
        sep=args.sep,
        date_format=args.date_format,
        dayfirst=args.dayfirst,
        drop_exact_duplicates=args.drop_exact_duplicates,
    )

    print(summary_text)
    print(f"\n[Done] Excel workbook : {excel_path}")
    print(f"[Done] Summary report : {summary_path}")


if __name__ == "__main__":
    main()
