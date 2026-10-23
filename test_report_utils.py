from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

from pdf_export import build_pdf_report

from report_utils import (
    NONE_LABEL,
    build_excel_report,
    guess_column,
    parse_numeric,
    prepare_analysis,
    read_data_file,
    visible_data,
)

BASE = Path(__file__).parent


def test_demo_file_loads_and_has_expected_rows():
    frame = read_data_file(BASE / "sample_sales.csv", "sample_sales.csv")
    assert len(frame) == 25
    assert "Total Sales" in frame.columns


def test_exact_duplicates_are_detected_and_removed():
    frame = pd.DataFrame({"Order": [1, 2, 2], "Sales": [100, 200, 200]})
    cleaned, analysis, quality = prepare_analysis(frame, sales_col="Sales", remove_duplicates=True)
    assert len(cleaned) == 2
    assert quality["duplicate_rows_detected"] == 1
    assert quality["duplicate_rows_removed"] == 1
    assert float(analysis["_report_sales"].sum()) == 300


def test_currency_parser_supports_commas_and_parentheses():
    values = parse_numeric(pd.Series(["₹1,200.50", "(20.00)", "bad", "3,500"]))
    assert values.iloc[0] == pytest.approx(1200.5)
    assert values.iloc[1] == pytest.approx(-20.0)
    assert pd.isna(values.iloc[2])
    assert values.iloc[3] == pytest.approx(3500)


def test_quantity_times_price_revenue_calculation():
    frame = pd.DataFrame({"Qty": [2, 3], "Price": [10, 5]})
    _, analysis, quality = prepare_analysis(
        frame,
        quantity_col="Qty",
        unit_price_col="Price",
        revenue_mode="quantity_price",
        remove_duplicates=False,
    )
    assert analysis["_report_sales"].tolist() == [20, 15]
    assert quality["invalid_sales_values"] == 0


def test_day_first_dates_can_be_parsed():
    frame = pd.DataFrame({"Date": ["13/10/2026"], "Sales": [75]})
    _, analysis, quality = prepare_analysis(frame, date_col="Date", sales_col="Sales", dayfirst_dates=True)
    assert analysis["_report_date"].iloc[0] == pd.Timestamp("2026-10-13")
    assert quality["invalid_dates"] == 0


def test_excel_report_is_a_valid_xlsx_and_does_not_expose_internal_columns():
    raw = read_data_file(BASE / "sample_sales.csv", "sample_sales.csv")
    columns = list(raw.columns)
    _, analysis, quality = prepare_analysis(
        raw,
        sales_col=guess_column(columns, ["total sales", "revenue"]),
        date_col=guess_column(columns, ["order date", "date"]),
        category_col=guess_column(columns, ["category"]),
        product_col=guess_column(columns, ["product"]),
        order_col=guess_column(columns, ["order id"]),
    )
    xlsx = build_excel_report(analysis, source_name="sample_sales.csv", quality=quality)
    assert xlsx[:4] == b"PK\x03\x04"
    workbook = pd.ExcelFile(BytesIO(xlsx))
    assert {"Overview", "Monthly trend", "Categories", "Top products", "Data quality", "Filtered data"}.issubset(set(workbook.sheet_names))
    exported_data = pd.read_excel(BytesIO(xlsx), sheet_name="Filtered data")
    assert not any(str(column).startswith("_report_") for column in exported_data.columns)


def test_guess_column_returns_not_mapped_for_unknown_headers():
    assert guess_column(["foo", "bar"], ["sales", "revenue"]) == NONE_LABEL


def test_pdf_report_is_generated():
    raw = read_data_file(BASE / "sample_sales.csv", "sample_sales.csv")
    columns = list(raw.columns)
    _, analysis, quality = prepare_analysis(
        raw,
        sales_col=guess_column(columns, ["total sales", "revenue"]),
        date_col=guess_column(columns, ["order date", "date"]),
        category_col=guess_column(columns, ["category"]),
        product_col=guess_column(columns, ["product"]),
        order_col=guess_column(columns, ["order id"]),
    )
    pdf = build_pdf_report(analysis, "sample_sales.csv", quality, "₹", "All dates", "All categories")
    assert pdf.startswith(b"%PDF-")
    assert len(pdf) > 1500
