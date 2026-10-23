"""PDF export helpers for Ledgerly Business Report Automator."""
from __future__ import annotations

from io import BytesIO
import html
import pandas as pd


def _safe_text(value: object) -> str:
    return html.escape(str(value))


def _money(value: float, symbol: str) -> str:
    if pd.isna(value):
        value = 0
    return f"{symbol}{value:,.2f}"


def build_pdf_report(
    analysis: pd.DataFrame,
    source_name: str,
    quality: dict,
    currency_symbol: str,
    date_range_label: str,
    categories_label: str,
) -> bytes:
    """Generate a compact PDF summary plus tables using ReportLab."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether

    buff = BytesIO()
    doc = SimpleDocTemplate(
        buff, pagesize=landscape(A4), rightMargin=16*mm, leftMargin=16*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Business performance report", author="Ledgerly Business Report Automator",
    )
    navy = colors.HexColor("#172B4D")
    teal = colors.HexColor("#0E9384")
    muted = colors.HexColor("#66758B")
    line = colors.HexColor("#E5EAF0")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=23, leading=27, textColor=navy, alignment=TA_LEFT, spaceAfter=5))
    styles.add(ParagraphStyle(name="Subtle", parent=styles["Normal"], fontSize=9, leading=13, textColor=muted, spaceAfter=4))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=13, leading=17, textColor=navy, spaceBefore=12, spaceAfter=6))
    styles.add(ParagraphStyle(name="Cell", parent=styles["BodyText"], fontSize=8, leading=10, textColor=navy))
    styles.add(ParagraphStyle(name="RightCell", parent=styles["BodyText"], fontSize=8, leading=10, textColor=navy, alignment=TA_RIGHT))

    valid = analysis.dropna(subset=["_report_sales"])
    total = float(valid["_report_sales"].sum()) if not valid.empty else 0.0
    mean = float(valid["_report_sales"].mean()) if not valid.empty else 0.0
    order_values = analysis.loc[analysis["_report_order"] != "Unspecified", "_report_order"] if "_report_order" in analysis else pd.Series(dtype=object)
    orders = int(order_values.nunique()) if not order_values.empty else len(analysis)
    title = [
        Paragraph("Business performance report", styles["ReportTitle"]),
        Paragraph(f"Source: {_safe_text(source_name)} &nbsp;&nbsp; | &nbsp;&nbsp; Rows in view: {len(analysis):,}", styles["Subtle"]),
        Paragraph(f"Date filter: {_safe_text(date_range_label)} &nbsp;&nbsp; | &nbsp;&nbsp; Categories: {_safe_text(categories_label)}", styles["Subtle"]),
        Spacer(1, 5*mm),
    ]

    kpi_data = [
        ["TOTAL SALES", "ROWS ANALYSED", "AVERAGE SALES / ROW", "UNIQUE ORDERS"],
        [_money(total, currency_symbol), f"{int(valid.shape[0]):,}", _money(mean, currency_symbol), f"{orders:,}"],
    ]
    kpi_table = Table(kpi_data, colWidths=[61*mm]*4, rowHeights=[8*mm, 13*mm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,0), 8),
        ("BACKGROUND", (0,1), (-1,1), colors.HexColor("#F4F8FA")),
        ("TEXTCOLOR", (0,1), (-1,1), navy), ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE", (0,1), (-1,1), 14), ("BOX", (0,0), (-1,-1), .5, line),
        ("INNERGRID", (0,0), (-1,-1), .5, line), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 9),
    ]))

    story = title + [kpi_table, Paragraph("Key observations", styles["Section"])]
    observations = []
    dated = analysis.dropna(subset=["_report_date", "_report_sales"]).copy()
    if not dated.empty:
        dated["Month"] = dated["_report_date"].dt.to_period("M").astype(str)
        monthly = dated.groupby("Month")["_report_sales"].sum().sort_index()
        if not monthly.empty:
            peak_month = monthly.idxmax()
            observations.append(f"Highest-sales month in the current view: <b>{_safe_text(peak_month)}</b> ({_money(float(monthly.max()), currency_symbol)}).")
    cat = analysis.dropna(subset=["_report_sales"]).groupby("_report_category")["_report_sales"].sum().sort_values(ascending=False)
    if not cat.empty:
        observations.append(f"Top category by sales in the current view: <b>{_safe_text(cat.index[0])}</b> ({_money(float(cat.iloc[0]), currency_symbol)}).")
    if quality.get("duplicate_rows_detected", 0):
        observations.append(f"The source contains {quality['duplicate_rows_detected']:,} exact duplicate row(s); removal setting: {quality['duplicate_rows_removed']:,} removed.")
    if quality.get("invalid_sales_values", 0):
        observations.append(f"{quality['invalid_sales_values']:,} row(s) have sales values that could not be parsed as numbers.")
    if not observations:
        observations.append("No additional pattern could be generated from the currently selected data.")
    for observation in observations[:5]:
        story.append(Paragraph("• " + observation, styles["BodyText"]))
        story.append(Spacer(1, 2*mm))

    story.append(Paragraph("Category performance", styles["Section"]))
    if cat.empty:
        story.append(Paragraph("No valid sales values are available for category analysis.", styles["Subtle"]))
    else:
        cat_rows = [["Category", "Sales", "Share of visible sales"]]
        denom = float(cat.sum())
        for label, value in cat.head(12).items():
            share = (float(value) / denom * 100) if denom else 0
            cat_rows.append([_safe_text(label), _money(float(value), currency_symbol), f"{share:.1f}%"])
        cat_table = Table(cat_rows, colWidths=[100*mm, 45*mm, 45*mm], repeatRows=1)
        cat_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), teal), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8),
            ("GRID", (0,0), (-1,-1), .35, line), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F9FB")]),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(cat_table)

    story.append(Paragraph("Data quality summary", styles["Section"]))
    qrows = [["Check", "Count"],
        ["Exact duplicate rows detected", f"{quality.get('duplicate_rows_detected', 0):,}"],
        ["Duplicate rows removed", f"{quality.get('duplicate_rows_removed', 0):,}"],
        ["Rows with missing values", f"{quality.get('rows_with_missing_values', 0):,}"],
        ["Missing cells", f"{quality.get('missing_cells', 0):,}"],
        ["Invalid sales values", f"{quality.get('invalid_sales_values', 0):,}"],
        ["Invalid / unreadable dates", f"{quality.get('invalid_dates', 0):,}"],
        ["Negative sales rows (possibly refunds / adjustments)", f"{quality.get('negative_sales_rows', 0):,}"],]
    qtable = Table(qrows, colWidths=[120*mm, 35*mm], repeatRows=1)
    qtable.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), .35, line), ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F9FB")]),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(qtable)
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("Note: This report is descriptive analysis, not an accounting statement. Verify figures against the source system before making business decisions.", styles["Subtle"]))

    def add_page_number(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(line)
        canvas.line(16*mm, 11*mm, landscape(A4)[0] - 16*mm, 11*mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(muted)
        canvas.drawString(16*mm, 7*mm, "Generated by Ledgerly Business Report Automator")
        canvas.drawRightString(landscape(A4)[0] - 16*mm, 7*mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buff.getvalue()
