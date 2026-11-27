"""
report_builder.py
------------------
Builds the Error Report .xlsx from the downloaded JSON. Rewritten to match
the logic in your reference script (error_report_extract.txt) exactly:

    records.append({
        "LineId": row.get("lineId"),
        "HandlingUnit": row.get("handlingUnit"),
        "Article": str(row.get("article")),
        "Location": str(row.get("location")),
        "ErrorDescription": err.get("errorDescription"),
    })
    df.groupby("Article").size()   # sorted ascending by Article - pandas' default
    df.groupby("Location").size()  # sorted ascending by Location - pandas' default

Sheet layout (matches the reference script's pd.ExcelWriter calls):

    Summary      | rows 1-4: Metric/Count (Total Errors, Unique Articles,
                   Unique Locations); rows 5-6 blank; row 7: Article/Error_Count
                   headers at col A/B and Location/Error_Count at col E/F
                   (two-column gap, matching the reference's startcol=4);
                   row 8+: both group summaries side by side.
    All_Errors   | LineId, HandlingUnit, Article, Location, ErrorDescription -
                   one row per individual error entry, in original order.
    By_Article   | Article, Error_Count - sorted ascending by Article.
    By_Location  | Location, Error_Count - sorted ascending by Location.

Source JSON shape (confirmed): {"header": {...}, "data": [...], "IsSuccess": bool}
Each record: {"lineId", "handlingUnit" (sometimes absent), "article", "location",
"errors": [{"errorDescription": ...}]}.
"""

from __future__ import annotations

from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

import config
import pdf_builder

THIN = Side(style="thin", color="000000")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _get_records(json_data: dict) -> list[dict]:
    data = json_data.get("data")
    return data if isinstance(data, list) else []


def _flatten_errors(records: list[dict]) -> list[dict]:
    """One row per individual error entry, with the field names/casing the
    reference script uses. A record with no "errors" (or an empty list)
    still produces one row (ErrorDescription=None), so it isn't dropped."""
    flattened = []
    for rec in records:
        if not isinstance(rec, dict):
            continue
        errors = rec.get(config.ERRORS_FIELD) or [{}]
        for err in errors:
            flattened.append({
                "LineId": rec.get("lineId"),
                "HandlingUnit": rec.get("handlingUnit"),
                "Article": str(rec.get(config.ARTICLE_FIELD)),
                "Location": str(rec.get(config.LOCATION_FIELD)),
                "ErrorDescription": (err or {}).get("errorDescription"),
            })
    return flattened


def _group_counts(flat_rows: list[dict], field: str) -> list[tuple[str, int]]:
    """Mirrors df.groupby(field).size(): sorted ascending by the group value
    (pandas' default groupby sort), not by count."""
    counts = Counter(r[field] for r in flat_rows)
    return sorted(counts.items(), key=lambda kv: kv[0])


def _write_plain_header(ws, headers: list[str], row: int = 1, start_col: int = 1):
    for offset, header in enumerate(headers):
        ws.cell(row=row, column=start_col + offset, value=header).font = Font(bold=True)


def _write_summary_header(ws, headers: list[str]):
    fill = PatternFill(start_color=config.COLOR_SUMMARY_HEADER_FILL,
                        end_color=config.COLOR_SUMMARY_HEADER_FILL, fill_type="solid")
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color=config.COLOR_WHITE)
        cell.fill = fill


def _autosize_columns(ws, min_width=10, max_width=50):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = col_cells[0].column_letter
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, length + 2))


# ---------------------------------------------------------------------------
# main builder
# ---------------------------------------------------------------------------

def build_error_report_xlsx(json_data: dict, xlsx_path: str, pdf_path: str | None = None,
                            title: str = "Error Report") -> dict:
    """Returns the metrics shown on the dashboard:
        {"total_errors": int, "unique_articles": int, "unique_locations": int}"""
    records = _get_records(json_data)
    flat_rows = _flatten_errors(records)

    total_errors = len(flat_rows)
    unique_articles = len({r["Article"] for r in flat_rows}) if flat_rows else 0
    unique_locations = len({r["Location"] for r in flat_rows}) if flat_rows else 0

    article_summary = _group_counts(flat_rows, "Article") if flat_rows else []
    location_summary = _group_counts(flat_rows, "Location") if flat_rows else []

    wb = Workbook()
    wb.remove(wb.active)

    # --- Summary ---
    ws_summary = wb.create_sheet("Summary")
    _write_summary_header(ws_summary, ["Metric", "Count"])
    for row_idx, (metric, value) in enumerate(
        [("Total Errors", total_errors), ("Unique Articles", unique_articles),
         ("Unique Locations", unique_locations)],
        start=2,
    ):
        ws_summary.cell(row=row_idx, column=1, value=metric)
        ws_summary.cell(row=row_idx, column=2, value=value)

    if article_summary or location_summary:
        header_row = 7  # rows 5-6 blank, matching the reference script's startrow=6 (0-indexed)
        _write_plain_header(ws_summary, ["Article", "Error_Count"], row=header_row, start_col=1)
        _write_plain_header(ws_summary, ["Location", "Error_Count"], row=header_row, start_col=5)

        data_start = header_row + 1
        for i in range(max(len(article_summary), len(location_summary))):
            row = data_start + i
            if i < len(article_summary):
                article, count = article_summary[i]
                ws_summary.cell(row=row, column=1, value=article)
                ws_summary.cell(row=row, column=2, value=count)
            if i < len(location_summary):
                location, count = location_summary[i]
                ws_summary.cell(row=row, column=5, value=location)
                ws_summary.cell(row=row, column=6, value=count)
    _autosize_columns(ws_summary)

    # --- All_Errors ---
    ws_all = wb.create_sheet("All_Errors")
    all_fields = ["LineId", "HandlingUnit", "Article", "Location", "ErrorDescription"]
    if flat_rows:
        _write_plain_header(ws_all, all_fields)
        for row_idx, row in enumerate(flat_rows, start=2):
            for col_idx, field in enumerate(all_fields, start=1):
                ws_all.cell(row=row_idx, column=col_idx, value=row.get(field))
    else:
        _write_plain_header(ws_all, ["Status", "Details"])
        ws_all.cell(row=2, column=1, value="No Errors")
        ws_all.cell(row=2, column=2, value="The JSON file contains an empty data array")
    _autosize_columns(ws_all)

    # --- By_Article ---
    ws_article = wb.create_sheet("By_Article")
    _write_plain_header(ws_article, ["Article", "Error_Count"])
    if article_summary:
        for row_idx, (article, count) in enumerate(article_summary, start=2):
            ws_article.cell(row=row_idx, column=1, value=article)
            ws_article.cell(row=row_idx, column=2, value=count)
    else:
        ws_article.cell(row=2, column=1, value="No data found in source JSON")
        ws_article.cell(row=2, column=2, value=0)
    _autosize_columns(ws_article)

    # --- By_Location ---
    ws_location = wb.create_sheet("By_Location")
    _write_plain_header(ws_location, ["Location", "Error_Count"])
    if location_summary:
        for row_idx, (location, count) in enumerate(location_summary, start=2):
            ws_location.cell(row=row_idx, column=1, value=location)
            ws_location.cell(row=row_idx, column=2, value=count)
    else:
        ws_location.cell(row=2, column=1, value="No data found in source JSON")
        ws_location.cell(row=2, column=2, value=0)
    _autosize_columns(ws_location)

    wb.save(xlsx_path)

    if pdf_path:
        writer = pdf_builder.PdfWriter(pdf_path)
        writer.add_title(title)
        writer.add_section("Summary", ["Metric", "Count"], [
            ["Total Errors", total_errors],
            ["Unique Articles", unique_articles],
            ["Unique Locations", unique_locations],
        ])
        writer.force_page_break()  # Summary stands alone as its own page
        writer.add_section("By_Article", ["Article", "Error_Count"],
                            [[a, c] for a, c in article_summary])
        writer.add_section("By_Location", ["Location", "Error_Count"],
                            [[l, c] for l, c in location_summary])
        writer.add_section("All_Errors", all_fields,
                            [[row.get(f) for f in all_fields] for row in flat_rows])
        writer.save()

    return {
        "total_errors": total_errors,
        "unique_articles": unique_articles,
        "unique_locations": unique_locations,
    }


# ---------------------------------------------------------------------------
# Comparison report - structure/colors confirmed from real reference:
# 268_inventory_comparison_20260911.json / 268_Comparison_Report_20260911.xlsx
# ---------------------------------------------------------------------------

def _join_reason_codes(value):
    """reasonCodes is a list in the JSON - joined with ", " in the xlsx,
    blank (None) when empty, matching the reference file exactly."""
    if isinstance(value, list):
        return ", ".join(str(v) for v in value) if value else None
    return value


def _write_comparison_header(ws, headers: list[str], fill_hex: str, font_color_hex: str | None = "FFFFFF"):
    fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color=font_color_hex) if font_color_hex else Font(bold=True)
        cell.fill = fill
        cell.border = BORDER_ALL


def build_comparison_report_xlsx(json_data: dict, xlsx_path: str, pdf_path: str | None = None) -> dict:
    """Returns the metrics shown on the dashboard:
        {"percentage_quantity_match": float|None, "percentage_value_match": float|None,
         "quantity_mismatch_count": int, "missing_in_scx_count": int, "missing_in_mhs_count": int}"""
    data = json_data.get("data", {}) or {}
    quantity_mismatches = data.get("quantityMismatches", []) or []
    missing_in_mhs = data.get("missingInMHS", []) or []
    missing_in_scx = data.get("missingInSCX", []) or []
    summary = data.get("summary", {}) or {}

    wb = Workbook()
    wb.remove(wb.active)

    # --- Quantity Mismatches ---
    ws_qm = wb.create_sheet("Quantity Mismatches")
    _write_comparison_header(ws_qm, config.COMPARISON_QM_COLUMNS, config.COLOR_QM_HEADER_FILL)
    for row_idx, rec in enumerate(quantity_mismatches, start=2):
        ws_qm.cell(row=row_idx, column=1, value=rec.get("articleNumber")).border = BORDER_ALL
        ws_qm.cell(row=row_idx, column=2, value=rec.get("mhsQuantity")).border = BORDER_ALL
        ws_qm.cell(row=row_idx, column=3, value=rec.get("scxQuantity")).border = BORDER_ALL
        ws_qm.cell(row=row_idx, column=4, value=f"=B{row_idx}-C{row_idx}").border = BORDER_ALL
        ws_qm.cell(row=row_idx, column=5, value=rec.get("price")).border = BORDER_ALL
        ws_qm.cell(row=row_idx, column=6, value=_join_reason_codes(rec.get("reasonCodes"))).border = BORDER_ALL
    _autosize_columns(ws_qm)

    # --- Missing In SCX / Missing In MHS (same layout, different data) ---
    def _write_missing_sheet(name: str, records: list[dict]):
        ws = wb.create_sheet(name)
        _write_comparison_header(ws, config.COMPARISON_MISSING_COLUMNS, config.COLOR_MISSING_HEADER_FILL)
        for row_idx, rec in enumerate(records, start=2):
            ws.cell(row=row_idx, column=1, value=rec.get("articleNumber")).border = BORDER_ALL
            ws.cell(row=row_idx, column=2, value=rec.get("quantity")).border = BORDER_ALL
            ws.cell(row=row_idx, column=3, value=rec.get("price")).border = BORDER_ALL
            ws.cell(row=row_idx, column=4, value=_join_reason_codes(rec.get("reasonCodes"))).border = BORDER_ALL
        _autosize_columns(ws)

    _write_missing_sheet("Missing In SCX", missing_in_scx)
    _write_missing_sheet("Missing In MHS", missing_in_mhs)

    # --- Summary ---
    ws_summary = wb.create_sheet("Summary")
    _write_comparison_header(ws_summary, ["Metric", "Value"], config.COLOR_COMPARISON_SUMMARY_HEADER_FILL,
                              font_color_hex=None)  # reference file's Summary header text is default (not white)
    for row_idx, (key, value) in enumerate(summary.items(), start=2):
        ws_summary.cell(row=row_idx, column=1, value=key).border = BORDER_ALL
        ws_summary.cell(row=row_idx, column=2, value=value).border = BORDER_ALL
    _autosize_columns(ws_summary)

    wb.save(xlsx_path)

    if pdf_path:
        def qty_diff(rec):
            mhs, scx = rec.get("mhsQuantity"), rec.get("scxQuantity")
            return (mhs - scx) if isinstance(mhs, (int, float)) and isinstance(scx, (int, float)) else None

        writer = pdf_builder.PdfWriter(pdf_path)
        writer.add_title("Comparison Report")
        writer.add_section("Summary", ["Metric", "Value"], [[k, v] for k, v in summary.items()])
        writer.force_page_break()  # Summary stands alone as its own first page
        writer.add_section("Quantity Mismatches", config.COMPARISON_QM_COLUMNS, [
            [rec.get("articleNumber"), rec.get("mhsQuantity"), rec.get("scxQuantity"), qty_diff(rec),
             rec.get("price"), _join_reason_codes(rec.get("reasonCodes"))]
            for rec in quantity_mismatches
        ])
        writer.add_section("Missing In SCX", config.COMPARISON_MISSING_COLUMNS, [
            [rec.get("articleNumber"), rec.get("quantity"), rec.get("price"),
             _join_reason_codes(rec.get("reasonCodes"))]
            for rec in missing_in_scx
        ])
        writer.add_section("Missing In MHS", config.COMPARISON_MISSING_COLUMNS, [
            [rec.get("articleNumber"), rec.get("quantity"), rec.get("price"),
             _join_reason_codes(rec.get("reasonCodes"))]
            for rec in missing_in_mhs
        ])
        writer.save()

    return {
        "percentage_quantity_match": summary.get("percentageQuantityMatch"),
        "percentage_value_match": summary.get("percentageValueMatch"),
        "quantity_mismatch_count": len(quantity_mismatches),
        "missing_in_scx_count": len(missing_in_scx),
        "missing_in_mhs_count": len(missing_in_mhs),
    }


# ---------------------------------------------------------------------------
# Fusion Inventory report - structure/colors confirmed from real reference:
# 268_Inventory_Fusion_Articles_20260911.json / 268_Inventory_Article_Report_20260911.xlsx
# (15653 real records) - not shown on the dashboard, saved to disk only.
# ---------------------------------------------------------------------------

def build_fusion_report_xlsx(json_data: dict, xlsx_path: str, pdf_path: str | None = None) -> dict:
    """Returns {"record_count": int} - used only for the run log, not the dashboard."""
    records = json_data.get("data", []) or []

    # Column order taken from the records themselves, in first-seen order -
    # matches the reference file exactly (handlingUnit, status, article,
    # location, quantity, ... lineId) without needing a fixed list.
    columns: list[str] = []
    for rec in records:
        if isinstance(rec, dict):
            for key in rec.keys():
                if key not in columns:
                    columns.append(key)

    # Reference file stores every value as text (str(value)), except it
    # renders a fully-empty column's nulls as the literal text "nan" - a
    # pandas artifact (an all-None column becomes float NaN, then str(NaN)
    # == "nan"). We use a clean blank cell for any null instead, which is
    # the same thing every other report in this app does and is easier to
    # read/filter in Excel than literal "nan" text. Computed once here and
    # reused for both the xlsx and the PDF.
    value_rows: list[list] = []
    for rec in records:
        row = []
        for col_name in columns:
            value = rec.get(col_name) if isinstance(rec, dict) else None
            row.append(str(value) if value is not None else None)
        value_rows.append(row)

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory_Report"

    title_cell = ws.cell(row=1, column=1, value="Inventory_Article_Report")
    title_cell.font = Font(bold=True)
    # row 2 intentionally blank

    header_row = 3
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color=config.COLOR_WHITE)
        cell.fill = PatternFill(start_color=config.COLOR_FUSION_HEADER_FILL,
                                 end_color=config.COLOR_FUSION_HEADER_FILL, fill_type="solid")
        cell.border = BORDER_ALL

    for row_offset, row_values in enumerate(value_rows):
        row = header_row + 1 + row_offset
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.border = BORDER_ALL

    # Print setup: landscape, fit to one page wide, repeat the header row on
    # every printed page - matches the reference file exactly.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"

    # No column auto-sizing here on purpose - at ~15k+ rows it's a
    # noticeable slowdown for a report nobody reviews column-by-column in
    # this app (it's saved to disk only, not previewed).

    wb.save(xlsx_path)

    if pdf_path:
        writer = pdf_builder.PdfWriter(pdf_path)
        writer.add_section("Inventory_Article_Report", columns, value_rows)
        writer.save()

    return {"record_count": len(records)}


# ---------------------------------------------------------------------------
# MHS Snapshot report - structure confirmed from real reference:
# 268_MHS_Snapshot_report_20260911_183613.json / .xlsx (17818 real events)
# ---------------------------------------------------------------------------

def _format_quantity(value):
    """Reference file shows whole-number float quantities as plain ints
    (132.0 -> 132) in the main totalQuantity column - matched here."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _join_storage_areas(storage_areas: list[dict] | None) -> str | None:
    """Each entry as "id:totalQuantity", multiple entries "; "-joined,
    blank when empty - matches the reference file exactly."""
    if not storage_areas:
        return None
    return "; ".join(f"{sa.get('id')}:{sa.get('totalQuantity')}" for sa in storage_areas)


def _join_reservations(reservations: list[dict] | None) -> str | None:
    """Each entry as "clientName:totalQuantity", "; "-joined, blank when empty."""
    if not reservations:
        return None
    return "; ".join(f"{r.get('clientName')}:{r.get('totalQuantity')}" for r in reservations)


def build_mhs_snapshot_xlsx(json_data: dict, xlsx_path: str, pdf_path: str | None = None) -> dict:
    """Returns {"record_count": int}."""
    events = ((json_data.get("data") or {}).get("inventoryEvents")) or []

    value_rows: list[list] = []
    for event in events:
        business_unit = event.get("businessUnit", {}) or {}
        item = event.get("item", {}) or {}
        stock_on_hand = event.get("stockOnHand", {}) or {}
        value_rows.append([
            business_unit.get("type"),
            business_unit.get("code"),
            item.get("number"),
            item.get("type"),
            item.get("uomCode"),
            _format_quantity(stock_on_hand.get("totalQuantity")),
            stock_on_hand.get("creationTimestamp"),
            _join_storage_areas(stock_on_hand.get("storageAreas")),
            _join_reservations(stock_on_hand.get("reservations")),
        ])

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"  # matches the reference file's (default) sheet name

    for col_idx, col_name in enumerate(config.MHS_SNAPSHOT_COLUMNS, start=1):
        ws.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)

    for row_offset, row_values in enumerate(value_rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_offset, column=col_idx, value=value)

    # No column auto-sizing / borders here on purpose - same performance
    # reasoning as the Fusion Inventory report (~18k rows), and the
    # reference file itself has no borders or fill on this sheet either.

    wb.save(xlsx_path)

    if pdf_path:
        writer = pdf_builder.PdfWriter(pdf_path)
        writer.add_section("MHS Snapshot", config.MHS_SNAPSHOT_COLUMNS, value_rows)
        writer.save()

    return {"record_count": len(events)}


# ---------------------------------------------------------------------------
# SCE Snapshot report - structure confirmed from real reference:
# 268_SCE_Snapshot_20260911.json / .xlsx (9911 real events)
# ---------------------------------------------------------------------------

def build_sce_snapshot_xlsx(json_data: dict, xlsx_path: str, pdf_path: str | None = None) -> dict:
    """Returns {"record_count": int}."""
    events = ((json_data.get("data") or {}).get("inventoryEvents")) or []

    value_rows: list[list] = []
    for event in events:
        business_unit = event.get("businessUnit", {}) or {}
        item = event.get("item", {}) or {}
        stock_on_hand = event.get("stockOnHand", {}) or {}
        storage_areas = stock_on_hand.get("storageAreas") or []
        value_rows.append([
            business_unit.get("type"),
            business_unit.get("code"),
            item.get("number"),
            item.get("type"),
            item.get("uomCode"),
            _format_quantity(stock_on_hand.get("totalQuantity")),
            stock_on_hand.get("creationTimestamp"),
            str(storage_areas) if storage_areas else None,  # matches reference: literal Python list-of-dicts text
        ])

    wb = Workbook()
    ws = wb.active
    ws.title = "Snapshot"  # matches the reference file's sheet name

    fill = PatternFill(start_color=config.COLOR_SCE_HEADER_FILL,
                        end_color=config.COLOR_SCE_HEADER_FILL, fill_type="solid")
    for col_idx, col_name in enumerate(config.SCE_SNAPSHOT_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color=config.COLOR_WHITE)
        cell.fill = fill

    for row_offset, row_values in enumerate(value_rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_offset, column=col_idx, value=value)

    # No column auto-sizing / borders here - same performance reasoning as
    # Fusion Inventory / MHS Snapshot, and the reference file has none either.

    wb.save(xlsx_path)

    if pdf_path:
        writer = pdf_builder.PdfWriter(pdf_path)
        writer.add_section("SCE Snapshot", config.SCE_SNAPSHOT_COLUMNS, value_rows)
        writer.save()

    return {"record_count": len(events)}


# ---------------------------------------------------------------------------
# Last Count Date Error Report - the second message on the error subscription.
# Same record layout and same workbook/PDF structure as the Error Report, so it
# reuses that builder; only the PDF title differs.
# ---------------------------------------------------------------------------

def build_last_count_date_error_report_xlsx(json_data: dict, xlsx_path: str,
                                             pdf_path: str | None = None) -> dict:
    return build_error_report_xlsx(json_data, xlsx_path, pdf_path,
                                   title="Last Count Date Error Report")
