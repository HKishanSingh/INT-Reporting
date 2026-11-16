# ============================================================
# FINAL PERFORMANCE REPORT + CREATIVE LEVEL BREAKDOWN
# Streamlit version
# Run: streamlit run app.py
# ============================================================

import io
import os
import re
import tempfile
import warnings
from datetime import datetime, date
from copy import copy

import numpy as np
import pandas as pd
import streamlit as st

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

warnings.filterwarnings("ignore", category=FutureWarning)

CURRENCY_FORMATS = {
    "1": '"$"#,##0.00',
    "2": '"€"#,##0.00',
    "3": '"£"#,##0.00',
    "4": '"₹"#,##0.00',
    "5": '"C$"#,##0.00',
    "6": '"A$"#,##0.00',
    "7": '"¥"#,##0',
    "8": '#,##0.00',
}

CURRENCY_NAMES = {
    "1": "USD ($)",
    "2": "EUR (€)",
    "3": "GBP (£)",
    "4": "INR (₹)",
    "5": "CAD (C$)",
    "6": "AUD (A$)",
    "7": "JPY (¥)",
    "8": "None (plain)",
}

header_fill = PatternFill(fill_type="solid", fgColor="BFBFBF")
total_fill = PatternFill(fill_type="solid", fgColor="BFBFBF")
white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
black_font = Font(name="Calibri", color="000000")
header_font = Font(name="Calibri", bold=True, color="000000")
total_font = Font(name="Calibri", bold=True, color="000000")
thin_side = Side(style="thin", color="000000")
table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)


def clean_column_names(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    duplicate_names = df.columns[df.columns.duplicated(keep=False)].unique()
    for col in duplicate_names:
        duplicate_data = df.loc[:, df.columns == col]
        combined = duplicate_data.bfill(axis=1).iloc[:, 0]
        df = df.loc[:, df.columns != col]
        df[col] = combined
    return df


def require_columns(df, required_columns, sheet_name):
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"{sheet_name} is missing required columns: " + ", ".join(missing))


def safe_ratio(numerator, denominator):
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    if isinstance(denominator, pd.Series):
        if not isinstance(numerator, pd.Series):
            numerator = pd.Series(numerator, index=denominator.index)
        else:
            numerator = numerator.reindex(denominator.index)
        result = numerator.div(denominator.where(denominator.notna() & denominator.ne(0)))
        return result.replace([np.inf, -np.inf], np.nan)
    if pd.isna(numerator) or pd.isna(denominator) or denominator == 0:
        return np.nan
    result = numerator / denominator
    return result if np.isfinite(result) else np.nan


def safe_ratio_zero(numerator, denominator):
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    if isinstance(denominator, pd.Series):
        if not isinstance(numerator, pd.Series):
            numerator = pd.Series(numerator, index=denominator.index)
        else:
            numerator = numerator.reindex(denominator.index)
        result = numerator.div(denominator.where(denominator.notna() & denominator.ne(0)))
        return result.replace([np.inf, -np.inf], np.nan).fillna(0)
    if pd.isna(numerator) or pd.isna(denominator) or denominator == 0:
        return 0
    result = numerator / denominator
    return result if np.isfinite(result) else 0


def format_date_columns(df):
    df = df.copy()
    for col in ["Campaign Start", "Campaign End"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    return df


def convert_numeric_columns(df, columns):
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def clean_and_convert_numeric(series):
    cleaned = (
        series.astype("string").str.strip()
        .str.replace(",", "", regex=False)
        .str.replace(r"[$€£₹¥]", "", regex=True)
        .str.replace(r"^[CA]\$", "", regex=True)
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def clean_percentage_series(series):
    s = series.astype("string").str.strip()
    s = (
        s.str.replace("%", "", regex=False)
         .str.replace(",", "", regex=False)
         .str.replace(r"[$€£₹¥]", "", regex=True)
         .str.replace(r"^[CA]\$", "", regex=True)
         .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    numeric = pd.to_numeric(s, errors="coerce")
    mask_large = numeric > 2
    numeric = numeric.where(~mask_large, numeric / 100)
    return numeric


def missing_values_to_zero(df):
    df = df.copy()
    missing_labels = {"na", "n/a", "nan", "none", "null", "<na>"}
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype(object)
            missing_mask = df[col].map(
                lambda value: isinstance(value, str) and value.strip().lower() in missing_labels
            )
            df.loc[missing_mask, col] = 0
    return df.fillna(0)


def first_existing_column(df, candidates):
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    return None


def to_excel_safe(value):
    if value is None:
        return None
    if value is pd.NA or value is pd.NaT:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    return value


def write_dataframe(ws, df, start_row=1):
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=start_row, column=c_idx, value=str(col_name))
    for r_offset, row in enumerate(df.itertuples(index=False, name=None), start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=start_row + r_offset, column=c_idx, value=to_excel_safe(value))


def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def header_label_value(worksheet, label):
    target = label.strip().lower()
    max_col = worksheet.max_column or 1
    for r in range(1, 9):
        for c in range(1, max_col + 1):
            v = worksheet.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().lower().startswith(target):
                for cc in range(c + 1, max_col + 1):
                    right = worksheet.cell(row=r, column=cc).value
                    if right not in (None, ""):
                        return right
                return None
    return None


def as_filename_token(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
    text = re.sub(r'[\\/:*?"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def as_filename_date(value):
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        parsed = pd.to_datetime(value, errors="coerce")
    else:
        text = str(value).strip()
        if not re.search(r"\d", text):
            return ""
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if pd.isna(parsed):
            parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        return as_filename_token(value)
    return parsed.strftime("%d-%m-%Y")


def apply_excel_totals(ws, header_row, data_start, data_end, sum_columns,
                       ratio_pairs=None, max_columns=None, sumif_rules=None):
    last_data = data_end - 1
    if last_data < data_start:
        return
    header_map = {cell.value: cell.column for cell in ws[header_row] if cell.value is not None}

    def letter(name):
        return get_column_letter(header_map[name])

    for col_name in sum_columns:
        if col_name not in header_map:
            continue
        col_letter = letter(col_name)
        ws.cell(row=data_end, column=header_map[col_name]).value = (
            f"=SUM({col_letter}{data_start}:{col_letter}{last_data})"
        )

    for col_name in (max_columns or []):
        if col_name not in header_map:
            continue
        col_letter = letter(col_name)
        ws.cell(row=data_end, column=header_map[col_name]).value = (
            f"=MAX({col_letter}{data_start}:{col_letter}{last_data})"
        )

    for result_col, num_col, den_col in (ratio_pairs or []):
        if result_col not in header_map or num_col not in header_map or den_col not in header_map:
            continue
        num_letter = letter(num_col)
        den_letter = letter(den_col)
        ws.cell(row=data_end, column=header_map[result_col]).value = (
            f'=IF({den_letter}{data_end}=0,"",{num_letter}{data_end}/{den_letter}{data_end})'
        )

    for rule in (sumif_rules or []):
        result_col = rule["result"]
        criteria_col = rule["criteria"]
        sum_col = rule["sum"]
        if result_col not in header_map or criteria_col not in header_map or sum_col not in header_map:
            continue
        criteria = rule["criteria_text"]
        c_letter = letter(criteria_col)
        s_letter = letter(sum_col)
        formula = (
            f'=SUMIF({c_letter}{data_start}:{c_letter}{last_data},"{criteria}",'
            f"{s_letter}{data_start}:{s_letter}{last_data})"
        )
        if rule.get("divide_by"):
            den_col = rule["divide_by"]
            if den_col not in header_map:
                continue
            d_letter = letter(den_col)
            den_formula = (
                f'SUMIF({c_letter}{data_start}:{c_letter}{last_data},"{criteria}",'
                f"{d_letter}{data_start}:{d_letter}{last_data})"
            )
            formula = f'=IF({den_formula}=0,"",{formula[1:]}/{den_formula})'
        ws.cell(row=data_end, column=header_map[result_col]).value = formula


def apply_delivery_formulas(ws, header_row, data_start, data_end):
    header_map = {cell.value: cell.column for cell in ws[header_row] if cell.value is not None}
    needed = ("Delivery", "Delivery Impressions", "Contracted Impressions")
    if any(name not in header_map for name in needed):
        return
    num = get_column_letter(header_map["Delivery Impressions"])
    den = get_column_letter(header_map["Contracted Impressions"])
    col = header_map["Delivery"]
    for r in range(data_start, data_end + 1):
        ws.cell(row=r, column=col).value = f'=IF({den}{r}=0,"",{num}{r}/{den}{r})'


def paint_black_columns(ws, header_row, data_start, data_end, column_names):
    header_map = {cell.value: cell.column for cell in ws[header_row] if cell.value is not None}
    for col_name in column_names:
        if col_name not in header_map:
            continue
        col = header_map[col_name]
        header_cell = ws.cell(row=header_row, column=col)
        header_cell.fill = header_fill
        header_cell.font = header_font
        header_cell.border = table_border
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(data_start, data_end + 1):
            cell = ws.cell(row=r, column=col)
            is_total = r == data_end
            cell.fill = total_fill if is_total else white_fill
            cell.font = total_font if is_total else black_font
            cell.border = table_border
            cell.alignment = Alignment(horizontal="center", vertical="center")


def build_report(file_bytes, currency_choice="1", log=print):
    def src():
        return io.BytesIO(file_bytes)

    currency_format = CURRENCY_FORMATS.get(currency_choice, CURRENCY_FORMATS["1"])
    log(f"Selected currency: {CURRENCY_NAMES.get(currency_choice, 'USD ($)')}")

    try:
        workbook_source = pd.ExcelFile(src())
    except Exception as exc:
        raise ValueError(f"Unable to open the uploaded Excel workbook: {exc}") from exc

    if "Performance Report" not in workbook_source.sheet_names:
        raise ValueError("The uploaded workbook must contain a sheet named 'Performance Report'.")
    if "Creative Level" not in workbook_source.sheet_names:
        raise ValueError("The uploaded workbook must contain a sheet named 'Creative Level'.")

    log("Reading Performance Report data starting from row 9 ...")
    df_raw = pd.read_excel(src(), sheet_name="Performance Report", header=8)
    df_raw = df_raw.dropna(axis=1, how="all")
    df_raw = clean_column_names(df_raw)

    require_columns(
        df_raw,
        [
            "Line Item Name", "Proposal ID", "Proposal Name", "Campaign Start",
            "Campaign End", "Line Item ID", "Contracted Impressions", "Billable Impressions",
        ],
        "Performance Report",
    )

    df_raw = df_raw[df_raw["Line Item Name"].notna()].copy()
    df_raw = df_raw[~df_raw["Line Item Name"].astype(str).str.strip().str.lower().eq("total")].copy()
    df_raw = df_raw.rename(columns={"Local Impressions": "Delivery Impressions", "Local Clicks": "Clicks"})
    df_raw = clean_column_names(df_raw)

    delivery_pct_candidates = [
        "% Delivered", "% Delivery", "Delivery %", "Delivered %",
        "%Delivered", "%Delivery", "Delivery Percent", "Delivered Percent",
        "Delivery", "Delivered",
    ]
    source_delivery_col = first_existing_column(df_raw, delivery_pct_candidates)
    if source_delivery_col:
        log(f"Found user Delivery column: '{source_delivery_col}'")
        df_raw["_User_Delivery_Pct"] = clean_percentage_series(df_raw[source_delivery_col])
    else:
        log("No user '% Delivered' style column found - Delivery will be blank")
        df_raw["_User_Delivery_Pct"] = np.nan

    optional_performance_columns = [
        "Delivery Impressions", "Clicks", "Third Party Impressions", "Third Party Clicks",
        "CTR Benchmark", "Video Starts", "Video Start Rate", "Video 25", "Video 50",
        "Video 75", "Video Completions", "Video Completion Rate", "Campaign Type",
        "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
        "Contracted Rate", "Total Active View Viewable Impressions",
        "Total Active View % Viewable Impressions",
    ]
    for col in optional_performance_columns:
        if col not in df_raw.columns:
            df_raw[col] = np.nan

    # Keep Third Party columns always (do not drop even if all zeros)
    for col in ["Third Party Impressions", "Third Party Clicks"]:
        if col not in df_raw.columns:
            df_raw[col] = np.nan
        else:
            df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce")

    line_item_lower = df_raw["Line Item Name"].fillna("").astype(str).str.strip().str.lower()
    has_vcpm = bool(line_item_lower.str.contains("vcpm", regex=False, na=False).any())
    log(f"VCPM detected: {'YES' if has_vcpm else 'NO'}")

    special_mask = line_item_lower.str.match(r"^\(\d+\)", na=False)
    df_main_raw = df_raw[~special_mask].copy()
    df_special_raw = df_raw[special_mask].copy()

    main_columns = [
        "Proposal ID", "Proposal Name", "Campaign Start", "Campaign End",
        "Line Item ID", "Line Item Name", "Contracted Impressions",
        "Delivery Impressions", "Third Party Impressions", "Billable Impressions",
        "Clicks", "Third Party Clicks", "CTR",
        "CTR Benchmark", "Delivery", "Video Starts", "Video Start Rate",
        "Video 25", "Video 50", "Video 75", "Video Completions", "Video Completion Rate",
    ]
    if has_vcpm:
        main_columns.extend([
            "Total Active View Viewable Impressions",
            "Total Active View % Viewable Impressions",
        ])
    main_columns.extend([
        "Contracted Rate", "Campaign Type", "Net Cost",
        "Campaign Budget Delivered", "Campaign Budget Remaining", "Campaign Length",
    ])

    special_columns = [
        "Proposal ID", "Proposal Name", "Campaign Start", "Campaign End",
        "Line Item ID", "Line Item Name", "Contracted Impressions",
        "Delivery Impressions", "Third Party Impressions", "Billable Impressions",
        "Clicks", "Third Party Clicks", "CTR",
        "Delivery",
        "Contracted Rate", "Campaign Type", "Net Cost",
        "Campaign Budget Delivered", "Campaign Budget Remaining", "Campaign Length",
    ]

    def prepare_main(df):
        df = format_date_columns(df.copy())
        for col in ["Contracted Impressions", "Delivery Impressions", "Clicks",
                    "Third Party Impressions", "Third Party Clicks"]:
            if col in df.columns:
                df[col] = clean_and_convert_numeric(df[col])
        other_numeric = [
            "Billable Impressions", "CTR Benchmark", "Video Starts", "Video 25",
            "Video 50", "Video 75", "Video Completions", "Contracted Rate",
            "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
        ]
        if has_vcpm:
            other_numeric.extend([
                "Total Active View Viewable Impressions",
                "Total Active View % Viewable Impressions",
            ])
        df = convert_numeric_columns(df, other_numeric)
        df["CTR"] = safe_ratio(df["Clicks"], df["Delivery Impressions"])
        df["Delivery"] = df["_User_Delivery_Pct"] if "_User_Delivery_Pct" in df.columns else np.nan
        df["Video Start Rate"] = safe_ratio(df["Video Starts"], df["Delivery Impressions"])
        df["Video Completion Rate"] = safe_ratio(df["Video Completions"], df["Video Starts"])
        start = pd.to_datetime(df["Campaign Start"], errors="coerce")
        end = pd.to_datetime(df["Campaign End"], errors="coerce")
        df["Campaign Length"] = (end - start).dt.days + 1
        if has_vcpm:
            current_line_items = df["Line Item Name"].fillna("").astype(str).str.strip().str.lower()
            current_vcpm_mask = current_line_items.str.contains("vcpm", regex=False, na=False)
            df.loc[~current_vcpm_mask, "Total Active View Viewable Impressions"] = np.nan
            df.loc[~current_vcpm_mask, "Total Active View % Viewable Impressions"] = np.nan
        campaign_type_lower = df["Campaign Type"].fillna("").astype(str).str.strip().str.lower()
        cpd_mask = campaign_type_lower.eq("cpd")
        cpd_length = pd.to_numeric(df.loc[cpd_mask, "Campaign Length"], errors="coerce")
        cpd_cost = pd.to_numeric(df.loc[cpd_mask, "Net Cost"], errors="coerce")
        df.loc[cpd_mask, "Contracted Rate"] = np.where(
            cpd_length.notna() & cpd_length.ne(0), cpd_cost / cpd_length, np.nan
        )
        for col in main_columns:
            if col not in df.columns:
                df[col] = np.nan
        return df[main_columns].copy()

    def prepare_special(df):
        if df.empty:
            return pd.DataFrame(columns=special_columns)
        df = format_date_columns(df.copy())
        for col in ["Contracted Impressions", "Delivery Impressions", "Clicks",
                    "Third Party Impressions", "Third Party Clicks"]:
            if col in df.columns:
                df[col] = clean_and_convert_numeric(df[col])
        df = convert_numeric_columns(df, [
            "Billable Impressions", "Contracted Rate", "Net Cost",
            "Campaign Budget Delivered", "Campaign Budget Remaining",
        ])
        df["CTR"] = safe_ratio(df["Clicks"], df["Delivery Impressions"])
        df["Delivery"] = df["_User_Delivery_Pct"] if "_User_Delivery_Pct" in df.columns else np.nan
        start = pd.to_datetime(df["Campaign Start"], errors="coerce")
        end = pd.to_datetime(df["Campaign End"], errors="coerce")
        df["Campaign Length"] = (end - start).dt.days + 1
        campaign_type_lower = df["Campaign Type"].fillna("").astype(str).str.strip().str.lower()
        cpd_mask = campaign_type_lower.eq("cpd")
        cpd_length = pd.to_numeric(df.loc[cpd_mask, "Campaign Length"], errors="coerce")
        cpd_cost = pd.to_numeric(df.loc[cpd_mask, "Net Cost"], errors="coerce")
        df.loc[cpd_mask, "Contracted Rate"] = np.where(
            cpd_length.notna() & cpd_length.ne(0), cpd_cost / cpd_length, np.nan
        )
        for col in special_columns:
            if col not in df.columns:
                df[col] = np.nan
        return df[special_columns].copy()

    df_main = prepare_main(df_main_raw)
    df_special = prepare_special(df_special_raw)

    def create_main_total(df):
        total = {col: np.nan for col in main_columns}
        total["Line Item Name"] = "Total"
        for col in [
            "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
            "Clicks", "Third Party Impressions", "Third Party Clicks",
            "Video Starts", "Video 25", "Video 50", "Video 75",
            "Video Completions", "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
        ]:
            total[col] = pd.to_numeric(df[col], errors="coerce").sum()
        total["CTR"] = safe_ratio(total["Clicks"], total["Delivery Impressions"])
        total["Delivery"] = np.nan
        total["Video Start Rate"] = safe_ratio(total["Video Starts"], total["Delivery Impressions"])
        total["Video Completion Rate"] = safe_ratio(total["Video Completions"], total["Video Starts"])
        if has_vcpm:
            line_items = df["Line Item Name"].fillna("").astype(str).str.strip().str.lower()
            vcpm_rows = df[line_items.str.contains("vcpm", regex=False, na=False)]
            vcpm_viewable = pd.to_numeric(vcpm_rows["Total Active View Viewable Impressions"], errors="coerce").sum()
            vcpm_delivery = pd.to_numeric(vcpm_rows["Delivery Impressions"], errors="coerce").sum()
            total["Total Active View Viewable Impressions"] = vcpm_viewable
            total["Total Active View % Viewable Impressions"] = safe_ratio(vcpm_viewable, vcpm_delivery)
        total["Campaign Length"] = pd.to_numeric(df["Campaign Length"], errors="coerce").max()
        return pd.DataFrame([total], columns=main_columns)

    def create_special_total(df):
        total = {col: np.nan for col in special_columns}
        total["Line Item Name"] = "Total"
        for col in [
            "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
            "Clicks", "Third Party Impressions", "Third Party Clicks",
            "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
        ]:
            total[col] = pd.to_numeric(df[col], errors="coerce").sum()
        total["CTR"] = safe_ratio(total["Clicks"], total["Delivery Impressions"])
        total["Delivery"] = np.nan
        total["Campaign Length"] = pd.to_numeric(df["Campaign Length"], errors="coerce").max()
        return pd.DataFrame([total], columns=special_columns)

    df_main_final = pd.concat([df_main, create_main_total(df_main)], ignore_index=True)
    if not df_special.empty:
        df_special_final = pd.concat([df_special, create_special_total(df_special)], ignore_index=True)
        has_special = True
    else:
        df_special_final = pd.DataFrame(columns=special_columns)
        has_special = False
    log(f"Special section included: {'YES' if has_special else 'NO'}")

    log("Reading Creative Level data starting from row 9 ...")
    df_creative_original = pd.read_excel(src(), sheet_name="Creative Level", header=8)
    df_creative = df_creative_original.copy().dropna(how="all").copy()
    df_creative = clean_column_names(df_creative)
    df_creative = df_creative.rename(columns={
        "Local Impressions": "Delivered Impressions",
        "Local Clicks": "Clicks",
        "Video Completion Rate": "Video Completions Rate",
    })
    creative_columns = [
        "Report Start", "Line Item Name", "Creative Name",
        "Delivered Impressions", "Third Party Impressions", "Clicks",
        "Third Party Clicks", "CTR", "Video Starts", "Video 25",
        "Video 50", "Video 75", "Video Completions",
        "Video Completions Rate", "Campaign Budget Delivered",
    ]
    for col in creative_columns:
        if col not in df_creative.columns:
            df_creative[col] = np.nan
    df_creative["Creative Name"] = (
        df_creative["Creative Name"].astype("string")
        .str.replace(r"\s*\(copy\)", "", case=False, regex=True)
        .str.strip()
    )
    report_start_lookup = df_raw[["Line Item Name", "Campaign Start"]].copy()
    report_start_lookup["Campaign Start"] = pd.to_datetime(report_start_lookup["Campaign Start"], errors="coerce")
    report_start_lookup = (
        report_start_lookup.dropna(subset=["Line Item Name"])
        .groupby("Line Item Name", as_index=False)["Campaign Start"].min()
        .rename(columns={"Campaign Start": "Report Start"})
    )
    report_start_lookup["Report Start"] = pd.to_datetime(report_start_lookup["Report Start"], errors="coerce").dt.date
    df_creative = df_creative.drop(columns=["Report Start"], errors="ignore").merge(
        report_start_lookup, on="Line Item Name", how="left"
    )
    if "Campaign Budget Delivered" not in df_creative.columns:
        df_creative["Campaign Budget Delivered"] = np.nan
    df_creative["Campaign Budget Delivered"] = pd.to_numeric(df_creative["Campaign Budget Delivered"], errors="coerce")
    for col in [
        "Delivered Impressions", "Third Party Impressions", "Clicks", "Third Party Clicks",
        "Video Starts", "Video 25", "Video 50", "Video 75", "Video Completions", "Campaign Budget Delivered",
    ]:
        if col in df_creative.columns:
            cleaned = (
                df_creative[col].astype("string").str.strip()
                .str.replace(",", "", regex=False)
                .str.replace(r"[$€£₹¥]", "", regex=True)
                .str.replace(r"^[CA]\$", "", regex=True)
                .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
            )
            df_creative[col] = pd.to_numeric(cleaned, errors="coerce").fillna(0)
    df_creative["Line Item Name"] = df_creative["Line Item Name"].astype("string").str.strip()
    df_creative["Creative Name"] = df_creative["Creative Name"].astype("string").str.strip()
    df_creative["Report Start"] = pd.to_datetime(df_creative["Report Start"], errors="coerce").dt.date
    df_creative["CTR"] = safe_ratio_zero(df_creative["Clicks"], df_creative["Delivered Impressions"])
    df_creative["Video Completions Rate"] = safe_ratio_zero(df_creative["Video Completions"], df_creative["Video Starts"])
    creative_name_text = df_creative["Creative Name"].astype("string").str.strip()
    invalid_creative_names = {"", "nan", "none", "null", "<na>", "na", "n/a"}
    valid_mask = creative_name_text.notna() & ~creative_name_text.str.lower().isin(invalid_creative_names)
    df_creative = df_creative.loc[valid_mask].copy()
    group_columns = ["Report Start", "Line Item Name", "Creative Name"]
    sum_metric_columns = [
        "Delivered Impressions", "Third Party Impressions", "Clicks", "Third Party Clicks",
        "Video Starts", "Video 25", "Video 50", "Video 75", "Video Completions", "Campaign Budget Delivered",
    ]
    df_creative = df_creative.groupby(group_columns, as_index=False, dropna=False, sort=False)[sum_metric_columns].sum()
    df_creative["CTR"] = safe_ratio_zero(df_creative["Clicks"], df_creative["Delivered Impressions"])
    df_creative["Video Completions Rate"] = safe_ratio_zero(df_creative["Video Completions"], df_creative["Video Starts"])
    df_creative = df_creative[creative_columns].copy()

    def create_creative_total(df):
        total = {col: 0 for col in creative_columns}
        total["Report Start"] = ""
        total["Line Item Name"] = ""
        total["Creative Name"] = "Total"
        for col in sum_metric_columns:
            total[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
        total["CTR"] = safe_ratio_zero(total["Clicks"], total["Delivered Impressions"])
        total["Video Completions Rate"] = safe_ratio_zero(total["Video Completions"], total["Video Starts"])
        return pd.DataFrame([total], columns=creative_columns)

    df_creative_final = pd.concat([df_creative, create_creative_total(df_creative)], ignore_index=True)
    df_creative_final = missing_values_to_zero(df_creative_final)
    total_row_mask = df_creative_final["Creative Name"].astype(str).str.strip().str.lower().eq("total")
    df_creative_final.loc[total_row_mask, ["Report Start", "Line Item Name"]] = ""
    df_creative_final.loc[total_row_mask, "CTR"] = safe_ratio_zero(
        df_creative_final.loc[total_row_mask, "Clicks"],
        df_creative_final.loc[total_row_mask, "Delivered Impressions"],
    )
    df_creative_final.loc[total_row_mask, "Video Completions Rate"] = safe_ratio_zero(
        df_creative_final.loc[total_row_mask, "Video Completions"],
        df_creative_final.loc[total_row_mask, "Video Starts"],
    )
    df_creative_final = df_creative_final[creative_columns].copy()
    log("Creative Level breakdown aggregation completed.")

    df_device = pd.DataFrame({"Device": [""] * 5, "Impressions": [0] * 5, "Clicks": [0] * 5, "CTR": [0] * 5})
    df_country = pd.DataFrame({"Country": [""] * 5, "Impressions": [0] * 5, "Clicks": [0] * 5, "CTR": [0] * 5})

    src_wb = load_workbook(src())
    src_ws = src_wb["Performance Report"]
    src_wb_values = load_workbook(src(), data_only=True)
    src_ws_values = src_wb_values["Performance Report"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Performance Report"
    ws.sheet_view.showGridLines = False
    HEADER_ROWS = 8

    for row in src_ws.iter_rows(min_row=1, max_row=HEADER_ROWS, max_col=src_ws.max_column):
        for cell in row:
            calc_value = src_ws_values.cell(row=cell.row, column=cell.column).value
            raw_value = cell.value
            if calc_value is not None:
                value = calc_value
            elif isinstance(raw_value, str) and raw_value.startswith("="):
                value = None
            else:
                value = raw_value
            new_cell = ws.cell(row=cell.row, column=cell.column, value=to_excel_safe(value))
            copy_cell_style(cell, new_cell)

    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.min_row <= HEADER_ROWS:
            ws.merge_cells(str(merged_range))
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[col_letter].width = dim.width
    for r in range(1, HEADER_ROWS + 1):
        if src_ws.row_dimensions[r].height:
            ws.row_dimensions[r].height = src_ws.row_dimensions[r].height

    tmp_dir = tempfile.gettempdir()
    for img in src_ws._images:
        try:
            img_path = os.path.join(tmp_dir, f"header_img_{id(img)}.png")
            with open(img_path, "wb") as f:
                f.write(img._data())
            new_img = XLImage(img_path)
            new_img.anchor = img.anchor
            ws.add_image(new_img)
        except Exception as e:
            log(f"Could not copy one image: {e}")

    def fill_blank_header_value(label_text, fallback_value):
        if fallback_value in (None, ""):
            return
        for r in range(1, HEADER_ROWS + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip().lower().startswith(label_text.lower()):
                    has_value_to_right = any(
                        ws.cell(row=r, column=cc).value not in (None, "")
                        for cc in range(c + 1, ws.max_column + 1)
                    )
                    if not has_value_to_right:
                        ws.cell(row=r, column=c + 1, value=fallback_value)
                    return

    proposal_names = df_main_final["Proposal Name"].dropna().astype(str).str.strip()
    proposal_names = proposal_names[proposal_names.ne("") & proposal_names.ne("Total")]
    proposal_ids = df_main_final["Proposal ID"].dropna().astype(str).str.strip()
    proposal_ids = proposal_ids[proposal_ids.ne("") & proposal_ids.ne("Total")]
    fallback_proposal_name = proposal_names.iloc[0] if not proposal_names.empty else None
    fallback_proposal_id = proposal_ids.iloc[0] if not proposal_ids.empty else None
    fill_blank_header_value("Report Proposal Name", fallback_proposal_name)
    fill_blank_header_value("Report Proposal ID", fallback_proposal_id)
    campaign_ends = pd.to_datetime(df_main["Campaign End"], errors="coerce").dropna()
    fallback_end = campaign_ends.max() if not campaign_ends.empty else None
    fill_blank_header_value("Report End Date", fallback_end.date() if fallback_end is not None else None)
    report_proposal_name = header_label_value(ws, "Report Proposal Name") or fallback_proposal_name
    report_proposal_id = header_label_value(ws, "Report Proposal ID") or fallback_proposal_id
    report_end_date = (
        header_label_value(ws, "Report End Date")
        or header_label_value(ws, "Report End")
        or (fallback_end.date() if fallback_end is not None else None)
    )
    download_name = (
        f"{as_filename_token(report_proposal_name) or 'Report'}_"
        f"{as_filename_token(report_proposal_id) or 'ID'}_"
        f"{as_filename_date(report_end_date) or 'EndDate'}.xlsx"
    )

    start_row = HEADER_ROWS + 1
    write_dataframe(ws, df_main_final, start_row=start_row)
    main_header_row = start_row
    main_data_start = start_row + 1
    main_data_end = start_row + len(df_main_final)
    main_sumif_rules = []
    if has_vcpm:
        main_sumif_rules = [
            {
                "result": "Total Active View Viewable Impressions",
                "criteria": "Line Item Name",
                "criteria_text": "*vcpm*",
                "sum": "Total Active View Viewable Impressions",
            },
            {
                "result": "Total Active View % Viewable Impressions",
                "criteria": "Line Item Name",
                "criteria_text": "*vcpm*",
                "sum": "Total Active View Viewable Impressions",
                "divide_by": "Delivery Impressions",
            },
        ]
    apply_excel_totals(
        ws, header_row=main_header_row, data_start=main_data_start, data_end=main_data_end,
        sum_columns=[
            "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
            "Clicks", "Third Party Impressions", "Third Party Clicks",
            "Video Starts", "Video 25", "Video 50", "Video 75",
            "Video Completions", "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
        ],
        ratio_pairs=[
            ("CTR", "Clicks", "Delivery Impressions"),
            ("Video Start Rate", "Video Starts", "Delivery Impressions"),
            ("Video Completion Rate", "Video Completions", "Video Starts"),
        ],
        max_columns=["Campaign Length"],
        sumif_rules=main_sumif_rules,
    )
    if has_special:
        special_start = main_data_end + 4
        write_dataframe(ws, df_special_final, start_row=special_start)
        special_header_row = special_start
        special_data_start = special_start + 1
        special_data_end = special_start + len(df_special_final)
        apply_excel_totals(
            ws, header_row=special_header_row, data_start=special_data_start, data_end=special_data_end,
            sum_columns=[
                "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
                "Clicks", "Third Party Impressions", "Third Party Clicks",
                "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
            ],
            ratio_pairs=[
                ("CTR", "Clicks", "Delivery Impressions"),
            ],
            max_columns=["Campaign Length"],
        )

    ws_cl = wb.create_sheet("Creative Level")
    write_dataframe(ws_cl, df_creative_original.dropna(how="all"), start_row=1)
    ws_bd = wb.create_sheet("Creative Level breakdown")
    write_dataframe(ws_bd, df_creative_final, start_row=1)
    if len(df_creative_final) >= 1:
        apply_excel_totals(
            ws_bd, header_row=1, data_start=2, data_end=1 + len(df_creative_final),
            sum_columns=sum_metric_columns,
            ratio_pairs=[
                ("CTR", "Clicks", "Delivered Impressions"),
                ("Video Completions Rate", "Video Completions", "Video Starts"),
            ],
        )
    ws_dev = wb.create_sheet("Device")
    write_dataframe(ws_dev, df_device, start_row=1)
    ws_cty = wb.create_sheet("Country")
    write_dataframe(ws_cty, df_country, start_row=1)
    src_wb.close()
    src_wb_values.close()

    def format_table(ws, header_row, data_start, data_end,
                     percentage_columns=None, integer_columns=None,
                     currency_columns=None, date_columns=None,
                     center_columns=None, force_total_fill_columns=None,
                     wide_columns=None, has_total_row=False):
        percentage_columns = percentage_columns or []
        integer_columns = integer_columns or []
        currency_columns = currency_columns or []
        date_columns = date_columns or []
        center_columns = center_columns or []
        wide_columns = wide_columns or {}
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=header_row, column=col_num)
            if cell.value is not None:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = table_border
        normal_end = data_end - 1 if has_total_row else data_end
        for row_num in range(data_start, normal_end + 1):
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_num, column=col_num)
                if cell.value is not None:
                    cell.fill = white_fill
                    cell.font = black_font
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = table_border
            # Set data row height to 25
            ws.row_dimensions[row_num].height = 25
        if has_total_row:
            header_map_temp = {cell.value: cell.column for cell in ws[header_row] if cell.value is not None}
            no_total_box = {"Proposal ID", "Proposal Name", "Campaign Start", "Campaign End", "Line Item ID"}
            empty_border = Border()
            empty_fill = PatternFill(fill_type=None)
            for col_name, col_num in header_map_temp.items():
                cell = ws.cell(row=data_end, column=col_num)
                if col_name in no_total_box:
                    cell.value = None
                    cell.fill = empty_fill
                    cell.border = empty_border
                    cell.font = black_font
                    continue
                cell.fill = total_fill
                cell.font = total_font
                cell.border = table_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            # Set total row height to 25 as well
            ws.row_dimensions[data_end].height = 25
        header_map = {cell.value: cell.column for cell in ws[header_row] if cell.value is not None}
        for col_name in integer_columns:
            if col_name in header_map:
                for r in range(data_start, data_end + 1):
                    ws.cell(row=r, column=header_map[col_name]).number_format = "#,##0"
        for col_name in percentage_columns:
            if col_name in header_map:
                for r in range(data_start, data_end + 1):
                    ws.cell(row=r, column=header_map[col_name]).number_format = "0.00%"
        for col_name in currency_columns:
            if col_name in header_map:
                for r in range(data_start, data_end + 1):
                    ws.cell(row=r, column=header_map[col_name]).number_format = currency_format
        for col_name in date_columns:
            if col_name in header_map:
                for r in range(data_start, data_end + 1):
                    ws.cell(row=r, column=header_map[col_name]).number_format = "m/d/yyyy"
        for col_name in center_columns:
            if col_name in header_map:
                for r in range(data_start, data_end + 1):
                    ws.cell(row=r, column=header_map[col_name]).alignment = Alignment(horizontal="center", vertical="center")
        for col_num in range(1, ws.max_column + 1):
            letter = get_column_letter(col_num)
            max_len = 0
            for r in range(header_row, data_end + 1):
                val = ws.cell(row=r, column=col_num).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
        for col_name, width in wide_columns.items():
            if col_name in header_map:
                ws.column_dimensions[get_column_letter(header_map[col_name])].width = width
        ws.row_dimensions[header_row].height = 30

    main_percentage = ["CTR", "CTR Benchmark", "Delivery", "Video Start Rate", "Video Completion Rate"]
    main_integer = [
        "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
        "Clicks", "Third Party Impressions", "Third Party Clicks",
        "Video Starts", "Video 25", "Video 50", "Video 75", "Video Completions", "Campaign Length",
    ]
    main_center = main_integer + main_percentage + [
        "Contracted Rate", "Campaign Type", "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining",
    ]
    if has_vcpm:
        main_percentage.append("Total Active View % Viewable Impressions")
        main_integer.append("Total Active View Viewable Impressions")
        main_center.extend([
            "Total Active View Viewable Impressions",
            "Total Active View % Viewable Impressions",
        ])
    format_table(
        ws=ws, header_row=main_header_row, data_start=main_data_start, data_end=main_data_end,
        percentage_columns=main_percentage, integer_columns=main_integer,
        currency_columns=["Contracted Rate", "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining"],
        date_columns=["Campaign Start", "Campaign End"], center_columns=main_center,
        wide_columns={"Proposal Name": 45, "Line Item Name": 55}, has_total_row=True,
    )
    if has_vcpm:
        paint_black_columns(
            ws, header_row=main_header_row, data_start=main_data_start, data_end=main_data_end,
            column_names=["Total Active View Viewable Impressions", "Total Active View % Viewable Impressions"],
        )
    header_map = {cell.value: cell.column for cell in ws[main_header_row] if cell.value is not None}
    for col_name in ["CTR Benchmark", "Video Completion Rate"]:
        if col_name in header_map:
            for r in range(main_data_start, main_data_end + 1):
                ws.cell(row=r, column=header_map[col_name]).border = table_border
    for col_name in ["Delivery", "CTR"]:
        if col_name in header_map:
            for r in range(main_data_start, main_data_end + 1):
                ws.cell(row=r, column=header_map[col_name]).number_format = "0.00%"
    if has_special:
        format_table(
            ws=ws, header_row=special_header_row, data_start=special_data_start, data_end=special_data_end,
            percentage_columns=["Delivery", "CTR"],
            integer_columns=[
                "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
                "Clicks", "Third Party Impressions", "Third Party Clicks", "Campaign Length",
            ],
            currency_columns=["Contracted Rate", "Net Cost", "Campaign Budget Delivered", "Campaign Budget Remaining"],
            date_columns=["Campaign Start", "Campaign End"],
            center_columns=[
                "Contracted Impressions", "Delivery Impressions", "Billable Impressions",
                "Clicks", "Third Party Impressions", "Third Party Clicks", "CTR", "Delivery",
                "Contracted Rate", "Campaign Type", "Net Cost", "Campaign Budget Delivered",
                "Campaign Budget Remaining", "Campaign Length",
            ],
            wide_columns={"Proposal Name": 45, "Line Item Name": 55}, has_total_row=True,
        )
        special_header_map = {cell.value: cell.column for cell in ws[special_header_row] if cell.value is not None}
        for col_name in ["Delivery", "CTR"]:
            if col_name in special_header_map:
                for r in range(special_data_start, special_data_end + 1):
                    ws.cell(row=r, column=special_header_map[col_name]).number_format = "0.00%"

    # Creative Level (raw) formatting + row height 25
    ws_cl.sheet_view.showGridLines = False
    for col in range(1, ws_cl.max_column + 1):
        cell = ws_cl.cell(row=1, column=col)
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = table_border
    for r in range(2, ws_cl.max_row + 1):
        for c in range(1, ws_cl.max_column + 1):
            cell = ws_cl.cell(row=r, column=c)
            if cell.value is not None:
                cell.fill = white_fill
                cell.font = black_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = table_border
        ws_cl.row_dimensions[r].height = 25
    for c in range(1, ws_cl.max_column + 1):
        letter = get_column_letter(c)
        max_len = 0
        for r in range(1, ws_cl.max_row + 1):
            val = ws_cl.cell(row=r, column=c).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_cl.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    ws_cl.row_dimensions[1].height = 30

    # Creative Level breakdown formatting + row height 25
    ws_bd.sheet_view.showGridLines = False
    breakdown_header_map = {cell.value: cell.column for cell in ws_bd[1] if cell.value is not None}
    for col in range(1, ws_bd.max_column + 1):
        cell = ws_bd.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = table_border
    for r in range(2, ws_bd.max_row + 1):
        is_total = str(ws_bd.cell(row=r, column=3).value).strip().lower() == "total"
        for c in range(1, ws_bd.max_column + 1):
            cell = ws_bd.cell(row=r, column=c)
            cell.border = table_border
            cell.fill = total_fill if is_total else white_fill
            cell.font = total_font if is_total else black_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if is_total and c in [breakdown_header_map.get("Report Start"), breakdown_header_map.get("Line Item Name")]:
                cell.value = ""
        ws_bd.row_dimensions[r].height = 25
    for col_name in ["CTR", "Video Completions Rate"]:
        if col_name in breakdown_header_map:
            for r in range(2, ws_bd.max_row + 1):
                ws_bd.cell(row=r, column=breakdown_header_map[col_name]).number_format = "0.00%"
    for col_name in [
        "Delivered Impressions", "Third Party Impressions", "Clicks", "Third Party Clicks",
        "Video Starts", "Video 25", "Video 50", "Video 75", "Video Completions",
    ]:
        if col_name in breakdown_header_map:
            for r in range(2, ws_bd.max_row + 1):
                ws_bd.cell(row=r, column=breakdown_header_map[col_name]).number_format = "#,##0"
    if "Campaign Budget Delivered" in breakdown_header_map:
        for r in range(2, ws_bd.max_row + 1):
            ws_bd.cell(row=r, column=breakdown_header_map["Campaign Budget Delivered"]).number_format = currency_format
    if "Report Start" in breakdown_header_map:
        for r in range(2, ws_bd.max_row + 1):
            cell = ws_bd.cell(row=r, column=breakdown_header_map["Report Start"])
            if cell.value not in (None, ""):
                cell.number_format = "m/d/yyyy"
    for col_name in [
        "Delivered Impressions", "Third Party Impressions", "Clicks", "Third Party Clicks",
        "Video Starts", "Video 25", "Video 50", "Video 75", "Video Completions",
        "CTR", "Video Completions Rate", "Campaign Budget Delivered",
    ]:
        if col_name in breakdown_header_map:
            for r in range(2, ws_bd.max_row + 1):
                ws_bd.cell(row=r, column=breakdown_header_map[col_name]).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, ws_bd.max_column + 1):
        letter = get_column_letter(c)
        max_len = 0
        for r in range(1, ws_bd.max_row + 1):
            val = ws_bd.cell(row=r, column=c).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_bd.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    for col_name in ["Line Item Name", "Creative Name"]:
        if col_name in breakdown_header_map:
            ws_bd.column_dimensions[get_column_letter(breakdown_header_map[col_name])].width = 45
    ws_bd.row_dimensions[1].height = 30

    # Device & Country sheets – row height 25
    for sheet in [ws_dev, ws_cty]:
        sheet.sheet_view.showGridLines = False
        for c in range(1, 5):
            cell = sheet.cell(row=1, column=c)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = table_border
        for r in range(2, 7):
            for c in range(1, 5):
                cell = sheet.cell(row=r, column=c)
                cell.border = table_border
                cell.fill = white_fill
                cell.font = black_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            sheet.cell(row=r, column=2).number_format = "#,##0"
            sheet.cell(row=r, column=3).number_format = "#,##0"
            ctr_cell = sheet.cell(row=r, column=4)
            ctr_cell.value = f"=IF(B{r}=0,0,C{r}/B{r})"
            ctr_cell.number_format = "0.00%"
            sheet.row_dimensions[r].height = 25
        sheet.column_dimensions["A"].width = 18
        sheet.column_dimensions["B"].width = 15
        sheet.column_dimensions["C"].width = 12
        sheet.column_dimensions["D"].width = 12
        sheet.row_dimensions[1].height = 30

    buffer = io.BytesIO()
    wb.save(buffer)
    output_bytes = buffer.getvalue()
    check_wb = load_workbook(io.BytesIO(output_bytes), read_only=True)
    expected = ["Performance Report", "Creative Level", "Creative Level breakdown", "Device", "Country"]
    missing = [s for s in expected if s not in check_wb.sheetnames]
    check_wb.close()
    if missing:
        raise RuntimeError("Missing sheets: " + ", ".join(missing))
    log("FINAL REPORT CREATED SUCCESSFULLY")
    log("Header block (rows 1-8 + logo) copied from source file")
    log("Report Proposal Name / ID filled only if source is blank")
    log("Delivery column = user-uploaded values only")
    log(f"Download name: {download_name}")
    return output_bytes, download_name, df_creative_final


st.set_page_config(page_title="Performance Report", layout="wide", initial_sidebar_state="collapsed")
st.markdown(
    """
    <link href="https://fonts.googleapis.com/css2?family=Libre+Franklin:wght@500;700;800&display=swap" rel="stylesheet">
    <style>
    html, body, [class*="css"], .stApp, .stMarkdown, label, p, h1, h2, h3, span, div {
        font-family: "Libre Franklin", Helvetica, Arial, sans-serif;
    }
    .stApp { background: #FFE500; }
    .block-container { padding-top: 1.2rem; max-width: 1180px; }
    header[data-testid="stHeader"] { background: transparent; }
    #MainMenu, footer { visibility: hidden; }
    .topbar {
        display: flex; align-items: baseline; justify-content: space-between;
        border-bottom: 2px solid #000; padding-bottom: 14px; margin-bottom: 28px;
    }
    .wordmark { color: #000; font-weight: 800; font-size: 1.55rem; line-height: 1; }
    .wordmark span { display: block; font-size: 0.72rem; letter-spacing: 0.22em; margin-top: 4px; }
    .nav { color: #000; font-weight: 800; font-size: 0.78rem; letter-spacing: 0.14em; }
    .panel h2 { color: #fff !important; font-weight: 800; font-size: 1.7rem; margin: 0 0 8px 0; }
    .panel p { color: #fff !important; font-weight: 500; margin: 0 0 18px 0; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child {
        background: #000; padding: 28px 22px 18px 22px;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child label p {
        color: #FFE500 !important;
    }
    div[data-testid="stFileUploader"] section { background: #111; border: 1px solid #FFE500; }
    div[data-testid="stFileUploader"] label,
    div[data-testid="stFileUploader"] span,
    div[data-testid="stFileUploader"] small { color: #FFE500 !important; font-weight: 700 !important; }
    div[data-baseweb="select"] > div { background: #000; color: #FFE500; border-color: #000; font-weight: 700; }
    .stButton > button, .stDownloadButton > button {
        background: #FFE500; color: #000; font-weight: 800; border: 2px solid #FFE500;
        border-radius: 999px; letter-spacing: 0.08em; width: 100%;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #000; color: #FFE500; border-color: #FFE500;
    }
    .log-title { color: #000; font-weight: 800; font-size: 1.7rem; margin: 0; }
    .log-rule { border: 0; border-top: 1px solid #000; margin: 8px 0 16px 0; }
    div[data-testid="stTextArea"] textarea {
        background: #111 !important; color: #FFE500 !important; font-weight: 600 !important;
        border: none !important; border-radius: 0 !important;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div class="topbar">
        <div class="wordmark">PERFORMANCE REPORT<span>ADVERTISING</span></div>
        <div class="nav">BUILD &nbsp;&nbsp; DOWNLOAD &nbsp;&nbsp; LOG</div>
    </div>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([0.92, 1.25], gap="large")
with left:
    st.markdown(
        """
        <div class="panel">
            <h2>BUILD THE REPORT</h2>
            <p>Upload the source workbook. Choose the currency, then build.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    currency_choice = st.selectbox(
        "CURRENCY",
        options=list(CURRENCY_NAMES.keys()),
        format_func=lambda k: CURRENCY_NAMES[k],
        index=0,
    )
    uploaded_file = st.file_uploader("SOURCE WORKBOOK", type=["xlsx", "xlsm"])
    generate = st.button("BUILD REPORT")

with right:
    st.markdown('<div class="log-title">RUN LOG</div><hr class="log-rule">', unsafe_allow_html=True)

if "result" not in st.session_state:
    st.session_state["result"] = None
if "log_text" not in st.session_state:
    st.session_state["log_text"] = "Waiting for a workbook."

if generate:
    if uploaded_file is None:
        st.session_state["result"] = None
        st.session_state["log_text"] = "Upload an Excel file before building the report."
    else:
        logs = []
        try:
            with st.spinner("Building report..."):
                out_bytes, out_name, _preview = build_report(
                    uploaded_file.getvalue(), currency_choice, log=logs.append
                )
            st.session_state["result"] = {"bytes": out_bytes, "name": out_name}
            st.session_state["log_text"] = "\n".join(str(line) for line in logs) or "No output."
        except Exception as exc:
            st.session_state["result"] = None
            st.session_state["log_text"] = "\n".join(str(line) for line in logs)
            if st.session_state["log_text"]:
                st.session_state["log_text"] += "\n\n"
            st.session_state["log_text"] += f"ERROR: {exc}"

with right:
    st.text_area("Final print output", value=st.session_state["log_text"], height=340, label_visibility="collapsed")
    result = st.session_state.get("result")
    if result:
        st.download_button(
            label="DOWNLOAD REPORT",
            data=result["bytes"],
            file_name=result["name"],
            mime="application/vnd.openxmlformats.spreadsheetml.sheet",
        )
