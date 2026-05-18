import io
import re
from copy import copy
from datetime import date
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator


APP_TITLE = "Purchase Order to Working File Dashboard"
BRAND = "@BAJRABHANU"
WORKING_SHEET_NAME = "Working"
OUTPUT_FILENAME = "Filled_Working_File.xlsx"


st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
        .main-title {
            padding: 18px 22px;
            border-radius: 18px;
            background: linear-gradient(135deg, #0f172a, #1d4ed8, #f97316);
            color: white;
            box-shadow: 0 14px 35px rgba(15, 23, 42, 0.22);
            margin-bottom: 18px;
        }
        .main-title h1 { margin: 0; font-size: 30px; }
        .main-title p { margin: 6px 0 0 0; font-size: 15px; opacity: 0.95; }
        .metric-card {
            padding: 16px;
            border-radius: 16px;
            border: 1px solid #e5e7eb;
            background: #ffffff;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
        }
        .metric-label { color:#64748b; font-size:13px; font-weight:700; }
        .metric-value { color:#0f172a; font-size:28px; font-weight:800; margin-top:4px; }
        .note-box {
            padding: 14px 16px;
            border-radius: 14px;
            background: #fff7ed;
            border: 1px solid #fed7aa;
            color:#7c2d12;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class='main-title'>
        <h1>📦 Purchase Order Automation Dashboard</h1>
        <p>Upload your formulated working file + multiple Flipkart PO files, then download the filled working Excel file. {BRAND}</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def clean_header(value) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean_text(value).upper())


def to_number(value):
    """Convert quantities or INR/% text to number where possible."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return value

    text = clean_text(value)
    if not text:
        return None

    text = text.replace(",", "")
    text = re.sub(r"\bINR\b|%", "", text, flags=re.IGNORECASE).strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    number = float(match.group(0))
    return int(number) if number.is_integer() else number


def read_po_as_dataframe(uploaded_file) -> pd.DataFrame:
    """Read .xls/.xlsx PO file without assuming fixed headers."""
    file_bytes = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"

    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=0,
        header=None,
        dtype=object,
        engine=engine,
    )


def find_label_value(df: pd.DataFrame, label: str) -> str:
    """Find a label such as PO# and return the next non-empty cell on the same row."""
    target = clean_header(label)

    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if clean_header(df.iat[r, c]) == target:
                for cc in range(c + 1, df.shape[1]):
                    value = clean_text(df.iat[r, cc])
                    if value:
                        return value

    return ""


def find_order_header_row(df: pd.DataFrame) -> Optional[int]:
    for r in range(df.shape[0]):
        headers = [clean_header(x) for x in df.iloc[r].tolist()]
        if "FSNISBN13" in headers and "QUANTITY" in headers:
            return r

    return None


def parse_po_file(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """Extract PO#, FSN/ISBN13, Quantity and useful optional fields from one PO file."""
    errors: List[str] = []

    try:
        df = read_po_as_dataframe(uploaded_file)
    except Exception as exc:
        return pd.DataFrame(), [f"{uploaded_file.name}: unable to read file. {exc}"]

    po_no = find_label_value(df, "PO#")

    if not po_no:
        all_text = " ".join(clean_text(x) for x in df.to_numpy().flatten())
        match = re.search(
            r"PURCHASE\s+ORDER\s*#\s*([A-Z0-9-]+)",
            all_text,
            flags=re.IGNORECASE,
        )
        if match:
            po_no = match.group(1).strip()

    header_row = find_order_header_row(df)

    if header_row is None:
        return pd.DataFrame(), [f"{uploaded_file.name}: order details table not found."]

    header_map: Dict[str, int] = {}

    for c, value in enumerate(df.iloc[header_row].tolist()):
        header_map[clean_header(value)] = c

    fsn_col = header_map.get("FSNISBN13")
    qty_col = header_map.get("QUANTITY")
    uom_col = header_map.get("UOM")
    title_col = header_map.get("TITLE")
    price_col = header_map.get("SUPPLIERPRICE")
    taxable_col = header_map.get("TAXABLEVALUE")
    req_date_col = header_map.get("REQUIREDBYDATE")

    if fsn_col is None or qty_col is None:
        return pd.DataFrame(), [f"{uploaded_file.name}: FSN/ISBN13 or Quantity column missing."]

    rows: List[Dict[str, object]] = []

    for r in range(header_row + 1, df.shape[0]):
        fsn = clean_text(df.iat[r, fsn_col])
        qty = to_number(df.iat[r, qty_col])

        if not fsn:
            continue

        if "TOTAL" in fsn.upper() or "IMPORTANT" in fsn.upper():
            break

        if qty is None:
            continue

        rows.append(
            {
                "PO#": po_no,
                "FSN/ISBN13": fsn,
                "Quantity": qty,
                "UOM": clean_text(df.iat[r, uom_col]) if uom_col is not None else "",
                "Title": clean_text(df.iat[r, title_col]) if title_col is not None else "",
                "Supplier Price": to_number(df.iat[r, price_col]) if price_col is not None else None,
                "Taxable Value": to_number(df.iat[r, taxable_col]) if taxable_col is not None else None,
                "Required by Date": clean_text(df.iat[r, req_date_col]) if req_date_col is not None else "",
                "Source File": uploaded_file.name,
            }
        )

    if not rows:
        errors.append(f"{uploaded_file.name}: no item rows found.")

    if not po_no:
        errors.append(f"{uploaded_file.name}: PO number not found; PO# will remain blank.")

    return pd.DataFrame(rows), errors


def copy_row_format_and_formulas(ws, source_row: int, target_row: int, max_col: int) -> None:
    """Copy template row style/formulas, translating formulas to the target row."""
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        if src.number_format:
            dst.number_format = src.number_format

        if src.alignment:
            dst.alignment = copy(src.alignment)

        if src.border:
            dst.border = copy(src.border)

        if src.fill:
            dst.fill = copy(src.fill)

        if src.font:
            dst.font = copy(src.font)

        if src.protection:
            dst.protection = copy(src.protection)

        if isinstance(src.value, str) and src.value.startswith("="):
            try:
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            except Exception:
                dst.value = src.value
        else:
            dst.value = src.value

    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def fill_working_file(template_file, extracted_df: pd.DataFrame, po_date_value: Optional[date]) -> bytes:
    wb = load_workbook(template_file)

    if WORKING_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"'{WORKING_SHEET_NAME}' sheet not found in the working file.")

    ws = wb[WORKING_SHEET_NAME]
    max_col = ws.max_column
    required_rows = max(2, len(extracted_df) + 1)

    # Make sure enough template rows are available.
    for target_row in range(2, required_rows + 1):
        source_row = 2 if target_row == 2 else 3
        copy_row_format_and_formulas(ws, source_row, target_row, max_col)

    # Clear old manual input cells from columns A:C.
    for r in range(2, max(ws.max_row, required_rows) + 1):
        ws.cell(r, 1).value = None
        ws.cell(r, 2).value = None
        ws.cell(r, 3).value = None

    # Fill orange/manual input columns.
    for idx, row in extracted_df.reset_index(drop=True).iterrows():
        excel_row = idx + 2
        ws.cell(excel_row, 1).value = row.get("PO#", "")
        ws.cell(excel_row, 2).value = row.get("FSN/ISBN13", "")
        ws.cell(excel_row, 3).value = row.get("Quantity", "")

        # Optional: Column Q PO Date
        if po_date_value:
            ws.cell(excel_row, 17).value = po_date_value.strftime("%d.%m.%Y")

    # Add audit summary sheet.
    if "PO Import Summary" in wb.sheetnames:
        del wb["PO Import Summary"]

    summary = wb.create_sheet("PO Import Summary")
    summary_headers = list(extracted_df.columns)
    summary.append(summary_headers)

    for _, row in extracted_df.iterrows():
        summary.append([row.get(col, "") for col in summary_headers])

    for col_cells in summary.columns:
        max_len = max(len(clean_text(cell.value)) for cell in col_cells)
        summary.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 38)

    summary.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


with st.sidebar:
    st.header("Upload Files")

    working_file = st.file_uploader(
        "Upload formulated working file (.xlsx)",
        type=["xlsx"],
    )

    po_files = st.file_uploader(
        "Upload Purchase Order files (.xls / .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

    st.divider()

    use_po_date = st.checkbox(
        "Overwrite PO Date in Working sheet column Q",
        value=False,
    )

    po_date_value = st.date_input(
        "PO Date",
        value=date.today(),
        disabled=not use_po_date,
    )

    process_clicked = st.button(
        "🚀 Process PO Files",
        type="primary",
        use_container_width=True,
    )


st.markdown(
    """
    <div class='note-box'>
    This app fills the manual/orange input area in the <b>Working</b> sheet:
    <b>PO#</b>, <b>FSN/ISBN13</b>, and <b>Quantity</b>.
    Your existing formulas and master sheets are preserved.
    </div>
    """,
    unsafe_allow_html=True,
)


if process_clicked:
    if working_file is None:
        st.error("Please upload the formulated working file first.")
        st.stop()

    if not po_files:
        st.error("Please upload at least one purchase order file.")
        st.stop()

    all_frames: List[pd.DataFrame] = []
    all_errors: List[str] = []

    with st.spinner("Reading PO files and preparing output workbook..."):
        for po in po_files:
            po_df, errors = parse_po_file(po)

            if not po_df.empty:
                all_frames.append(po_df)

            all_errors.extend(errors)

        if not all_frames:
            st.error("No valid PO item rows could be extracted.")

            for err in all_errors:
                st.warning(err)

            st.stop()

        extracted = pd.concat(all_frames, ignore_index=True)

        extracted = extracted[
            [
                "PO#",
                "FSN/ISBN13",
                "Quantity",
                "UOM",
                "Title",
                "Supplier Price",
                "Taxable Value",
                "Required by Date",
                "Source File",
            ]
        ]

        output_bytes = fill_working_file(
            working_file,
            extracted,
            po_date_value if use_po_date else None,
        )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(
            f"""
            <div class='metric-card'>
                <div class='metric-label'>PO Files</div>
                <div class='metric-value'>{len(po_files)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class='metric-card'>
                <div class='metric-label'>Rows Extracted</div>
                <div class='metric-value'>{len(extracted)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class='metric-card'>
                <div class='metric-label'>Total Quantity</div>
                <div class='metric-value'>{extracted['Quantity'].sum():,.0f}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c4:
        st.markdown(
            f"""
            <div class='metric-card'>
                <div class='metric-label'>Unique PO#</div>
                <div class='metric-value'>{extracted['PO#'].nunique()}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("Preview of extracted PO data")
    st.dataframe(extracted, use_container_width=True, hide_index=True)

    if all_errors:
        with st.expander("Warnings / files needing review"):
            for err in all_errors:
                st.warning(err)

    st.download_button(
        label="⬇️ Download Filled Working Excel File",
        data=output_bytes,
        file_name=OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("Upload the working file and PO files from the left side, then click Process PO Files.")