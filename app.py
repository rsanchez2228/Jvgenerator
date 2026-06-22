import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# Set up the look and feel of the web page
st.set_page_config(page_title="JV Generation & Analytics Console", page_icon="📊", layout="wide")

st.title("📊 Treasury Management Web Console")
st.markdown("Generate fresh templates, run automated structural audits, and export balanced journals straight to SAP.")

# Initialize a session state key to handle file uploader resetting without page refreshes
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# ==========================================
# ENGINE 1: THE EXCEL GENERATOR
# ==========================================
def create_jv_template():
    """Builds a 300-row template in memory and returns it as a downloadable file object."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "JV Upload Form"
    ws1.views.sheetView[0].showGridLines = True

    # Styles definition
    HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    LABEL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    FONT_HEADER = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    FONT_LABEL = Font(name="Segoe UI", size=10, bold=True, color="1F4E78")
    FONT_BODY = Font(name="Segoe UI", size=11)
    
    THIN_BORDER = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    header_fields = [
        ("Document Date (YYYYMMDD):", "B1"),
        ("Posting Date (YYYYMMDD):", "B2"),
        ("Reference:", "E1"),
        ("Header Text:", "E2"),
        ("Control Total (Credits):", "H1"),
        ("Posting Period (MM):", "H2")
    ]
    
    for label_text, cell_coord in header_fields:
        cell = ws1[cell_coord]
        cell.value = label_text
        cell.font = FONT_LABEL
        cell.fill = LABEL_FILL
        cell.alignment = Alignment(horizontal="right", vertical="center")
        
    ws1["I1"] = "=SUMIF(I7:I306, \"CR\", H7:H306)"
    ws1["I1"].number_format = "$#,##0.00"
    ws1["I1"].font = Font(name="Segoe UI", size=11, bold=True, color="1F4E78")
    ws1["I1"].border = THIN_BORDER

    ws1["I2"] = ""
    ws1["I2"].number_format = "@"
    ws1["I2"].font = FONT_BODY
    ws1["I2"].border = THIN_BORDER
    ws1["I2"].alignment = Alignment(horizontal="center", vertical="center")
    
    for input_cell in ["C1", "C2", "F1", "F2"]:
        ws1[input_cell].number_format = "@"
        ws1[input_cell].font = FONT_BODY
        ws1[input_cell].border = THIN_BORDER

    headers = ["Line No", "Fund", "GL acct", "Business area", "Functional area", "Cost Center", 
               "WBS Element", "Amount", "DR/CR", "Line Description", "Validation Status"]

    for col_idx, header in enumerate(headers, 1):
        cell = ws1.cell(row=6, column=col_idx, value=header)
        cell.font = FONT_HEADER
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    ws1.row_dimensions[6].height = 28

    for row_idx in range(7, 307):
        ws1.row_dimensions[row_idx].height = 20
        ws1.cell(row=row_idx, column=1, value=(row_idx - 6))
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws1.cell(row=row_idx, column=col_idx)
            cell.font = FONT_BODY
            cell.border = THIN_BORDER
            if row_idx % 2 == 0:
                cell.fill = ZEBRA_FILL
            
            if col_idx == 8:
                cell.number_format = "#######0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in [1, 2, 3, 4, 5, 6, 7]:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "@"
            elif col_idx in [9, 11]:
                cell.alignment = Alignment(horizontal="center")

    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == "A":
            ws1.column_dimensions[col_letter].width = 10
            continue
        max_len = 0
        for cell in col:
            if cell.value is not None:
                val_str = str(cell.value)
                if val_str.startswith('='):
                    val_str = " $999,999.99 "
                max_len = max(max_len, len(val_str))
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 16)

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ==========================================
# ENGINE 2: THE AUDITOR & CONVERTER
# ==========================================
def process_and_validate_jv(uploaded_file):
    """Audits the uploaded spreadsheet stream and compiles data directly into SAP text lines."""
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        if "JV Upload Form" not in wb.sheetnames:
            return False, ["Workbook Structure Error: Missing sheet named 'JV Upload Form'."], 0, 0, None

        ws = wb["JV Upload Form"]
        
        doc_date_raw = str(ws["C1"].value or "").strip()
        post_date_raw = str(ws["C2"].value or "").strip()
        reference = str(ws["F1"].value or "T1919").strip()
        header_text = str(ws["F2"].value or f"Transmittal Accrual {datetime.today().strftime('%Y')}").strip()
        posting_period = str(ws["I2"].value or "").strip()
        
        def is_invalid_date(date_str):
            if len(date_str) != 8 or not date_str.isdigit():
                return True
            try:
                datetime.strptime(date_str, "%Y%m%d")
                return False
            except ValueError:
                return True

        # UPDATED: Comprehensive, structural field error messaging explicitly pinpointing cell coordinates
        if not posting_period:
            return False, ["Missing Required Field: 'Posting Period' value is completely empty. Location: Cell I2 (Row 2, Column 9)"], 0, 0, None

        if not doc_date_raw:
            return False, ["Missing Required Field: 'Document Date' value is completely empty. Location: Cell C1 (Row 1, Column 3)"], 0, 0, None
        if is_invalid_date(doc_date_raw):
            return False, [f"Invalid Value Error: 'Document Date' [{doc_date_raw}] must be exactly 8 digits format (YYYYMMDD). Location: Cell C1 (Row 1, Column 3)"], 0, 0, None

        if not post_date_raw:
            return False, ["Missing Required Field: 'Posting Date' value is completely empty. Location: Cell C2 (Row 2, Column 3)"], 0, 0, None
        if is_invalid_date(post_date_raw):
            return False, [f"Invalid Value Error: 'Posting Date' [{post_date_raw}] must be exactly 8 digits format (YYYYMMDD). Location: Cell C2 (Row 2, Column 3)"], 0, 0, None
            
        total_debit = 0.0
        total_credit = 0.0
        detailed_errors = []
        
        sap_lines = [f"H|EJ|{reference}|{doc_date_raw}|{post_date_raw}|{header_text}|{posting_period}|X\n"]

        # Scan rows 7 to 306
        for row in range(7, 307):
            fund = str(ws.cell(row=row, column=2).value or "").strip()
            gl_account = str(ws.cell(row=row, column=3).value or "").strip()
            bus_area = str(ws.cell(row=row, column=4).value or "").strip()
            func_area = str(ws.cell(row=row, column=5).value or "").strip()
            cost_center = str(ws.cell(row=row, column=6).value or "").strip()
            wbs_element = str(ws.cell(row=row, column=7).value or "").strip()
            amount_val = ws.cell(row=row, column=8).value
            drcr_flag = str(ws.cell(row=row, column=9).value or "").strip().upper()
            description = ws.cell(row=row, column=10).value or "Transmittal Entry"

            if not fund and not gl_account and not bus_area and not func_area and not cost_center and not amount_val and drcr_flag == "":
                continue

            if fund and (len(fund) != 6 or not fund.isdigit()):
                detailed_errors.append(f"Format Constraint Mismatch: 'Fund' must be exactly 6 digits. Location: Cell B{row} (Row {row}, Column 2) - Found: '{fund}'")
                continue
                
            if not gl_account:
                detailed_errors.append(f"Missing Value: Field 'GL acct' cannot be empty. Location: Cell C{row} (Row {row}, Column 3)")
                continue
            elif len(gl_account) != 6 or not gl_account.isdigit():
                detailed_errors.append(f"Format Constraint Mismatch: 'GL acct' must be exactly 6 digits. Location: Cell C{row} (Row {row}, Column 3) - Found: '{gl_account}'")
                continue

            if bus_area and (len(bus_area) != 4 or not bus_area.isdigit()):
                detailed_errors.append(f"Format Constraint Mismatch: 'Business area' must be exactly 4 digits. Location: Cell D{row} (Row {row}, Column 4) - Found: '{bus_area}'")
                continue

            if func_area and (len(func_area) != 7 or not func_area.isdigit()):
                detailed_errors.append(f"Format Constraint Mismatch: 'Functional area' must be exactly 7 digits. Location: Cell E{row} (Row {row}, Column 5) - Found: '{func_area}'")
                continue

            if cost_center and not re.match(r"^\d{8}-\d{6}$", cost_center):
                detailed_errors.append(f"Format Constraint Mismatch: 'Cost Center' must adhere to 8digits-6digits mask. Location: Cell F{row} (Row {row}, Column 6) - Found: '{cost_center}'")
                continue

            try:
                amount = float(amount_val or 0.0)
            except ValueError:
                detailed_errors.append(f"Data Type Error: 'Amount' column contains a non-numeric string value. Location: Cell H{row} (Row {row}, Column 8)")
                continue

            if amount < 0:
                detailed_errors.append(f"Value Constraint Error: 'Amount' value cannot be negative numbers. Location: Cell H{row} (Row {row}, Column 8)")
            elif drcr_flag not in ["DR", "CR"]:
                detailed_errors.append(f"Value Constraint Error: 'DR/CR' column choice must be explicitly written as 'DR' or 'CR'. Location: Cell I{row} (Row {row}, Column 9)")
            else:
                if drcr_flag == "DR":
                    total_debit += amount
                elif drcr_flag == "CR":
                    total_credit += amount
                
                amount_formatted = f"{amount:.2f}"
                sap_line = f"D|{row-6}|{fund}|{gl_account}|{bus_area}|{func_area}|{cost_center}|{wbs_element}|{amount_formatted}|{drcr_flag}||||||{description}\n"
                sap_lines.append(sap_line)

        if len(detailed_errors) > 0:
            return False, detailed_errors, total_debit, total_credit, None

        if abs(total_debit - total_credit) > 0.01:
            return False, ["Calculation Discrepancy Error: Out of Balance! Aggregate Sum of Debits must match Credits exactly."], total_debit, total_credit, None

        txt_content = "".join(sap_lines)
        return True, [], total_debit, total_credit, txt_content

    except Exception as e:
        return False, [f"Fatal Parse Error: Layout schema cannot be checked. Technical details: {str(e)}"], 0, 0, None


# ==========================================
# PHASE 3: THE WEB USER INTERFACE (SIDEBAR & PANELS)
# ==========================================
with st.sidebar:
    st.header("🗂 Template Operations")
    st.write("Need a fresh entry form? Click below to download a blank configured tracking worksheet.")
    
    template_data = create_jv_template()
    st.download_button(
        label="📥 Download Blank Form Template",
        data=template_data,
        file_name="JV_Upload_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.divider()
    st.caption("System Version: 2026.WebStream")

# Main Page Action Area
st.subheader("📁 Drag and Drop Validation & SAP Export")
st.write("Upload your completed Excel worksheet here to instantly test calculations and structure rules.")

uploaded_file = st.file_uploader(
    "Choose a completed JV Excel File (.xlsx)", 
    type=["xlsx"],
    key=f"jv_uploader_{st.session_state['uploader_key']}"
)

if uploaded_file is not None:
    st.info("File loaded successfully. Analyzing data structures...")
    
    success, errors, debits, credits, sap_txt_output = process_and_validate_jv(uploaded_file)
    
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.metric(label="Calculated Total Debits", value=f"${debits:,.2f}")
    with col2:
        st.metric(label="Calculated Total Credits", value=f"${credits:,.2f}")
    with col3:
        st.write("")
        st.write("")
        if st.button("🔄 Upload New JV", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.rerun()
        
    if success:
        st.write("") 
        st.success("🎉 Perfect! Your spreadsheet passed all date constraints, digit matching structure layouts, and balances perfectly!")
        
        st.download_button(
            label="💾 Download Ready SAP Text File (.txt)",
            data=sap_txt_output,
            file_name=f"SAP_JV_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain"
        )
    else:
        st.write("") 
        st.error("❌ Structural Validation Failed! Please resolve the following errors inside your spreadsheet and re-upload:")
        for err in errors:
            st.markdown(f"* {err}")
