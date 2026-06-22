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
        ("Control Total (Credits):", "H1")
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
        
        for col_idx in range(1, len(headers) +
