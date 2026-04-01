import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import shutil
import uuid

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

st.set_page_config(page_title="TPN TOOL ⚡", layout="centered")

# =========================
# CSS
# =========================
st.markdown("""
<style>
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {padding-top: 0rem !important;}

html, body {background-color: #f1f5f9;}

.header {
    text-align: center;
    padding: 8px 0;
}

.header h1 {color: #0284c7; margin: 0;}
.header p {color: #64748b; margin: 0;}

.card {
    background: white;
    padding: 20px;
    border-radius: 12px;
}

.stButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
}

.stDownloadButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: #16a34a;
    color: white;
}

/* FOOTER */
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background: white;
    color: #64748b;
    text-align: center;
    padding: 10px;
    font-size: 12px;
    border-top: 1px solid #e2e8f0;
    z-index: 9999;
}
</style>
""", unsafe_allow_html=True)

# =========================
# FIX EXCEL
# =========================
def fix_excel_styles(path):
    tmp_dir = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as zin:
        zin.extractall(tmp_dir)

    style_path = os.path.join(tmp_dir, "xl", "styles.xml")
    if os.path.exists(style_path):
        os.remove(style_path)

    sheet_dir = os.path.join(tmp_dir, "xl", "worksheets")

    if os.path.exists(sheet_dir):
        for file in os.listdir(sheet_dir):
            if file.endswith(".xml"):
                fpath = os.path.join(sheet_dir, file)

                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()

                content = re.sub(r'\s*s="\d+"', '', content)

                with open(fpath, "w", encoding="utf-8") as f:
                    f.write(content)

    fixed_path = path.replace(".xlsx", "_fixed.xlsx")
    shutil.make_archive(fixed_path.replace(".xlsx", ""), 'zip', tmp_dir)
    os.rename(fixed_path.replace(".xlsx", ".zip"), fixed_path)

    return fixed_path

def safe_load(path):
    try:
        return load_workbook(path)
    except:
        return load_workbook(fix_excel_styles(path))

# =========================
# FIND COLUMN
# =========================
def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).replace("\xa0", " ").strip()
            if "Shipment Nbr" in v:
                return cell.column
    return None

# =========================
# AUTO WIDTH
# =========================
def auto_adjust_column_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_len + 3

# =========================
# UI
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="card">', unsafe_allow_html=True)

files = st.file_uploader(
    "📂 Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("🚀 RUN TOOL"):

    if not files or len(files) != 2:
        st.error("Vui lòng upload đúng 2 file")
        st.stop()

    tmp_dir = tempfile.gettempdir()

    path_tpn = None
    path_book1 = None

    # save files
    for f in files:
        path = os.path.join(tmp_dir, f.name)
        with open(path, "wb") as x:
            x.write(f.read())

        wb = safe_load(path)
        ws = wb.active

        header = [str(c.value).strip() if c.value else "" for c in ws[1]]

        if any("Shipment Nbr" in h for h in header):
            path_tpn = path
        else:
            path_book1 = path

    # =========================
    # FILE 2 DATA
    # =========================
    df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

    all_numbers = set()
    for v in df.iloc[:, 0].dropna().astype(str):
        all_numbers.update(re.findall(r"\d{4}", v))

    # =========================
    # FILE 1
    # =========================
    wb = safe_load(path_tpn)
    ws = wb.active

    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    col_index = find_shipment_col(ws)

    yellow = PatternFill("solid", fgColor="FFFF00")
    header_fill = PatternFill("solid", fgColor="000080")
    header_font = Font(color="FFFFFF", bold=True)

    ketqua_numbers = set()
    count = 0

    # header style
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    # bold all
    bold_font = Font(bold=True)
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                cell.font = bold_font

    for i in range(2, ws.max_row + 1):
        val = ws.cell(i, col_index).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            ketqua_numbers.update(nums)

            if nums & all_numbers:
                ws.cell(i, col_index).fill = yellow
                count += 1

    save1 = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
    wb.save(save1)
    wb.close()

    # =========================
    # FILE 2 (FIXED COLOR HERE)
    # =========================
    wb2 = safe_load(path_book1)
    ws2 = wb2.active

    ws2.sheet_view.topLeftCell = "A1"
    ws2.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    red_font = Font(color="FF0000", bold=True)

    for i in range(2, ws2.max_row + 1):
        val = ws2.cell(i, 1).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))   # ✅ FIXED BUG
            if nums & ketqua_numbers:
                ws2.cell(i, 1).font = red_font

    auto_adjust_column_width(ws2)

    save2 = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")
    wb2.save(save2)
    wb2.close()

    # =========================
    # ZIP
    # =========================
    zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

    with zipfile.ZipFile(zip_path, "w") as z:
        z.write(save1, "TPN_KET_QUA.xlsx")
        z.write(save2, "TPN_KE_HOACH_XE.xlsx")

    with open(zip_path, "rb") as f:
        data = f.read()

    st.success(f"Done! Matched: {count}")

    st.download_button("📥 Download ZIP", data, file_name="TPN_COMPLETE.zip")

st.markdown('</div>', unsafe_allow_html=True)

# =========================
# FOOTER
# =========================
st.markdown("""
<div class="footer">
© 2026 TPN TOOL • Built with Streamlit • All rights reserved
</div>
""", unsafe_allow_html=True)
