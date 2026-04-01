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

.card {
    background: white;
    padding: 20px;
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# FIX FILE (REMOVE STYLE)
# =========================
def fix_excel_styles(path):
    import re

    tmp_dir = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as zin:
        zin.extractall(tmp_dir)

    # remove styles.xml
    style_path = os.path.join(tmp_dir, "xl", "styles.xml")
    if os.path.exists(style_path):
        os.remove(style_path)

    # remove style refs
    sheet_dir = os.path.join(tmp_dir, "xl", "worksheets")
    if os.path.exists(sheet_dir):
        for file in os.listdir(sheet_dir):
            if file.endswith(".xml"):
                p = os.path.join(sheet_dir, file)
                content = open(p, encoding="utf-8").read()
                content = re.sub(r'\s*s="\d+"', '', content)
                open(p, "w", encoding="utf-8").write(content)

    fixed_path = path.replace(".xlsx", "_fixed.xlsx")
    shutil.make_archive(fixed_path.replace(".xlsx", ""), 'zip', tmp_dir)
    os.rename(fixed_path.replace(".xlsx", ".zip"), fixed_path)

    return fixed_path

def safe_load(path):
    try:
        return load_workbook(path, data_only=True)
    except:
        fixed = fix_excel_styles(path)
        return load_workbook(fixed, data_only=True)

# =========================
# FIND COLUMN
# =========================
def find_col(ws):
    for c in ws[1]:
        if c.value and "Shipment Nbr" in str(c.value):
            return c.column
    return None

# =========================
# UI
# =========================
uploaded_files = st.file_uploader(
    "📂 Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("🚀 RUN TOOL"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("Chọn đúng 2 file")
        st.stop()

    tmp = tempfile.gettempdir()

    path_tpn = None
    path_book = None

    # =========================
    # SAVE FILES + DETECT
    # =========================
    for f in uploaded_files:
        path = os.path.join(tmp, f.name)
        open(path, "wb").write(f.read())

        wb = safe_load(path)
        ws = wb.active
        header = [str(c.value) if c.value else "" for c in ws[1]]

        if any("Shipment Nbr" in h for h in header):
            path_tpn = path
        else:
            path_book = path

        wb.close()

    # =========================
    # 🔥 CLONE FILE GỐC (GIỮ FORMAT)
    # =========================
    output_tpn = os.path.join(tmp, "TPN_KET_QUA.xlsx")
    shutil.copy(path_tpn, output_tpn)

    output_book = os.path.join(tmp, "TPN_KE_HOACH_XE.xlsx")
    shutil.copy(path_book, output_book)

    # =========================
    # READ DATA FROM FIXED FILE
    # =========================
    wb_fixed = safe_load(path_tpn)
    ws_fixed = wb_fixed.active

    col = find_col(ws_fixed)

    df = pd.read_excel(path_book, usecols=[0])

    all_nums = set()
    for v in df.iloc[:, 0].dropna().astype(str):
        all_nums.update(re.findall(r"\d{4}", v))

    match_rows = []
    ketqua_numbers = set()

    for i in range(2, ws_fixed.max_row + 1):
        val = ws_fixed.cell(i, col).value
        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            ketqua_numbers.update(nums)

            if nums & all_nums:
                match_rows.append(i)

    wb_fixed.close()

    # =========================
    # APPLY LÊN FILE COPY (KHÔNG ĐỤNG FILE GỐC)
    # =========================
    wb = load_workbook(output_tpn)
    ws = wb.active

    yellow = PatternFill("solid", fgColor="FFFF00")

    for i in match_rows:
        ws.cell(i, col).fill = yellow

    wb.save(output_tpn)
    wb.close()

    # =========================
    # FILE 2
    # =========================
    wb2 = load_workbook(output_book)
    ws2 = wb2.active

    red = Font(color="FF0000")

    for i in range(2, ws2.max_row + 1):
        val = ws2.cell(i, 1).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            if nums & ketqua_numbers:
                ws2.cell(i, 1).font = red

    wb2.save(output_book)
    wb2.close()

    # =========================
    # ZIP
    # =========================
    zip_path = os.path.join(tmp, "TPN_RESULT.zip")

    with zipfile.ZipFile(zip_path, "w") as z:
        z.write(output_tpn, "TPN_KET_QUA.xlsx")
        z.write(output_book, "TPN_KE_HOACH_XE.xlsx")

    st.download_button(
        "📥 Download",
        data=open(zip_path, "rb"),
        file_name="TPN_RESULT.zip"
    )
