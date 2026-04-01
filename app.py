import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
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
# HEADER
# =========================
st.markdown("""
<div style="text-align:center">
<h2>⚡ TPN TOOL</h2>
<p>Xử lý Shipment</p>
</div>
""", unsafe_allow_html=True)

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
# MAIN UI
# =========================
uploaded_files = st.file_uploader(
    "📂 Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("RUN TOOL"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("Cần đúng 2 file")
        st.stop()

    tmp_dir = tempfile.gettempdir()

    path_tpn = None
    path_book1 = None

    # =========================
    # SAVE FILES
    # =========================
    for f in uploaded_files:
        path = os.path.join(tmp_dir, f.name)
        with open(path, "wb") as out:
            out.write(f.read())

        wb_check = load_workbook(path, data_only=True)
        ws_check = wb_check.active

        header = [str(c.value).strip() if c.value else "" for c in ws_check[1]]

        if any("Shipment Nbr" in h for h in header):
            path_tpn = path
        else:
            path_book1 = path

        wb_check.close()

    # =========================
    # OUTPUT PATH
    # =========================
    save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
    kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

    # =========================
    # READ BOOK1
    # =========================
    df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

    all_numbers = set()
    for v in df.iloc[:, 0].dropna().astype(str):
        all_numbers.update(re.findall(r"\d{4}", v))

    # =========================
    # 🔥 LOAD TPN (GIỮ FULL FORMAT)
    # =========================
    wb = load_workbook(path_tpn)   # ❗ KHÔNG read_only
    ws = wb.active

    col_index = find_shipment_col(ws)

    if not col_index:
        st.error("Không tìm thấy Shipment Nbr")
        st.stop()

    yellow = PatternFill("solid", fgColor="FFFF00")

    ketqua_numbers = set()
    count = 0

    for i in range(2, ws.max_row + 1):
        val = ws.cell(i, col_index).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            ketqua_numbers.update(nums)

            if nums & all_numbers:
                ws.cell(i, col_index).fill = yellow
                count += 1

    wb.save(save_path)
    wb.close()

    # =========================
    # BOOK1 OUTPUT (GIỮ FORMAT GỐC)
    # =========================
    wb2 = load_workbook(path_book1)  # ❗ KHÔNG read_only
    ws2 = wb2.active

    red = Font(color="FF0000")

    for i in range(2, ws2.max_row + 1):
        val = ws2.cell(i, 1).value

        if val:
            nums = set(re.findall(r"\d\{4}", str(val)))
            if nums & ketqua_numbers:
                ws2.cell(i, 1).font = red

    wb2.save(kehoach_path)
    wb2.close()

    # =========================
    # ZIP
    # =========================
    zip_path = os.path.join(tmp_dir, "TPN_RESULT.zip")

    with zipfile.ZipFile(zip_path, "w") as z:
        z.write(save_path, "TPN_KET_QUA.xlsx")
        z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

    with open(zip_path, "rb") as f:
        zip_data = f.read()

    st.success(f"Done! Matched: {count}")

    st.download_button(
        "Download ZIP",
        data=zip_data,
        file_name="TPN_RESULT.zip"
    )
