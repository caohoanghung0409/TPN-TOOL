import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile

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

st.title("⚡ TPN TOOL")

# =========================
# SAFE REPAIR EXCEL (QUAN TRỌNG)
# =========================
def repair_excel(path):
    """
    FIX FILE EXCEL EXPORT LỖI FORMAT
    """
    wb = load_workbook(path, data_only=False, keep_links=False)
    clean_path = path.replace(".xlsx", "_CLEAN.xlsx")
    wb.save(clean_path)
    wb.close()
    return clean_path

# =========================
# GET HEADER SAFE
# =========================
def get_header(ws):
    header = []
    for cell in ws[1]:
        val = cell.value
        if val is not None:
            val = str(val).replace("\xa0", " ").strip()
        header.append(val)
    return header

# =========================
# UI
# =========================
uploaded_files = st.file_uploader(
    "Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("RUN"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("Cần đúng 2 file")
        st.stop()

    tmp_dir = tempfile.gettempdir()

    path_tpn = None
    path_book1 = None

    # =========================
    # SAVE FILES + REPAIR
    # =========================
    for file in uploaded_files:
        path = os.path.join(tmp_dir, file.name)

        with open(path, "wb") as f:
            f.write(file.read())

        try:
            clean_path = repair_excel(path)
        except Exception as e:
            st.error(f"File lỗi không đọc được: {file.name}")
            st.stop()

        ws_test = load_workbook(clean_path, read_only=True).active
        header = [str(c.value).strip() if c.value else "" for c in ws_test[1]]

        if any("Shipment Nbr" in h for h in header):
            path_tpn = clean_path
        else:
            path_book1 = clean_path

    # =========================
    # OUTPUT PATH
    # =========================
    save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
    kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

    # =========================
    # FILE 1 PROCESS
    # =========================
    df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

    all_numbers = set()
    for v in df.iloc[:, 0].dropna().astype(str):
        all_numbers.update(re.findall(r"\d{4}", v))

    wb = load_workbook(path_tpn)
    ws = wb.active

    header = get_header(ws)

    col_index = next(
        (i + 1 for i, v in enumerate(header)
         if v and "Shipment Nbr" in v),
        None
    )

    if not col_index:
        st.error("Không tìm thấy cột Shipment Nbr")
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
    # FILE 2 PROCESS
    # =========================
    wb2 = load_workbook(path_book1)
    ws2 = wb2.active

    red = Font(color="FF0000")

    for i in range(2, ws2.max_row + 1):
        val = ws2.cell(i, 1).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            if nums & ketqua_numbers:
                ws2.cell(i, 1).font = red

    wb2.save(kehoach_path)
    wb2.close()

    # =========================
    # ZIP
    # =========================
    zip_path = os.path.join(tmp_dir, "TPN.zip")

    with zipfile.ZipFile(zip_path, "w") as z:
        z.write(save_path, "TPN_KET_QUA.xlsx")
        z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

    with open(zip_path, "rb") as f:
        data = f.read()

    st.success(f"Done! matched={count}")

    st.download_button(
        "Download ZIP",
        data=data,
        file_name="TPN_RESULT.zip"
    )
