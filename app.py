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
# SAFE EXCEL READER (FIX CRASH STYLE)
# =========================
def safe_read_excel(path, **kwargs):
    try:
        return pd.read_excel(path, engine="openpyxl", **kwargs)

    except Exception:
        # remove broken style
        tmp_fixed = path.replace(".xlsx", "_fixed.xlsx")

        with zipfile.ZipFile(path, "r") as zin:
            with zipfile.ZipFile(tmp_fixed, "w") as zout:
                for item in zin.infolist():
                    if item.filename != "xl/styles.xml":
                        zout.writestr(item, zin.read(item.filename))

        return pd.read_excel(tmp_fixed, engine="openpyxl", **kwargs)


# =========================
# UI
# =========================
st.markdown("""
<style>
header {display:none;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
.block-container {padding-top:0rem;}
</style>
""", unsafe_allow_html=True)

st.title("⚡ TPN TOOL")

uploaded_files = st.file_uploader(
    "Chọn 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("RUN TOOL"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("Cần đúng 2 file!")
        st.stop()

    tmp_dir = tempfile.gettempdir()

    path_tpn = None
    path_book1 = None

    # =========================
    # SAVE FILES + DETECT
    # =========================
    for file in uploaded_files:
        path = os.path.join(tmp_dir, file.name)

        with open(path, "wb") as f:
            f.write(file.read())

        df_check = safe_read_excel(path, nrows=1)
        header = [str(x).strip() for x in df_check.columns]

        if "Shipment Nbr" in header:
            path_tpn = path
        else:
            path_book1 = path

    save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
    kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

    # =========================
    # FILE 1
    # =========================
    df = safe_read_excel(path_book1, usecols=[0])

    all_numbers = set()
    for v in df.iloc[:, 0].dropna().astype(str):
        all_numbers.update(re.findall(r"\d{4}", v))

    # =========================
    # FILE TPN
    # =========================
    wb = load_workbook(path_tpn)
    ws = wb.active

    header = [cell.value for cell in ws[1]]
    col_index = next((i + 1 for i, v in enumerate(header)
                      if v and str(v).strip() == "Shipment Nbr"), None)

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ketqua_numbers = set()
    count = 0

    for i in range(2, ws.max_row + 1):
        val = ws.cell(row=i, column=col_index).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            ketqua_numbers.update(nums)

            if nums & all_numbers:
                ws.cell(row=i, column=col_index).fill = yellow_fill
                count += 1

    wb.save(save_path)
    wb.close()

    # =========================
    # FILE 2
    # =========================
    wb2 = load_workbook(path_book1)
    ws2 = wb2.active

    red_font = Font(color="FF0000")

    for i in range(2, ws2.max_row + 1):
        val = ws2.cell(row=i, column=1).value

        if val:
            nums = set(re.findall(r"\d{4}", str(val)))
            if nums & ketqua_numbers:
                ws2.cell(row=i, column=1).font = red_font

    wb2.save(kehoach_path)
    wb2.close()

    # =========================
    # ZIP
    # =========================
    zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

    with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
        zipf.write(save_path, "TPN_KET_QUA.xlsx")
        zipf.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

    with open(zip_buffer.name, "rb") as f:
        zip_data = f.read()

    st.success(f"Done! Match: {count}")

    st.download_button(
        "Download ZIP",
        data=zip_data,
        file_name="TPN_COMPLETE.zip"
    )
