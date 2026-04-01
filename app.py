import streamlit as st
import pandas as pd
import re
import tempfile
import os

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="TPN TOOL ⚡", layout="centered")

st.title("TPN TOOL ⚡")
st.write("Upload 2 file để xử lý")

file_tpn = st.file_uploader("Upload file TPN (có Shipment Nbr)", type=["xlsx"])
file_book1 = st.file_uploader("Upload file Book1", type=["xlsx"])

if st.button("RUN TOOL"):

    if not file_tpn or not file_book1:
        st.error("Thiếu file!")
        st.stop()

    with st.spinner("Đang xử lý..."):

        tmp_dir = tempfile.gettempdir()

        path_tpn = os.path.join(tmp_dir, "tpn.xlsx")
        path_book1 = os.path.join(tmp_dir, "book1.xlsx")

        with open(path_tpn, "wb") as f:
            f.write(file_tpn.read())

        with open(path_book1, "wb") as f:
            f.write(file_book1.read())

        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        # =========================
        # STEP 1: READ BOOK1
        # =========================
        df = pd.read_excel(path_book1, usecols=[0])

        all_numbers = set()
        for v in df.iloc[:, 0].dropna().astype(str):
            all_numbers.update(re.findall(r"\d{4}", v))

        # =========================
        # STEP 2: PROCESS TPN (KET_QUA)
        # =========================
        wb = load_workbook(path_tpn)
        ws = wb.active

        # tìm cột Shipment Nbr
        header = [cell.value for cell in ws[1]]
        col_index = None

        for i, v in enumerate(header):
            if v and str(v).strip() == "Shipment Nbr":
                col_index = i + 1
                break

        if not col_index:
            st.error("Không tìm thấy cột Shipment Nbr")
            st.stop()

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

        # FIX A1 (set active cell)
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"

        wb.save(save_path)
        wb.close()

        # =========================
        # STEP 3: PROCESS KE_HOACH
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

        # FIX A1
        ws2.sheet_view.selection[0].activeCell = "A1"
        ws2.sheet_view.selection[0].sqref = "A1"

        wb2.save(kehoach_path)
        wb2.close()

    st.success(f"Xong! Matched: {count}")

    # =========================
    # DOWNLOAD FILE
    # =========================
    with open(save_path, "rb") as f:
        st.download_button("Download KET_QUA", f, file_name="TPN_KET_QUA.xlsx")

    with open(kehoach_path, "rb") as f:
        st.download_button("Download KE_HOACH_XE", f, file_name="TPN_KE_HOACH_XE.xlsx")