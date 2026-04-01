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

.block-container {
    padding-top: 0rem !important;
}

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
</style>
""", unsafe_allow_html=True)

# =========================
# STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# =========================
# HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

# =========================
# SAFE LOAD (KHÔNG XOÁ STYLE NỮA)
# =========================
def safe_load(path):
    return load_workbook(
        path,
        data_only=False,   # GIỮ FORMULA + FORMAT
        keep_links=False
    )

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
# UI
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel cần xử lý",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    st.markdown('<p style="font-size:12px;color:#64748b;">📌 Chỉ upload file .xlsx</p>', unsafe_allow_html=True)

    if st.button("🚀 RUN TOOL"):

        if not uploaded_files or len(uploaded_files) != 2:
            st.error("⚠️ Vui lòng chọn đúng 2 file!")
            st.stop()

        with st.spinner("⏳ Đang xử lý..."):

            tmp_dir = tempfile.gettempdir()

            path_tpn = None
            path_book1 = None

            # =========================
            # SAVE FILES
            # =========================
            for file in uploaded_files:
                path = os.path.join(tmp_dir, file.name)

                with open(path, "wb") as f:
                    f.write(file.read())

                wb_check = safe_load(path)
                ws_check = wb_check.active

                header = [
                    str(c.value).replace("\xa0", " ").strip()
                    if c.value else ""
                    for c in ws_check[1]
                ]

                wb_check.close()

                if any("Shipment Nbr" in h for h in header):
                    path_tpn = path
                else:
                    path_book1 = path

            # =========================
            # OUTPUT PATH
            # =========================
            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # FILE 2 - LẤY 4 SỐ
            # =========================
            df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

            all_numbers = set()
            for v in df.iloc[:, 0].dropna().astype(str):
                all_numbers.update(re.findall(r"\d{4}", v))

            # =========================
            # FILE 1 - GIỮ FORMAT GỐC
            # =========================
            wb = safe_load(path_tpn)
            ws = wb.active

            col_index = find_shipment_col(ws)

            if not col_index:
                st.error("❌ Không tìm thấy cột Shipment Nbr")
                st.stop()

            yellow = PatternFill("solid", fgColor="FFFF00")

            ketqua_numbers = set()
            count = 0

            for i in range(2, ws.max_row + 1):
                cell = ws.cell(i, col_index)
                val = cell.value

                if val:
                    nums = set(re.findall(r"\d{4}", str(val)))
                    ketqua_numbers.update(nums)

                    if nums & all_numbers:
                        # CHỈ THÊM FILL, KHÔNG ĐỤNG FORMAT KHÁC
                        cell.fill = yellow
                        count += 1

            wb.save(save_path)
            wb.close()

            # =========================
            # FILE 2 - GIỮ FORMAT GỐC
            # =========================
            wb2 = safe_load(path_book1)
            ws2 = wb2.active

            red = Font(color="FF0000")

            for i in range(2, ws2.max_row + 1):
                cell = ws2.cell(i, 1)
                val = cell.value

                if val:
                    nums = set(re.findall(r"\d{4}", str(val)))
                    if nums & ketqua_numbers:
                        cell.font = red

            wb2.save(kehoach_path)
            wb2.close()

            # =========================
            # ZIP OUTPUT
            # =========================
            zip_path = os.path.join(tmp_dir, "TPN_RESULT.zip")

            with zipfile.ZipFile(zip_path, "w") as z:
                z.write(save_path, "TPN_KET_QUA.xlsx")
                z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

            with open(zip_path, "rb") as f:
                zip_data = f.read()

        st.success(f"✅ Hoàn tất! Matched: {count}")

        st.download_button(
            "📥 Download ALL (ZIP)",
            data=zip_data,
            file_name="TPN_RESULT.zip"
        )

        st.session_state["uploader_key"] += 1

    st.markdown('</div>', unsafe_allow_html=True)
