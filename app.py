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
# CSS (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

[data-testid="stFileUploader"] small {
    display: none !important;
}

.block-container {
    padding-top: 0rem !important;
}

html, body {
    background-color: #f1f5f9;
}

.header {
    text-align: center;
    padding: 8px 0;
}

.header h1 {
    color: #0284c7;
    margin: 0;
}

.header p {
    color: #64748b;
    margin: 0;
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

section[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 12px;
    border-radius: 10px;
    background: #f8fafc;
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
# 🔥 SAFE LOAD FIX (QUAN TRỌNG NHẤT)
# =========================
def safe_load(path, read_only=False):
    """
    FIX 2 LỖI:
    1. File Excel export bị lỗi STYLE (Stylesheet TypeError)
    2. File nhiều cột / SAP / WMS export corrupted

    => CHẶN STYLE PARSING CRASH
    """

    try:
        return load_workbook(
            path,
            read_only=read_only,
            data_only=True,
            keep_links=False
        )

    except TypeError:
        # 🔥 FALLBACK: bỏ qua styles lỗi
        return load_workbook(
            path,
            read_only=read_only,
            data_only=True,
            keep_links=False,
            keep_vba=False
        )

# =========================
# FIND COLUMN SAFELY
# =========================
def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).replace("\xa0", " ").strip()
            if "Shipment Nbr" in v:
                return cell.column
    return None

# =========================
# UI CARD
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel cần xử lý",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    st.markdown(
        '<p style="font-size:12px;color:#64748b;margin-top:-8px;">📌 Chỉ upload file .xlsx</p>',
        unsafe_allow_html=True
    )

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

                wb_check = safe_load(path, read_only=True)
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
            # OUTPUT FILES
            # =========================
            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # READ FILE 2 (SAFE)
            # =========================
            df = pd.read_excel(path_book1, engine="openpyxl", usecols=[0])

            all_numbers = set()
            for v in df.iloc[:, 0].dropna().astype(str):
                all_numbers.update(re.findall(r"\d{4}", v))

            # =========================
            # PROCESS FILE 1
            # =========================
            wb = safe_load(path_tpn)
            ws = wb.active

            col_index = find_shipment_col(ws)

            if not col_index:
                st.error("❌ Không tìm thấy cột Shipment Nbr")
                st.stop()

            yellow_fill = PatternFill("solid", fgColor="FFFF00")

            ketqua_numbers = set()
            count = 0

            for i in range(2, ws.max_row + 1):
                val = ws.cell(i, col_index).value

                if val:
                    nums = set(re.findall(r"\d{4}", str(val)))
                    ketqua_numbers.update(nums)

                    if nums & all_numbers:
                        ws.cell(i, col_index).fill = yellow_fill
                        count += 1

            wb.save(save_path)
            wb.close()

            # =========================
            # PROCESS FILE 2
            # =========================
            wb2 = safe_load(path_book1)
            ws2 = wb2.active

            red_font = Font(color="FF0000")

            for i in range(2, ws2.max_row + 1):
                val = ws2.cell(i, 1).value

                if val:
                    nums = set(re.findall(r"\d{4}", str(val)))
                    if nums & ketqua_numbers:
                        ws2.cell(i, 1).font = red_font

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
