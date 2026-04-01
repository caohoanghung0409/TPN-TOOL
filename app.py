import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="TPN TOOL ⚡", layout="centered")

# =========================
# CSS FIX FULL (QUAN TRỌNG)
# =========================
st.markdown("""
<style>

/* 🔥 ẨN HEADER STREAMLIT */
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* 🔥 FIX KHOẢNG TRẮNG TRIỆT ĐỂ */
[data-testid="stAppViewContainer"] {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

[data-testid="stAppViewBlockContainer"] {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

section.main {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

section.main > div {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

.block-container {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

html, body {
    margin: 0 !important;
    padding: 0 !important;
}

/* ===== UI ===== */
body {
    background-color: #f1f5f9;
}

/* header custom */
.header {
    text-align: center;
    padding: 5px 0 15px 0;
}

.header h1 {
    color: #0284c7;
    font-size: 28px;
    margin: 0;
}

.header p {
    color: #64748b;
    font-size: 14px;
    margin: 0;
}

/* card */
.card {
    background: white;
    padding: 25px;
    border-radius: 12px;
    box-shadow: 0 5px 20px rgba(0,0,0,0.08);
}

/* button RUN */
.stButton>button {
    width: 100%;
    height: 45px;
    border-radius: 10px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
    font-size: 16px;
}

/* button DOWNLOAD */
.stDownloadButton>button {
    width: 100%;
    height: 45px;
    border-radius: 10px;
    background: #16a34a;
    color: white;
    font-size: 16px;
}

/* uploader */
section[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 15px;
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
# CARD
# =========================
st.markdown('<div class="card">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📂 Chọn 2 file Excel (TPN + Book1)",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

# =========================
# RUN TOOL
# =========================
if st.button("🚀 RUN TOOL"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("⚠️ Vui lòng chọn đúng 2 file!")
        st.stop()

    with st.spinner("⏳ Đang xử lý dữ liệu..."):

        tmp_dir = tempfile.gettempdir()

        path_tpn = None
        path_book1 = None

        # SAVE + DETECT FILE
        for file in uploaded_files:
            path = os.path.join(tmp_dir, file.name)

            with open(path, "wb") as f:
                f.write(file.read())

            df_check = pd.read_excel(path, nrows=1)
            header = [str(x).strip() for x in df_check.columns]

            if "Shipment Nbr" in header:
                path_tpn = path
            else:
                path_book1 = path

        if not path_tpn or not path_book1:
            st.error("❌ Không detect được file!")
            st.stop()

        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        # =========================
        # READ BOOK1
        # =========================
        df = pd.read_excel(path_book1, usecols=[0])

        all_numbers = set()
        for v in df.iloc[:, 0].dropna().astype(str):
            all_numbers.update(re.findall(r"\d{4}", v))

        # =========================
        # PROCESS TPN
        # =========================
        wb = load_workbook(path_tpn)
        ws = wb.active

        header = [cell.value for cell in ws[1]]
        col_index = next((i + 1 for i, v in enumerate(header)
                          if v and str(v).strip() == "Shipment Nbr"), None)

        if not col_index:
            st.error("❌ Không tìm thấy Shipment Nbr")
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

        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"
        ws.sheet_view.topLeftCell = "A1"

        wb.save(save_path)
        wb.close()

        # =========================
        # PROCESS KE_HOACH
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

        # AUTO FIT
        for col in ws2.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            ws2.column_dimensions[col_letter].width = min(max_length + 2, 50)

        ws2.sheet_view.selection[0].activeCell = "A1"
        ws2.sheet_view.selection[0].sqref = "A1"
        ws2.sheet_view.topLeftCell = "A1"

        wb2.save(kehoach_path)
        wb2.close()

        # =========================
        # ZIP FILE
        # =========================
        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            zipf.write(save_path, "TPN_KET_QUA.xlsx")
            zipf.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_buffer.name, "rb") as f:
            zip_data = f.read()

    st.success(f"✅ Hoàn tất! Matched: {count}")

    st.download_button(
        "📥 Download ALL (ZIP)",
        data=zip_data,
        file_name="TPN_RESULT.zip"
    )

    # RESET
    st.session_state["uploader_key"] += 1

st.markdown('</div>', unsafe_allow_html=True)
