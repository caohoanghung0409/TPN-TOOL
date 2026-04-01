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
# MODERN UI - CSS UPGRADE
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

html, body {
    font-family: 'Inter', sans-serif;
    background: linear-gradient(135deg, #eef2ff, #f0f9ff, #ecfeff);
}

/* hide default streamlit UI */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {display: none !important;}

/* center container */
.block-container {
    padding-top: 1.2rem !important;
    max-width: 900px;
}

/* HEADER */
.header {
    text-align: center;
    margin-bottom: 20px;
    animation: fadeIn 0.6s ease-in-out;
}

.header h1 {
    font-size: 34px;
    font-weight: 800;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 6px;
}

.header p {
    color: #475569;
    font-size: 14px;
}

/* CARD */
.card {
    background: white;
    padding: 24px;
    border-radius: 18px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.08);
    border: 1px solid #e2e8f0;
    animation: fadeInUp 0.5s ease-in-out;
}

/* buttons */
.stButton>button {
    width: 100%;
    height: 44px;
    border-radius: 12px;
    border: none;
    font-weight: 600;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
    transition: 0.2s;
}

.stButton>button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 15px rgba(34,197,94,0.25);
}

/* download button */
.stDownloadButton>button {
    width: 100%;
    height: 44px;
    border-radius: 12px;
    font-weight: 600;
    background: linear-gradient(90deg, #16a34a, #22c55e);
    color: white;
}

/* file uploader */
section[data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    padding: 14px;
    border-radius: 12px;
    background: #f8fafc;
}

/* hide small note */
[data-testid="stFileUploader"] small {
    display: none !important;
}

/* animation */
@keyframes fadeIn {
    from {opacity: 0;}
    to {opacity: 1;}
}

@keyframes fadeInUp {
    from {opacity: 0; transform: translateY(10px);}
    to {opacity: 1; transform: translateY(0);}
}
</style>
""", unsafe_allow_html=True)

# =========================
# STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# =========================
# HEADER UI
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

# =========================
# FIX EXCEL CORRUPT
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

# =========================
# SAFE LOAD
# =========================
def safe_load(path, read_only=False):
    try:
        return load_workbook(path, read_only=read_only, data_only=True, keep_links=False)
    except Exception:
        fixed = fix_excel_styles(path)
        return load_workbook(fixed, read_only=read_only, data_only=True, keep_links=False)

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
        '<p style="font-size:12px;color:#64748b;">📌 Chỉ upload file .xlsx</p>',
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

            for file in uploaded_files:
                path = os.path.join(tmp_dir, file.name)

                with open(path, "wb") as f:
                    f.write(file.read())

                wb_check = safe_load(path, read_only=True)
                ws_check = wb_check.active

                header = [
                    str(c.value).replace("\xa0", " ").strip() if c.value else ""
                    for c in ws_check[1]
                ]

                wb_check.close()

                if any("Shipment Nbr" in h for h in header):
                    path_tpn = path
                else:
                    path_book1 = path

            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # FILE 1
            # =========================
            df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

            all_numbers = set()
            for v in df.iloc[:, 0].dropna().astype(str):
                all_numbers.update(re.findall(r"\d{4}", v))

            wb = safe_load(path_tpn)
            ws = wb.active

            ws.sheet_view.topLeftCell = "A1"
            ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

            col_index = find_shipment_col(ws)

            if not col_index:
                st.error("❌ Không tìm thấy cột Shipment Nbr")
                st.stop()

            yellow = PatternFill("solid", fgColor="FFFF00")

            ketqua_numbers = set()
            count = 0

            header_fill = PatternFill("solid", fgColor="0B3D91")
            header_font = Font(color="FFFFFF", bold=True)

            for cell in ws[1]:
                if cell.value:
                    cell.fill = header_fill
                    cell.font = header_font

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

            wb.save(save_path)
            wb.close()

            # =========================
            # FILE 2
            # =========================
            wb2 = safe_load(path_book1)
            ws2 = wb2.active

            ws2.sheet_view.topLeftCell = "A1"
            ws2.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

            red = Font(color="FF0000")

            for i in range(2, ws2.max_row + 1):
                val = ws2.cell(i, 1).value
                if val:
                    nums = set(re.findall(r"\d{4}", str(val)))
                    if nums & ketqua_numbers:
                        ws2.cell(i, 1).font = red

            auto_adjust_column_width(ws2)

            wb2.save(kehoach_path)
            wb2.close()

            # =========================
            # ZIP
            # =========================
            zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

            with zipfile.ZipFile(zip_path, "w") as z:
                z.write(save_path, "TPN_KET_QUA.xlsx")
                z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

            with open(zip_path, "rb") as f:
                zip_data = f.read()

        st.success(f"✅ COMPLETE !!! Matched: {count}")

        st.download_button(
            "📥 Download ALL (ZIP)",
            data=zip_data,
            file_name="TPN_COMPLETE.zip"
        )

        st.session_state["uploader_key"] += 1

    st.markdown('</div>', unsafe_allow_html=True)
