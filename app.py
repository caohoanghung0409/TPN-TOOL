import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import shutil
import uuid
import xml.etree.ElementTree as ET

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

.stButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
}
</style>
""", unsafe_allow_html=True)


# =========================
# SAFE EXCEL FIX (KHÔNG MẤT FORMAT)
# =========================
def fix_excel_styles(path):
    tmp_dir = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as zin:
        zin.extractall(tmp_dir)

    style_path = os.path.join(tmp_dir, "xl", "styles.xml")

    if os.path.exists(style_path):
        try:
            tree = ET.parse(style_path)
            root = tree.getroot()

            # remove broken nodes if exist
            for tag in ["fillCount", "cellXfs", "cellStyleXfs"]:
                for elem in root.findall(f".//{tag}"):
                    root.remove(elem)

            tree.write(style_path, encoding="utf-8", xml_declaration=True)

        except:
            pass

    fixed_path = path.replace(".xlsx", "_fixed.xlsx")
    shutil.make_archive(fixed_path.replace(".xlsx", ""), 'zip', tmp_dir)
    os.rename(fixed_path.replace(".xlsx", ".zip"), fixed_path)

    return fixed_path


# =========================
# SAFE LOAD (AUTO REPAIR)
# =========================
def safe_load(path):
    try:
        return load_workbook(
            path,
            data_only=False,
            keep_links=False
        )
    except Exception:
        fixed = fix_excel_styles(path)
        return load_workbook(
            fixed,
            data_only=False,
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
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

st.markdown("""
<div class="card">
<h2 style="text-align:center;">⚡ TPN TOOL</h2>
<p style="text-align:center;">Xử lý & đối soát Shipment</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📂 Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

if st.button("🚀 RUN TOOL"):

    if not uploaded_files or len(uploaded_files) != 2:
        st.error("⚠️ Cần đúng 2 file Excel")
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
                str(c.value).replace("\xa0", " ").strip() if c.value else ""
                for c in ws_check[1]
            ]

            wb_check.close()

            if any("Shipment Nbr" in h for h in header):
                path_tpn = path
            else:
                path_book1 = path

        # =========================
        # OUTPUT
        # =========================
        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        # =========================
        # FILE 2 - GET DATA
        # =========================
        df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl")

        all_numbers = set()
        for v in df.iloc[:, 0].dropna().astype(str):
            all_numbers.update(re.findall(r"\d{4}", v))

        # =========================
        # FILE 1 PROCESS (GIỮ FORMAT)
        # =========================
        wb = safe_load(path_tpn)
        ws = wb.active

        col_index = find_shipment_col(ws)

        if not col_index:
            st.error("❌ Không tìm thấy Shipment Nbr")
            st.stop()

        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        ketqua_numbers = set()
        count = 0

        for i in range(2, ws.max_row + 1):
            cell = ws.cell(i, col_index)

            if cell.value:
                nums = set(re.findall(r"\d{4}", str(cell.value)))
                ketqua_numbers.update(nums)

                if nums & all_numbers:
                    cell.fill = yellow
                    count += 1

        wb.save(save_path)
        wb.close()

        # =========================
        # FILE 2 PROCESS (GIỮ FORMAT)
        # =========================
        wb2 = safe_load(path_book1)
        ws2 = wb2.active

        red = Font(color="FF0000")

        for i in range(2, ws2.max_row + 1):
            cell = ws2.cell(i, 1)

            if cell.value:
                nums = set(re.findall(r"\d{4}", str(cell.value)))

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

    st.success(f"✅ Done! Matched: {count}")

    st.download_button(
        "📥 Download ZIP",
        data=zip_data,
        file_name="TPN_RESULT.zip"
    )

    st.session_state["uploader_key"] += 1
