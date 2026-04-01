import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="TPN TOOL ⚡", layout="centered")

# =========================
# SUPER SAFE EXCEL READER
# =========================
def read_excel_safe(path, nrows=None):
    """
    4-layer fallback reader:
    1. pandas + openpyxl
    2. pandas auto engine
    3. openpyxl read_only
    4. XML fallback (xlsx zip)
    """

    # =========================
    # LAYER 1
    # =========================
    try:
        return pd.read_excel(path, nrows=nrows, engine="openpyxl")
    except Exception:
        pass

    # =========================
    # LAYER 2
    # =========================
    try:
        return pd.read_excel(path, nrows=nrows)
    except Exception:
        pass

    # =========================
    # LAYER 3 (openpyxl safe mode)
    # =========================
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active

        data = list(ws.values)

        if not data:
            return pd.DataFrame()

        df = pd.DataFrame(data[1:], columns=data[0])
        if nrows:
            df = df.head(nrows)

        return df

    except Exception:
        pass

    # =========================
    # LAYER 4 (ZIP XML RAW)
    # =========================
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/worksheets/sheet1.xml")

        root = ET.fromstring(xml)

        rows = []
        for row in root.iter():
            if row.tag.endswith("row"):
                values = []
                for c in row:
                    values.append(c.text)
                rows.append(values)

        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        return df.head(nrows if nrows else 1)

    except Exception:
        return pd.DataFrame()


# =========================
# HEADER DETECT SAFE
# =========================
def get_header(path):
    df = read_excel_safe(path, nrows=1)
    if df.empty:
        return []
    return [str(x).strip() for x in df.columns]


# =========================
# UI
# =========================
st.markdown("""
<style>
header {display:none;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

.block-container {padding-top:0rem;}

html, body {background:#f1f5f9;}

.card {
    background:white;
    padding:20px;
    border-radius:12px;
}
</style>
""", unsafe_allow_html=True)

st.title("⚡ TPN TOOL FIX PRO")

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0


with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Upload 2 Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"up_{st.session_state['uploader_key']}"
    )

    if st.button("🚀 RUN TOOL"):

        if not uploaded_files or len(uploaded_files) != 2:
            st.error("Chọn đúng 2 file")
            st.stop()

        tmp_dir = tempfile.gettempdir()

        path_tpn = None
        path_book1 = None

        # =========================
        # SAVE + DETECT
        # =========================
        for file in uploaded_files:
            path = os.path.join(tmp_dir, file.name)

            with open(path, "wb") as f:
                f.write(file.read())

            header = get_header(path)

            st.write(f"📄 {file.name} → {header}")

            if "Shipment Nbr" in header:
                path_tpn = path
            else:
                path_book1 = path

        if not path_tpn or not path_book1:
            st.error("❌ Không đọc được file hoặc không nhận diện đúng TPN")
            st.stop()

        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        # =========================
        # FILE 1
        # =========================
        df = pd.read_excel(path_book1, usecols=[0])

        all_numbers = set()
        for v in df.iloc[:, 0].dropna().astype(str):
            all_numbers.update(re.findall(r"\d{4}", v))

        wb = load_workbook(path_tpn)
        ws = wb.active

        header = [cell.value for cell in ws[1]]
        col_index = next(
            (i + 1 for i, v in enumerate(header)
             if v and str(v).strip() == "Shipment Nbr"),
            None
        )

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
        # FIX EXCEL VIEW
        # =========================
        def fix_excel(path, auto_fit=False):
            wb_fix = load_workbook(path)
            ws_fix = wb_fix.active

            ws_fix.sheet_view.topLeftCell = "A1"

            if auto_fit:
                for col in ws_fix.columns:
                    max_length = 0
                    col_letter = col[0].column_letter

                    for cell in col:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))

                    ws_fix.column_dimensions[col_letter].width = max_length + 2

            wb_fix.save(path)
            wb_fix.close()

        fix_excel(save_path, False)
        fix_excel(kehoach_path, True)

        # =========================
        # ZIP
        # =========================
        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")

        with zipfile.ZipFile(zip_buffer.name, "w") as zipf:
            zipf.write(save_path, "TPN_KET_QUA.xlsx")
            zipf.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_buffer.name, "rb") as f:
            zip_data = f.read()

        st.success(f"✅ Done! Matched: {count}")

        st.download_button(
            "📥 Download ZIP",
            data=zip_data,
            file_name="TPN_RESULT.zip"
        )

        st.session_state["uploader_key"] += 1

    st.markdown('</div>', unsafe_allow_html=True)
