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
# UI STYLE
# =========================
st.markdown("""
<style>

header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 0rem !important;
}

.header {
    text-align: center;
    padding: 10px 0;
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
# HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý Shipment an toàn (anti-crash)</p>
</div>
""", unsafe_allow_html=True)


# =========================
# GET HEADER (NO OPENPYXL STYLE PARSE)
# =========================
def get_excel_header(path):
    """
    Đọc header trực tiếp từ XML -> tránh crash style openpyxl
    """
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("xl/worksheets/sheet1.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()

                ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

                rows = root.findall(".//a:row", ns)
                if not rows:
                    return []

                first_row = rows[0]
                cells = first_row.findall(".//a:c", ns)

                header = []
                for c in cells:
                    v = c.find("a:v", ns)
                    header.append(v.text if v is not None else "")

                return header
    except:
        return []


# =========================
# SESSION
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0


# =========================
# UI
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn đúng 2 file Excel",
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

                # 🔥 SAFE HEADER DETECT (NO OPENPYXL STYLE CRASH)
                header = get_excel_header(path)
                header = [str(x).strip() for x in header]

                if "Shipment Nbr" in header:
                    path_tpn = path
                else:
                    path_book1 = path

            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # FILE 1
            # =========================
            df = pd.read_excel(path_book1, usecols=[0])

            all_numbers = set()
            for v in df.iloc[:, 0].dropna().astype(str):
                all_numbers.update(re.findall(r"\d{4}", v))

            # =========================
            # FILE 2
            # =========================
            wb = load_workbook(path_tpn)
            ws = wb.active

            header_row = [cell.value for cell in ws[1]]

            col_index = next(
                (i + 1 for i, v in enumerate(header_row)
                 if v and str(v).strip() == "Shipment Nbr"),
                None
            )

            yellow_fill = PatternFill(
                start_color="FFFF00",
                end_color="FFFF00",
                fill_type="solid"
            )

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
            # FILE 3
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
            # ZIP OUTPUT
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
            file_name="TPN_COMPLETE.zip"
        )

        st.session_state["uploader_key"] += 1

    st.markdown('</div>', unsafe_allow_html=True)
