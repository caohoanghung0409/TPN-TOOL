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

# =========================
# HEADER
# =========================
st.markdown("""
<div style="text-align:center">
<h2>⚡ TPN TOOL</h2>
<p>Xử lý Shipment</p>
</div>
""", unsafe_allow_html=True)


# =========================
# DETECT FILE TPN
# =========================
def detect_tpn_file(path):
    try:
        df = pd.read_excel(path, nrows=5)
        cols = [str(c).strip() for c in df.columns]
        return any("Shipment Nbr" in c for c in cols)
    except:
        return False


# =========================
# UI
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if st.button("🚀 RUN TOOL"):

        if not uploaded_files or len(uploaded_files) != 2:
            st.error("⚠️ Vui lòng chọn đúng 2 file Excel!")
            st.stop()

        with st.spinner("⏳ Đang xử lý..."):

            tmp_dir = tempfile.gettempdir()

            path_tpn = None
            path_book1 = None

            # =========================
            # SAVE + DETECT FILES
            # =========================
            for file in uploaded_files:
                path = os.path.join(tmp_dir, file.name)

                with open(path, "wb") as f:
                    f.write(file.read())

                if detect_tpn_file(path):
                    path_tpn = path
                else:
                    path_book1 = path

            if not path_tpn or not path_book1:
                st.error("❌ Không nhận diện được file TPN hoặc file dữ liệu!")
                st.stop()

            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # READ BOOK1
            # =========================
            df = pd.read_excel(path_book1)

            all_numbers = set()
            for v in df.astype(str).values.flatten():
                all_numbers.update(re.findall(r"\d{4}", v))

            # =========================
            # PROCESS TPN FILE
            # =========================
            wb = load_workbook(path_tpn)
            ws = wb.active

            header = [str(c.value).strip() if c.value else "" for c in ws[1]]

            col_index = None
            for i, v in enumerate(header):
                if "Shipment Nbr" in v:
                    col_index = i + 1
                    break

            if not col_index:
                st.error("❌ Không tìm thấy cột Shipment Nbr")
                st.stop()

            yellow_fill = PatternFill("solid", fgColor="FFFF00")

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
            # PROCESS BOOK1 OUTPUT
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
            # FORCE A1 VIEW
            # =========================
            def fix_excel_view(path):
                wb_fix = load_workbook(path)
                ws_fix = wb_fix.active

                ws_fix.sheet_view.topLeftCell = "A1"
                ws_fix.sheet_view.selection.clear()

                wb_fix.save(path)
                wb_fix.close()

            fix_excel_view(save_path)
            fix_excel_view(kehoach_path)

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

        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
