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
# HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

# =========================
# SAFE DETECT TPN
# =========================
def is_tpn(path):
    try:
        df = pd.read_excel(path, nrows=10)
        cols = [str(c).lower() for c in df.columns if c is not None]

        # chỉ cần chứa shipment là đủ
        return any("shipment" in c for c in cols)

    except Exception:
        return False


with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel cần xử lý",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if st.button("🚀 RUN TOOL"):

        if not uploaded_files or len(uploaded_files) != 2:
            st.error("⚠️ Vui lòng chọn đúng 2 file!")
            st.stop()

        with st.spinner("⏳ Đang xử lý..."):

            tmp_dir = tempfile.gettempdir()

            path1 = os.path.join(tmp_dir, uploaded_files[0].name)
            path2 = os.path.join(tmp_dir, uploaded_files[1].name)

            with open(path1, "wb") as f:
                f.write(uploaded_files[0].read())

            with open(path2, "wb") as f:
                f.write(uploaded_files[1].read())

            # =========================
            # DETECT (FIX CỐT LÕI)
            # =========================
            if is_tpn(path1):
                path_tpn = path1
                path_book1 = path2
            elif is_tpn(path2):
                path_tpn = path2
                path_book1 = path1
            else:
                st.error("❌ Không detect được file TPN (không thấy Shipment Nbr)")
                st.stop()

            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # =========================
            # BOOK1 READ SAFE
            # =========================
            df = pd.read_excel(path_book1)

            all_numbers = set()
            for v in df.astype(str).values.flatten():
                all_numbers.update(re.findall(r"\d{4}", str(v)))

            # =========================
            # TPN PROCESS
            # =========================
            wb = load_workbook(path_tpn)
            ws = wb.active

            header = [str(c.value).strip() if c.value else "" for c in ws[1]]

            col_index = None
            for i, v in enumerate(header):
                if "Shipment" in v:
                    col_index = i + 1
                    break

            if col_index is None:
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
            # BOOK1 OUTPUT
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
            # FIX VIEW
            # =========================
            def fix_view(p):
                wbv = load_workbook(p)
                wbv.active.sheet_view.topLeftCell = "A1"
                wbv.save(p)
                wbv.close()

            fix_view(save_path)
            fix_view(kehoach_path)

            # =========================
            # ZIP
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

    st.markdown('</div>', unsafe_allow_html=True)
