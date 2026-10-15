from io import BytesIO
from pathlib import Path
import streamlit as st
from docx import Document
from docx.shared import RGBColor

st.set_page_config(page_title="Pengisian SPPD", page_icon="📄", layout="centered")
TEMPLATE = Path(__file__).parent / "template_sppd.docx"

st.title("Pengisian SPPD")
st.caption("Hanya tulisan berwarna merah pada template yang dapat diisi. Label biru/pertanyaan dan teks hitam tetap mengikuti dokumen asli.")
st.info("Template Word asli digunakan sebagai dasar. Aplikasi hanya mengganti teks jawaban merah yang telah ditentukan; teks hitam, label pertanyaan, tanda baca, tabel, dan garis tidak dibuat ulang.")

with st.form("sppd"):
    st.subheader("I. Data perjalanan")
    asal = st.text_input("Berangkat dari", "Tangerang")
    keperluan = st.text_input("Ke / keperluan", "Try Out Catur")
    tiba_kembali = st.text_input("Tiba di", "Tangerang")
    tanggal_kembali = st.text_input("Tanggal tiba", "18 Oktober 2026")

    st.subheader("II. Lokasi perjalanan")
    lokasi_kiri = st.text_input("Lokasi tiba — kolom kiri", "Piset Square. Jl. Pelajar Pejuang No. 119 Bandung")
    tanggal_kiri = st.text_input("Tanggal — kolom kiri", "17 Oktober 2026")
    lokasi_kanan = st.text_input("Lokasi tiba — kolom kanan", "Piset Square. Jl. Pelajar Pejuang No. 119 Bandung")
    tanggal_kanan = st.text_input("Tanggal — kolom kanan", "18 Oktober 2026")
    submit = st.form_submit_button("Buat dokumen Word", type="primary")

# Pemetaan hanya pada paragraf yang memang berisi jawaban berwarna merah.
# Setiap entri: (baris tabel, kolom tabel, indeks paragraf dalam sel, nilai formulir)
RED_FIELDS = [
    (0, 3, 0, "asal"),
    (0, 3, 1, "keperluan"),
    (0, 5, 0, "tiba_kembali"),
    (0, 5, 1, "tanggal_kembali"),
    (1, 3, 0, "lokasi_kiri"),
    (1, 3, 2, "tanggal_kiri"),
    (1, 5, 0, "lokasi_kanan"),
    (1, 5, 1, "tanggal_kanan"),
]

if submit:
    if not TEMPLATE.exists():
        st.error("File template_sppd.docx tidak ditemukan. Pastikan file ini ada di repository yang sama dengan app.py.")
        st.stop()
    doc = Document(str(TEMPLATE))
    values = {
        "asal": asal, "keperluan": keperluan,
        "tiba_kembali": tiba_kembali, "tanggal_kembali": tanggal_kembali,
        "lokasi_kiri": lokasi_kiri, "tanggal_kiri": tanggal_kiri,
        "lokasi_kanan": lokasi_kanan, "tanggal_kanan": tanggal_kanan,
    }
    table = doc.tables[0]
    for row_idx, col_idx, para_idx, key in RED_FIELDS:
        cell = table.cell(row_idx, col_idx)
        if para_idx >= len(cell.paragraphs):
            continue
        para = cell.paragraphs[para_idx]
        # Hanya mengganti run yang memang berwarna merah. Semua teks non-merah dibiarkan.
        red_runs = []
        for run in para.runs:
            color = run.font.color
            if color and color.type is not None and color.rgb is not None and str(color.rgb).upper() in {"EE0000", "FF0000", "C00000", "E60000"}:
                red_runs.append(run)
        if red_runs:
            red_runs[0].text = values[key]
            for run in red_runs[1:]:
                run.text = ""
        # Jika run merah tidak terdeteksi karena perbedaan format, jangan sentuh paragrafnya.
    output = BytesIO()
    doc.save(output)
    st.success("Dokumen Word selesai dibuat. Teks selain jawaban merah dipertahankan dari template.")
    st.download_button("Unduh SPPD (.docx)", output.getvalue(), file_name="SPPD_KONI.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with st.expander("Catatan"):
    st.write("Versi ini membatasi formulir pada delapan jawaban merah yang terdeteksi di template. Jika ada tulisan merah lain yang ingin dijadikan input, pemetaan dapat ditambahkan tanpa mengubah bagian hitam/biru.")
