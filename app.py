from io import BytesIO
from pathlib import Path
import streamlit as st
from docx import Document

st.set_page_config(page_title="Pengisian SPPD KONI", page_icon="📄", layout="wide")
TEMPLATE = Path(__file__).parent / "template_sppd.docx"

st.title("Pengisian SPPD — KONI Kabupaten Tangerang")
st.write("Isi hanya jawaban yang ditandai warna merah pada dokumen Word. Pertanyaan/label biru dan teks baku hitam tetap mengikuti template asli.")

def is_red(run):
    try:
        rgb = run.font.color.rgb
        return rgb is not None and str(rgb).upper() in {"EE0000", "FF0000", "C00000", "E60000"}
    except Exception:
        return False

def replace_red_paragraph(paragraph, value):
    """Replace only red-colored runs in a paragraph; preserve all non-red text and paragraph formatting."""
    red_runs = [run for run in paragraph.runs if is_red(run)]
    if not red_runs:
        return False
    red_runs[0].text = value or ""
    for run in red_runs[1:]:
        run.text = ""
    return True

def replace_red_cell_paragraph(doc, row, col, paragraph_index, value):
    table = doc.tables[0]
    cell = table.rows[row].cells[col]
    if paragraph_index >= len(cell.paragraphs):
        return False
    return replace_red_paragraph(cell.paragraphs[paragraph_index], value)

if not TEMPLATE.exists():
    st.error("File template_sppd.docx tidak ditemukan. Pastikan file ini berada di folder yang sama dengan app.py.")
    st.stop()

with st.form("sppd_form"):
    st.subheader("I. Data perjalanan")
    col1, col2 = st.columns(2)
    with col1:
        asal = st.text_input("Berangkat dari", "Tangerang")
        keperluan = st.text_input("Ke / keperluan", "Try Out Catur")
        tanggal_berangkat = st.text_input("Tanggal keberangkatan", "17 Oktober 2026")
    with col2:
        tempat_tiba = st.text_input("Tiba di", "Tangerang")
        tanggal_tiba = st.text_input("Tanggal tiba", "18 Oktober 2026")

    st.subheader("II. Lokasi perjalanan")
    col3, col4 = st.columns(2)
    with col3:
        lokasi_kiri = st.text_input("Lokasi tiba — kolom kiri", "Piset Square. Jl. Pelajar Pejuang No. 119 Bandung")
        tanggal_lokasi_kiri = st.text_input("Tanggal — kolom kiri", "17 Oktober 2026")
    with col4:
        lokasi_kanan = st.text_input("Lokasi tiba — kolom kanan", "Piset Square. Jl. Pelajar Pejuang No. 119 Bandung")
        tanggal_lokasi_kanan = st.text_input("Tanggal — kolom kanan", "18 Oktober 2026")

    submitted = st.form_submit_button("Buat dokumen Word", type="primary")

if submitted:
    doc = Document(str(TEMPLATE))
    # Pemetaan hanya ke paragraf yang pada dokumen asli mengandung teks merah.
    # Semua label/pertanyaan dan teks hitam tidak disentuh.
    targets = [
        (0, 3, 0, asal),
        (0, 3, 1, keperluan),
        (0, 3, 2, tanggal_berangkat),  # Tanggal keberangkatan: input terpisah
        (0, 5, 0, tempat_tiba),
        (0, 5, 1, tanggal_tiba),
        (1, 3, 0, lokasi_kiri),
        (1, 3, 2, tanggal_lokasi_kiri),
        (1, 5, 0, lokasi_kanan),
        (1, 5, 1, tanggal_lokasi_kanan),
    ]
    for row, col, paragraph_index, value in targets:
        replace_red_cell_paragraph(doc, row, col, paragraph_index, value)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    st.success("Dokumen selesai dibuat. Unduh dan periksa hasilnya di Microsoft Word sebelum dicetak.")
    st.download_button(
        "Unduh SPPD (.docx)",
        data=output.getvalue(),
        file_name="SPPD_KONI.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

with st.expander("Cara menjalankan aplikasi"):
    st.markdown("""
1. Unggah `app.py`, `requirements.txt`, dan `template_sppd.docx` ke repository GitHub yang sama.
2. Di Streamlit Community Cloud, pilih repository tersebut dan set **Main file path** ke `app.py`.
3. Aplikasi akan memakai `template_sppd.docx` sebagai format asli.
""")
