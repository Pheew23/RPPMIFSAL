import streamlit as st
import requests
import os
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import markdown
import re

# --- KONFIGURASI ---
load_dotenv()
API_KEY = os.getenv("NVIDIA_API_KEY")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- CSS ---
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #FF4B4B; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #555; text-align: center; margin-bottom: 2rem; }
    .stButton>button { background-color: #FF4B4B; color: white; font-weight: bold; width: 100%; border: none; padding: 10px; border-radius: 5px; }
    .status-badge { display: inline-block; padding: 5px 10px; border-radius: 15px; font-size: 0.8rem; font-weight: bold; margin-bottom: 10px; background-color: #d4edda; color: #155724; }
</style>
""", unsafe_allow_html=True)

# --- FUNGSI PEMBERSIH TEKS & PEMBUAT WORD ---
def create_clean_word_document(content, guru_name, nip_guru, kepala_madrasah):
    doc = Document()

    # Setup Style Font (Arial/Times New Roman)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # 1. Judul Utama
    title = doc.add_heading('PERANGKAT PEMBELAJARAN KBC 2026', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    # 2. Info Guru
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run(f"Guru Pengampu: {guru_name}\n").bold = True
    if nip_guru:
        info.add_run(f"NIP: {nip_guru}")

    doc.add_paragraph() # Spasi

    # 3. Proses Konten AI (Menghapus simbol *, #, dll)
    # Kita pecah per baris untuk dianalisis
    lines = content.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            doc.add_paragraph() # Spasi antar paragraf
            continue

        # Deteksi Heading (## Judul)
        if line.startswith('## '):
            text = line.replace('##', '').replace('**', '').strip()
            h = doc.add_heading(text, level=2)
            h.runs[0].bold = True
        elif line.startswith('### '):
            text = line.replace('###', '').replace('**', '').strip()
            doc.add_heading(text, level=3)
        # Deteksi List (- atau *)
        elif line.startswith('- ') or line.startswith('* '):
            text = line[2:].replace('**', '').strip() # Hapus simbol list dan bold markdown
            p = doc.add_paragraph(text, style='List Bullet')
            # Jika ada teks bold di tengah kalimat (misal: **Konsep**)
            for run in p.runs:
                if '**' in run.text:
                    run.text = run.text.replace('**', '')
                    run.bold = True
        # Paragraf Biasa
        else:
            # Hapus simbol markdown yang tersisa
            clean_text = line.replace('**', '').replace('__', '').strip()
            p = doc.add_paragraph(clean_text)

            # Cek jika ada teks tebal di tengah kalimat
            # Ini sederhana, untuk kompleks butuh regex lebih advance
            if '**' in line:
                parts = line.split('**')
                p.clear()
                for i, part in enumerate(parts):
                    if i % 2 == 1: # Bagian ganjil adalah yang diapit **
                        run = p.add_run(part)
                        run.bold = True
                    else:
                        p.add_run(part)

    # 4. Bagian Tanda Tangan
    doc.add_page_break()
    doc.add_paragraph("Mengetahui,", style='Normal')

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False

    # Kolom Kiri (Guru)
    cell_guru = table.cell(0, 0)
    p_guru = cell_guru.paragraphs[0]
    p_guru.add_run(f"Guru Mata Pelajaran,\n\n\n\n\n( {guru_name} )\nNIP. {nip_guru if nip_guru else '-'}")

    # Kolom Kanan (Kepala Madrasah)
    cell_kepala = table.cell(0, 1)
    p_kepala = cell_kepala.paragraphs[0]
    p_kepala.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_kepala.add_run(f"Kepala Madrasah,\n\n\n\n\n( {kepala_madrasah} )\nNIP. ...........................")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- FUNGSI UTAMA ---
def main():
    if not API_KEY:
        st.error("🚫 **ERROR:** File `.env` tidak ditemukan.")
        st.stop()

    st.markdown("<h1 class='main-header'>❤️ Generator RPP KBC 2026</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Format Word Profesional | Lagos AI 9.1</p>", unsafe_allow_html=True)
    st.markdown(f"<div style='text-align: right;'><span class='status-badge'>✅ API Aktif</span></div>", unsafe_allow_html=True)

    with st.form("kbc_form"):
        # Data Guru
        st.subheader("👤 Data Guru & Madrasah")
        col1, col2, col3 = st.columns(3)
        with col1:
            guru_name = st.text_input("Nama Guru", placeholder="Nama Lengkap")
        with col2:
            nip_guru = st.text_input("NIP Guru", placeholder="Nomor Induk Pegawai")
        with col3:
            kepala_madrasah = st.text_input("Nama Kepala Madrasah", placeholder="Nama Kepala Sekolah")

        st.divider()

        # Data Pembelajaran
        st.subheader("📚 Data Pembelajaran")
        col4, col5 = st.columns(2)
        with col4:
            mapel = st.text_input("Mata Pelajaran", placeholder="Contoh: Fiqih")
            materi = st.text_area("Materi Pokok", placeholder="Contoh: Thaharah", height=100)
        with col5:
            kelas_options = [f"Kelas {i}" for i in range(1, 13)]
            kelas = st.selectbox("Kelas", kelas_options)
            semester = st.selectbox("Semester", ["Ganjil", "Genap"])

        st.divider()

        # Komponen Dokumen
        st.subheader("📦 Komponen Dokumen")
        st.write("Pilih dokumen yang ingin dibuat:")
        # PERBAIKAN MULTISELECT: Label dipisah dari widget
        komponen = st.multiselect(
            "", 
            options=["RPP Lengkap", "ATP & CP", "KKTP", "Prota & Promes", "LKPD"],
            default=["RPP Lengkap"]
        )

        submitted = st.form_submit_button("🚀 Buat & Download Word")

    if submitted:
        if not guru_name or not kepala_madrasah:
            st.warning("⚠️ Mohon lengkapi Nama Guru dan Kepala Madrasah.")
        elif not mapel or not materi:
            st.warning("⚠️ Mohon lengkapi Mata Pelajaran dan Materi.")
        else:
            with st.spinner('⏳ AI sedang menyusun dokumen rapi (mohon tunggu)...'):
                prompt = f"""
                Buatkan perangkat pembelajaran KBC 2026 untuk:
                Mapel: {mapel}, Kelas: {kelas}, Materi: {materi}.
                Komponen: {', '.join(komponen)}.

                INSTRUKSI FORMAT PENTING:
                1. Gunakan Markdown untuk struktur.
                2. Gunakan ## untuk Judul Bagian (misal: ## Tujuan Pembelajaran).
                3. Gunakan ### untuk Sub-Judul.
                4. Gunakan - untuk daftar poin.
                5. Gunakan **teks** untuk menebalkan kata penting.
                6. JANGAN gunakan simbol lain yang tidak perlu.
                7. Langsung mulai konten, jangan ada pembuka "Tentu ini dokumen Anda".
                """

                headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
                payload = {
                    "model": MODEL_NAME,
                    "messages": [{"role": "system", "content": "Anda ahli kurikulum KBC 2026."}, {"role": "user", "content": prompt}],
                    "temperature": 0.7,
                    "max_tokens": 4096
                }

                try:
                    response = requests.post(API_URL, headers=headers, json=payload, timeout=300)
                    response.raise_for_status()
                    content = response.json()['choices'][0]['message']['content']

                    # Buat File Word
                    doc_file = create_clean_word_document(content, guru_name, nip_guru, kepala_madrasah)

                    st.success("✅ Dokumen Siap Diunduh!")

                    file_name = f"RPP_{mapel}_{kelas}_{guru_name.replace(' ', '_')}.docx"
                    st.download_button(
                        label="📥 Download File Word (.docx)",
                        data=doc_file,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()
