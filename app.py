import streamlit as st
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import re
import time

# --- KONFIGURASI SISTEM ---
st.set_page_config(
    page_title="Generator Dokumen Madrasah (Clean Text)",
    page_icon="✨",
    layout="wide"
)

# Konstanta API
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response_robust(prompt, system_instruction):
    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7,
        "top_p": 0.7,
        "max_tokens": 4096,
        "stream": False
    }

    max_retries = 2
    attempt = 0

    while attempt < max_retries:
        try:
            with st.status(f"🧠 AI Sedang Menyusun Dokumen (Tunggu...)...", expanded=True) as status:
                response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=None)

                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ Dokumen Disusun!", state="complete", expanded=False)
                    return content
                else:
                    st.warning(f"API Error: {response.status_code}")
                    status.update(label="❌ Gagal", state="error")
                    attempt += 1
                    time.sleep(2)
        except Exception as e:
            st.error(f"Koneksi Error: {str(e)}")
            attempt += 1
            time.sleep(2)

    return None

def set_font_safe(run, font_name='Times New Roman', size=12, bold=False):
    """Helper function untuk set font dengan benar"""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def clean_markdown_symbols(text):
    """Menghapus simbol markdown yang tidak diinginkan dari teks biasa"""
    # Hapus bintang yang mengapit teks tebal/miring (contoh: **teks** atau *teks*)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    # Hapus tanda pagar berlebih di awal jika lolos deteksi heading
    text = re.sub(r'^#+\s*', '', text)
    return text.strip()

def create_word_doc_clean(content, doc_type, school_data):
    """Versi final: Tanpa bintang, tanpa simbol markdown sisa"""
    try:
        doc = Document()

        # Set Default Style
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

        # --- 1. KOPI SURAT ---
        if doc_type in ["RPP", "Modul Ajar"]:
            p1 = doc.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(f"PEMERINTAH KABUPATEN {school_data['kabupaten'].upper()}\n")
            set_font_safe(r1, size=14, bold=True)

            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(f"DINAS PENDIDIKAN DAN KEBUDAYAAN\nMADRASAH {school_data['jenis'].upper()} {school_data['nama_madrasah'].upper()}\n")
            set_font_safe(r2, size=12, bold=True)

            p3 = doc.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r3 = p3.add_run(f"Alamat: {school_data['alamat']} | Telp: {school_data['telp']}")
            set_font_safe(r3, size=10)

            p_line = doc.add_paragraph()
            r_line = p_line.add_run("_" * 70)
            set_font_safe(r_line, size=10)

            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(f"{doc_type.upper()}\n{school_data['mapel']} - {school_data['kelas']}")
            set_font_safe(r_title, size=14, bold=True)
            r_title.underline = True

            doc.add_paragraph()

        # --- 2. PARSE KONTEN & TABEL ---
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False

        for line in lines:
            stripped = line.strip()

            # Skip separator tabel markdown
            if re.match(r'^\|[\s\-:]+\|$', stripped):
                continue

            # Deteksi Tabel
            if stripped.startswith('|') and stripped.endswith('|'):
                in_table_mode = True
                # Ambil isi sel, bersihkan dari simbol markdown
                raw_cells = stripped.split('|')[1:-1]
                cells = []
                for c in raw_cells:
                    clean_c = clean_markdown_symbols(c.strip())
                    if clean_c: cells.append(clean_c)

                if cells:
                    table_buffer.append(cells)
            else:
                # Render Tabel jika mode tabel berakhir
                if in_table_mode and table_buffer:
                    if len(table_buffer) > 1: 
                        try:
                            max_cols = max(len(row) for row in table_buffer)
                            for row in table_buffer:
                                while len(row) < max_cols: row.append("")

                            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
                            table.style = 'Table Grid'

                            for i, row_data in enumerate(table_buffer):
                                row = table.rows[i]
                                for j, cell_text in enumerate(row_data):
                                    cell = row.cells[j]
                                    cell.text = str(cell_text) # Teks sudah bersih
                                    # Pastikan font di dalam tabel juga bersih
                                    for paragraph in cell.paragraphs:
                                        for run in paragraph.runs:
                                            set_font_safe(run, size=11, bold=(i==0))
                        except Exception as e:
                            doc.add_paragraph(f"[Error Tabel: {str(e)}]")

                    table_buffer = []
                    in_table_mode = False

                if not stripped:
                    doc.add_paragraph()
                    continue

                # Handle Headings & List (DENGAN PEMBERSIHAN SIMBOL)
                p = doc.add_paragraph()
                final_text = ""
                is_bold = False
                font_size = 12
                is_list = False

                # Cek Heading
                if stripped.startswith('# '):
                    final_text = clean_markdown_symbols(stripped[2:])
                    is_bold = True
                    font_size = 14
                elif stripped.startswith('## '):
                    final_text = clean_markdown_symbols(stripped[3:])
                    is_bold = True
                    font_size = 13
                # Cek List (Bullet points)
                elif stripped.startswith('- ') or stripped.startswith('* '):
                    is_list = True
                    # Hapus simbol - atau * di depan
                    temp_text = stripped[2:]
                    final_text = clean_markdown_symbols(temp_text)
                else:
                    final_text = clean_markdown_symbols(stripped)

                # Terapkan ke dokumen
                if is_list:
                    p.style = 'List Bullet' # Gunakan bullet asli Word

                r = p.add_run(final_text)
                set_font_safe(r, size=font_size, bold=is_bold)

        # Flush sisa tabel
        if in_table_mode and table_buffer and len(table_buffer) > 1:
            max_cols = max(len(row) for row in table_buffer)
            for row in table_buffer:
                while len(row) < max_cols: row.append("")
            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
            table.style = 'Table Grid'
            for i, row_data in enumerate(table_buffer):
                for j, cell_text in enumerate(row_data):
                    cell = table.rows[i].cells[j]
                    cell.text = str(cell_text)
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            set_font_safe(run, size=11, bold=(i==0))

        # --- 3. TANDA TANGAN ---
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n\n")
            sig_table = doc.add_table(rows=1, cols=2)
            sig_table.autofit = False

            cell_kiri = sig_table.cell(0, 0)
            cell_kiri.width = Inches(3.0)
            p_kiri = cell_kiri.paragraphs[0]
            p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r_kiri_1 = p_kiri.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
            set_font_safe(r_kiri_1, size=12)

            r_kiri_2 = p_kiri.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")
            set_font_safe(r_kiri_2, size=12, bold=True)

            cell_kanan = sig_table.cell(0, 1)
            cell_kanan.width = Inches(3.0)
            p_kanan = cell_kanan.paragraphs[0]
            p_kanan.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r_kanan_1 = p_kanan.add_run(f"{school_data['kota']}, {school_data['tanggal_buat']}\nGuru Mata Pelajaran\n\n\n\n\n")
            set_font_safe(r_kanan_1, size=12)

            r_kanan_2 = p_kanan.add_run(f"( {school_data['nama_guru']} )\nNIP. {school_data['nip_guru']}")
            set_font_safe(r_kanan_2, size=12, bold=True)

        # Simpan
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"❌ Error Fatal: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

# --- UI STREAMLIT ---

st.title("✨ Generator Dokumen Madrasah (Tanpa Simbol Bintang)")
st.markdown(f"**Model:** `{MODEL_NAME}` | **Output:** Bersih & Rapi")

with st.sidebar:
    st.header("🏫 Data Madrasah")
    nama_madrasah = st.text_input("Nama Madrasah", "MI/MTs/MA Al-Hikmah")
    jenis = st.selectbox("Jenjang", ["MI", "MTs", "MA"])
    kabupaten = st.text_input("Kabupaten/Kota", "Cirebon")
    alamat = st.text_area("Alamat", "Jl. Pendidikan No. 1")
    telp = st.text_input("Telp", "0231-1234567")

    st.subheader("👤 Data Guru")
    kepala_madrasah = st.text_input("Kepala Madrasah", "Drs. H. Ahmad Fulan, M.Pd")
    nip_kepala = st.text_input("NIP Kepala", "19700101 199001 1 001")
    nama_guru = st.text_input("Nama Guru", "Fulan bin Fulan, S.Pd")
    nip_guru = st.text_input("NIP Guru", "19900101 202001 1 001")
    kota = st.text_input("Kota", "Cirebon")
    tanggal_buat = st.date_input("Tanggal")

    st.subheader("📝 Dokumen")
    doc_type = st.selectbox("Jenis", ["RPP", "Modul Ajar", "ATP", "CP", "KKTP", "Prota", "Promes"])
    mapel = st.text_input("Mapel", "Pendidikan Agama Islam")
    kelas = st.text_input("Kelas", "VII / Ganjil")
    materi = st.text_area("Materi", "Akhlak Terpuji", height=100)

    btn = st.button("🚀 Buat Dokumen Bersih", type="primary", use_container_width=True)

if btn:
    if not nama_madrasah or not materi:
        st.warning("Lengkapi data dulu!")
    else:
        data = {
            "nama_madrasah": nama_madrasah, "jenis": jenis, "kabupaten": kabupaten,
            "alamat": alamat, "telp": telp, "kepala_madrasah": kepala_madrasah,
            "nip_kepala": nip_kepala, "nama_guru": nama_guru, "nip_guru": nip_guru,
            "kota": kota, "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
            "mapel": mapel, "kelas": kelas
        }

        sys_prompt = f"Buat {doc_type} lengkap format markdown. Gunakan tabel | col | col |. Gunakan - untuk list. Jangan ada pembuka chat."
        user_prompt = f"Buat {doc_type} {mapel} kelas {kelas} materi: {materi}"

        content = get_ai_response_robust(user_prompt, sys_prompt)

        if content:
            st.success("Konten AI Siap! Sedang membersihkan simbol...")
            buffer = create_word_doc_clean(content, doc_type, data)
            if buffer:
                fname = f"{doc_type}_{mapel}_{kelas}_Bersih.docx"
                st.download_button("📥 Download Word (Tanpa Bintang)", buffer, fname, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                st.info("💡 File yang diunduh sudah bebas dari simbol *, #, dan - yang tidak perlu. List otomatis menjadi bullet point Word.")
