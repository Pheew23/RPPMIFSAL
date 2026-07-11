import streamlit as st
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import re
import time

# --- KONFIGURASI SISTEM & UI ---
st.set_page_config(
    page_title="MI MIFTAHUSSALAM - Generator Modul Ajar KBC",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOM UNTUK UI/UX & ANIMASI ---
# Ini hanya mengubah tampilan, tidak mengubah logika kode sama sekali.
st.markdown("""
<style>
    /* Import Font Modern */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    /* Global Font & Background */
    .stApp {
        font-family: 'Poppins', sans-serif;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        color: #2d3436;
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
        box-shadow: 2px 0 10px rgba(0,0,0,0.05);
    }

    /* Input Fields Styling */
    .stTextInput > div > div > input, 
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > select {
        border-radius: 10px;
        border: 1px solid #dfe6e9;
        background-color: #ffffff;
        transition: all 0.3s ease;
        box-shadow: 0 2px 5px rgba(0,0,0,0.02);
    }

    .stTextInput > div > div > input:focus, 
    .stTextArea > div > div > textarea:focus {
        border-color: #00b894;
        box-shadow: 0 0 8px rgba(0, 184, 148, 0.3);
        transform: translateY(-2px);
    }

    /* Button Primary Styling & Animation */
    .stButton > button {
        background: linear-gradient(90deg, #00b894 0%, #00cec9 100%);
        color: white;
        border: none;
        border-radius: 50px;
        padding: 12px 24px;
        font-weight: 600;
        font-size: 16px;
        letter-spacing: 0.5px;
        box-shadow: 0 4px 15px rgba(0, 184, 148, 0.4);
        transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        width: 100%;
    }

    .stButton > button:hover {
        transform: scale(1.05) translateY(-3px);
        box-shadow: 0 8px 25px rgba(0, 184, 148, 0.6);
    }

    .stButton > button:active {
        transform: scale(0.95);
    }

    /* Title Styling */
    h1 {
        text-align: center;
        background: -webkit-linear-gradient(45deg, #00b894, #0984e3);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
        margin-bottom: 10px;
        animation: fadeInDown 1s ease-out;
    }

    /* Card Effect for Containers */
    .css-1r6slb0 {
        background-color: rgba(255, 255, 255, 0.6);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1);
    }

    /* Status Message Styling */
    .stStatus {
        border-radius: 15px;
    }

    /* Animations */
    @keyframes fadeInDown {
        from { opacity: 0; transform: translateY(-20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.02); }
        100% { transform: scale(1); }
    }

    /* Hide default footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Konstanta API
# PENTING: Dalam produksi, gunakan st.secrets("NVIDIA_API_KEY")
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response_kbc(prompt, system_instruction):
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
        "top_p": 0.9,
        "max_tokens": 6096,
        "stream": False
    }

    max_retries = 1
    attempt = 0

    while attempt < max_retries:
        try:
            # Status box dengan animasi custom via CSS sudah diterapkan global
            with st.status(f"❤️ AI (Lagos AI 9.1) Sedang Merenung... (Sabar ya, sedang memikirkan cinta)", expanded=True) as status:
                # Timeout 300 detik
                response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=300)

                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ Kerjaan Saya Beres! Dokumen siap cinta.", state="complete", expanded=False)
                    return content
                else:
                    st.error(f"API Error: {response.status_code}")
                    status.update(label="❌ Gagal", state="error")
                    return None
        except requests.exceptions.ReadTimeout:
            st.error("Timeout: Server AI terlalu sibuk. Silakan coba lagi nanti kalau gak males.")
            return None
        except Exception as e:
            st.error(f"Koneksi Error: {str(e)}")
            return None

    return None

def set_font_safe(run, font_name='Times New Roman', size=12, bold=False):
    if run is None: return
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def clean_markdown_symbols(text):
    if not text: return ""
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'^#+\s*', '', text)
    text = re.sub(r'^[-*]\s+', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def create_word_doc_kbc(content, doc_type, school_data):
    try:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        style.paragraph_format.space_after = Pt(6)

        # --- 1. HEADER MADRASAH ---
        if doc_type in ["RPP", "Modul Ajar", "ATP", "CP", "Prota", "Promes", "LKPD"]:
            # Baris 1
            p1 = doc.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(f"PEMERINTAH {school_data['kabupaten'].upper()}\n")
            set_font_safe(r1, size=14, bold=True)

            # Baris 2
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(f"KEMENTRIAN AGAMA ISLAM\nMADRASAH {school_data['jenis'].upper()} {school_data['nama_madrasah'].upper()}\n")
            set_font_safe(r2, size=12, bold=True)

            # Baris 3
            p3 = doc.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r3 = p3.add_run(f"Alamat: {school_data['alamat']} | Telp: {school_data['telp']}")
            set_font_safe(r3, size=10)

            # Garis
            p_line = doc.add_paragraph()
            r_line = p_line.add_run("_" * 70)
            set_font_safe(r_line, size=10)
            p_line.paragraph_format.space_after = Pt(0)
            p_line.paragraph_format.space_before = Pt(0)

            # Judul Dokumen
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_title.paragraph_format.space_before = Pt(12)

            r_title_1 = p_title.add_run(f"{doc_type.upper()}\n")
            set_font_safe(r_title_1, size=14, bold=True)
            r_title_1.underline = True

            r_title_2 = p_title.add_run(f"Materi: {school_data['mapel']} - {school_data['kelas']}")
            set_font_safe(r_title_2, size=14, bold=True)
            r_title_2.underline = True

            doc.add_paragraph() # Spacer

        # --- 2. PARSE KONTEN ---
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # Deteksi Tabel
            if stripped.startswith('|') and not re.match(r'^\|[\s\-:]+\|$', stripped):
                in_table_mode = True

            if in_table_mode:
                if re.match(r'^\|[\s\-:]+\|$', stripped):
                    i += 1
                    continue

                if stripped.startswith('|') and stripped.endswith('|'):
                    raw_cells = stripped.split('|')[1:-1]
                    cells = [clean_markdown_symbols(c.strip()) for c in raw_cells]
                    table_buffer.append(cells)
                    i += 1
                    continue
                else:
                    # Render Tabel
                    if table_buffer:
                        try:
                            max_cols = max(len(row) for row in table_buffer)
                            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
                            table.style = 'Table Grid'
                            table.autofit = False
                            for col in table.columns:
                                col.width = Inches(6.0 / max_cols)

                            for r_idx, row_data in enumerate(table_buffer):
                                row = table.rows[r_idx]
                                for c_idx, cell_text in enumerate(row_data):
                                    if c_idx < len(row.cells):
                                        cell = row.cells[c_idx]
                                        cell.text = cell_text
                                        for paragraph in cell.paragraphs:
                                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                            for run in paragraph.runs:
                                                set_font_safe(run, size=11, bold=(r_idx==0))
                        except Exception as e:
                            doc.add_paragraph(f"[Error Tabel: {str(e)}]")

                        table_buffer = []
                        in_table_mode = False

            if not in_table_mode:
                if not stripped:
                    doc.add_paragraph()
                    i += 1
                    continue

                # Buat Paragraf Baru
                p = doc.add_paragraph()

                is_bold = False
                font_size = 12
                final_text = stripped
                is_list = False

                if stripped.startswith('# '):
                    final_text = clean_markdown_symbols(stripped[2:])
                    is_bold = True
                    font_size = 14
                    p.paragraph_format.space_before = Pt(12)
                elif stripped.startswith('## '):
                    final_text = clean_markdown_symbols(stripped[3:])
                    is_bold = True
                    font_size = 13
                    p.paragraph_format.space_before = Pt(6)

                if stripped.startswith('- ') or stripped.startswith('* '):
                    is_list = True
                    p.style = 'List Bullet'
                    final_text = clean_markdown_symbols(stripped[2:])
                else:
                    final_text = clean_markdown_symbols(stripped)

                # Tambahkan ke Paragraf
                r = p.add_run(final_text)
                set_font_safe(r, size=font_size, bold=is_bold)

            i += 1

        # Flush sisa tabel
        if in_table_mode and table_buffer:
            try:
                max_cols = max(len(row) for row in table_buffer)
                table = doc.add_table(rows=len(table_buffer), cols=max_cols)
                table.style = 'Table Grid'
                for r_idx, row_data in enumerate(table_buffer):
                    row = table.rows[r_idx]
                    for c_idx, cell_text in enumerate(row_data):
                         if c_idx < len(row.cells):
                            cell = row.cells[c_idx]
                            cell.text = cell_text
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    set_font_safe(run, size=11, bold=(r_idx==0))
            except Exception as e:
                st.error(f"Error tabel akhir: {e}")

        # --- 3. TANDA TANGAN ---
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n")
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

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"❌ Error Fatal: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

# --- UI STREAMLIT (Tampilan Diperbarui) ---

# Header dengan animasi
st.title("❤️ MI MIFTAHUSSALAM GENERATED Modul Ajar")
st.markdown("""
<div style='text-align: center; color: #636e72; margin-bottom: 20px;'>
    <p>✨ Bug <code>AttributeError</code> sudah diperbaiki. Menggunakan Qwen 397B dengan cinta sepenuh hati.</p>
</div>
""", unsafe_allow_html=True)

# Layout Kolom untuk memisahkan Sidebar dan Konten Utama dengan lebih rapi
col_sidebar, col_main = st.columns([1, 3])

with col_sidebar:
    with st.container():
        st.markdown("### 🏫 Data Madrasah")
        nama_madrasah = st.text_input("Nama Madrasah", "MI MIFTAHUSSALAM")
        jenis = st.selectbox("Jenjang", ["MI/SD", "MTs/SMP", "MA/SMA"])
        kabupaten = st.text_input("Kabupaten/Kota", "Kota Bogor")
        alamat = st.text_area("Alamat", "Jl. Rimba Mulya II No.46,Pasirmulya, Bogor Barat")
        telp = st.text_input("Telp", "")

        st.markdown("### 👤 Data Guru")
        kepala_madrasah = st.text_input("Kepala Madrasah", "Drs.Andi Supriadi")
        nip_kepala = st.text_input("NIP Kepala", "-")
        nama_guru = st.text_input("Nama Guru", "..")
        nip_guru = st.text_input("NIP Guru", "-")
        kota = st.text_input("Kota", "Bogor")
        tanggal_buat = st.date_input("Tanggal")

with col_main:
    with st.container():
        st.markdown("### 📝 Konten Pembelajaran")
        # Layout grid kecil untuk input konten
        c1, c2 = st.columns(2)
        with c1:
            doc_type = st.selectbox("Jenis Dokumen", ["Modul Ajar", "RPP", "ATP", "CP", "Prota", "Promes", "LKPD"])
            mapel = st.text_input("Mata Pelajaran", "Pendidikan Agama Islam & Budi Pekerti")
        with c2:
            kelas = st.text_input("Kelas / Fase", "VI / Fase C")
            tahun_ajaran = st.text_input("Tahun Ajaran", "2026/2027")

        materi = st.text_area("Topik / Materi", "Akhlak Terpuji: Kasih Sayang Terhadap Sesama", height=150)

        # Tombol dengan animasi CSS
        btn = st.button("✨ Generate Dokumen KBC 2026", type="primary", use_container_width=True)

        if btn:
            if not nama_madrasah or not materi:
                st.warning("⚠️ Mohon lengkapi data madrasah dan materi ya.")
            else:
                data = {
                    "nama_madrasah": nama_madrasah, "jenis": jenis, "kabupaten": kabupaten,
                    "alamat": alamat, "telp": telp, "kepala_madrasah": kepala_madrasah,
                    "nip_kepala": nip_kepala, "nama_guru": nama_guru, "nip_guru": nip_guru,
                    "kota": kota, "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
                    "mapel": mapel, "kelas": kelas, "tahun_ajaran": tahun_ajaran
                }

                sys_prompt = """
                Anda adalah Ahli Kurikulum Berbasis Cinta (KBC) 2026. 
                Fokus: Cinta kepada Allah dan Rasul-Nya, Cinta kepada Ilmu, Cinta Kepada Diri Sendiri, Cinta Kepada Sesama, Cinta Kepada Alam dan Lingkungan.
                Format: Markdown, Gunakan Tabel untuk langkah pembelajaran, List untuk tujuan.
                Jangan ada kode blok (```). Langsung isi dokumen.
                Jika CP: Fokus pada elemen dan kata kunci operasional.
                Penyusun:ganti sesuaikan dengan input nama guru.
                satuan pendidikan sesuaikan dengan input nama madrasah.
                """

                user_prompt = f"Buat {doc_type} untuk {mapel} kelas {kelas} topik: {materi}. Sertakan tabel kegiatan dengan kolom 'Nilai Cinta'."

                content = get_ai_response_kbc(user_prompt, sys_prompt)

                if content:
                    st.success("🎉 Sedang membuat file Word dengan penuh kasih sayang...")
                    buffer = create_word_doc_kbc(content, doc_type, data)
                    if buffer:
                        fname = f"KBC2026_{doc_type}_{mapel}_{kelas}.docx"
                        st.download_button(
                            label="📥 Download Word Sekarang", 
                            data=buffer, 
                            file_name=fname, 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
