import streamlit as st
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import re
import time
import json # Tambahan untuk Lottie

# --- FUNGSI ANIMASI LOTTIE ---
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

# --- KONFIGURASI SISTEM ---
st.set_page_config(
    page_title="MI MIFTAHUSSALAM - KBC 2026",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- INJEKSI CSS KUSTOM (ELEGAN & MODERN) ---
st.markdown("""
    <style>
    /* Mengimpor font elegan */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Efek Glassmorphism untuk Card */
    .stApp {
        background: transparent;
    }
    
    .main-card {
        background: rgba(255, 255, 255, 0.05);
        border-radius: 20px;
        padding: 30px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.15);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        margin-bottom: 20px;
    }

    /* Styling tombol utama agar lebih beranimasi */
    div.stButton > button:first-child {
        background: linear-gradient(45deg, #FF4B4B, #FF8E8E);
        color: white;
        border: none;
        padding: 15px 30px;
        border-radius: 12px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(255, 75, 75, 0.3);
    }
    
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255, 75, 75, 0.5);
    }

    /* Sidebar Styling */
    .css-1d391kg {
        background-color: #0E1117;
    }
    </style>
    """, unsafe_allow_html=True)

# Konstanta API (Tetap)
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER (LOGIKA ASLI - TIDAK DIRUBAH) ---

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
            with st.status(f"❤️ AI (Lagos AI 9.1) Sedang Merenung... (Tunggu ya sekitar 15-45 hari)", expanded=True) as status:
                response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=300)
                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ Kerjaan Saya Beres ya!", state="complete", expanded=False)
                    return content
                else:
                    st.error(f"API Error: {response.status_code}")
                    status.update(label="❌ Gagal", state="error")
                    return None
        except requests.exceptions.ReadTimeout:
            st.error("Timeout: Server AI terlalu sibuk. Silakan coba lagi nanti.")
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
            p1 = doc.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(f"PEMERINTAH {school_data['kabupaten'].upper()}\n")
            set_font_safe(r1, size=14, bold=True)
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(f"KEMENTRIAN AGAMA ISLAM\nMADRASAH {school_data['jenis'].upper()} {school_data['nama_madrasah'].upper()}\n")
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
            p_title.paragraph_format.space_before = Pt(12)
            r_title_1 = p_title.add_run(f"{doc_type.upper()}\n")
            set_font_safe(r_title_1, size=14, bold=True)
            r_title_1.underline = True
            r_title_2 = p_title.add_run(f"Materi: {school_data['mapel']} - {school_data['kelas']}")
            set_font_safe(r_title_2, size=14, bold=True)
            r_title_2.underline = True
            doc.add_paragraph()

        # --- 2. PARSE KONTEN ---
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
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
                    if table_buffer:
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
                        except: pass
                        table_buffer = []
                        in_table_mode = False
            if not in_table_mode:
                if not stripped:
                    doc.add_paragraph()
                    i += 1
                    continue
                p = doc.add_paragraph()
                is_bold, font_size = False, 12
                final_text = stripped
                if stripped.startswith('# '):
                    final_text = clean_markdown_symbols(stripped[2:])
                    is_bold, font_size = True, 14
                elif stripped.startswith('## '):
                    final_text = clean_markdown_symbols(stripped[3:])
                    is_bold, font_size = True, 13
                if stripped.startswith('- ') or stripped.startswith('* '):
                    p.style = 'List Bullet'
                    final_text = clean_markdown_symbols(stripped[2:])
                else:
                    final_text = clean_markdown_symbols(stripped)
                r = p.add_run(final_text)
                set_font_safe(r, size=font_size, bold=is_bold)
            i += 1
        
        # Sisa tabel & Tanda Tangan
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n")
            sig_table = doc.add_table(rows=1, cols=2)
            cell_kiri = sig_table.cell(0, 0)
            p_kiri = cell_kiri.paragraphs[0]
            p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_kiri_1 = p_kiri.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
            set_font_safe(r_kiri_1, size=12)
            r_kiri_2 = p_kiri.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")
            set_font_safe(r_kiri_2, size=12, bold=True)
            cell_kanan = sig_table.cell(0, 1)
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
        return None

# --- UI STREAMLIT (NEW DESIGN) ---

# Memuat animasi
from streamlit_lottie import st_lottie
lottie_edu = load_lottieurl("https://lottie.host/e2d09c31-7290-482a-9957-c59800742f1f/o86M61f87s.json")
lottie_heart = load_lottieurl("https://assets5.lottiefiles.com/packages/lf20_v9t630.json")

# Layouting dengan kolom untuk Header
head_col1, head_col2 = st.columns([0.8, 0.2])
with head_col1:
    st.title("❤️ MI MIFTAHUSSALAM Workspace")
    st.subheader("Kurikulum Berbasis Cinta (KBC) 2026")
    st.write("Sistem cerdas penghasil Modul Ajar otomatis yang dirancang dengan kasih sayang.")

with head_col2:
    if lottie_heart:
        st_lottie(lottie_heart, height=150, key="heart")

st.divider()

# SIDEBAR MODERISASI
with st.sidebar:
    if lottie_edu:
        st_lottie(lottie_edu, height=120)
    
    st.header("⚙️ Konfigurasi")
    
    with st.expander("🏫 Data Madrasah", expanded=False):
        nama_madrasah = st.text_input("Nama Madrasah", "MI MIFTAHUSSALAM")
        jenis = st.selectbox("Jenjang", ["MI/SD", "MTs/SMP", "MA/SMA"])
        kabupaten = st.text_input("Kabupaten/Kota", "Kota Bogor")
        alamat = st.text_area("Alamat", "Jl. Rimba Mulya II No.46, Pasirmulya, Bogor Barat")
        telp = st.text_input("Telp", "")

    with st.expander("👤 Data Personel", expanded=False):
        kepala_madrasah = st.text_input("Kepala Madrasah", "Drs.Andi Supriadi")
        nip_kepala = st.text_input("NIP Kepala", "-")
        nama_guru = st.text_input("Nama Guru", "..")
        nip_guru = st.text_input("NIP Guru", "-")
        kota = st.text_input("Kota", "Bogor")
        tanggal_buat = st.date_input("Tanggal")

    st.info("Aplikasi ini mendukung Mode Gelap/Terang secara otomatis.")

# AREA KERJA UTAMA
col_main, col_preview = st.columns([0.6, 0.4], gap="large")

with col_main:
    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    st.markdown("### 📝 Detail Pembelajaran")
    
    doc_type = st.selectbox("Jenis Dokumen", ["Modul Ajar", "RPP", "ATP", "CP", "Prota", "Promes", "LKPD"])
    
    c1, c2 = st.columns(2)
    with c1:
        mapel = st.text_input("Mata Pelajaran", "Pendidikan Agama Islam")
        kelas = st.text_input("Kelas / Fase", "VI / Fase C")
    with c2:
        tahun_ajaran = st.text_input("Tahun Ajaran", "2026/2027")
    
    materi = st.text_area("Topik Utama / Materi", "Akhlak Terpuji: Kasih Sayang Terhadap Sesama", height=120)
    
    btn = st.button("✨ JALANKAN GENERASI CERDAS", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col_preview:
    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    st.markdown("### 🔍 Info Kurikulum KBC")
    st.write("Kurikulum Berbasis Cinta menekankan 5 Pilar Utama:")
    st.markdown("""
    - 💖 Cinta kepada Allah & Rasul
    - 📚 Cinta kepada Ilmu
    - 🧘 Cinta kepada Diri Sendiri
    - 🤝 Cinta kepada Sesama
    - 🌿 Cinta kepada Lingkungan
    """)
    st.divider()
    st.caption("Pastikan API Key Anda aktif untuk menghindari error.")
    st.markdown('</div>', unsafe_allow_html=True)

# LOGIKA GENERASI (TETAP SAMA)
if btn:
    if not nama_madrasah or not materi:
        st.warning("⚠️ Mohon isi data madrasah dan topik materi.")
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
        Format: Markdown, Gunakan Tabel untuk langkah pembelajaran, List untuk tujuan. Sertakan 'Nilai Cinta' di tiap tabel.
        Jangan ada kode blok (```). Langsung isi dokumen.
        """

        user_prompt = f"Buat {doc_type} untuk {mapel} kelas {kelas} topik: {materi}. Sertakan tabel kegiatan dengan kolom 'Nilai Cinta'."

        content = get_ai_response_kbc(user_prompt, sys_prompt)

        if content:
            st.balloons()
            st.success("✅ Konten Berhasil Dibuat!")
            buffer = create_word_doc_kbc(content, doc_type, data)
            if buffer:
                fname = f"KBC2026_{doc_type}_{mapel}.docx"
                st.download_button(
                    label="📥 DOWNLOAD DOKUMEN WORD",
                    data=buffer,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
