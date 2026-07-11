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
    page_title="MI MIFTAHUSSALAM - Generator Modul Ajar",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOM DENGAN DUKUNGAN DARK MODE ---
st.markdown("""
<style>
    /* Import Font Inter */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* DEFINISI VARIABEL WARNA (DEFAULT: LIGHT MODE) */
    :root {
        --bg-color: #F8F9FA;
        --sidebar-bg: #FFFFFF;
        --text-main: #212529;
        --text-muted: #6C757D;
        --border-color: #DEE2E6;
        --input-bg: #FFFFFF;
        --card-bg: #FFFFFF;
        --primary-color: #0D6EFD;
        --primary-hover: #0B5ED7;
        --success-color: #198754;
        --shadow: 0 2px 8px rgba(0,0,0,0.08);
        --code-bg: #F8F9FA;
    }

    /* DARK MODE OVERRIDES */
    [data-theme="dark"] {
        --bg-color: #121212;
        --sidebar-bg: #1E1E1E;
        --text-main: #E0E0E0;
        --text-muted: #A0A0A0;
        --border-color: #333333;
        --input-bg: #2C2C2C;
        --card-bg: #1E1E1E;
        --primary-color: #3D8BFD;
        --primary-hover: #5A9FFF;
        --success-color: #2ECC71;
        --shadow: 0 2px 8px rgba(0,0,0,0.3);
        --code-bg: #2C2C2C;
    }

    /* TERAPKAN VARIABEL KE APLIKASI */
    .stApp {
        font-family: 'Inter', sans-serif;
        background-color: var(--bg-color);
        color: var(--text-main);
        transition: background-color 0.3s ease, color 0.3s ease;
    }

    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: var(--sidebar-bg);
        border-right: 1px solid var(--border-color);
        transition: background-color 0.3s ease, border-color 0.3s ease;
    }

    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label {
        color: var(--text-main);
    }

    /* INPUT FIELDS */
    .stTextInput > div > div > input, 
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > select,
    .stDateInput > div > div > input {
        background-color: var(--input-bg);
        border: 1px solid var(--border-color);
        color: var(--text-main);
        border-radius: 8px;
        box-shadow: none;
        transition: all 0.2s ease;
    }

    .stTextInput > div > div > input:focus, 
    .stTextArea > div > div > textarea:focus,
    .stSelectbox > div > div > select:focus {
        border-color: var(--primary-color);
        box-shadow: 0 0 0 3px rgba(13, 110, 253, 0.2);
        outline: none;
    }

    /* BUTTONS */
    .stButton > button {
        background-color: var(--primary-color);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
        box-shadow: var(--shadow);
        transition: all 0.2s ease;
        width: 100%;
    }

    .stButton > button:hover {
        background-color: var(--primary-hover);
        transform: translateY(-1px);
    }

    .stDownloadButton > button {
        background-color: var(--success-color);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        box-shadow: var(--shadow);
        transition: all 0.2s ease;
    }
    .stDownloadButton > button:hover {
        filter: brightness(1.1);
        transform: translateY(-1px);
    }

    /* HEADINGS */
    h1, h2, h3, h4 {
        color: var(--text-main);
        font-weight: 700;
    }

    .stCaption {
        color: var(--text-muted);
    }

    /* TOGGLE SWITCH CUSTOM */
    .theme-toggle-container {
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        background: var(--card-bg);
        padding: 8px 15px;
        border-radius: 50px;
        box-shadow: var(--shadow);
        border: 1px solid var(--border-color);
        display: flex;
        align-items: center;
        gap: 10px;
        cursor: pointer;
        transition: all 0.3s ease;
    }

    .theme-toggle-container:hover {
        transform: scale(1.05);
    }

    .theme-icon {
        font-size: 1.2rem;
    }

    /* ANIMASI FADE IN */
    .element-container {
        animation: fadeIn 0.4s ease-out;
    }

    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* HIDE DEFAULT STREAMLIT FOOTER */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>

<!-- JAVASCRIPT UNTUK TOGGLE MODE -->
<script>
    // Cek apakah ada preferensi tersimpan
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);

    function toggleTheme() {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';

        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);

        // Update icon text (opsional, handled by React rerender usually, but this helps visual feedback)
        const btn = document.getElementById('theme-toggle-btn');
        if(btn) {
            btn.innerHTML = newTheme === 'light' ? '☀️ Mode Terang' : '🌙 Mode Gelap';
        }
    }
</script>
""", unsafe_allow_html=True)

# --- UI TOGGLE BUTTON (HTML Custom) ---
# Kita buat tombol custom yang memanggil fungsi JS di atas
st.markdown("""
<div style="text-align: right;">
    <div id="theme-toggle-btn" class="theme-toggle-container" onclick="toggleTheme()">
        <span class="theme-icon">🌙</span>
        <span style="font-weight: 1000; font-size: 1.5rem; color: var(--text-main);">Mode Gelap</span>
    </div>
</div>
<br>
""", unsafe_allow_html=True)

# Konstanta API
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER (LOGIKA TETAP SAMA) ---

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
            with st.status(f"🤖 AI sedang menyusun dokumen...", expanded=True) as status:
                response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=300)

                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ Dokumen berhasil dibuat!", state="complete", expanded=False)
                    return content
                else:
                    st.error(f"API Error: {response.status_code}")
                    status.update(label="❌ Gagal", state="error")
                    return None
        except requests.exceptions.ReadTimeout:
            st.error("Timeout: Server AI sedang sibuk. Coba lagi sebentar.")
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
            p_line.paragraph_format.space_after = Pt(0)
            p_line.paragraph_format.space_before = Pt(0)

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

                r = p.add_run(final_text)
                set_font_safe(r, size=font_size, bold=is_bold)

            i += 1

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

# --- UI STREAMLIT ---

# Header
col_header, _ = st.columns([3, 1])
with col_header:
    st.title("📚 Generator Modul Ajar KBC")
    st.caption("MI MIFTAHUSSALAM | Sistem Otomatisasi Dokumen Pendidikan")

st.markdown("---")

# Layout Utama
col_sidebar, col_main = st.columns([1, 2.5])

with col_sidebar:
    st.markdown("### 🏫 Data Madrasah")
    nama_madrasah = st.text_input("Nama Madrasah", "MI MIFTAHUSSALAM")
    jenis = st.selectbox("Jenjang", ["MI/SD", "MTs/SMP", "MA/SMA"])
    kabupaten = st.text_input("Kabupaten/Kota", "Kota Bogor")
    alamat = st.text_area("Alamat", "Jl. Rimba Mulya II No.46,Pasirmulya, Bogor Barat", height=80)
    telp = st.text_input("Telp", "")

    st.markdown("### 👤 Data Guru")
    kepala_madrasah = st.text_input("Kepala Madrasah", "Drs.Andi Supriadi")
    nip_kepala = st.text_input("NIP Kepala", "-")
    nama_guru = st.text_input("Nama Guru", "..")
    nip_guru = st.text_input("NIP Guru", "-")
    kota = st.text_input("Kota", "Bogor")
    tanggal_buat = st.date_input("Tanggal")

with col_main:
    st.markdown("### 📝 Konten Pembelajaran")

    c1, c2 = st.columns(2)
    with c1:
        doc_type = st.selectbox("Jenis Dokumen", ["Modul Ajar", "RPP", "ATP", "CP", "Prota", "Promes", "LKPD"])
        mapel = st.text_input("Mata Pelajaran", "Pendidikan Agama Islam & Budi Pekerti")

    with c2:
        kelas = st.text_input("Kelas / Fase", "VI / Fase C")
        tahun_ajaran = st.text_input("Tahun Ajaran", "2026/2027")

    materi = st.text_area("Topik / Materi", "Akhlak Terpuji: Kasih Sayang Terhadap Sesama", height=150)

    st.markdown("<br>", unsafe_allow_html=True)

    btn = st.button("✨ Buat Dokumen Sekarang", type="primary", use_container_width=True)

    if btn:
        if not nama_madrasah or not materi:
            st.warning("⚠️ Mohon lengkapi data madrasah dan materi terlebih dahulu.")
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
            penyusun:ganti sesuai nama guru.
            """

            user_prompt = f"Buat {doc_type} untuk {mapel} kelas {kelas} topik: {materi}. Sertakan tabel kegiatan dengan kolom 'Nilai Cinta'."

            content = get_ai_response_kbc(user_prompt, sys_prompt)

            if content:
                st.success("✅ Dokumen berhasil dibuat! Silakan unduh di bawah ini.")
                buffer = create_word_doc_kbc(content, doc_type, data)
                if buffer:
                    fname = f"KBC2026_{doc_type}_{mapel}_{kelas}.docx"
                    st.download_button(
                        label="📥 Unduh File Word (.docx)", 
                        data=buffer, 
                        file_name=fname, 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
