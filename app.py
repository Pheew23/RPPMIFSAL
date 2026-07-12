import streamlit as st
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import re
import time

# --- FUNGSI ANIMASI LOTTIE ---
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

# --- KONFIGURASI SISTEM ---
st.set_page_config(
    page_title="MI MIFTAHUSSALAM - KBC Premium",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- PREMIUM ADAPTIVE GLASSMORPHISM CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Plus Jakarta Sans', sans-serif; 
    }
    
    .premium-card {
        background: rgba(255, 255, 255, 0.04);
        border-radius: 16px;
        padding: 28px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        margin-bottom: 25px;
    }
    
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #FF4B4B 0%, #FF7676 100%);
        color: white;
        border: none;
        padding: 14px 28px;
        border-radius: 10px;
        font-weight: 600;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 18px rgba(255, 75, 75, 0.25);
    }
    
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 22px rgba(255, 75, 75, 0.45);
    }

    .p-badge {
        background: rgba(255, 75, 75, 0.1);
        color: #FF4B4B;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Konstanta API
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- LOGIKA CORE & HELPER ---
def get_ai_response_kbc(prompt, system_instruction):
    headers = {"Authorization": f"Bearer {NVIDIA_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7, "top_p": 0.9, "max_tokens": 6096, "stream": False
    }
    try:
        with st.status(f"❤️ Lagos AI Sedang Merakit Lembar Konten...", expanded=True) as status:
            response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=300)
            if response.status_code == 200:
                status.update(label="✅ Perumusan Selesai!", state="complete", expanded=False)
                return response.json()['choices'][0]['message']['content']
            else:
                status.update(label="❌ Gagal Terhubung ke AI", state="error")
                return None
    except Exception as e:
        st.error(f"Koneksi Bermasalah: {str(e)}")
        return None

def set_font_safe(run, font_name='Times New Roman', size=11, bold=False):
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
    return text.strip()

# --- UTILITY UNTUK MENGAKTIFKAN LAYOUT 2 KOLOM ---
def set_document_to_two_columns(section):
    sectPr = section._sectPr
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '2')  
    cols.set(qn('w:space'), '500')  # Margin jarak antar kolom yang seimbang
    sectPr.append(cols)

def create_word_doc_kbc(content, doc_type, school_data):
    try:
        doc = Document()
        
        # Pengaturan Margin Kertas Resmi (Ramping & Efisien)
        current_section = doc.sections[0]
        current_section.top_margin = Inches(0.6)
        current_section.bottom_margin = Inches(0.6)
        current_section.left_margin = Inches(0.6)
        current_section.right_margin = Inches(0.6)

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(10.5) 
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        
        # --- [1 KOLOM] KOP UTAMA & PANEL INFORMASI ---
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(2)
        set_font_safe(p_title.add_run(f"{doc_type.upper()}\n"), size=13, bold=True)
        set_font_safe(p_title.add_run(f"MATA PELAJARAN : {school_data['mapel'].upper()}\n"), size=11, bold=True)
        set_font_safe(p_title.add_run(f"MATERI: {school_data['materi'].upper()}\n"), size=11, bold=True)
        
        # Garis Pembatas Kop
        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_after = Pt(8)
        set_font_safe(p_line.add_run("=" * 85), size=10, bold=True)
        
        # Identitas Naskah Soal Ringkas Berdampingan (Menggunakan Tabel Tanpa Border agar Rapi)
        table_id = doc.add_table(rows=3, cols=2)
        table_id.autofit = True
        
        labels_kiri = [
            f"Satuan Pendidikan : {school_data['nama_madrasah']}",
            f"Mata Pelajaran    : {school_data['mapel']}",
            f"Kelas / Fase      : {school_data['kelas']}"
        ]
        labels_kanan = [
            f"Tahun Pelajaran  : {school_data['tahun_ajaran']}",
            f"Penyusun/Guru    : {school_data['nama_guru']}",
            f"Sifat Ujian      : Mandiri / Tutup Buku"
        ]
        
        for idx in range(3):
            row = table_id.rows[idx]
            # Kolom Kiri
            p_cell_l = row.cells[0].paragraphs[0]
            p_cell_l.paragraph_format.space_after = Pt(2)
            set_font_safe(p_cell_l.add_run(labels_kiri[idx]), size=10)
            # Kolom Kanan
            p_cell_r = row.cells[1].paragraphs[0]
            p_cell_r.paragraph_format.space_after = Pt(2)
            set_font_safe(p_cell_r.add_run(labels_kanan[idx]), size=10)

        # Hapus garis border tabel identitas agar bersih
        for row in table_id.rows:
            for cell in row.cells:
                cell.width = Inches(3.5)

        doc.add_paragraph()

        # --- MASUK KE MODUL [2 KOLOM] AREA NASKAH SOAL ---
        soal_section = doc.add_section()
        set_document_to_two_columns(soal_section)

        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False
        
        for line in lines:
            stripped = line.strip()
            if stripped.startswith('|') and not re.match(r'^\|[\s\-:]+\|$', stripped):
                in_table_mode = True
            if in_table_mode:
                if re.match(r'^\|[\s\-:]+\|$', stripped): continue
                if stripped.startswith('|') and stripped.endswith('|'):
                    table_buffer.append([clean_markdown_symbols(c.strip()) for c in stripped.split('|')[1:-1]])
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
                                        row.cells[c_idx].text = cell_text
                                        for p in row.cells[c_idx].paragraphs:
                                            for run in p.runs: set_font_safe(run, size=9.5, bold=(r_idx==0))
                        except: pass
                        table_buffer = []
                        in_table_mode = False
            if not in_table_mode:
                if not stripped:
                    doc.add_paragraph()
                    continue
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(2) # Spasi antar baris pertanyaan sangat rapat
                is_bold, font_size = False, 10.5
                if stripped.startswith('# '):
                    stripped, is_bold, font_size = stripped[2:], True, 11.5
                    p.paragraph_format.space_before = Pt(6)
                elif stripped.startswith('## '):
                    stripped, is_bold, font_size = stripped[3:], True, 10.5
                    p.paragraph_format.space_before = Pt(4)
                if stripped.startswith('- ') or stripped.startswith('* '):
                    p.style = 'List Bullet'
                    stripped = stripped[2:]
                set_font_safe(p.add_run(clean_markdown_symbols(stripped)), size=font_size, bold=is_bold)

        # NOTE: Bagian tabel Tanda Tangan Kepala Sekolah & Guru sengaja DIBUANG sepenuhnya di sini agar naskah bersih.

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error Pembuatan Dokumen Word: {e}")
        return None

# --- DIALOG POP-UP DOKUMEN ---
@st.dialog("📥 Berkas Anda Telah Siap!")
def show_download_popup(buffer, filename, details):
    st.markdown('<div class="p-badge">Status: Sukses</div>', unsafe_allow_html=True)
    st.write(f"### 🎉 Naskah Soal Terbentuk!")
    st.markdown(f"""
    Lembar ujian Anda telah diformat menggunakan layout **Dual Kolom Kertas Resmi (Tanpa Kolom Tanda Tangan)**.
    * **Nama Berkas:** `{filename}`
    * **Ukuran Font Dasar:** 10.5 pt (Rapi & Padat)
    """)
    st.divider()
    st.download_button(
        label="📥 UNDUH SEKARANG",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# --- UI INTERFACE CORE ---
from streamlit_lottie import st_lottie
lottie_heart = load_lottieurl("https://assets5.lottiefiles.com/packages/lf20_v9t630.json")

c_title, c_logo = st.columns([0.85, 0.15])
with c_title:
    st.title("Workspace KBC 2026 ❤️")
    st.caption("Aplikasi Administrasi Lembaga & Bank Soal Otomatis Tingkat Lanjut")
with c_logo:
    if lottie_heart: st_lottie(lottie_heart, height=80, key="main_heart")

st.divider()

# TIGA TAB UTAMA
tab_setup, tab_materi, tab_soal = st.tabs([
    "🏛️ Langkah 1: Profil Lembaga", 
    "📚 Langkah 2: Modul Pembelajaran", 
    "📝 Langkah 3: Modul Bank Soal (Dual Column & Tanpa TTD)"
])

with tab_setup:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("Informasi Instansi & Guru")
    col1, col2 = st.columns(2)
    with col1:
        nama_madrasah = st.text_input("Nama Sekolah/Madrasah", "MI MIFTAHUSSALAM")
        jenis = st.selectbox("Jenjang Satuan", ["MI/SD", "MTs/SMP", "MA/SMA"])
        kabupaten = st.text_input("Kabupaten/Kota Domisili", "Kota Bogor")
        alamat = st.text_area("Alamat Lengkap", "Jl. Rimba Mulya II No.46, Pasirmulya, Bogor Barat", height=68)
    with col2:
        nama_guru = st.text_input("Nama Lengkap Guru (Penyusun)", "Guru Pengajar")
        nip_guru = st.text_input("NIP Guru", "-")
        kepala_madrasah = st.text_input("Nama Kepala Madrasah", "Drs. Andi Supriadi")
        nip_kepala = st.text_input("NIP Kepala", "-")
        kota = st.text_input("Kota Pembuatan Dokumen", "Bogor")
    st.markdown('</div>', unsafe_allow_html=True)

with tab_materi:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("Kurikulum Perangkat Mengajar")
    cx, cy = st.columns(2)
    with cx:
        doc_type_materi = st.selectbox("Pilih Output Dokumen Perangkat", ["Modul Ajar", "RPP", "ATP", "CP", "Prota", "Promes", "LKPD"])
        mapel_materi = st.text_input("Mata Pelajaran", "Akidah Akhlak", key="mp_materi")
    with cy:
        kelas_materi = st.text_input("Kelas / Fase Tingkat", "IV / Fase B", key="kls_materi")
        tahun_ajaran_materi = st.text_input("Periode Tahun Ajaran", "2026/2027", key="ta_materi")
    materi_pembelajaran = st.text_area("Topik Pokok Pembelajaran / Bab Bahasan", "Sifat Wajib Allah", height=100, key="mat_materi")
    st.divider()
    btn_materi = st.button("🚀 GENERASIKAN MODUL PERANGKAT", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab_soal:
    st.markdown('<div class="premium-card">', unsafe_allow_html=True)
    st.subheader("Konfigurasi Evaluasi Ujian / Ulangan Harian (Hasil Akhir Dual Kolom)")
    
    c_soal1, c_soal2 = st.columns(2)
    with c_soal1:
        doc_type_soal = st.selectbox("Jenis Asesmen Ulangan", ["Asesmen Tengah Semester (ATS)", "Sumatif Akhir Semester (SAS)", "Sumatif Akhir Tahun (SAT)"])
        mapel_soal = st.text_input("Mata Pelajaran Ujian", "Akidah Akhlak", key="mp_soal")
        kelas_soal = st.text_input("Jenjang / Kelas / Fase", "Kelas 4 SD / Fase B", key="kls_soal")
    with c_soal2:
        materi_soal = st.text_input("Materi Pokok Ujian", "Sifat Wajib Allah", key="mat_soal")
        kesulitan_soal = st.selectbox("Tingkat Kesulitan", ["Mudah", "Sedang", "Sulit", "Campuran / HOTS"])
        tahun_ajaran_soal = st.text_input("Tahun Ajaran Ujian", "2026/2027", key="ta_soal")
        
    st.markdown("##### 📊 Komposisi Distribusi Soal")
    cg1, cg2, cg3 = st.columns(3)
    with cg1:
        jml_pg = st.number_input("Jumlah Pilihan Ganda (PG)", min_value=0, max_value=50, value=10)
    with cg2:
        jml_isian = st.number_input("Jumlah Soal Isian Singkat", min_value=0, max_value=30, value=5)
    with cg3:
        jml_uraian = st.number_input("Jumlah Soal Uraian / Essay", min_value=0, max_value=20, value=2)
        
    st.divider()
    btn_soal = st.button("📝 GENERASIKAN LEMBAR BANK SOAL", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("📋 Kontrol Administrasi")
    st.markdown("Pilih tab di halaman tengah untuk menentukan tipe file output berkas Anda.")
    tanggal_buat = st.date_input("Tanggal Cetak Dokumen")

# --- PROSES EKSEKUSI TAB 2 ---
if btn_materi:
    if not nama_madrasah or not materi_pembelajaran:
        st.warning("⚠️ Harap lengkapi Profil Instansi pada Langkah 1 dan Materi pada Langkah 2.")
    else:
        data = {
            "nama_madrasah": nama_madrasah, "jenis": jenis, "kabupaten": kabupaten, "alamat": alamat, 
            "kepala_madrasah": kepala_madrasah, "nip_kepala": nip_kepala, "nama_guru": nama_guru, "nip_guru": nip_guru,
            "kota": kota, "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
            "mapel": mapel_materi, "kelas": kelas_materi, "tahun_ajaran": tahun_ajaran_materi, "materi": materi_pembelajaran
        }
        sys_prompt = "Anda adalah Ahli Kurikulum Berbasis Cinta (KBC) 2026. Format: Markdown. Langsung isi dokumen inti."
        user_prompt = f"Buat dokumen {doc_type_materi} pelajaran {mapel_materi} kelas {kelas_materi} topik: {materi_pembelajaran}. Sertakan tabel langkah kegiatan."
        
        content = get_ai_response_kbc(user_prompt, sys_prompt)
        if content:
            buffer = create_word_doc_kbc(content, doc_type_materi, data)
            if buffer:
                show_download_popup(buffer, f"KBC2026_{doc_type_materi}_{mapel_materi}.docx", data)

# --- PROSES EKSEKUSI TAB 3 ---
if btn_soal:
    if not nama_madrasah or not materi_soal:
        st.warning("⚠️ Harap lengkapi instansi pada Langkah 1 dan detail komposisi soal pada Langkah 3.")
    else:
        data = {
            "nama_madrasah": nama_madrasah, "jenis": jenis, "kabupaten": kabupaten, "alamat": alamat, 
            "kepala_madrasah": kepala_madrasah, "nip_kepala": nip_kepala, 
            "nama_guru": nama_guru, "nip_guru": nip_guru, 
            "kota": kota, "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
            "mapel": mapel_soal, "kelas": kelas_soal, "tahun_ajaran": tahun_ajaran_soal, 
            "materi": materi_soal, "tingkat_kesulitan": kesulitan_soal
        }
        
        sys_prompt = """
        Anda adalah Ahli Evaluasi Akademik Kurikulum Berbasis Cinta (KBC) 2026.
        Tugas Anda adalah memproduksi naskah lembar ujian/soal yang siap dikerjakan siswa.
        Format Output Harus Terstruktur Menggunakan Penomoran Rapi:
        - BAGIAN I: SOAL PILIHAN GANDA (berikan opsi A, B, C, D)
        - BAGIAN II: SOAL ISIAN SINGKAT
        - BAGIAN III: SOAL URAIAN / ESSAY
        Jangan memunculkan jawaban di lembar ini. Buatkan Kunci Jawaban terpisah di paling bawah halaman dokumen.
        Jangan gunakan kode blok markdown (```).
        """
        
        user_prompt = f"""
        Buatkan naskah soal ujian {doc_type_soal} untuk Mata Pelajaran {mapel_soal}, {kelas_soal}.
        Materi Pokok Bahasan: {materi_soal}.
        Tingkat Kesulitan Soal: {kesulitan_soal}.
        Komposisi Soal Yang Harus Dibuat:
        - {jml_pg} Soal Pilihan Ganda (PG)
        - {jml_isian} Soal Isian Singkat
        - {jml_uraian} Soal Uraian/Essay
        Sertakan Kunci Jawaban lengkap di bagian paling akhir halaman!
        """
        
        content = get_ai_response_kbc(user_prompt, sys_prompt)
        if content:
            buffer = create_word_doc_kbc(content, doc_type_soal, data)
            if buffer:
                show_download_popup(buffer, f"Soal_{doc_type_soal.split(' ')[0]}_{mapel_soal}_{kelas_soal.replace(' ', '')}.docx", data)
