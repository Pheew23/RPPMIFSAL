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
    page_title="Generator KBC 2026 (Clean Word)",
    page_icon="❤️",
    layout="wide"
)

# Konstanta API
# PERINGATAN: Jangan pernah menaruh API Key langsung di kode production. 
# Gunakan st.secrets atau environment variable. Ini hanya untuk demo sesuai request.
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response_kbc(prompt, system_instruction):
    """Mengambil respon AI dengan instruksi khusus KBC 2026"""
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
        "temperature": 0.8, # Slightly higher for creativity in KBC context
        "top_p": 0.9,
        "max_tokens": 4096,
        "stream": False
    }

    max_retries = 2
    attempt = 0

    while attempt < max_retries:
        try:
            with st.status(f"❤️ AI Sedang Merancang Pembelajaran Berbasis Cinta...", expanded=True) as status:
                response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=60)

                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ Konsep KBC 2026 Siap!", state="complete", expanded=False)
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
    """Helper function untuk set font dengan aman (mencegah bug font berubah)"""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    # Fix bug font Asia/Timur agar konsisten
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def clean_markdown_symbols(text):
    """Membersihkan simbol markdown agar tidak muncul di Word"""
    if not text: return ""

    # 1. Hapus Bold/Italic markers (** atau *)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)

    # 2. Hapus Heading markers (#)
    text = re.sub(r'^#+\s*', '', text)

    # 3. Hapus simbol list (- atau *) di awal baris (akan ditangani terpisah di logic utama)
    # Tapi jika tersisa, hapus saja
    text = re.sub(r'^[-*]\s+', '', text)

    # 4. Bersihkan spasi berlebih
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def create_word_doc_kbc(content, doc_type, school_data):
    """
    Versi Robust: 
    1. Parsing tabel lebih ketat.
    2. List conversion ke Native Word Bullets.
    3. Header/Footer standar madrasah.
    """
    try:
        doc = Document()

        # Set Default Style Global
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        style.paragraph_format.space_after = Pt(6) # Spasi antar paragraf rapi

        # --- 1. HEADER MADRASAH (KOP SURAT) ---
        if doc_type in ["RPP", "Modul Ajar", "ATP", "Prota", "Promes"]:
            # Baris 1: PEMERINTAH
            p1 = doc.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(f"PEMERINTAH KABUPATEN {school_data['kabupaten'].upper()}\n")
            set_font_safe(r1, size=14, bold=True)

            # Baris 2: DINAS & NAMA MADRASAH
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(f"DINAS PENDIDIKAN DAN KEBUDAYAAN\nMADRASAH {school_data['jenis'].upper()} {school_data['nama_madrasah'].upper()}\n")
            set_font_safe(r2, size=12, bold=True)

            # Baris 3: Alamat
            p3 = doc.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r3 = p3.add_run(f"Alamat: {school_data['alamat']} | Telp: {school_data['telp']}")
            set_font_safe(r3, size=10)

            # Garis Pembatas
            p_line = doc.add_paragraph()
            r_line = p_line.add_run("_" * 70)
            set_font_safe(r_line, size=10)
            p_line.paragraph_format.space_after = Pt(0)
            p_line.paragraph_format.space_before = Pt(0)

            # Judul Dokumen
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_title.paragraph_format.space_before = Pt(12)
            r_title = p_title.add_run(f"{doc_type.upper()}\n")
            r_title.add_break()
            r_title.add_run(f"Materi: {school_data['mapel']} - {school_data['kelas']}")
            set_font_safe(r_title, size=14, bold=True)
            r_title.underline = True

            doc.add_paragraph() # Spacer

        # --- 2. PARSE KONTEN AI KE WORD ---
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # LOGIKA TABEL (Perbaikan Bug Utama)
            # Deteksi awal tabel
            if stripped.startswith('|') and not re.match(r'^\|[\s\-:]+\|$', stripped):
                in_table_mode = True

            if in_table_mode:
                # Jika baris adalah pemisah tabel (|---|), skip
                if re.match(r'^\|[\s\-:]+\|$', stripped):
                    i += 1
                    continue

                # Jika masih format tabel
                if stripped.startswith('|') and stripped.endswith('|'):
                    raw_cells = stripped.split('|')[1:-1] # Hapus pipe pertama & terakhir
                    cells = [clean_markdown_symbols(c.strip()) for c in raw_cells]
                    table_buffer.append(cells)
                    i += 1
                    continue
                else:
                    # Tabel berakhir, render tabel
                    if table_buffer:
                        try:
                            max_cols = max(len(row) for row in table_buffer)
                            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
                            table.style = 'Table Grid'
                            table.autofit = False

                            # Set lebar kolom otomatis
                            for col in table.columns:
                                col.width = Inches(6.0 / max_cols)

                            for r_idx, row_data in enumerate(table_buffer):
                                row = table.rows[r_idx]
                                for c_idx, cell_text in enumerate(row_data):
                                    if c_idx < len(row.cells):
                                        cell = row.cells[c_idx]
                                        cell.text = cell_text
                                        # Styling sel tabel
                                        for paragraph in cell.paragraphs:
                                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                            for run in paragraph.runs:
                                                set_font_safe(run, size=11, bold=(r_idx==0)) # Header bold
                        except Exception as e:
                            doc.add_paragraph(f"[Error Render Tabel: {str(e)}]")

                        table_buffer = []
                        in_table_mode = False
                    # Lanjut proses baris ini sebagai teks biasa (fall-through)

            # Jika bukan tabel atau tabel sudah selesai
            if not in_table_mode:
                if not stripped:
                    doc.add_paragraph() # Empty line
                    i += 1
                    continue

                p = doc.add_paragraph()
                final_text = clean_markdown_symbols(stripped)

                # Deteksi Heading
                is_bold = False
                font_size = 12

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

                # Deteksi List (Bullet Points) - PERBAIKAN BUG LIST
                if stripped.startswith('- ') or stripped.startswith('* '):
                    p.style = 'List Bullet' # Gunakan style native Word
                    # Hapus tanda - atau * dari teks
                    final_text = clean_markdown_symbols(stripped[2:])

                r = p.add_run(final_text)
                set_font_safe(r, size=font_size, bold=is_bold)

            i += 1

        # Flush sisa tabel jika ada di akhir dokumen
        if in_table_mode and table_buffer:
            # (Logika render tabel sama seperti di atas)
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

        # --- 3. TANDA TANGAN ---
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n")
            sig_table = doc.add_table(rows=1, cols=2)
            sig_table.autofit = False

            # Kolom Kiri (Kepala Madrasah)
            cell_kiri = sig_table.cell(0, 0)
            cell_kiri.width = Inches(3.0)
            p_kiri = cell_kiri.paragraphs[0]
            p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r_kiri_1 = p_kiri.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
            set_font_safe(r_kiri_1, size=12)
            r_kiri_2 = p_kiri.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")
            set_font_safe(r_kiri_2, size=12, bold=True)

            # Kolom Kanan (Guru)
            cell_kanan = sig_table.cell(0, 1)
            cell_kanan.width = Inches(3.0)
            p_kanan = cell_kanan.paragraphs[0]
            p_kanan.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r_kanan_1 = p_kanan.add_run(f"{school_data['kota']}, {school_data['tanggal_buat']}\nGuru Mata Pelajaran\n\n\n\n\n")
            set_font_safe(r_kanan_1, size=12)
            r_kanan_2 = p_kanan.add_run(f"( {school_data['nama_guru']} )\nNIP. {school_data['nip_guru']}")
            set_font_safe(r_kanan_2, size=12, bold=True)

        # Save
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"❌ Error Fatal saat membuat Word: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

# --- UI STREAMLIT ---

st.title("❤️ Generator Dokumen KBC 2026")
st.markdown("""
**Kurikulum Berbasis Cinta (KBC) 2026**
Fokus pada *Well-being* (Kesejahteraan), *Relationship* (Hubungan), dan *Meaning* (Makna).
Dokumen dihasilkan bersih, tanpa simbol aneh, dan format tabel rapi.
""")

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

    st.subheader("📝 Konten Pembelajaran")
    doc_type = st.selectbox("Jenis Dokumen", ["Modul Ajar", "RPP", "ATP", "Prota", "Promes"])
    mapel = st.text_input("Mata Pelajaran", "Pendidikan Agama Islam & Budi Pekerti")
    kelas = st.text_input("Kelas / Fase", "VII / Fase D")

    st.info("💡 KBC 2026 akan otomatis menyusun langkah pembelajaran yang menyentuh hati.")
    materi = st.text_area("Topik / Materi", "Akhlak Terpuji: Kasih Sayang Terhadap Sesama", height=100)

btn = st.button("✨ Generate Dokumen KBC 2026", type="primary", use_container_width=True)

if btn:
    if not nama_madrasah or not materi:
        st.warning("Mohon lengkapi Nama Madrasah dan Topik Materi.")
    else:
        data = {
            "nama_madrasah": nama_madrasah, "jenis": jenis, "kabupaten": kabupaten,
            "alamat": alamat, "telp": telp, "kepala_madrasah": kepala_madrasah,
            "nip_kepala": nip_kepala, "nama_guru": nama_guru, "nip_guru": nip_guru,
            "kota": kota, "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
            "mapel": mapel, "kelas": kelas
        }

        # --- SYSTEM PROMPT KBC 2026 ---
        sys_prompt = """
        Anda adalah Ahli Kurikulum Berbasis Cinta (KBC) 2026. 
        Tugas Anda membuat dokumen pendidikan (RPP/Modul Ajar) yang berfokus pada:
        1. Well-being (Kesejahteraan psikologis siswa).
        2. Relationship (Hubungan harmonis guru-siswa dan antar siswa).
        3. Meaning (Makna pembelajaran dalam kehidupan nyata).

        FORMAT OUTPUT WAJIB:
        - Gunakan Markdown.
        - Gunakan Tabel Markdown (| Col | Col |) untuk bagian 'Langkah Pembelajaran' atau 'Asesmen'.
        - Gunakan List (- item) untuk tujuan pembelajaran.
        - Jangan gunakan kode blok (```).
        - Jangan ada kata pembuka seperti "Tentu, ini dokumennya". Langsung isi dokumen.
        - Pastikan bahasa menyentuh hati, humanis, dan tidak kaku.
        """

        user_prompt = f"""
        Buatkan {doc_type} untuk:
        Mapel: {mapel}
        Kelas: {kelas}
        Topik: {materi}

        Struktur yang diharapkan:
        1. Informasi Umum (Termasuk Profil Pelajar Pancasila & Dimensi KBC).
        2. Komponen Inti (Tujuan, Pemahaman Bermakna, Pertanyaan Pemantik).
        3. Langkah Pembelajaran (Pembuka, Inti, Penutup) - WAJIB DALAM BENTUK TABEL.
           Kolom Tabel: Kegiatan, Waktu, Nilai Cinta yang Ditanamkan.
        4. Asesmen (Formatif & Sumatif).
        """

        content = get_ai_response_kbc(user_prompt, sys_prompt)

        if content:
            st.success("Dokumen berhasil disusun! Sedang membersihkan format Word...")
            buffer = create_word_doc_kbc(content, doc_type, data)

            if buffer:
                fname = f"KBC2026_{doc_type}_{mapel}_{kelas}.docx"
                st.download_button(
                    label="📥 Download Dokumen Word (Siap Cetak)",
                    data=buffer,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                with st.expander("Lihat Preview Teks Mentah"):
                    st.text(content)
