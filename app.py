import streamlit as st
import requests
import json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import base64
from io import BytesIO

# --- KONFIGURASI SISTEM ---
st.set_page_config(
    page_title="Generator Dokumen Madrasah AI",
    page_icon="📚",
    layout="wide"
)

# Konstanta API NVIDIA
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response(prompt, system_instruction):
    """Mengirim request ke NVIDIA NIM API"""
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

    try:
        response = requests.post(NVIDIA_API_URL, headers=headers, json=payload)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except Exception as e:
        st.error(f"Terjadi kesalahan saat menghubungi AI: {str(e)}")
        return None

def create_word_doc(content, doc_type, school_data):
    """Membuat file .docx yang rapi dari konten Markdown/Text"""
    doc = Document()

    # Style Dasar
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # 1. KOPI SURAT (Khusus RPP/Modul Ajar)
    if doc_type in ["RPP", "Modul Ajar"]:
        # Judul Kop
        header = doc.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header.add_run(f"PEMERINTAH KABUPATEN {school_data['kabupaten'].upper()}\n")
        run.bold = True
        run.font.size = Pt(14)

        sub_header = doc.add_paragraph()
        sub_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = sub_header.add_run(f"DINAS PENDIDIKAN DAN KEBUDAYAAN\nMADRASAH {school_data['jenis'].upper()} {school_data['nama_madrasah'].upper()}\n")
        run2.bold = True
        run2.font.size = Pt(12)

        address = doc.add_paragraph()
        address.alignment = WD_ALIGN_PARAGRAPH.CENTER
        address.add_run(f"Alamat: {school_data['alamat']} | Telp: {school_data['telp']}")
        address.font.size = Pt(10)

        doc.add_paragraph("_" * 70) # Garis pemisah

        # Judul Dokumen
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.add_run(f"{doc_type.upper()}\n{school_data['mapel']} - {school_data['kelas']}")
        title_run.bold = True
        title_run.font.size = Pt(14)
        title_run.underline = True
        doc.add_paragraph() # Spasi

    else:
        # Judul untuk dokumen lain (Prota, Promes, dll)
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.add_run(f"{doc_type.upper()}\n{school_data['nama_madrasah']}")
        title_run.bold = True
        title_run.font.size = Pt(14)
        doc.add_paragraph()

    # 2. ISI KONTEN (Parsing sederhana untuk Markdown ke DOCX)
    lines = content.split('\n')
    current_table = None
    table_rows = []
    is_table = False

    for line in lines:
        line = line.strip()
        if not line:
            if is_table and table_rows:
                # Finalisasi tabel jika ada baris kosong
                add_table_to_doc(doc, table_rows)
                table_rows = []
                is_table = False
            doc.add_paragraph()
            continue

        # Deteksi Tabel Sederhana (Format Markdown: | Col | Col |)
        if line.startswith('|') and line.endswith('|'):
            is_table = True
            cells = [c.strip() for c in line.split('|')[1:-1]]
            table_rows.append(cells)
        else:
            if is_table and table_rows:
                add_table_to_doc(doc, table_rows)
                table_rows = []
                is_table = False

            # Deteksi Heading
            if line.startswith('# '):
                p = doc.add_paragraph()
                run = p.add_run(line[2:])
                run.bold = True
                run.font.size = Pt(14)
            elif line.startswith('## '):
                p = doc.add_paragraph()
                run = p.add_run(line[3:])
                run.bold = True
                run.font.size = Pt(12)
            elif line.startswith('- ') or line.startswith('* '):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line[2:])
            else:
                p = doc.add_paragraph(line)

    # Jika tabel berakhir di akhir dokumen
    if is_table and table_rows:
        add_table_to_doc(doc, table_rows)

    # 3. TANDA TANGAN (Khusus RPP/Modul Ajar)
    if doc_type in ["RPP", "Modul Ajar"]:
        doc.add_paragraph("\n\n")

        # Layout 2 Kolom untuk Tanda Tangan
        # Karena python-docx sulit membuat kolom side-by-side tanpa section break kompleks,
        # kita gunakan tabel transparan 2 kolom untuk positioning.
        sig_table = doc.add_table(rows=1, cols=2)
        sig_table.autofit = False
        sig_table.allow_autofit = False

        # Kolom Kiri: Guru
        cell_guru = sig_table.cell(0, 0)
        cell_guru.width = Inches(3.0)
        p_guru = cell_guru.paragraphs[0]
        p_guru.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_guru.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
        p_guru.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")

        # Kolom Kanan: Guru Mapel
        cell_guru_mapel = sig_table.cell(0, 1)
        cell_guru_mapel.width = Inches(3.0)
        p_guru_mapel = cell_guru_mapel.paragraphs[0]
        p_guru_mapel.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_guru_mapel.add_run(f"{school_data['kota']}, {school_data['tanggal_buat']}\nGuru Mata Pelajaran\n\n\n\n\n")
        p_guru_mapel.add_run(f"( {school_data['nama_guru']} )\nNIP. {school_data['nip_guru']}")

    # Simpan ke BytesIO
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def add_table_to_doc(doc, rows):
    """Helper untuk menambahkan tabel ke dokumen"""
    if not rows: return
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'

    for i, row_data in enumerate(rows):
        row = table.rows[i]
        for j, cell_text in enumerate(row_data):
            cell = row.cells[j]
            cell.text = cell_text
            # Bold header row
            if i == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

# --- UI STREAMLIT ---

st.title("📚 Generator Dokumen Pendidikan Madrasah AI")
st.markdown(f"**Model:** `{MODEL_NAME}` via **NVIDIA Builder** | **Dev:** Lagos AI 9.1 (rian dev)")

with st.sidebar:
    st.header("🏫 Data Madrasah")
    nama_madrasah = st.text_input("Nama Madrasah", "MI/MTs/MA Al-Hikmah")
    jenis = st.selectbox("Jenjang", ["MI", "MTs", "MA"])
    kabupaten = st.text_input("Kabupaten/Kota", "Contoh: Cirebon")
    alamat = st.text_area("Alamat Lengkap", "Jl. Pendidikan No. 1")
    telp = st.text_input("Nomor Telepon", "0231-1234567")

    st.subheader("👤 Data Guru & Kepala")
    kepala_madrasah = st.text_input("Nama Kepala Madrasah", "Drs. H. Ahmad Fulan, M.Pd")
    nip_kepala = st.text_input("NIP Kepala", "19700101 199001 1 001")
    nama_guru = st.text_input("Nama Guru Penyusun", "Fulan bin Fulan, S.Pd")
    nip_guru = st.text_input("NIP Guru", "19900101 202001 1 001")
    kota = st.text_input("Kota Tanda Tangan", "Cirebon")
    tanggal_buat = st.date_input("Tanggal Pembuatan")

    st.subheader("📝 Detail Dokumen")
    doc_type = st.selectbox(
        "Jenis Dokumen",
        ["RPP", "Modul Ajar", "ATP", "CP", "KKTP", "Prota", "Promes"]
    )
    mapel = st.text_input("Mata Pelajaran", "Pendidikan Agama Islam")
    kelas = st.text_input("Kelas/Semester", "VII (Tujuh) / Ganjil")
    materi = st.text_area("Topik/Materi Spesifik", "Contoh: Akhlak Terpuji, Shalat Berjamaah")

    generate_btn = st.button("🚀 Buat Dokumen", type="primary")

# --- LOGIKA UTAMA ---

if generate_btn:
    if not nama_madrasah or not materi:
        st.warning("Mohon lengkapi Nama Madrasah dan Topik Materi.")
    else:
        school_data = {
            "nama_madrasah": nama_madrasah,
            "jenis": jenis,
            "kabupaten": kabupaten,
            "alamat": alamat,
            "telp": telp,
            "kepala_madrasah": kepala_madrasah,
            "nip_kepala": nip_kepala,
            "nama_guru": nama_guru,
            "nip_guru": nip_guru,
            "kota": kota,
            "tanggal_buat": tanggal_buat.strftime("%d %B %Y"),
            "mapel": mapel,
            "kelas": kelas
        }

        with st.spinner(f"Sedang mengakses {MODEL_NAME} untuk menyusun {doc_type}..."):
            # System Prompt yang Sangat Spesifik
            system_prompt = f"""
            Anda adalah asisten ahli penyusun kurikulum pendidikan Indonesia (Kurikulum Merdeka) untuk tingkat Madrasah.
            Tugas Anda adalah membuat dokumen {doc_type} yang sangat lengkap, terstruktur, dan siap pakai.

            PENTING:
            1. Gunakan format Markdown yang rapi.
            2. Untuk tabel, WAJIB menggunakan format Markdown standar (| Header | Header |) agar bisa dikonversi ke Word.
            3. Isi konten harus mendalam, pedagogis, dan sesuai dengan nilai-nilai Islam jika relevan.
            4. Jangan sertakan kata pengantar seperti "Berikut adalah dokumen...", langsung mulai dari Judul Dokumen (Level 2 Heading).
            5. Untuk RPP/Modul Ajar, pastikan mencakup: Tujuan Pembelajaran, Langkah Kegiatan (Pendahuluan, Inti, Penutup), Asesmen, dan Pengayaan/Remedial.
            6. Untuk ATP/Prota/Promes, buat dalam format tabel rincian minggu/semester.

            Data Konteks:
            - Mapel: {mapel}
            - Kelas: {kelas}
            - Materi: {materi}
            """

            user_prompt = f"Buatkan {doc_type} lengkap untuk materi '{materi}' kelas {kelas} mata pelajaran {mapel}."

            ai_content = get_ai_response(user_prompt, system_prompt)

            if ai_content:
                st.success("Dokumen berhasil dibuat oleh AI!")

                # Tampilkan Preview
                st.subheader("Preview Konten")
                st.markdown(ai_content)

                # Generate Word File
                doc_buffer = create_word_doc(ai_content, doc_type, school_data)

                # Download Button
                filename = f"{doc_type}_{mapel}_{kelas}_{nama_guru.replace(' ', '_')}.docx"
                st.download_button(
                    label="📥 Download File Word (.docx) Siap Cetak",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                st.info("💡 Tips: File Word yang diunduh sudah memiliki Kop Surat, Tanda Tangan, dan Format Tabel yang rapi. Anda hanya perlu menyesuaikan sedikit jika ada data spesifik yang berubah.")

# Footer
st.markdown("---")
st.caption("Dikembangkan dengan Lagos AI 9.1 (rian dev) menggunakan NVIDIA NIM & Qwen3.5-397B-A17B")
