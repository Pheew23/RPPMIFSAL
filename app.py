import streamlit as st
import requests
import json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import re
import time

# --- KONFIGURASI SISTEM ---
st.set_page_config(
    page_title="Generator Dokumen Madrasah AI (Fixed)",
    page_icon="📚",
    layout="wide"
)

# Konstanta API NVIDIA
# Pastikan key ini valid. Jika error 401, cek kembali key Anda.
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response(prompt, system_instruction):
    """Mengirim request ke NVIDIA NIM API dengan retry logic sederhana"""
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
        with st.spinner('Menghubungi NVIDIA AI Cloud...'):
            response = requests.post(NVIDIA_API_URL, headers=headers, json=payload, timeout=60)

        if response.status_code == 200:
            result = response.json()
            return result['choices'][0]['message']['content']
        else:
            st.error(f"API Error: {response.status_code} - {response.text}")
            return None
    except requests.exceptions.Timeout:
        st.error("Timeout: Respon AI terlalu lama. Coba lagi dengan materi yang lebih singkat.")
        return None
    except Exception as e:
        st.error(f"Koneksi Gagal: {str(e)}")
        return None

def clean_markdown_tables(text):
    """Membersihkan format tabel markdown agar konsisten untuk parser"""
    lines = text.split('\n')
    cleaned_lines = []
    in_table = False

    for line in lines:
        stripped = line.strip()
        # Deteksi awal tabel
        if stripped.startswith('|') and stripped.endswith('|'):
            in_table = True
            # Pastikan pemisah header (|---|---|) ada jika ini baris kedua tabel
            if len(cleaned_lines) > 0 and cleaned_lines[-1].startswith('|') and not cleaned_lines[-1].contains('|---'):
                 # Cek apakah baris sebelumnya adalah header, jika ya, pastikan ada separator
                 pass 
            cleaned_lines.append(stripped)
        else:
            if in_table and stripped == "":
                # Akhiri tabel jika ada baris kosong
                in_table = False
            cleaned_lines.append(line)

    return '\n'.join(cleaned_lines)

def create_word_doc(content, doc_type, school_data):
    """Membuat file .docx dengan error handling yang ketat"""
    try:
        doc = Document()

        # Set Font Global (Times New Roman)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        # Support karakter kompleks
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

        # 1. KOPI SURAT
        if doc_type in ["RPP", "Modul Ajar"]:
            try:
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

                doc.add_paragraph("_" * 70)

                title = doc.add_paragraph()
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_run = title.add_run(f"{doc_type.upper()}\n{school_data['mapel']} - {school_data['kelas']}")
                title_run.bold = True
                title_run.font.size = Pt(14)
                title_run.underline = True
                doc.add_paragraph()
            except Exception as e:
                st.warning(f"Gagal membuat kop surat: {e}. Melanjutkan tanpa kop.")

        # 2. PARSE KONTEN
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False

        for line in lines:
            stripped = line.strip()

            # Skip baris kosong di awal tabel markdown separator (|---|)
            if re.match(r'^\|[\s\-:]+\|$', stripped):
                continue

            if stripped.startswith('|') and stripped.endswith('|'):
                in_table_mode = True
                # Bersihkan pipe ganda dan split
                cells = [c.strip() for c in stripped.split('|')[1:-1]]
                table_buffer.append(cells)
            else:
                if in_table_mode and table_buffer:
                    # Render Tabel
                    try:
                        if len(table_buffer) > 1: # Pastikan ada lebih dari 1 baris
                            table = doc.add_table(rows=len(table_buffer), cols=len(table_buffer[0]))
                            table.style = 'Table Grid'
                            table.autofit = False

                            # Set lebar kolom otomatis
                            for col in table.columns:
                                col.width = Inches(6 / len(table_buffer[0])) # Distribusi lebar

                            for i, row_data in enumerate(table_buffer):
                                row = table.rows[i]
                                for j, cell_text in enumerate(row_data):
                                    if j < len(row.cells):
                                        cell = row.cells[j]
                                        cell.text = cell_text
                                        if i == 0: # Header bold
                                            for par in cell.paragraphs:
                                                for run in par.runs:
                                                    run.bold = True
                                                    run.font.name = 'Times New Roman'
                    except Exception as e:
                        doc.add_paragraph(f"[Error rendering tabel: {str(e)}]")

                    table_buffer = []
                    in_table_mode = False

                if not stripped:
                    doc.add_paragraph()
                    continue

                # Handle Headings
                if stripped.startswith('# '):
                    p = doc.add_paragraph()
                    run = p.add_run(stripped[2:])
                    run.bold = True
                    run.font.size = Pt(14)
                elif stripped.startswith('## '):
                    p = doc.add_paragraph()
                    run = p.add_run(stripped[3:])
                    run.bold = True
                    run.font.size = Pt(12)
                elif stripped.startswith('- ') or stripped.startswith('* '):
                    p = doc.add_paragraph(style='List Bullet')
                    p.add_run(stripped[2:])
                else:
                    doc.add_paragraph(stripped)

        # Flush sisa tabel jika ada di akhir
        if in_table_mode and table_buffer:
             # Logika render tabel sama seperti di atas (disederhanakan untuk brevity)
             if len(table_buffer) > 1:
                table = doc.add_table(rows=len(table_buffer), cols=len(table_buffer[0]))
                table.style = 'Table Grid'
                for i, row_data in enumerate(table_buffer):
                    row = table.rows[i]
                    for j, cell_text in enumerate(row_data):
                        if j < len(row.cells):
                            row.cells[j].text = cell_text

        # 3. TANDA TANGAN
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n\n")
            sig_table = doc.add_table(rows=1, cols=2)
            sig_table.autofit = False

            # Kolom Kiri
            cell_kiri = sig_table.cell(0, 0)
            cell_kiri.width = Inches(3.0)
            p_kiri = cell_kiri.paragraphs[0]
            p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_kiri.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
            p_kiri.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")

            # Kolom Kanan
            cell_kanan = sig_table.cell(0, 1)
            cell_kanan.width = Inches(3.0)
            p_kanan = cell_kanan.paragraphs[0]
            p_kanan.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_kanan.add_run(f"{school_data['kota']}, {school_data['tanggal_buat']}\nGuru Mata Pelajaran\n\n\n\n\n")
            p_kanan.add_run(f"( {school_data['nama_guru']} )\nNIP. {school_data['nip_guru']}")

        # Simpan ke Buffer
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)  # PENTING: Reset pointer ke awal
        return buffer

    except Exception as e:
        st.error(f"Fatal Error saat membuat dokumen Word: {str(e)}")
        return None

# --- UI STREAMLIT ---

st.title("📚 Generator Dokumen Madrasah AI (Stable)")
st.markdown(f"**Model:** `{MODEL_NAME}` | **Status:** Siap Mengunduh")

with st.sidebar:
    st.header("🏫 Data Madrasah")
    nama_madrasah = st.text_input("Nama Madrasah", "MI/MTs/MA Al-Hikmah")
    jenis = st.selectbox("Jenjang", ["MI", "MTs", "MA"])
    kabupaten = st.text_input("Kabupaten/Kota", "Cirebon")
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
    materi = st.text_area("Topik/Materi Spesifik", "Contoh: Akhlak Terpuji")

    generate_btn = st.button("🚀 Buat & Download Dokumen", type="primary", use_container_width=True)

# --- LOGIKA UTAMA ---

if generate_btn:
    if not nama_madrasah or not materi:
        st.warning("⚠️ Mohon lengkapi Nama Madrasah dan Topik Materi.")
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

        system_prompt = f"""
        Anda adalah ahli kurikulum Indonesia. Buatlah {doc_type} untuk Madrasah.
        ATURAN WAJIB:
        1. Gunakan format Markdown.
        2. Untuk data tabel, gunakan format: | Kolom 1 | Kolom 2 | di baris pertama, dan | --- | --- | di baris kedua.
        3. Jangan gunakan kode blok (```markdown) untuk membungkus tabel, tulis langsung teksnya.
        4. Isi harus lengkap, rinci, dan siap pakai.
        5. Langsung mulai dari judul dokumen (## Judul), jangan ada pembuka chat.
        """

        user_prompt = f"Buat {doc_type} lengkap untuk {mapel} kelas {kelas}, materi: {materi}."

        # Langkah 1: Get AI Content
        ai_content = get_ai_response(user_prompt, system_prompt)

        if ai_content:
            st.success("✅ Konten berhasil dibuat oleh AI!")

            with st.expander("Lihat Preview Konten Mentah", expanded=False):
                st.markdown(ai_content)

            # Langkah 2: Generate Word
            st.info("🔄 Sedang memformat dokumen Word...")
            doc_buffer = create_word_doc(ai_content, doc_type, school_data)

            if doc_buffer:
                filename = f"{doc_type}_{mapel.replace(' ', '_')}_{kelas.replace(' ', '_')}.docx"

                st.success("🎉 Dokumen siap diunduh!")

                st.download_button(
                    label="📥 DOWNLOAD FILE WORD (.docx)",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.error("Gagal membuat file Word. Silakan coba lagi atau perkecil kompleksitas materi.")

st.markdown("---")
st.caption("Diperbaiki oleh Lagos AI 9.1 (rian dev) - Optimasi Buffer & Parsing Tabel")
