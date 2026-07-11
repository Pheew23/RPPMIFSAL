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
    page_title="Generator Dokumen Madrasah AI (High-Performance)",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Konstanta API NVIDIA
NVIDIA_API_KEY = "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"
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- FUNGSI HELPER ---

def get_ai_response_robust(prompt, system_instruction):
    """
    Mengirim request ke NVIDIA NIM dengan timeout panjang dan retry logic.
    Dioptimalkan untuk model besar (397B) yang butuh waktu inferensi lebih lama.
    """
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
        "max_tokens": 4096, # Memastikan ruang cukup untuk dokumen panjang
        "stream": False
    }

    max_retries = 3
    attempt = 0

    while attempt < max_retries:
        try:
            with st.status(f"🧠 Menghubungi Qwen3.5-397B (Percobaan {attempt + 1}/{max_retries})...", expanded=True) as status:
                st.write("Mengirim permintaan ke NVIDIA Cloud...")

                # TIMEOUT DISET KE NONE (Menunggu sampai selesai) atau sangat lama (300 detik)
                # Model 397B bisa butuh waktu untuk generate konten panjang
                response = requests.post(
                    NVIDIA_API_URL, 
                    headers=headers, 
                    json=payload, 
                    timeout=None  # Kunci: Tidak ada batas waktu putus paksa
                )

                if response.status_code == 200:
                    st.write("✅ Respon diterima. Memproses konten...")
                    result = response.json()
                    content = result['choices'][0]['message']['content']
                    status.update(label="✅ AI Selesai Berpikir!", state="complete", expanded=False)
                    return content
                else:
                    error_msg = f"API Error: {response.status_code} - {response.text}"
                    st.warning(error_msg)
                    status.update(label="❌ Gagal Menghubungi API", state="error")
                    attempt += 1
                    if attempt < max_retries:
                        st.info("Mencoba ulang dalam 2 detik...")
                        time.sleep(2)
                    else:
                        return None

        except requests.exceptions.Timeout:
            st.error("Timeout: Server AI terlalu sibuk. Mencoba ulang...")
            attempt += 1
            time.sleep(2)
        except Exception as e:
            st.error(f"Koneksi Gagal: {str(e)}")
            attempt += 1
            if attempt < max_retries:
                time.sleep(2)
            else:
                return None

    return None

def create_word_doc_safe(content, doc_type, school_data):
    """Membuat file .docx dengan penanganan error yang ketat"""
    try:
        doc = Document()

        # Set Font Global
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

        # 1. KOPI SURAT
        if doc_type in ["RPP", "Modul Ajar"]:
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

        # 2. PARSE KONTEN & TABEL
        lines = content.split('\n')
        table_buffer = []
        in_table_mode = False

        for line in lines:
            stripped = line.strip()

            # Skip separator baris tabel markdown (|---|)
            if re.match(r'^\|[\s\-:]+\|$', stripped):
                continue

            if stripped.startswith('|') and stripped.endswith('|'):
                in_table_mode = True
                cells = [c.strip() for c in stripped.split('|')[1:-1]]
                # Filter sel kosong yang mungkin terjadi karena format aneh
                cells = [c for c in cells if c != ""]
                if cells: # Hanya tambah jika ada isi
                    table_buffer.append(cells)
            else:
                if in_table_mode and table_buffer:
                    # Render Tabel
                    if len(table_buffer) > 1: 
                        try:
                            # Deteksi jumlah kolom maksimal
                            max_cols = max(len(row) for row in table_buffer)
                            # Pad row yang pendek agar sesuai jumlah kolom
                            for row in table_buffer:
                                while len(row) < max_cols:
                                    row.append("")

                            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
                            table.style = 'Table Grid'
                            table.autofit = False

                            for i, row_data in enumerate(table_buffer):
                                row = table.rows[i]
                                for j, cell_text in enumerate(row_data):
                                    cell = row.cells[j]
                                    cell.text = str(cell_text)
                                    if i == 0: # Header bold
                                        for par in cell.paragraphs:
                                            for run in par.runs:
                                                run.bold = True
                                                run.font.name = 'Times New Roman'
                        except Exception as e:
                            doc.add_paragraph(f"[Tabel tidak dapat dirender: {str(e)}]")

                    table_buffer = []
                    in_table_mode = False

                if not stripped:
                    doc.add_paragraph()
                    continue

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

        # Flush sisa tabel
        if in_table_mode and table_buffer and len(table_buffer) > 1:
            max_cols = max(len(row) for row in table_buffer)
            for row in table_buffer:
                while len(row) < max_cols: row.append("")
            table = doc.add_table(rows=len(table_buffer), cols=max_cols)
            table.style = 'Table Grid'
            for i, row_data in enumerate(table_buffer):
                for j, cell_text in enumerate(row_data):
                    table.rows[i].cells[j].text = str(cell_text)

        # 3. TANDA TANGAN
        if doc_type in ["RPP", "Modul Ajar"]:
            doc.add_paragraph("\n\n")
            sig_table = doc.add_table(rows=1, cols=2)
            sig_table.autofit = False

            cell_kiri = sig_table.cell(0, 0)
            cell_kiri.width = Inches(3.0)
            p_kiri = cell_kiri.paragraphs[0]
            p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_kiri.add_run(f"Mengetahui,\nKepala Madrasah\n\n\n\n\n")
            p_kiri.add_run(f"( {school_data['kepala_madrasah']} )\nNIP. {school_data['nip_kepala']}")

            cell_kanan = sig_table.cell(0, 1)
            cell_kanan.width = Inches(3.0)
            p_kanan = cell_kanan.paragraphs[0]
            p_kanan.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_kanan.add_run(f"{school_data['kota']}, {school_data['tanggal_buat']}\nGuru Mata Pelajaran\n\n\n\n\n")
            p_kanan.add_run(f"( {school_data['nama_guru']} )\nNIP. {school_data['nip_guru']}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"❌ Error Fatal saat membuat Word: {str(e)}")
        return None

# --- UI STREAMLIT ---

st.title("🚀 Generator Dokumen Madrasah AI (Long-Running Optimized)")
st.markdown(f"**Model:** `{MODEL_NAME}` (397B Parameters) | **Status:** Siap Menunggu Respon Panjang")

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
    materi = st.text_area("Topik/Materi Spesifik", "Contoh: Akhlak Terpuji, Shalat Berjamaah", height=100)

    generate_btn = st.button("🚀 Mulai Generasi Dokumen (Tunggu Beberapa Detik)", type="primary", use_container_width=True)

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
        Anda adalah ahli kurikulum Indonesia tingkat madrasah. 
        Buatlah dokumen {doc_type} yang SANGAT LENGKAP, DETAIL, dan SIAP PAKAI.

        PENTING:
        1. Gunakan format Markdown.
        2. Tabel HARUS menggunakan format: | Kolom 1 | Kolom 2 | (dengan baris pemisah |---|---|).
        3. Jangan gunakan kode blok (```markdown) untuk membungkus seluruh output. Tulis langsung teksnya.
        4. Isi konten harus pedagogis, mendalam, dan sesuai nilai Islam.
        5. Langsung mulai dari judul (## Judul), tanpa pembuka chat.
        """

        user_prompt = f"Buat {doc_type} lengkap untuk {mapel} kelas {kelas}, materi: {materi}. Pastikan semua komponen kurikulum terpenuhi."

        # Eksekusi dengan UI Progress
        ai_content = get_ai_response_robust(user_prompt, system_prompt)

        if ai_content:
            st.success("✅ Konten berhasil dibuat!")

            with st.expander("Lihat Preview Konten (Markdown)", expanded=False):
                st.markdown(ai_content)

            st.info("🔄 Sedang mengkonversi ke format Word (.docx) yang rapi...")
            doc_buffer = create_word_doc_safe(ai_content, doc_type, school_data)

            if doc_buffer:
                filename = f"{doc_type}_{mapel.replace(' ', '_')}_{kelas.replace(' ', '_')}.docx"

                st.balloons()
                st.success("🎉 Dokumen Siap Diunduh!")

                st.download_button(
                    label="📥 DOWNLOAD FILE WORD (.docx) - SIAP CETAK",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.error("Gagal membuat file Word. Silakan coba lagi.")
        else:
            st.error("❌ Gagal mendapatkan respon dari AI setelah beberapa percobaan. Periksa koneksi internet atau coba materi yang lebih sederhana.")

st.markdown("---")
st.caption("Dioptimalkan oleh Lagos AI 9.1 (rian dev) - Timeout Disabled & Retry Logic Enabled")
