import streamlit as st
import os
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- KONFIGURASI & LOAD ENV ---
load_dotenv()

# Konfigurasi Halaman
st.set_page_config(
    page_title="Generator KBC 2026 | NVIDIA NIM Powered",
    page_icon="🚀",
    layout="wide"
)

# Ambil Kredensial dari .env
api_key = os.getenv("NVIDIA_API_KEY")
base_url = os.getenv("NVIDIA_BASE_URL", "https://integrate.api.nvidia.com/v1")
model_name = os.getenv("NVIDIA_MODEL_NAME", "qwen/qwen3.5-397b-a17b")

# Validasi Awal
if not api_key:
    st.error("⚠️ **API Key NVIDIA tidak ditemukan!** Silakan cek file `.env` Anda.")
    st.stop()

# Inisialisasi Client OpenAI (Kompatibel dengan NVIDIA NIM)
client = OpenAI(
    api_key=api_key,
    base_url=base_url
)

# --- FUNGSI: GENERATOR DOKUMEN WORD (.docx) ---
def create_word_doc(title, content_markdown, teacher_name, head_name, doc_type):
    doc = Document()

    # Set Font Global (Times New Roman, 12pt)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # 1. Judul Dokumen (Tengah, Bold, Uppercase)
    header = doc.add_heading(title.upper(), 0)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header.runs[0].font.name = 'Times New Roman'
    header.runs[0].font.size = Pt(14)
    header.runs[0].font.bold = True

    # Garis Bawah Judul (Opsional, untuk estetika resmi)
    p_underline = doc.add_paragraph()
    p_underline.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_underline = p_underline.add_run("_" * 60)
    run_underline.font.size = Pt(10)

    # 2. Info Dokumen
    doc.add_paragraph(f"Jenis Dokumen: {doc_type}", style='Heading 2')
    doc.add_paragraph(f"Kurikulum: Kurikulum Berbasis Cinta (KBC) 2026")

    # 3. Parsing Konten Markdown ke DOCX
    lines = content_markdown.split('\n')
    current_list = None

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            doc.add_paragraph() # Spasi kosong
            continue

        # Handle Headings
        if line_stripped.startswith('# '):
            p = doc.add_heading(line_stripped[2:], level=1)
            for run in p.runs: run.font.name = 'Times New Roman'
        elif line_stripped.startswith('## '):
            p = doc.add_heading(line_stripped[3:], level=2)
            for run in p.runs: run.font.name = 'Times New Roman'
        elif line_stripped.startswith('### '):
            p = doc.add_heading(line_stripped[4:], level=3)
            for run in p.runs: run.font.name = 'Times New Roman'
        # Handle Lists
        elif line_stripped.startswith('- ') or line_stripped.startswith('* '):
            p = doc.add_paragraph(line_stripped[2:], style='List Bullet')
            for run in p.runs: run.font.name = 'Times New Roman'
        elif line_stripped[0].isdigit() and line_stripped[1] == '.':
            p = doc.add_paragraph(line_stripped, style='List Number')
            for run in p.runs: run.font.name = 'Times New Roman'
        # Handle Teks Biasa
        else:
            p = doc.add_paragraph(line_stripped)
            for run in p.runs: run.font.name = 'Times New Roman'

    # 4. Bagian Tanda Tangan (Layout Tabel Transparan)
    doc.add_paragraph('\n' * 3) # Spasi besar sebelum ttd

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.style = None # Hilangkan border tabel

    # Kolom Kiri: Guru
    cell_guru = table.cell(0, 0)
    cell_guru.width = Inches(3.5)
    p_guru = cell_guru.paragraphs[0]
    p_guru.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_guru = p_guru.add_run(f"Guru Mata Pelajaran,\n\n\n\n\n")
    run_guru.font.bold = True
    run_name_guru = p_guru.add_run(f"({teacher_name})\nNIP. ...........................")
    run_name_guru.font.bold = True

    # Kolom Kanan: Kepala Madrasah
    cell_head = table.cell(0, 1)
    cell_head.width = Inches(3.5)
    p_head = cell_head.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_head = p_head.add_run(f"Kepala Madrasah,\n\n\n\n\n")
    run_head.font.bold = True
    run_name_head = p_head.add_run(f"({head_name})\nNIP. ...........................")
    run_name_head.font.bold = True

    # Simpan ke Buffer Memory
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- FUNGSI: AI GENERATION (NVIDIA NIM) ---
def generate_kbc_content(doc_type, subject, grade, topic, teacher, head, duration, goals):
    system_instruction = """
    Anda adalah asisten ahli pendidikan Indonesia yang beroperasi di bawah 'Kurikulum Berbasis Cinta (KBC) 2026'.
    Filosofi Utama:
    1. Humanis & Empatik: Setiap interaksi belajar harus membangun kasih sayang.
    2. Psikologis Aman: Siswa harus merasa dihargai sebelum dinilai.
    3. Kontekstual & Relevan: Materi dikaitkan dengan kehidupan nyata dan nilai kebaikan.

    TUGAS:
    Buat dokumen administratif ({doc_type}) yang sangat terstruktur, formal, namun hangat.
    Format output WAJIB menggunakan Markdown standar (# untuk judul, ## untuk subjudul, - untuk list).
    Jangan sertakan kata pembuka seperti 'Tentu, ini dokumennya'. Langsung berikan isi dokumen.
    Pastikan bagian 'Langkah Pembelajaran' (jika RPP) menyertakan ice-breaking berbasis cinta dan refleksi perasaan.
    """

    user_prompt = f"""
    Buatkan {doc_type} dengan detail berikut:
    - Mata Pelajaran: {subject}
    - Kelas/Fase: {grade}
    - Topik: {topic}
    - Guru Pengampu: {teacher}
    - Kepala Madrasah: {head}
    - Waktu: {duration}
    - Fokus Tujuan: {goals}

    Standar Format 2026:
    - Gunakan bahasa Indonesia baku yang elegan.
    - Untuk RPP: Sertakan Pendahuluan (Apersepsi Emosional), Inti (Eksplorasi Kolaboratif), Penutup (Refleksi Hati).
    - Untuk KKTP: Gunakan deskripsi kualitatif yang memotivasi, bukan hanya angka.
    - Pastikan struktur rapi seperti dokumen resmi dinas.
    """

    try:
        # Panggilan ke NVIDIA NIM
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_instruction},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=4096,
            top_p=0.9,
            frequency_penalty=0.1
        )

        return response.choices[0].message.content
    except Exception as e:
        error_msg = str(e)
        if "401" in error_msg:
            return "ERROR: API Key tidak valid atau kedaluwarsa. Cek file .env."
        elif "404" in error_msg:
            return f"ERROR: Model '{model_name}' tidak ditemukan di endpoint ini. Periksa nama model atau URL base."
        else:
            return f"ERROR KONEKSI NVIDIA NIM: {error_msg}"

# --- UI STREAMLIT ---
st.title("🚀 Generator Administrasi KBC 2026 (NVIDIA NIM)")
st.markdown(f"**Model:** `{model_name}` | **Engine:** NVIDIA Builder | **Dev:** Lagos AI 9.1")

with st.sidebar:
    st.header("📝 Input Data Dokumen")

    doc_type = st.selectbox(
        "Pilih Jenis Dokumen",
        ["RPP (Modul Ajar)", "ATP (Alur Tujuan Pembelajaran)", "CP (Capaian Pembelajaran)", 
         "KKTP (Kriteria Ketuntasan)", "PROTA (Program Tahunan)", "PROMES (Program Semester)"]
    )

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        subject = st.text_input("Mata Pelajaran", "Akidah Akhlak")
        grade = st.text_input("Kelas / Fase", "Kelas 7 (Fase D)")
    with col2:
        topic = st.text_input("Topik Materi", "Beriman kepada Kitab-kitab Allah")

    st.divider()

    teacher_name = st.text_input("Nama Guru Lengkap", "Ahmad Fauzi, S.Pd.I")
    head_name = st.text_input("Nama Kepala Madrasah", "Dr. H. Abdullah, M.Ag")
    duration = st.text_input("Alokasi Waktu / Semester", "2 JP x Pertemuan / Semester Ganjil 2026")

    goals = st.text_area(
        "Tujuan Pembelajaran / Fokus KBC", 
        placeholder="Contoh: Siswa mampu menjelaskan keyakinan terhadap kitab suci dengan penuh rasa cinta dan tanggung jawab moral."
    )

    btn_generate = st.button("✨ Generate dengan AI", type="primary", use_container_width=True)

# Logika Utama
if btn_generate:
    if not all([subject, grade, topic, teacher_name, head_name]):
        st.warning("⚠️ Mohon lengkapi semua data di sidebar.")
    else:
        with st.spinner('🤖 NVIDIA NIM sedang menyusun kurikulum berbasis cinta...'):
            # 1. Generate Teks
            ai_content = generate_kbc_content(
                doc_type, subject, grade, topic, teacher_name, head_name, duration, goals
            )

            if ai_content.startswith("ERROR"):
                st.error(ai_content)
            else:
                st.success("✅ Dokumen berhasil disusun!")

                tab_preview, tab_download = st.tabs(["👁️ Preview Markdown", "📥 Download Word (.docx)"])

                with tab_preview:
                    st.markdown(ai_content)

                with tab_download:
                    st.info("File Word yang dihasilkan sudah diformat sesuai standar dinas (Times New Roman, Layout Tanda Tangan Otomatis).")

                    # 2. Buat File Word
                    doc_file = create_word_doc(
                        title=f"{doc_type} {subject} {grade}",
                        content_markdown=ai_content,
                        teacher_name=teacher_name,
                        head_name=head_name,
                        doc_type=doc_type
                    )

                    filename_safe = f"{doc_type.replace(' ', '_')}_{subject.replace(' ', '_')}_{grade}.docx"

                    st.download_button(
                        label="⬇️ Unduh File Word (.docx)",
                        data=doc_file,
                        file_name=filename_safe,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

st.markdown("---")
st.caption("Dikembangkan oleh **Lagos AI 9.1 (rian dev)** | Powered by **NVIDIA NIM** & **Streamlit**")
