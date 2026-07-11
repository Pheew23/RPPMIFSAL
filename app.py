import streamlit as st
import requests
import os
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
import io

# --- KONFIGURASI ---
load_dotenv()
API_KEY = os.getenv("NVIDIA_API_KEY")
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- CSS ---
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #FF4B4B; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #555; text-align: center; margin-bottom: 2rem; }
    .stButton>button { background-color: #FF4B4B; color: white; font-weight: bold; width: 100%; border: none; padding: 10px; border-radius: 5px; }
    .stButton>button:hover { background-color: #ff2b2b; }
    .status-badge { display: inline-block; padding: 5px 10px; border-radius: 15px; font-size: 0.8rem; font-weight: bold; margin-bottom: 10px; }
    .status-ok { background-color: #d4edda; color: #155724; }
    .status-err { background-color: #f8d7da; color: #721c24; }
</style>
""", unsafe_allow_html=True)

# --- FUNGSI AI ---
def generate_content(prompt_text):
    if not API_KEY:
        return None, "API Key tidak ditemukan."

    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": "Anda adalah ahli kurikulum KBC 2026. Buatlah perangkat pembelajaran (RPP, ATP, dll) yang sangat rinci, terstruktur, dan humanis. Gunakan format Markdown dengan header (#, ##) yang jelas untuk setiap bagian."},
            {"role": "user", "content": prompt_text}
        ],
        "temperature": 0.7,
        "max_tokens": 4096
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        content = response.json()['choices'][0]['message']['content']
        return content, None
    except Exception as e:
        return None, str(e)

# --- FUNGSI PEMBUAT WORD ---
def create_word_document(content, filename):
    doc = Document()

    # Style Judul Utama
    heading = doc.add_heading(content.split('\n')[0].replace('#', '').strip(), 0)
    heading.alignment = 1  # Center

    # Process sisa konten
    lines = content.split('\n')
    current_paragraph = None

    for line in lines[1:]: # Skip judul utama yang sudah diambil
        line = line.strip()
        if not line:
            continue

        if line.startswith('## '):
            doc.add_heading(line.replace('##', '').strip(), level=2)
        elif line.startswith('### '):
            doc.add_heading(line.replace('###', '').strip(), level=3)
        elif line.startswith('- ') or line.startswith('* '):
            if not current_paragraph:
                current_paragraph = doc.add_paragraph(style='List Bullet')
            current_paragraph.add_run(line[2:].strip())
            current_paragraph = None # Reset agar item berikutnya jadi paragraf baru jika ada jarak
        else:
            p = doc.add_paragraph(line)
            p.formatting.space_after = Pt(6)
            current_paragraph = p

    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- UI UTAMA ---
def main():
    if not API_KEY:
        st.error("🚫 **ERROR:** File `.env` tidak ditemukan. Pastikan file `.env` ada di folder yang sama.")
        st.stop()

    st.markdown("<h1 class='main-header'>❤️ Generator RPP KBC 2026</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Format Word (.docx) | Lagos AI 9.1</p>", unsafe_allow_html=True)
    st.markdown(f"<div style='text-align: right;'><span class='status-badge status-ok'>✅ API Aktif</span></div>", unsafe_allow_html=True)

    with st.form("kbc_form"):
        col1, col2 = st.columns(2)

        with col1:
            mapel = st.text_input("Mata Pelajaran", placeholder="Contoh: Matematika")
            materi = st.text_area("Materi Pokok", placeholder="Contoh: Pecahan Senilai", height=100)

        with col2:
            # Revisi: Pilihan Kelas Spesifik 1-12
            kelas_options = [f"Kelas {i}" for i in range(1, 13)]
            kelas = st.selectbox("Pilih Kelas", kelas_options)
            semester = st.selectbox("Semester", ["Ganjil", "Genap"])

        st.write("**Komponen Dokumen:**")
        komponen = st.multiselect(
            "", 
            options=["RPP Lengkap", "ATP & CP", "KKTP", "Prota & Promes", "LKPD"],
            default=["RPP Lengkap"]
        )

        submitted = st.form_submit_button("🚀 Buat Dokumen Word")

    if submitted:
        if not mapel or not materi:
            st.warning("⚠️ Mohon lengkapi Mata Pelajaran dan Materi.")
        elif not komponen:
            st.warning("⚠️ Mohon pilih minimal satu komponen.")
        else:
            with st.spinner('✨ AI sedang menyusun dokumen Word yang rapi...'):
                prompt = f"Buatkan {', '.join(komponen)} untuk {mapel} {kelas} materi '{materi}' semester {semester}. Kurikulum KBC 2026. Format harus sangat rapi dengan judul dan sub-judul yang jelas."

                content, error = generate_content(prompt)

                if error:
                    st.error(f"❌ {error}")
                else:
                    st.success("✅ Dokumen berhasil dibuat!")

                    # Tampilkan Preview Markdown
                    with st.expander("Lihat Preview Teks"):
                        st.markdown(content)

                    # Buat File Word
                    doc_buffer = create_word_document(content, f"RPP_{mapel}_{kelas}")

                    # Tombol Download
                    st.download_button(
                        label="📥 Download File Word (.docx)",
                        data=doc_buffer,
                        file_name=f"RPP_KBC2026_{mapel}_{kelas}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

if __name__ == "__main__":
    main()
