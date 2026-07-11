import streamlit as st
import requests
import os
from dotenv import load_dotenv

# 1. Load Environment Variables
load_dotenv()
API_KEY = os.getenv("NVIDIA_API_KEY")

# 2. Konfigurasi Halaman
st.set_page_config(page_title="Generator KBC 2026", page_icon="❤️", layout="wide")

# 3. CSS Styling
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #FF4B4B; text-align: center; }
    .stButton>button { background-color: #FF4B4B; color: white; font-weight: bold; width: 100%; }
    .output-box { background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #FF4B4B; white-space: pre-wrap; }
</style>
""", unsafe_allow_html=True)

# 4. Fungsi API
def call_ai(prompt_text):
    if not API_KEY:
        return "ERROR: API Key tidak ditemukan di file .env"

    url = "https://integrate.api.nvidia.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "qwen/qwen3.5-397b-a17b",
        "messages": [
            {"role": "system", "content": "Anda adalah ahli kurikulum KBC 2026 (Kurikulum Berbasis Cinta). Buatlah RPP, ATP, LKPD yang humanis, penuh empati, dan mendetail. Gunakan format Markdown."},
            {"role": "user", "content": prompt_text}
        ],
        "temperature": 0.7,
        "max_tokens": 4096
    }

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        return f"Error API: {str(e)}"

# 5. UI Utama
def main():
    # Cek API Key di awal
    if not API_KEY:
        st.error("🚫 **CRITICAL ERROR:** File `.env` tidak ditemukan atau kosong. Pastikan file `.env` berisi `NVIDIA_API_KEY=...`")
        st.stop()

    st.markdown("<h1 class='main-header'>❤️ Generator RPP KBC 2026</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center'>Oleh Lagos AI 9.1 (Rian Dev)</p>")

    # Input Form
    with st.form("kbc_form"):
        col1, col2 = st.columns(2)
        with col1:
            mapel = st.text_input("Mata Pelajaran")
            materi = st.text_area("Materi Pokok", height=100)
        with col2:
            kelas = st.selectbox("Kelas", ["Fase A", "Fase B", "Fase C", "Fase D", "Fase E", "Fase F"])
            semester = st.selectbox("Semester", ["Ganjil", "Genap"])

        # PERBAIKAN UTAMA DI SINI: Menggunakan syntax paling aman untuk multiselect
        st.write("**Pilih Dokumen:**")
        komponen = st.multiselect(
            "", # Label kosong karena kita sudah pakai st.write di atasnya
            options=["RPP Lengkap", "ATP & CP", "KKTP", "Prota & Promes", "LKPD"],
            default=["RPP Lengkap", "LKPD"]
        )

        submitted = st.form_submit_button("🚀 Generate Sekarang")

    if submitted:
        if not mapel or not materi:
            st.warning("Mohon isi Mata Pelajaran dan Materi.")
        elif not komponen:
            st.warning("Mohon pilih minimal satu dokumen.")
        else:
            with st.spinner("Sedang menyusun perangkat pembelajaran berbasis cinta..."):
                prompt = f"Buatkan {', '.join(komponen)} untuk {mapel} kelas {kelas} materi {materi} semester {semester} dengan pendekatan KBC 2026."
                result = call_ai(prompt)

                if "Error" in result:
                    st.error(result)
                else:
                    st.success("Selesai!")
                    st.markdown(f"<div class='output-box'>{result}</div>", unsafe_allow_html=True)

                    st.download_button(
                        label="📥 Download (.md)",
                        data=result,
                        file_name=f"RPP_{mapel}.md",
                        mime="text/markdown"
                    )

if __name__ == "__main__":
    main()
