import streamlit as st
import requests
import os
from dotenv import load_dotenv

# --- MUAT KONFIGURASI OTOMATIS ---
load_dotenv()
API_KEY_OTOMATIS = os.getenv("NVIDIA_API_KEY")

# Konfigurasi Halaman
st.set_page_config(
    page_title="Generator KBC 2026 - Fixed",
    page_icon="❤️",
    layout="wide"
)

# --- KONFIGURASI API ---
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- CSS CUSTOM ---
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #FF4B4B; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #555; text-align: center; margin-bottom: 2rem; }
    .stButton>button { width: 100%; background-color: #FF4B4B; color: white; font-weight: bold; border: none; padding: 10px; border-radius: 5px; }
    .stButton>button:hover { background-color: #ff2b2b; }
    .output-box { background-color: #f9f9f9; padding: 25px; border-radius: 10px; border-left: 5px solid #FF4B4B; margin-top: 10px; line-height: 1.6; white-space: pre-wrap; }
    .status-badge { display: inline-block; padding: 5px 10px; border-radius: 15px; font-size: 0.8rem; font-weight: bold; margin-bottom: 10px; }
    .status-ok { background-color: #d4edda; color: #155724; }
    .status-err { background-color: #f8d7da; color: #721c24; }
</style>
""", unsafe_allow_html=True)

# --- FUNGSI API ---
def generate_kbc_content(api_key, system_prompt, user_prompt):
    if not api_key:
        return "❌ Error: API Key tidak ditemukan."

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.7,
        "top_p": 0.9,
        "max_tokens": 4096
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except Exception as e:
        return f"❌ Error Sistem: {str(e)}"

# --- PROMPT SYSTEM ---
SYSTEM_PROMPT = """
Anda adalah Ahli Kurikulum Pendidikan Indonesia spesialis 'Kurikulum Berbasis Cinta (KBC) 2026'.
Filosofi: Integrasi akademik dengan empati, kasih sayang, dan kesejahteraan emosional.
Tugas: Buat perangkat pembelajaran (RPP, ATP, LKPD, dll) yang HOLISTIK, MANUSIAWI, dan SIAP PAKAI.
Format: Gunakan Markdown (#, ##, ###) yang rapi. Langsung berikan isi dokumen tanpa basa-basi.
"""

def build_user_prompt(mapel, kelas, materi, semester, komponen):
    return f"""
    BUATKAN PERANGKAT PEMBELAJARAN KBC 2026 UNTUK:
    - Mapel: {mapel}
    - Kelas: {kelas}
    - Materi: {materi}
    - Semester: {semester}
    - Komponen Diminta: {komponen}

    INSTRUKSI:
    1. ATP: Alur logis & fleksibel.
    2. KKTP: Asesmen membangun, bukan menghakimi.
    3. RPP: Sertakan 'Ice Breaking Cinta' & 'Refleksi Hati'.
    4. LKPD: Aktivitas kolaboratif berbasis empati.
    5. Prota/Promes: Realistis dan manusiawi.

    HASIL HARUS DETAIL DAN SIAP CETAK.
    """

# --- UI UTAMA ---
def main():
    # Cek API Key
    if API_KEY_OTOMATIS:
        status_html = '<span class="status-badge status-ok">✅ API Terhubung Otomatis</span>'
    else:
        status_html = '<span class="status-badge status-err">⚠️ API Key Tidak Ditemukan</span>'

    st.markdown(f"<div style='text-align: right;'>{status_html}</div>", unsafe_allow_html=True)
    st.markdown("<h1 class='main-header'>❤️ Generator KBC 2026</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Perbaikan Error Streamlit | Lagos AI 9.1</p>", unsafe_allow_html=True)

    if not API_KEY_OTOMATIS:
        st.error("🚫 **GAGAL MEMUAT API KEY:** Pastikan file `.env` ada di folder yang sama dan berisi `NVIDIA_API_KEY=...`")
        st.stop()

    # Input Data
    col1, col2 = st.columns([2, 1])
    with col1:
        mapel = st.text_input("Mata Pelajaran", placeholder="Contoh: Matematika")
        materi = st.text_area("Materi Pokok", placeholder="Contoh: Pecahan", height=100)
    with col2:
        kelas = st.selectbox("Kelas / Fase", 
            ["Fase A (SD 1-2)", "Fase B (SD 3-4)", "Fase C (SD 5-6)", 
             "Fase D (SMP 7-9)", "Fase E (SMA 10)", "Fase F (SMA 11-12)"])
        semester = st.selectbox("Semester", ["Ganjil", "Genap"])

    st.divider()

    # PERBAIKAN DI SINI: Menggunakan hanya keyword arguments untuk multiselect
    st.subheader("📦 Pilih Komponen Dokumen")
    options = [
        "CP & ATP (Alur Tujuan)", 
        "KKTP (Kriteria Ketuntasan)", 
        "Prota & Promes (Program Tahun/Semester)", 
        "RPP Lengkap (Modul Ajar)", 
        "LKPD Interaktif (Lembar Kerja)"
    ]

    komponen_list = st.multiselect(
        label="Centang semua yang dibutuhkan:",
        options=options,
        default=["RPP Lengkap", "LKPD Interaktif"]
    )

    if st.button("🚀 Buat Perangkat Ajar Sekarang"):
        if not mapel or not materi:
            st.warning("⚠️ Mohon lengkapi Mata Pelajaran dan Materi Pokok.")
        elif not komponen_list:
            st.warning("⚠️ Mohon pilih minimal satu komponen.")
        else:
            target_komponen = ", ".join(komponen_list)

            with st.spinner('✨ Lagos AI sedang meracik kurikulum... (Tunggu sebentar)'):
                user_prompt = build_user_prompt(mapel, kelas, materi, semester, target_komponen)
                response_text = generate_kbc_content(API_KEY_OTOMATIS, SYSTEM_PROMPT, user_prompt)

                if "Error" in response_text or "❌" in response_text:
                    st.error(response_text)
                else:
                    st.balloons()
                    st.success("✅ Dokumen Berhasil Dibuat!")

                    st.markdown(f"""
                    <div class="output-box">
                    {response_text}
                    </div>
                    """, unsafe_allow_html=True)

                    file_name = f"KBC2026_{mapel.replace(' ', '_')}.md"
                    st.download_button(
                        label="📥 Unduh Dokumen (.md)",
                        data=response_text,
                        file_name=file_name,
                        mime="text/markdown"
                    )

if __name__ == "__main__":
    main()
