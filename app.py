import streamlit as st
import requests
import os
from dotenv import load_dotenv

# --- MUAT KONFIGURASI OTOMATIS ---
# Ini akan membaca file .env di folder yang sama secara otomatis
load_dotenv()

# Ambil API Key dari lingkungan sistem
API_KEY_OTOMATIS = os.getenv("NVIDIA_API_KEY")

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Generator KBC 2026 - Otomatis",
    page_icon="❤️",
    layout="wide"
)

# --- KONFIGURASI API & MODEL ---
API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- CSS CUSTOM ---
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #FF4B4B;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #555;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #FF4B4B;
        color: white;
        font-weight: bold;
        border: none;
        padding: 10px;
        border-radius: 5px;
    }
    .stButton>button:hover {
        background-color: #ff2b2b;
    }
    .output-box {
        background-color: #f9f9f9;
        padding: 25px;
        border-radius: 10px;
        border-left: 5px solid #FF4B4B;
        margin-top: 10px;
        line-height: 1.6;
    }
    .status-badge {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 15px;
        font-size: 0.8rem;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .status-ok { background-color: #d4edda; color: #155724; }
    .status-err { background-color: #f8d7da; color: #721c24; }
</style>
""", unsafe_allow_html=True)

# --- FUNGSI PEMANGGILAN AI ---
def generate_kbc_content(api_key, prompt_system, prompt_user):
    if not api_key:
        return "❌ Error: API Key tidak ditemukan. Pastikan file .env sudah dibuat dengan benar."

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": prompt_system},
            {"role": "user", "content": prompt_user}
        ],
        "temperature": 0.7,
        "top_p": 0.9,
        "max_tokens": 4096
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except requests.exceptions.Timeout:
        return "⏱️ Timeout: Permintaan memakan waktu terlalu lama. Coba lagi dengan materi yang lebih spesifik."
    except requests.exceptions.RequestException as e:
        return f"❌ Error Koneksi: {str(e)}"
    except KeyError as e:
        return f"❌ Error Format Respons: {str(e)}"

# --- PROMPT ENGINEERING KBC 2026 ---
SYSTEM_PROMPT = """
Anda adalah Ahli Kurikulum Pendidikan Indonesia tingkat nasional yang spesialis dalam 'Kurikulum Berbasis Cinta (KBC) 2026'.
Filosofi KBC 2026: Pembelajaran mengintegrasikan kecakapan akademik dengan kecerdasan emosional, empati, cinta lingkungan, dan kasih sayang.

Tugas Anda:
1. Hasilkan perangkat pembelajaran yang HOLISTIK, MANUSIAWI, dan SIAP PAKAI.
2. Bahasa: Indonesia baku, hangat, dan inspiratif.
3. Struktur: Gunakan Markdown (#, ##, ###) untuk keterbacaan maksimal.
4. Konten: Pastikan setiap bagian (ATP, KKTP, LKPD) memiliki sentuhan 'Cinta' (misal: refleksi perasaan, kerja sama kelompok yang inklusif).
5. Langsung berikan dokumen akhir tanpa basa-basi pembuka.
"""

def build_user_prompt(mapel, kelas, materi, semester, komponen):
    return f"""
    BUATKAN PERANGKAT PEMBELAJARAN KBC 2026 LENGKAP UNTUK:
    - Mata Pelajaran: {mapel}
    - Kelas/Fase: {kelas}
    - Materi Pokok: {materi}
    - Semester: {semester}

    KOMPONEN YANG DIMINTA: {komponen}

    INSTRUKSI SPESIFIK:
    - Jika meminta ATP: Sertakan alur yang logis namun fleksibel terhadap kondisi emosional siswa.
    - Jika meminta KKTP: Gunakan asesmen formatif yang membangun, bukan menghakimi.
    - Jika meminta RPP: Sertakan skenario 'Ice Breaking Cinta' dan 'Refleksi Hati' di langkah pembelajaran.
    - Jika meminta LKPD: Desain aktivitas kolaboratif yang memecahkan masalah nyata dengan empati.

    HASIL HARUS SANGAT DETAIL, TERSTRUKTUR, DAN SIAP DICETAK OLEH GURU.
    """

# --- UI UTAMA ---
def main():
    # Cek Status API Key
    if API_KEY_OTOMATIS:
        status_html = '<span class="status-badge status-ok">✅ API Terhubung Otomatis</span>'
    else:
        status_html = '<span class="status-badge status-err">⚠️ API Key Tidak Ditemukan (Cek file .env)</span>'

    st.markdown(f"<div style='text-align: right;'>{status_html}</div>", unsafe_allow_html=True)
    st.markdown("<h1 class='main-header'>❤️ Generator Perangkat Ajar KBC 2026</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Otomatis & Lengkap | Powered by Qwen 3.5 via NVIDIA NIM</p>", unsafe_allow_html=True)

    if not API_KEY_OTOMATIS:
        st.error("🚫 **GAGAL MEMUAT API KEY:** Sistem tidak menemukan file `.env` atau kunci API di dalamnya. Harap hubungi administrator untuk membuat file `.env` dengan kunci yang valid.")
        st.stop()

    # Layout Kolom untuk Input
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📝 Detail Pembelajaran")
        mapel = st.text_input("Mata Pelajaran", placeholder="Contoh: Matematika, IPA, Sejarah")
        materi = st.text_area("Materi Pokok / Topik", placeholder="Contoh: Pecahan Senilai, Rantai Makanan, Perjuangan Kemerdekaan", height=100)

    with col2:
        st.subheader("⚙️ Pengaturan")
        kelas = st.selectbox("Kelas / Fase", 
            ["Fase A (SD 1-2)", "Fase B (SD 3-4)", "Fase C (SD 5-6)", 
             "Fase D (SMP 7-9)", "Fase E (SMA 10)", "Fase F (SMA 11-12)"])
        semester = st.selectbox("Semester", ["Ganjil", "Genap"])

    st.divider()

    st.subheader("📦 Pilih Komponen Dokumen")
    komponen_list = st.multiselect(
        "Centang semua yang dibutuhkan:",
        ["CP & ATP (Alur Tujuan)", "KKTP (Kriteria Ketuntasan)", 
         "Prota & Promes (Program Tahun/Semester)", 
         "RPP Lengkap (Modul Ajar)", 
         "LKPD Interaktif (Lembar Kerja)"],
        default=["RPP Lengkap", "LKPD Interaktif"]
    )

    if st.button("🚀 Buat Perangkat Ajar Sekarang"):
        if not mapel or not materi:
            st.warning("⚠️ Mohon lengkapi Mata Pelajaran dan Materi Pokok.")
        elif not komponen_list:
            st.warning("⚠️ Mohon pilih minimal satu komponen dokumen.")
        else:
            target_komponen = ", ".join(komponen_list)

            with st.spinner('✨ Lagos AI sedang meracik kurikulum berbasis cinta untuk Anda... (Ini mungkin memakan waktu 10-20 detik)'):
                user_prompt = build_user_prompt(mapel, kelas, materi, semester, target_komponen)

                # Panggil API dengan kunci otomatis
                response_text = generate_kbc_content(API_KEY_OTOMATIS, SYSTEM_PROMPT, user_prompt)

                if "Error" in response_text or "❌" in response_text:
                    st.error(response_text)
                else:
                    st.balloons()
                    st.success("✅ Dokumen Berhasil Dibuat!")

                    # Tampilan Hasil
                    st.markdown(f"""
                    <div class="output-box">
                    {response_text}
                    </div>
                    """, unsafe_allow_html=True)

                    # Tombol Download
                    file_name = f"KBC2026_{mapel.replace(' ', '_')}_{kelas.replace(' ', '_')}.md"
                    st.download_button(
                        label="📥 Unduh Dokumen Lengkap (.md)",
                        data=response_text,
                        file_name=file_name,
                        mime="text/markdown"
                    )

    # Footer
    st.markdown("---")
    st.caption("Dikembangkan oleh Lagos AI 9.1 (Rian Dev) | Sistem Otomatisasi API Aktif")

if __name__ == "__main__":
    main()
