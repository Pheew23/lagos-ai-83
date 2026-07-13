import streamlit as st
from openai import OpenAI
import io
import re
import base64
import requests
from docx import Document

# --- FUNGSI ANIMASI LOTTIE (PEMALIS TAMPILAN) ---
def load_lottieurl(url: str):
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            return r.json()
    except:
        pass
    return None

# --- 1. KONFIGURASI HALAMAN & TEMA ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Premium Workspace",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- INJEKSI CUSTOM CSS UNTUK DARK & LIGHT MODE ELEGAN ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700&display=swap');

        /* Font Global */
        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', sans-serif;
        }

        /* Judul Gradasi Neon Modern */
        .main-title {
            font-size: 2.8rem;
            font-weight: 700;
            background: linear-gradient(90deg, #7d4eff, #00d2ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0rem;
        }

        .subtitle {
            font-size: 1.1rem;
            color: #888888;
            margin-bottom: 2rem;
            font-weight: 300;
        }

        /* Glassmorphism Panel untuk Kotak Input/Upload */
        .glass-card {
            background: rgba(255, 255, 255, 0.03);
            border-radius: 16px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            margin-bottom: 15px;
        }

        /* Customisasi Tombol Utama Premium */
        div.stButton > button:first-child {
            background: linear-gradient(90deg, #7d4eff, #5a32d6);
            color: white;
            border: none;
            border-radius: 10px;
            font-weight: 600;
            padding: 12px 24px;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(125, 78, 255, 0.25);
        }
        div.stButton > button:first-child:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(125, 78, 255, 0.45);
            background: linear-gradient(90deg, #8e5bff, #6b42e6);
        }

        /* Penyesuaian Komponen Chat Agar Lebih Kontras di Dua Mode */
        .stChatMessage {
            background-color: rgba(255, 255, 255, 0.02) !important;
            border: 1px solid rgba(255, 255, 255, 0.05) !important;
            border-radius: 12px !important;
            margin-bottom: 10px;
            padding: 15px !important;
        }
    </style>
    """, unsafe_allow_html=True)

# --- API KEY YANG DISEMBATKAN ---
API_KEY = "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"
BASE_URL = "https://integrate.api.nvidia.com/v1"
MODEL_NAME = "mistralai/mistral-large-3-675b-instruct-2512"

Build

Playground

Model Card"

# --- FUNGSI PEMBANTU (LOGIKA UTAMA TETAP SAMA) ---

@st.cache_data(show_spinner=False)
def konversi_gambar_ke_base64(uploaded_file):
    if uploaded_file is not None:
        return base64.b64encode(uploaded_file.read()).decode('utf-8')
    return None

@st.cache_data(show_spinner=False)
def ekstrak_teks_dari_dokumen(uploaded_file):
    teks_hasil = ""
    nama_file = uploaded_file.name.lower()
    try:
        # Jika user menggunakan library PdfReader (pastikan pypdf terinstal jika pakai PDF)
        if nama_file.endswith('.pdf'):
            from pypdf import PdfReader
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                teks = page.extract_text()
                if teks: teks_hasil += teks + "\n"
        elif nama_file.endswith('.txt'):
            teks_hasil = uploaded_file.read().decode("utf-8")
        return teks_hasil.strip()
    except Exception as e:
        st.error(f"Gagal membaca dokumen: {str(e)}")
        return ""

def buat_file_word(riwayat_pesan):
    doc = Document()
    doc.add_heading('Lagos AI - Laporan Analisis', 0)
    for msg in riwayat_pesan:
        if msg["role"] == "system": continue
        role_title = "User" if msg["role"] == "user" else "AI Assistant"
        doc.add_heading(f"{role_title}", level=2)
        content = msg["content"]
        text_content = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        
        lines = text_content.split('\n')
        for line in lines:
            line = line.strip()
            if not line: continue
            if line.startswith('# '): doc.add_heading(line[2:], 3)
            elif line.startswith('- '): doc.add_paragraph(line[2:], style='List Bullet')
            elif re.match(r'^\d+\.\s', line): doc.add_paragraph(line[3:], style='List Number')
            else: doc.add_paragraph(line)
        doc.add_paragraph("\n" + "_"*50)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- POP-UP MODAL UNDUH LAPORAN ---
@st.dialog("🔮 Unduh Laporan Lagos AI")
def tampilkan_popup_unduh(buffer):
    st.success("🎉 Berkas laporan berbasis teks berhasil dirangkum secara otomatis!")
    st.write("Klik tombol premium di bawah ini untuk menyimpan dokumen kerja Anda ke perangkat komputer.")
    st.divider()
    st.download_button(
        label="📥 DOWNLOAD SEKARANG (.DOCX)",
        data=buffer,
        file_name="Lagos_AI_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# --- SIDEBAR: INFO & RESET ---
with st.sidebar:
    # Memuat animasi Lottie pemanis di Sidebar
    from streamlit_lottie import st_lottie
    anim_sidebar = load_lottieurl("https://lottie.host/80e98031-6453-48b4-bb50-bf654c6ee1ff/t3Kx56jU2W.json")
    if anim_sidebar:
        st_lottie(anim_sidebar, height=120)

    st.markdown("### ⚙️ Engine Status")
    st.success(f"🤖 {MODEL_NAME}")
    
    st.divider()
    st.markdown("### 📥 Menu Aksi")
    
    # Tombol Ekspor Rangkuman Obrolan
    if "messages" in st.session_state and len(st.session_state.messages) > 1:
        if st.button("📝 Susun Laporan Ekspor", type="primary", use_container_width=True):
            with st.spinner("Merangkum berkas..."):
                file_word = buat_file_word(st.session_state.messages)
                tampilkan_popup_unduh(file_word)
    else:
        st.info("Ketik obrolan terlebih dahulu untuk membuka menu unduh laporan.")

    st.divider()
    if st.button("🗑️ Bersihkan Memori Chat", type="secondary", use_container_width=True):
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (rian dev), asisten analitik tingkat tinggi. Berikan jawaban cerdas, mendalam, dan terstruktur dalam Bahasa Indonesia."}]
        st.session_state.uploaded_file = None
        st.session_state.uploaded_image = None
        st.rerun()

# --- 2. INISIALISASI SESSION STATE ---
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": "Anda adalah Lagos AI 9.1 (rian dev), asisten analitik tingkat tinggi. Berikan jawaban cerdas, mendalam, dan terstruktur dalam Bahasa Indonesia."}
    ]
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "uploaded_image" not in st.session_state:
    st.session_state.uploaded_image = None

# --- 3. TAMPILAN UTAMA ---
st.markdown('<div class="main-title">🔮 Lagos AI 9.1</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Analisis Multimodal Cerdas: Gambar, Dokumen & Konteks</div>', unsafe_allow_html=True)

# --- 4. PANEL UNGGUH BERKAS MULTIMODAL ---
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
c_upload1, c_upload2 = st.columns(2)

with c_upload1:
    st.markdown("🔒 **Lampirkan Objek Gambar**")
    uploaded_image = st.file_uploader("Upload Image", type=["jpg", "jpeg", "png"], label_visibility="collapsed", key="img_up")
    if uploaded_image: st.session_state.uploaded_image = uploaded_image

with c_upload2:
    st.markdown("📁 **Lampirkan Dokumen Referensi**")
    uploaded_file = st.file_uploader("Upload Doc", type=["pdf", "txt"], label_visibility="collapsed", key="doc_up")
    if uploaded_file: st.session_state.uploaded_file = uploaded_file
st.markdown('</div>', unsafe_allow_html=True)

# Tampilan Status Unggahan (Preview Mini)
if st.session_state.uploaded_image or st.session_state.uploaded_file:
    c_prev1, c_prev2 = st.columns(2)
    with c_prev1:
        if st.session_state.uploaded_image:
            st.image(st.session_state.uploaded_image, caption="📷 Gambar Aktif", width=120)
            if st.button("🗑️ Hapus Gambar", key="del_img"):
                st.session_state.uploaded_image = None
                st.rerun()
    with c_prev2:
        if st.session_state.uploaded_file:
            st.success(f"📄 {st.session_state.uploaded_file.name}")
            if st.button("🗑️ Hapus Dokumen", key="del_doc"):
                st.session_state.uploaded_file = None
                st.rerun()

st.divider()

# --- 5. TAMPILAN RIWAYAT CHAT (MENGGUNAKAN ELEMEN BAWAAN YANG RESPONSIF TEMA) ---
for message in st.session_state.messages:
    if message["role"] == "system":
        continue
    
    with st.chat_message(message["role"]):
        content = message["content"]
        if isinstance(content, list):
            text_disp = next((item["text"] for item in content if item["type"] == "text"), "")
            st.markdown(text_disp)
        else:
            st.markdown(content)

# --- 6. KOLOM INPUT UTAMA ---
prompt = st.chat_input("Tanyakan analisis ke Lagos AI...")

# --- 7. LOGIKA PROSES (VERSI PERBAIKAN: TANPA DUPLIKAT) ---
if prompt:
    teks_perintah = prompt.strip()
    konten_payload = []

    # 1. Masukkan gambar ke request SEKARANG saja
    if st.session_state.uploaded_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.uploaded_image)
        konten_payload.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}})

    # 2. Masukkan konten dokumen jika ada
    teks_dokumen = ""
    if st.session_state.uploaded_file:
        with st.spinner("Membaca dokumen teks..."):
            teks_dokumen = ekstrak_teks_dari_dokumen(st.session_state.uploaded_file)
        if teks_dokumen:
            teks_dokumen = f"[KONTEN DOKUMEN: {st.session_state.uploaded_file.name}]\n{teks_dokumen}\n[AKHIR KONTEN]\n\n"

    final_prompt = teks_dokumen + teks_perintah
    konten_payload.append({"type": "text", "text": final_prompt})

    # Tampilkan ke UI user
    with st.chat_message("user"):
        st.markdown(prompt)

    # SIMPAN KE STATE: Kirim multimodal payload lengkap untuk request saat ini
    st.session_state.messages.append({"role": "user", "content": konten_payload})

    # Response Streaming dari Model AI (Cukup Satu Blok Ini Saja)
    with st.chat_message("assistant"):
        client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
        placeholder = st.empty()
        full_response = ""

        try:
            response_stream = client.chat.completions.create(
                model=MODEL_NAME,
                messages=st.session_state.messages,
                temperature=0.7,
                max_tokens=4096, 
                stream=True
            )

            for chunk in response_stream:
                if chunk.choices and len(chunk.choices) > 0:
                    delta = chunk.choices[0].delta.content
                    if delta:
                        full_response += delta
                        placeholder.markdown(full_response + "▌")

            placeholder.markdown(full_response)
            
            # --- STRATEGI PENYELAMAT MEMORI ---
            # Mengubah payload user terakhir menjadi TEKS SAJA agar chat berikutnya ringan
            st.session_state.messages[-1] = {"role": "user", "content": f"[User menanyakan gambar/dokumen] {prompt}"}
            
            # Masukkan respon AI ke dalam riwayat
            st.session_state.messages.append({"role": "assistant", "content": full_response})

            # Otomatis kosongkan file uploader di UI setelah sukses dianalisis
            st.session_state.uploaded_image = None
            st.session_state.uploaded_file = None
            
            # Memicu rerun agar tampilan UI pembatalan upload langsung sinkron
            st.rerun()

        except Exception as e:
            st.error(f"Terjadi kesalahan teknis API: {str(e)}")
            # Jika error, hapus input terakhir agar tidak merusak state
            if st.session_state.messages[-1]["role"] == "user":
                st.session_state.messages.pop()

    # Response Streaming dari Model AI
    with st.chat_message("assistant"):
        client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
        placeholder = st.empty()
        full_response = ""

        try:
            response_stream = client.chat.completions.create(
                model=MODEL_NAME,
                messages=st.session_state.messages,
                temperature=0.7,
                max_tokens=8096,
                stream=True
            )

            for chunk in response_stream:
                if chunk.choices and len(chunk.choices) > 0:
                    delta = chunk.choices[0].delta.content
                    if delta:
                        full_response += delta
                        placeholder.markdown(full_response + "▌")

            placeholder.markdown(full_response)
            st.session_state.messages.append({"role": "assistant", "content": full_response})

        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {str(e)}")
