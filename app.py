import streamlit as st
from openai import OpenAI
import io
import re
import base64
import requests
from docx import Document
from streamlit_lottie import st_lottie

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI | Chat Interface",
    page_icon="✨",
    layout="centered", # Menggunakan centered agar chatbox tidak terlalu melebar (seperti Gemini)
    initial_sidebar_state="collapsed" # Sidebar disembunyikan secara default agar lebih bersih
)

# --- 2. CUSTOM CSS (GAYA CLEAN & MINIMALIS) ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

        /* Font & Background meniru gaya modern yang bersih */
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        /* Menyembunyikan elemen default Streamlit yang mengganggu */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Mempercantik bubble chat AI agar lebih soft */
        .stChatMessage:nth-child(even) {
            background-color: rgba(255, 255, 255, 0.03) !important;
            border-radius: 12px;
            padding: 1rem;
        }
        
        /* Pill indikator file yang diunggah */
        .file-pill {
            display: inline-block;
            background: rgba(125, 78, 255, 0.2);
            color: #b59bf5;
            padding: 4px 12px;
            border-radius: 16px;
            font-size: 0.8rem;
            margin-right: 8px;
            margin-bottom: 15px;
            border: 1px solid rgba(125, 78, 255, 0.4);
        }
    </style>
""", unsafe_allow_html=True)

# --- KONFIGURASI API ---
API_KEY = "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" # Pindahkan ke st.secrets saat rilis
BASE_URL = "https://integrate.api.nvidia.com/v1"
MODEL_NAME = "mistralai/mistral-large-3-675b-instruct-2512"

# --- FUNGSI PEMBANTU ---
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
        return ""

def buat_file_word(riwayat_pesan):
    doc = Document()
    doc.add_heading('Lagos AI - Chat Export', 0)
    for msg in riwayat_pesan:
        if msg["role"] == "system": continue
        role_title = "User" if msg["role"] == "user" else "Lagos AI"
        doc.add_heading(f"{role_title}", level=2)
        content = msg["content"]
        text_content = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        
        for line in text_content.split('\n'):
            line = line.strip()
            if not line: continue
            if line.startswith('# '): doc.add_heading(line[2:], 3)
            elif line.startswith('- '): doc.add_paragraph(line[2:], style='List Bullet')
            else: doc.add_paragraph(line)
        doc.add_paragraph("\n" + "_"*40 + "\n")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 3. INISIALISASI SESSION STATE ---
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "system", "content": "Anda adalah asisten AI yang cerdas dan terstruktur."}]
if "temp_image" not in st.session_state:
    st.session_state.temp_image = None
if "temp_doc" not in st.session_state:
    st.session_state.temp_doc = None

# --- SIDEBAR (UNTUK FITUR EKSPOR & IKLAN) ---
with st.sidebar:
    st.markdown("### ⚙️ Lagos AI Settings")
    if st.button("🗑️ Chat Baru", use_container_width=True):
        st.session_state.messages = [{"role": "system", "content": "Anda adalah asisten AI yang cerdas dan terstruktur."}]
        st.rerun()
    
    st.divider()
    if len(st.session_state.messages) > 1:
        file_word = buat_file_word(st.session_state.messages)
        st.download_button(
            label="📥 Ekspor Obrolan (.DOCX)",
            data=file_word,
            file_name="Lagos_AI_Export.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    
    # Ruang Iklan Sidebar (Tidak mengganggu chat utama)
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.caption("Sponsored")
    st.markdown('<div style="background:rgba(255,255,255,0.05); height:150px; border-radius:8px; display:flex; align-items:center; justify-content:center; border:1px dashed #555;">[Ad Space]</div>', unsafe_allow_html=True)

# --- 4. AREA OBROLAN UTAMA ---
# Header sapaan jika chat masih kosong
if len(st.session_state.messages) == 1:
    st.markdown("<h1 style='text-align: center; margin-top: 10vh; font-weight: 600; background: -webkit-linear-gradient(45deg, #7d4eff, #00d2ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>Halo, ada yang bisa saya bantu hari ini?</h1>", unsafe_allow_html=True)

# Render Riwayat Chat
for message in st.session_state.messages:
    if message["role"] == "system": continue
    with st.chat_message(message["role"]):
        content = message["content"]
        text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        st.markdown(text_disp)

# Spasi kosong agar input tidak menutupi chat terakhir
st.markdown("<div style='height: 80px'></div>", unsafe_allow_html=True)

# --- 5. AREA INPUT TERPADU (BOTTOM ANCHORED) ---
# Menggunakan container untuk merapikan tombol lampiran dan kolom input
input_container = st.container()

with input_container:
    # Menampilkan indikator jika ada file yang siap dikirim
    if st.session_state.temp_image:
        st.markdown(f"<div class='file-pill'>📷 Gambar siap dilampirkan</div>", unsafe_allow_html=True)
    if st.session_state.temp_doc:
        st.markdown(f"<div class='file-pill'>📄 Dokumen siap dilampirkan</div>", unsafe_allow_html=True)

    # Baris yang memuat tombol lampiran (Popover) dan kolom obrolan
    col_attach, col_input = st.columns([1, 10])
    
    with col_attach:
        # Tombol Popover gaya Gemini untuk Attachments
        with st.popover("📎"):
            st.markdown("**Lampirkan File**")
            up_img = st.file_uploader("Gambar", type=["jpg", "png", "jpeg"], label_visibility="collapsed")
            up_doc = st.file_uploader("Dokumen", type=["pdf", "txt"], label_visibility="collapsed")
            
            if up_img: st.session_state.temp_image = up_img
            if up_doc: st.session_state.temp_doc = up_doc

    with col_input:
        prompt = st.chat_input("Tanya Lagos AI atau instruksikan file lampiran...")

# --- 6. LOGIKA PEMROSESAN UTAMA ---
if prompt:
    # Siapkan payload
    konten_payload = []
    
    if st.session_state.temp_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.temp_image)
        konten_payload.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}})
    
    teks_dokumen = ""
    if st.session_state.temp_doc:
        with st.spinner("Membaca dokumen..."):
            teks_dokumen = ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
        if teks_dokumen:
            teks_dokumen = f"[KONTEN DOKUMEN: {st.session_state.temp_doc.name}]\n{teks_dokumen}\n[AKHIR KONTEN]\n\n"

    final_prompt = teks_dokumen + prompt
    konten_payload.append({"type": "text", "text": final_prompt})

    # Tampilkan prompt pengguna di UI
    with st.chat_message("user"):
        st.markdown(prompt)

    # Simpan ke riwayat
    st.session_state.messages.append({"role": "user", "content": konten_payload})

    # Proses AI
    with st.chat_message("assistant"):
        client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
        placeholder = st.empty()
        full_response = ""

        try:
            response_stream = client.chat.completions.create(
                model=MODEL_NAME,
                messages=st.session_state.messages,
                temperature=0.2,
                max_tokens=2096,
                stream=True
            )

            for chunk in response_stream:
                if chunk.choices and len(chunk.choices) > 0:
                    delta = chunk.choices[0].delta.content
                    if delta:
                        full_response += delta
                        placeholder.markdown(full_response + "▌")

            placeholder.markdown(full_response)
            
            # Hemat memori: Ubah riwayat pengguna menjadi teks saja setelah diproses
            st.session_state.messages[-1] = {"role": "user", "content": prompt}
            st.session_state.messages.append({"role": "assistant", "content": full_response})

            # Bersihkan cache file setelah berhasil dikirim
            st.session_state.temp_image = None
            st.session_state.temp_doc = None
            
            st.rerun()

        except Exception as e:
            st.error(f"Kesalahan koneksi AI: {str(e)}")
            if st.session_state.messages[-1]["role"] == "user":
                st.session_state.messages.pop()
