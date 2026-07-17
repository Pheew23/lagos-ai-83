import streamlit as st
from openai import OpenAI
import io
import re
import base64
import requests
from docx import Document

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Premium Chat",
    page_icon="🔮",
    layout="centered", 
    initial_sidebar_state="collapsed"
)

# --- 2. CUSTOM CSS (GAYA CLEAN & BRANDING LAGOS) ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', sans-serif;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Branding Lagos AI di Tengah Atas */
        .header-title {
            text-align: center;
            font-size: 2.2rem;
            font-weight: 700;
            background: linear-gradient(90deg, #7d4eff, #00d2ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0px;
            padding-top: 10px;
        }
        .header-subtitle {
            text-align: center;
            color: #888888;
            font-size: 0.95rem;
            font-weight: 300;
            margin-bottom: 30px;
        }
        
        .stChatMessage:nth-child(even) {
            background-color: rgba(255, 255, 255, 0.03) !important;
            border-radius: 12px;
            padding: 1rem;
        }
        
        .file-pill {
            display: inline-block;
            background: rgba(125, 78, 255, 0.15);
            color: #b59bf5;
            padding: 4px 14px;
            border-radius: 20px;
            font-size: 0.8rem;
            margin-right: 8px;
            margin-bottom: 12px;
            border: 1px solid rgba(125, 78, 255, 0.3);
        }
    </style>
""", unsafe_allow_html=True)

# --- KONFIGURASI API ---
# Mengambil API Key secara aman dari st.secrets
API_KEY = st.secrets["NVIDIA_API_KEY"] 
BASE_URL = "https://integrate.api.nvidia.com/v1"

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
    doc.add_heading('Lagos AI 9.1 - Analisis Laporan', 0)
    for msg in riwayat_pesan:
        if msg["role"] == "system": continue
        role_title = "User" if msg["role"] == "user" else "Lagos AI 9.1"
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
    st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
if "temp_image" not in st.session_state:
    st.session_state.temp_image = None
if "temp_doc" not in st.session_state:
    st.session_state.temp_doc = None

# --- BRANDING UTAMA (SELALU TAMPIL DI ATAS) ---
st.markdown('<div class="header-title">🔮 Lagos AI 9.1</div>', unsafe_allow_html=True)
st.markdown('<div class="header-subtitle">Premium Multimodal Assistant</div>', unsafe_allow_html=True)

# --- SIDEBAR (TANPA IKLAN, MURNI UTILITIES) ---
with st.sidebar:
    st.markdown("### ⚙️ Engine Status")
    st.success("🤖 Lagos AI 9.1 Active")
    
    # Komponen Pemilihan Model AI
    st.markdown("### 🧠 Pilih Model AI")
    
    # Mapping nama singkat untuk UI ke ID asli untuk API
    MODEL_MAPPING = {
        "google/gemma-4-31b-it": "1. Stabil",
        "thinkingmachines/inkling": "2. Cepat",
        "mistralai/mistral-medium-3.5-128b": "3. Analisis Mendalam",
        "openai/gpt-oss-120b": "4. Sangat Cepat",
        "nvidia/nemotron-3-ultra-550b-a55b": "5. Projek Khusus"
    }
    
    MODEL_NAME = st.selectbox(
        label="Pilih model aktif:",
        options=list(MODEL_MAPPING.keys()),
        index=1, # Default ke mistral-small-4-119b-2603
        format_func=lambda x: MODEL_MAPPING[x], # Merubah visual teks menjadi nama singkat saja
        label_visibility="collapsed"
    )
    
    st.divider()
    if st.button("🗑️ Bersihkan Memori Chat", type="secondary", use_container_width=True):
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        st.rerun()
    
    if len(st.session_state.messages) > 1:
        file_word = buat_file_word(st.session_state.messages)
        st.download_button(
            label="📥 Unduh Laporan (.DOCX)",
            data=file_word,
            file_name="Lagos_AI_9.1_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary"
        )

# --- 4. AREA OBROLAN UTAMA ---
# Pesan selamat datang yang elegan jika baru memulai
if len(st.session_state.messages) == 1:
    st.markdown("<p style='text-align: center; margin-top: 5vh; color: #666;'>Sistem siap. Lampirkan gambar/dokumen atau ketik pertanyaan Anda di bawah.</p>", unsafe_allow_html=True)

for message in st.session_state.messages:
    if message["role"] == "system": continue
    with st.chat_message(message["role"]):
        content = message["content"]
        text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        st.markdown(text_disp)

st.markdown("<div style='height: 80px'></div>", unsafe_allow_html=True)

# --- 5. AREA INPUT TERPADU ---
input_container = st.container()

with input_container:
    if st.session_state.temp_image:
        st.markdown(f"<div class='file-pill'>📷 Gambar telah dilampirkan</div>", unsafe_allow_html=True)
    if st.session_state.temp_doc:
        st.markdown(f"<div class='file-pill'>📄 Dokumen telah dilampirkan</div>", unsafe_allow_html=True)

    col_attach, col_input = st.columns([1, 10])
    
    with col_attach:
        with st.popover("📎"):
            st.markdown("**Lampirkan File**")
            up_img = st.file_uploader("Upload Image", type=["jpg", "png", "jpeg"], label_visibility="collapsed")
            up_doc = st.file_uploader("Upload Doc", type=["pdf", "txt"], label_visibility="collapsed")
            
            if up_img: st.session_state.temp_image = up_img
            if up_doc: st.session_state.temp_doc = up_doc

    with col_input:
        prompt = st.chat_input("Tanyakan sesuatu pada Lagos AI 9.1...")

# ... (Kode bagian 1 sampai 5 tetap sama) ...

# --- 6. LOGIKA PEMROSESAN ---
if prompt:
    teks_dokumen = ""
    if st.session_state.temp_doc:
        with st.spinner("Membaca referensi dokumen..."):
            teks_dokumen = ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
        if teks_dokumen:
            teks_dokumen = f"[KONTEN DOKUMEN: {st.session_state.temp_doc.name}]\n{teks_dokumen}\n[AKHIR KONTEN]\n\n"

    final_prompt = teks_dokumen + prompt

    # PERBAIKAN: Bedakan format payload untuk teks biasa vs multimodal (gambar)
    if st.session_state.temp_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.temp_image)
        konten_payload = [
            {"type": "text", "text": final_prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
        ]
    else:
        # Kirim sebagai teks String biasa jika tidak ada gambar
        konten_payload = final_prompt 

    with st.chat_message("user"):
        st.markdown(prompt)

    st.session_state.messages.append({"role": "user", "content": konten_payload})

    with st.chat_message("assistant"):
        client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
        placeholder = st.empty()
        full_response = ""

        try:
            response_stream = client.chat.completions.create(
                model=MODEL_NAME, 
                messages=st.session_state.messages,
                temperature=0.4,
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
            
            # Format ulang log chat user menjadi String standar di riwayat setelah diproses
            st.session_state.messages[-1] = {"role": "user", "content": f"[User Query] {prompt}"}
            st.session_state.messages.append({"role": "assistant", "content": full_response})

            st.session_state.temp_image = None
            st.session_state.temp_doc = None
            
            st.rerun()

        except Exception as e:
            st.error(f"Kesalahan teknis pada engine Lagos AI: {str(e)}")
            if st.session_state.messages[-1]["role"] == "user":
                st.session_state.messages.pop()
