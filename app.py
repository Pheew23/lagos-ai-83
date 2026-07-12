import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document

# --- 1. KONFIGURASI HALAMAN & TEMA ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Advanced Workspace",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- API KEY YANG DISEMBATKAN (HANYA UNTUK PENGGUNAAN PRIBADI) ---
# ⚠️ JANGAN BAGIKAN KODE INI KE ORANG LAIN ATAU UPLOAD KE GITHUB
API_KEY = "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"
BASE_URL = "https://integrate.api.nvidia.com/v1"
MODEL_NAME = "qwen/qwen3.5-397b-a17b"

# --- CUSTOM CSS UNTUK UI/UX MODERN ---
def load_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

        .stApp {
            background-color: #0e1117;
            font-family: 'Inter', sans-serif;
        }

        .chat-message {
            padding: 1.5rem;
            border-radius: 1rem;
            margin-bottom: 1rem;
            display: flex;
            align-items: flex-start;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
            border: 1px solid #262730;
        }

        .chat-message.user {
            background-color: #2b2d31;
            border-left: 4px solid #4CAF50;
        }

        .chat-message.assistant {
            background-color: #1a1a1a;
            border-left: 4px solid #7d4eff;
        }

        .chat-avatar {
            font-size: 1.5rem;
            margin-right: 1rem;
            line-height: 1.2;
        }

        .chat-content {
            flex: 1;
            color: #e0e0e0;
            line-height: 1.6;
        }

        .main-title {
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(90deg, #7d4eff, #00d2ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }

        .subtitle {
            font-size: 1.1rem;
            color: #888;
            margin-bottom: 2rem;
            font-weight: 300;
        }

        .upload-box {
            background: #161b22;
            border: 2px dashed #30363d;
            border-radius: 1rem;
            padding: 2rem;
            text-align: center;
            transition: all 0.3s ease;
        }
        .upload-box:hover {
            border-color: #7d4eff;
            background: #1c2128;
        }

        section[data-testid="stSidebar"] {
            background-color: #0d0f14;
            border-right: 1px solid #262730;
        }

        .stButton>button {
            background: linear-gradient(90deg, #7d4eff, #5a32d6);
            color: white;
            border: none;
            border-radius: 0.5rem;
            font-weight: 600;
            transition: transform 0.2s;
        }
        .stButton>button:hover {
            transform: scale(1.02);
            background: linear-gradient(90deg, #8e5bff, #6b42e6);
        }

        .danger-btn > button {
            background: #3a1c1c;
            color: #ff6b6b;
            border: 1px solid #ff6b6b;
        }
        .danger-btn > button:hover {
            background: #ff6b6b;
            color: white;
        }
    </style>
    """, unsafe_allow_html=True)

load_custom_css()

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
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                teks = page.extract_text()
                if teks:
                    teks_hasil += teks + "\n"
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
        if isinstance(content, list):
            text_content = next((item["text"] for item in content if item["type"] == "text"), "")
        else:
            text_content = str(content)

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

# --- SIDEBAR: INFO & RESET ---
with st.sidebar:
    st.markdown("### ⚙️ Konfigurasi")
    st.success(f"✅ Model: {MODEL_NAME}")
    st.info("API Key sudah dikonfigurasi otomatis.")

    st.divider()
    st.markdown("### 📦 Ekspor Data")

    if "messages" in st.session_state and len(st.session_state.messages) > 1:
        with st.spinner("📝 Menyusun dokumen..."):
            file_word = buat_file_word(st.session_state.messages)
            st.download_button(
                label="📥 Download Laporan (.docx)",
                data=file_word,
                file_name="Lagos_AI_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
    else:
        st.info("Mulai percakapan untuk mengunduh laporan.")

    st.divider()
    if st.button("🗑️ Reset Memori", type="secondary", use_container_width=True):
        st.session_state.messages = []
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

# --- 4. AREA INPUT & UPLOAD ---
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("#### 📷 Unggah Gambar")
    uploaded_image = st.file_uploader(
        "Pilih file JPG/PNG", 
        type=["jpg", "jpeg", "png"],
        label_visibility="collapsed",
        key="img_uploader"
    )
    if uploaded_image:
        st.session_state.uploaded_image = uploaded_image

with col2:
    st.markdown("#### 📄 Unggah Dokumen")
    uploaded_file = st.file_uploader(
        "Pilih file PDF/TXT", 
        type=["pdf", "txt"],
        label_visibility="collapsed",
        key="doc_uploader"
    )
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file

# Tampilan Preview File
if st.session_state.uploaded_image or st.session_state.uploaded_file:
    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.session_state.uploaded_image:
            st.image(st.session_state.uploaded_image, caption="Gambar Terlampir", use_container_width=True, width=100)
            if st.button("Hapus Gambar", key="del_img_btn"):
                st.session_state.uploaded_image = None
                st.rerun()
    with c2:
        if st.session_state.uploaded_file:
            st.info(f"📄 {st.session_state.uploaded_file.name}")
            if st.button("Hapus Dokumen", key="del_doc_btn"):
                st.session_state.uploaded_file = None
                st.rerun()

# --- 5. CHAT HISTORY ---
st.divider()

for i, message in enumerate(st.session_state.messages):
    if message["role"] == "system":
        continue

    role = message["role"]
    content = message["content"]

    display_text = ""
    if isinstance(content, list):
        for item in content:
            if item["type"] == "text":
                display_text += item["text"] + "\n"
    else:
        display_text = content

    with st.container():
        css_class = "user" if role == "user" else "assistant"
        icon = "👤" if role == "user" else "🤖"

        st.markdown(f"""
        <div class="chat-message {css_class}">
            <div class="chat-avatar">{icon}</div>
            <div class="chat-content">
                {display_text.replace(chr(10), '<br>')}
            </div>
        </div>
        """, unsafe_allow_html=True)

# --- 6. INPUT AREA ---
st.divider()
prompt = st.chat_input("Ketik pertanyaan Anda di sini... (Contoh: 'Analisis data di dokumen ini')")

# --- 7. LOGIKA PROSES ---
if prompt:
    teks_perintah = prompt.strip()

    konten_payload = []

    # Proses Gambar
    if st.session_state.uploaded_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.uploaded_image)
        konten_payload.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}})

    # Proses Dokumen
    teks_dokumen = ""
    if st.session_state.uploaded_file:
        with st.spinner("📂 Membaca konten dokumen..."):
            teks_dokumen = ekstrak_teks_dari_dokumen(st.session_state.uploaded_file)

        if teks_dokumen:
            teks_dokumen = f"[ISI DOKUMEN {st.session_state.uploaded_file.name}]\n{teks_dokumen}\n[AKHIR DOKUMEN]\n\n"

    final_prompt = teks_dokumen + teks_perintah
    konten_payload.append({"type": "text", "text": final_prompt})

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
            st.error(f"Terjadi kesalahan: {str(e)}")
            st.session_state.messages.pop()
