import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document
from pypdf import PdfReader

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI 8.3 Workspace",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. CSS UNTUK TAMPILAN ELEGAN & MINIMALIS ---
st.markdown("""
<style>
    /* Font Modern & Bersih */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    :root {
        --bg-color: #0f1115;
        --card-bg: #161b22;
        --text-main: #e6edf3;
        --text-muted: #8b949e;
        --accent-color: #58a6ff; /* Biru soft */
        --border-color: #30363d;
        --input-bg: #0d1117;
    }

    body {
        background-color: var(--bg-color);
        color: var(--text-main);
        font-family: 'Inter', sans-serif;
        background-image: radial-gradient(circle at 10% 20%, rgba(88, 166, 255, 0.05) 0%, transparent 20%);
    }

    /* Container Login */
    .login-container {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 90vh;
        flex-direction: column;
    }

    .login-card {
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        padding: 3rem;
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        width: 100%;
        max-width: 380px;
        text-align: center;
        transition: transform 0.3s ease;
    }

    .login-card:hover {
        transform: translateY(-5px);
        border-color: var(--accent-color);
    }

    .login-title {
        font-weight: 700;
        font-size: 1.8rem;
        margin-bottom: 0.5rem;
        color: var(--text-main);
    }

    .login-subtitle {
        color: var(--text-muted);
        font-size: 0.9rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }

    /* Input Styling */
    .stTextInput > div > div > input {
        background-color: var(--input-bg) !important;
        color: var(--text-main) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 6px !important;
        padding: 10px !important;
        font-size: 1rem !important;
        transition: border-color 0.2s;
    }

    .stTextInput > div > div > input:focus {
        border-color: var(--accent-color) !important;
        box-shadow: 0 0 0 2px rgba(88, 166, 255, 0.2) !important;
    }

    .stButton > button {
        background-color: var(--accent-color) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        width: 100% !important;
        padding: 10px !important;
        transition: background-color 0.2s;
    }

    .stButton > button:hover {
        background-color: #3a8bd6 !important; /* Lebih gelap saat hover */
    }

    /* Chat Bubble Styling (Clean Look) */
    .stChatMessage {
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        margin-bottom: 1rem;
        padding: 1rem;
        box-shadow: none;
    }

    .stChatMessage[data-testid="stChatMessage"]:nth-child(odd) {
        border-left: 4px solid var(--accent-color);
    }

    .stChatMessage[data-testid="stChatMessage"]:nth-child(even) {
        border-left: 4px solid var(--text-muted);
    }

    /* Header Utama */
    .main-header {
        border-bottom: 1px solid var(--border-color);
        padding-bottom: 1rem;
        margin-bottom: 2rem;
    }

    .user-badge {
        background: rgba(88, 166, 255, 0.15);
        color: var(--accent-color);
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.9rem;
        font-weight: 600;
        border: 1px solid rgba(88, 166, 255, 0.3);
    }

    /* Sidebar */
    .sidebar .sidebar-content {
        background: var(--card-bg);
        border-right: 1px solid var(--border-color);
    }
</style>
""", unsafe_allow_html=True)

# --- 3. FUNGSI LOGIKA ---

def init_session_state():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "username" not in st.session_state:
        st.session_state.username = ""
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "system", "content": "Anda adalah Qwen 3.5, asisten analitik cerdas. Berikan jawaban yang mendalam, terstruktur, dan profesional dalam Bahasa Indonesia."}
        ]

def login_page():
    """Halaman Login Minimalis"""
    st.markdown("""
    <div class="login-container">
        <div class="login-card">
            <div class="login-title">🔮 LAGOS Workspace</div>
            <div class="login-subtitle">Masukkan identitas Anda untuk memulai</div>

            <form id="login-form">
                <input type="text" id="user-name" class="login-input" placeholder="Nama Pengguna" required>
                <button type="submit" class="login-btn">Masuk</button>
            </form>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Logika Streamlit untuk menangani submit
    with st.form("login_form", clear_on_submit=False):
        name_input = st.text_input(
            "", 
            placeholder="Nama Pengguna", 
            key="login_name_input", 
            label_visibility="collapsed"
        )
        submit_login = st.form_submit_button("Masuk", use_container_width=True)

        if submit_login:
            if name_input:
                st.session_state.logged_in = True
                st.session_state.username = name_input
                st.rerun()
            else:
                st.error("Harap masukkan nama Anda.")

def ekstrak_teks_dari_dokumen(uploaded_file):
    teks_hasil = ""
    nama_file = uploaded_file.name.lower()
    try:
        if nama_file.endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                teks_halaman = page.extract_text()
                if teks_halaman:
                    teks_hasil += teks_halaman + "\n"
        elif nama_file.endswith('.txt'):
            teks_hasil = uploaded_file.read().decode("utf-8")
    except Exception as e:
        st.error(f"Gagal membaca dokumen: {e}")
    return teks_hasil.strip()

def konversi_gambar_ke_base64(uploaded_file):
    if uploaded_file is not None:
        return base64.b64encode(uploaded_file.read()).decode('utf-8')
    return None

def buat_file_word(riwayat_pesan, username):
    doc = Document()
    doc.add_heading(f'Laporan Analitik - {username}', level=0)
    doc.add_heading('Qwen 3.5 Workspace', level=1)

    for msg in riwayat_pesan:
        if msg["role"] == "system": continue
        if msg["role"] == "user":
            doc.add_heading("Perintah User:", level=2)
            content = msg["content"]
            if isinstance(content, list):
                for item in content:
                    if item["type"] == "text": doc.add_paragraph(item["text"])
            else: doc.add_paragraph(content)
        elif msg["role"] == "assistant":
            doc.add_heading("Respons AI:", level=2)
            paragraf_list = msg["content"].split('\n')
            for p_text in paragraf_list:
                teks_bersih = p_text.strip()
                if not teks_bersih: continue
                match_heading = re.match(r'^(#{1,6})\s+(.*)$', teks_bersih)
                if match_heading:
                    level = min(len(match_heading.group(1)), 3)
                    doc.add_heading(match_heading.group(2).replace('**',''), level=level)
                    continue
                is_bullet = False
                if teks_bersih.startswith('- ') or teks_bersih.startswith('* '):
                    teks_bersih = teks_bersih[2:]
                    p = doc.add_paragraph(style='List Bullet')
                    is_bullet = True
                elif re.match(r'^\d+\.\s+', teks_bersih):
                    teks_bersih = re.sub(r'^\d+\.\s+', '', teks_bersih)
                    p = doc.add_paragraph(style='List Number')
                else:
                    p = doc.add_paragraph()

                parts = re.split(r'(\*\*.*?\*\*)', teks_bersih)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        p.add_run(part.replace('**', '')).bold = True
                    else:
                        p.add_run(part)
            doc.add_paragraph("_" * 40)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. APLIKASI UTAMA ---

init_session_state()

if not st.session_state.logged_in:
    login_page()
else:
    # Header dengan nama pengguna
    st.markdown(f"""
    <div class="main-header">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1 style="margin:0; font-size: 1.8rem; font-weight: 600;">Qwen 3.5 <span style="color: var(--text-muted); font-weight: 300;">| Analitik</span></h1>
            </div>
            <div class="user-badge">
                👤 {st.session_state.username}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### 🛠️ Panel Kontrol")
        st.markdown("---")

        if len(st.session_state.messages) > 1:
            if st.button("📥 Ekspor ke Word (.docx)", use_container_width=True):
                file_word = buat_file_word(st.session_state.messages, st.session_state.username)
                st.download_button(
                    label="Download",
                    data=file_word,
                    file_name=f"Laporan_{st.session_state.username}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        st.markdown("---")
        if st.button("🔒 Keluar & Reset", use_container_width=True, type="primary"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state.logged_in = False
            st.rerun()

    # Konfigurasi API
    BASE_URL = "https://integrate.api.nvidia.com/v1"
    nvidia_api_key = "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"
    MODEL_NAME = "qwen/qwen3.5-122b-a10b"
    client = OpenAI(base_url=BASE_URL, api_key=nvidia_api_key)

    # Tampilkan Chat History
    for message in st.session_state.messages:
        if message["role"] != "system":
            with st.chat_message(message["role"]):
                if isinstance(message["content"], list):
                    for item in message["content"]:
                        if item["type"] == "text":
                            st.markdown(item["text"])
                else:
                    st.markdown(message["content"])

    st.divider()

    # Input Area
    with st.container():
        col_img, col_doc = st.columns(2)

        with col_img:
            uploaded_image = st.file_uploader(
                "Upload Gambar", 
                type=["jpg", "jpeg", "png"],
                key="img_up",
                label_visibility="collapsed"
            )

        with col_doc:
            uploaded_file = st.file_uploader(
                "Upload Dokumen (PDF/TXT)",
                type=["pdf", "txt"],
                key="doc_up",
                label_visibility="collapsed"
            )

        if uploaded_image:
            c1, c2 = st.columns([1, 4])
            with c1: st.image(uploaded_image, width=80)
            with c2: 
                st.caption(f"File: {uploaded_image.name}")
                if st.button("❌ Hapus", key="del_img", use_container_width=True): st.rerun()

        if uploaded_file:
            c1, c2 = st.columns([1, 4])
            with c1: st.markdown("📄")
            with c2: 
                st.caption(f"File: {uploaded_file.name}")
                if st.button("❌ Hapus", key="del_doc", use_container_width=True): st.rerun()

        user_input = st.chat_input("Ketik pesan Anda di sini...")

    # Logika Chat
    if user_input:
        teks_perintah = str(user_input).strip()
        konteks_dokumen = ""

        if uploaded_file:
            with st.spinner("Menganalisis dokumen..."):
                isi_teks = ekstrak_teks_dari_dokumen(uploaded_file)
                if isi_teks:
                    konteks_dokumen = f"Dokumen ({uploaded_file.name}):\n\"\"\"\n{isi_teks}\n\"\"\"\n\n"

        payload_konten = []
        if uploaded_image:
            base64_img = konversi_gambar_ke_base64(uploaded_image)
            final_prompt = f"{konteks_dokumen}{teks_perintah}"
            payload_konten = [
                {"type": "text", "text": final_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
            ]
        else:
            final_prompt = f"{konteks_dokumen}{teks_perintah}"
            payload_konten = final_prompt

        with st.chat_message("user"):
            st.markdown(teks_perintah)
            if uploaded_image: st.image(uploaded_image, width=200)
            if uploaded_file: st.caption(f"Dokumen terlampir: {uploaded_file.name}")

        st.session_state.messages.append({"role": "user", "content": payload_konten})

        with st.chat_message("assistant"):
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
                        delta_content = chunk.choices[0].delta.content
                        if delta_content:
                            full_response += delta_content
                            placeholder.markdown(full_response + "▌")
                placeholder.markdown(full_response)
                st.session_state.messages.append({"role": "assistant", "content": full_response})
            except Exception as e:
                st.error(f"Error: {e}")
