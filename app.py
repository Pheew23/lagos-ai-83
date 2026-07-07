import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document
from pypdf import PdfReader

# --- 1. KONFIGURASI HALAMAN (DARK MODE DEFAULT) ---
st.set_page_config(
    page_title="Qwen 3.5 Workspace",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. CUSTOM CSS (DARK MODE & ELEGANT) ---
st.markdown("""
<style>
    /* Import Font Modern */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    :root {
        --bg-color: #0d1117;       /* Dark Background */
        --card-bg: #161b22;        /* Card Background */
        --text-main: #c9d1d9;      /* Main Text */
        --text-muted: #8b949e;     /* Muted Text */
        --accent-color: #58a6ff;   /* Soft Blue Accent */
        --border-color: #30363d;   /* Border Color */
        --input-bg: #0d1117;       /* Input Background */
    }

    body {
        background-color: var(--bg-color);
        color: var(--text-main);
        font-family: 'Inter', sans-serif;
        /* Background subtle gradient */
        background-image: radial-gradient(circle at 15% 50%, rgba(88, 166, 255, 0.08) 0%, transparent 25%);
    }

    /* --- LOGIN PAGE STYLES --- */
    .login-wrapper {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 90vh;
        flex-direction: column;
    }

    .login-card {
        background: var(--card-bg);
        border: 1px solid var(--border-color);
        padding: 2.5rem;
        border-radius: 12px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.6);
        width: 100%;
        max-width: 400px;
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }

    .login-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 50px rgba(0,0,0,0.8), 0 0 15px rgba(88, 166, 255, 0.1);
    }

    .login-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: var(--text-main);
        margin-bottom: 0.5rem;
    }

    .login-subtitle {
        color: var(--text-muted);
        font-size: 0.95rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }

    /* Input Styling di Login */
    .stTextInput > div > div > input {
        background-color: var(--input-bg) !important;
        color: var(--text-main) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 6px !important;
        padding: 12px !important;
        font-size: 1rem !important;
        transition: border-color 0.2s;
    }

    .stTextInput > div > div > input:focus {
        border-color: var(--accent-color) !important;
        box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.15) !important;
        outline: none !important;
    }

    .stButton > button {
        background-color: var(--accent-color) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        width: 100% !important;
        padding: 12px !important;
        font-size: 1rem !important;
        transition: background-color 0.2s, transform 0.1s;
    }

    .stButton > button:hover {
        background-color: #3a8bd6 !important; /* Lebih gelap saat hover */
        transform: scale(1.01);
    }

    /* --- MAIN APP STYLES --- */
    .main-header {
        border-bottom: 1px solid var(--border-color);
        padding-bottom: 1rem;
        margin-bottom: 2rem;
    }

    .user-badge {
        background: rgba(88, 166, 255, 0.15);
        color: var(--accent-color);
        padding: 6px 16px;
        border-radius: 20px;
        font-size: 0.9rem;
        font-weight: 600;
        border: 1px solid rgba(88, 166, 255, 0.2);
        display: inline-block;
    }

    /* Chat Bubbles */
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

    /* Sidebar */
    .sidebar .sidebar-content {
        background: var(--card-bg);
        border-right: 1px solid var(--border-color);
    }

    /* File Uploader Styling */
    .stFileUploader > div {
        background: rgba(255,255,255,0.02);
        border: 1px dashed var(--border-color);
        border-radius: 8px;
        transition: all 0.2s;
    }
    .stFileUploader > div:hover {
        border-color: var(--accent-color);
        background: rgba(88, 166, 255, 0.05);
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
            {"role": "system", "content": "Anda adalah Qwen 3.5, asisten analitik cerdas. Berikan jawaban yang mendalam, terstruktur, dan profesional dalam Bahasa Indonesia. Gunakan format Markdown untuk kerapian."}
        ]

def login_page():
    """Halaman Login Minimalis Mode Gelap"""
    st.markdown("""
    <div class="login-wrapper">
        <div class="login-card">
            <div class="login-title">🔮 Qwen Workspace</div>
            <div class="login-subtitle">Akses sistem analitik terintegrasi</div>

            <!-- Form HTML untuk tampilan yang lebih rapi -->
            <div style="text-align: left; margin-bottom: 1rem;">
                <label style="color: var(--text-muted); font-size: 0.9rem; margin-bottom: 5px; display: block;">Identitas Pengguna</label>
                <input type="text" id="user-name" class="stTextInput" placeholder="Masukkan nama Anda..." style="width: 100%; padding: 12px; background: var(--input-bg); border: 1px solid var(--border-color); color: white; border-radius: 6px; font-family: 'Inter', sans-serif;" required>
            </div>

            <button id="login-btn" style="width: 100%; padding: 12px; background: var(--accent-color); color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;">Masuk ke Sistem</button>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Logika Streamlit untuk menangani submit
    # Kita gunakan form Streamlit native tapi dengan label kosong agar sesuai CSS
    with st.form("login_form", clear_on_submit=False):
        name_input = st.text_input(
            "", 
            placeholder="Masukkan nama Anda...", 
            key="login_name_input", 
            label_visibility="collapsed"
        )
        submit_login = st.form_submit_button("Masuk", use_container_width=True)

        if submit_login:
            if name_input and name_input.strip():
                st.session_state.logged_in = True
                st.session_state.username = name_input.strip()
                st.success("Login berhasil!")
                st.rerun()
            else:
                st.error("Harap masukkan nama Anda untuk melanjutkan.")

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
    # --- HALAMAN LOGIN ---
    login_page()
else:
    # --- HALAMAN UTAMA (DARK MODE) ---

    # Header
    st.markdown(f"""
    <div class="main-header">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1 style="margin:0; font-size: 1.8rem; font-weight: 700; color: var(--text-main);">Qwen 3.5 <span style="color: var(--text-muted); font-weight: 400; font-size: 1rem;">| Analitik</span></h1>
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
            if st.button("📥 Ekspor Laporan (.docx)", use_container_width=True):
                file_word = buat_file_word(st.session_state.messages, st.session_state.username)
                st.download_button(
                    label="Download File",
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
                "Upload Gambar (JPG/PNG)", 
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

        # Preview File
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

        user_input = st.chat_input("Ketik perintah Anda di sini...")

    # Logika Chat
    if user_input:
        teks_perintah = str(user_input).strip()
        konteks_dokumen = ""

        if uploaded_file:
            with st.spinner("Menganalisis dokumen...
                isi_teks = ekstrak_teks_dari_dokumen(uploaded_file)
            if isi_teks:
                konteks_dokumen = f"Dokumen ({uploaded_file.name}):\n\"\"\"\n{isi_teks}\n\"\"\"\n\n"
            else:
                st.warning("Dokumen tidak memiliki teks yang dapat dibaca.")

        # Siapkan Payload
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

        # Tampilkan Pesan User
        with st.chat_message("user"):
            st.markdown(teks_perintah)
            if uploaded_image:
                st.image(uploaded_image, width=250, caption="Gambar Terlampir")
            if uploaded_file:
                st.caption(f"Dokumen terlampir: {uploaded_file.name}")

        # Simpan ke riwayat
        st.session_state.messages.append({"role": "user", "content": payload_konten})

        # Respons AI
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
                            # Efek kursor berkedip saat mengetik
                            placeholder.markdown(full_response + "▌")

                # Hapus kursor berkedip saat selesai
                placeholder.markdown(full_response)

                # Simpan respons AI ke riwayat
                st.session_state.messages.append({"role": "assistant", "content": full_response})

            except Exception as e:
                st.error(f"⚠️ Terjadi kesalahan: {e}")
                st.info("Silakan coba lagi atau periksa koneksi API.")
