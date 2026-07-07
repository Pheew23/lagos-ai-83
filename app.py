import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document
from pypdf import PdfReader

# --- 1. KONFIGURASI HALAMAN & CSS FUTURISTIK ---
st.set_page_config(
    page_title="Qwen 3.5 - Nexus Interface",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS untuk tampilan Futuristik & Login Page
st.markdown("""
<style>
    /* Import Font Futuristik */
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@300;500;700&display=swap');

    :root {
        --neon-blue: #00f3ff;
        --neon-purple: #bc13fe;
        --bg-dark: #05050a;
        --glass-bg: rgba(20, 20, 35, 0.7);
        --glass-border: rgba(0, 243, 255, 0.3);
        --text-main: #e0e0e0;
    }

    body {
        background-color: var(--bg-dark);
        color: var(--text-main);
        font-family: 'Rajdhani', sans-serif;
        background-image: radial-gradient(circle at 50% 50%, #1a1a2e 0%, #05050a 100%);
    }

    /* Judul & Heading */
    h1, h2, h3, h4 {
        font-family: 'Orbitron', sans-serif;
        color: var(--neon-blue);
        text-shadow: 0 0 10px rgba(0, 243, 255, 0.5);
        margin-bottom: 0.5rem;
    }

    /* Container Login */
    .login-container {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 90vh;
        flex-direction: column;
    }

    .login-box {
        background: var(--glass-bg);
        backdrop-filter: blur(15px);
        border: 1px solid var(--glass-border);
        padding: 3rem;
        border-radius: 20px;
        box-shadow: 0 0 30px rgba(0, 243, 255, 0.15), inset 0 0 20px rgba(0,0,0,0.5);
        text-align: center;
        width: 450px;
        max-width: 90%;
    }

    .login-input {
        background: rgba(0, 0, 0, 0.6);
        border: 1px solid var(--neon-purple);
        border-radius: 10px;
        color: white;
        padding: 12px;
        font-size: 1.2rem;
        text-align: center;
        margin-bottom: 20px;
        font-family: 'Rajdhani', sans-serif;
        width: 100%;
        box-sizing: border-box;
    }

    .login-input:focus {
        outline: none;
        border-color: var(--neon-blue);
        box-shadow: 0 0 10px var(--neon-blue);
    }

    .login-btn {
        background: linear-gradient(90deg, var(--neon-purple), var(--neon-blue));
        border: none;
        color: white;
        padding: 12px 40px;
        border-radius: 25px;
        font-weight: bold;
        font-family: 'Orbitron', sans-serif;
        cursor: pointer;
        transition: all 0.3s ease;
        width: 100%;
        font-size: 1.1rem;
        letter-spacing: 1px;
    }

    .login-btn:hover {
        transform: scale(1.02);
        box-shadow: 0 0 20px var(--neon-blue);
    }

    /* Chat Bubble Styling */
    .stChatMessage {
        background: rgba(255, 255, 255, 0.03);
        border-radius: 15px;
        border: 1px solid rgba(255, 255, 255, 0.05);
        margin-bottom: 1rem;
        padding: 1rem;
    }

    /* User Message Style */
    .stChatMessage[data-testid="stChatMessage"]:nth-child(odd) {
        background: linear-gradient(135deg, rgba(188, 19, 254, 0.1), rgba(188, 19, 254, 0.05));
        border-left: 3px solid var(--neon-purple);
    }

    /* AI Message Style */
    .stChatMessage[data-testid="stChatMessage"]:nth-child(even) {
        background: linear-gradient(135deg, rgba(0, 243, 255, 0.05), rgba(0, 243, 255, 0.02));
        border-left: 3px solid var(--neon-blue);
    }

    /* Input Chat Area */
    .stChatInput input {
        background: rgba(10, 10, 15, 0.8);
        border: 1px solid var(--neon-blue);
        color: white;
        border-radius: 12px;
        box-shadow: 0 0 10px rgba(0, 243, 255, 0.1);
        font-family: 'Rajdhani', sans-serif;
        font-size: 1.1rem;
    }

    .stChatInput input::placeholder {
        color: rgba(255, 255, 255, 0.5);
    }

    /* Sidebar Styling */
    .sidebar .sidebar-content {
        background: rgba(5, 5, 10, 0.95);
        border-right: 1px solid var(--neon-purple);
        backdrop-filter: blur(10px);
    }

    /* Upload Button Styling */
    .stFileUploader > div {
        background: rgba(255,255,255,0.03);
        border: 1px dashed var(--neon-blue);
        border-radius: 10px;
        transition: all 0.3s;
    }

    .stFileUploader > div:hover {
        border-color: var(--neon-purple);
        background: rgba(255,255,255,0.05);
    }

    /* Utility Classes */
    .highlight-text {
        color: var(--neon-blue);
        font-weight: bold;
    }

    .status-dot {
        height: 10px;
        width: 10px;
        background-color: #00ff00;
        border-radius: 50%;
        display: inline-block;
        box-shadow: 0 0 10px #00ff00;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. FUNGSI LOGIKA UTAMA ---

def init_session_state():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "username" not in st.session_state:
        st.session_state.username = ""
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "system", "content": "Anda adalah Qwen 3.5, asisten analitik tingkat tinggi dengan antarmuka futuristik. Berikan jawaban cerdas, mendalam, dan terstruktur dalam Bahasa Indonesia."}
        ]

def login_page():
    """Tampilan Halaman Login"""
    st.markdown("""
    <div class="login-container">
        <div class="login-box">
            <h1 style="font-size: 2.5rem; margin-bottom: 10px; color: var(--neon-blue);">🔮 QWEN NEXUS</h1>
            <p style="color: #aaa; margin-bottom: 30px; font-size: 1.1rem;">Sistem Analitik Cerdas Terintegrasi</p>

            <form id="login-form">
                <input type="text" id="user-name" class="login-input" placeholder="Masukkan Nama Identitas..." required>
                <button type="submit" class="login-btn">AKSES SISTEM</button>
            </form>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Menggunakan Streamlit form untuk menangkap submit
    with st.form("login_form", clear_on_submit=False):
        st.markdown("<h2 style='text-align:center; color:white; font-family:Orbitron;'>Identifikasi Pengguna</h2>", unsafe_allow_html=True)
        name_input = st.text_input(
            "", 
            placeholder="Nama Pengguna...", 
            key="login_name_input", 
            label_visibility="collapsed",
            help="Masukkan nama Anda untuk memulai sesi"
        )
        submit_login = st.form_submit_button("🚀 Masuk ke Sistem", use_container_width=True)

        if submit_login and name_input:
            st.session_state.logged_in = True
            st.session_state.username = name_input
            st.success(f"Selamat datang, {name_input}! Menginisialisasi sistem...")
            st.rerun()
        elif submit_login and not name_input:
            st.error("⚠️ Nama diperlukan untuk mengakses sistem.")

def ekstrak_teks_dari_dokumen(uploaded_file):
    """Ekstrak teks dari PDF atau TXT"""
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
    """Konversi gambar ke base64 untuk payload API"""
    if uploaded_file is not None:
        return base64.b64encode(uploaded_file.read()).decode('utf-8')
    return None

def buat_file_word(riwayat_pesan, username):
    """Membuat dokumen Word yang rapi dari riwayat chat"""
    doc = Document()
    # Header Dokumen
    doc.add_heading(f'Laporan Analitik - Pengguna: {username}', level=0)
    doc.add_heading('Qwen 3.5 Workspace', level=1)
    doc.add_paragraph("-" * 50)

    for msg in riwayat_pesan:
        if msg["role"] == "system": continue

        if msg["role"] == "user":
            doc.add_heading("👤 Perintah User:", level=2)
            content = msg["content"]
            if isinstance(content, list):
                for item in content:
                    if item["type"] == "text": doc.add_paragraph(item["text"])
            else: doc.add_paragraph(content)

        elif msg["role"] == "assistant":
            doc.add_heading("🤖 Respons AI:", level=2)
            paragraf_list = msg["content"].split('\n')
            for p_text in paragraf_list:
                teks_bersih = p_text.strip()
                if not teks_bersih: continue

                # Deteksi Heading Markdown
                match_heading = re.match(r'^(#{1,6})\s+(.*)$', teks_bersih)
                if match_heading:
                    level = min(len(match_heading.group(1)), 3)
                    doc.add_heading(match_heading.group(2).replace('**',''), level=level)
                    continue

                # Deteksi List
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

                # Format Bold (**teks**)
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

# --- 3. APLIKASI UTAMA ---

init_session_state()

if not st.session_state.logged_in:
    # --- TAMPILKAN HALAMAN LOGIN ---
    login_page()
else:
    # --- TAMPILKAN APLIKASI UTAMA ---

    # Header Futuristik
    st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid rgba(0, 243, 255, 0.2); margin-bottom: 20px;">
        <div>
            <h1 style="margin:0; font-size: 2.2rem;">🔮 QWEN 3.5 <span style="font-size:1rem; color:#aaa; font-weight:300;">// NEXUS</span></h1>
            <p style="margin:0; color: var(--neon-blue); font-size: 1.1rem;">Selamat datang, <span class="highlight-text">{st.session_state.username}</span></p>
        </div>
        <div style="text-align: right; font-size: 0.9rem; color: #888; font-family: 'Orbitron', sans-serif;">
            Model: Qwen/Qwen3.5-122B<br>
            Status: <span class="status-dot"></span> Online
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar (Kontrol)
    with st.sidebar:
        st.markdown("### 🛠️ Panel Kontrol")
        st.markdown("---")

        # Export
        if len(st.session_state.messages) > 1:
            if st.button("📥 Ekspor Laporan (.docx)", use_container_width=True, type="primary"):
                file_word = buat_file_word(st.session_state.messages, st.session_state.username)
                st.download_button(
                    label="Download File",
                    data=file_word,
                    file_name=f"Laporan_{st.session_state.username}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        st.markdown("---")
        if st.button("🗑️ Reset Memori Sistem", use_container_width=True, type="secondary"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state.logged_in = False
            st.info("Sistem telah di-reset. Mengalihkan ke halaman login...")
            st.rerun()

    # Konfigurasi API
    BASE_URL = "https://integrate.api.nvidia.com/v1"
    nvidia_api_key = "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"
    MODEL_NAME = "qwen/qwen3.5-122b-a10b"
    client = OpenAI(base_url=BASE_URL, api_key=nvidia_api_key)

    # Area Chat History
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

    # --- AREA INPUT MULTIMODAL & DOKUMEN ---
    input_container = st.container()
    with input_container:
        # Kolom untuk Upload Gambar dan Dokumen
        col_img, col_doc = st.columns(2)

        with col_img:
            uploaded_image = st.file_uploader(
                "📷 Upload Gambar (JPG/PNG)", 
                type=["jpg", "jpeg", "png"],
                key="img_up",
                label_visibility="collapsed"
            )

        with col_doc:
            uploaded_file = st.file_uploader(
                "📄 Upload Dokumen (PDF/TXT)",
                type=["pdf", "txt"],
                key="doc_up",
                label_visibility="collapsed"
            )

        # Preview Gambar dengan tombol hapus
        if uploaded_image:
            c1, c2 = st.columns([1, 4])
            with c1:
                st.image(uploaded_image, width=80, caption="Preview Gambar")
            with c2:
                st.info(f"File: {uploaded_image.name}")
                if st.button("❌ Hapus Gambar", key="del_img", use_container_width=True):
                    st.cache_data.clear()
                    # Kita tidak menghapus dari session_state secara langsung di sini karena
                    # Streamlit file_uploader bersifat stateless per render.
                    # Namun, jika user refresh atau pindah, file hilang.
                    # Untuk UX yang lebih baik, kita bisa set flag di session_state jika perlu.
                    st.rerun()

        # Preview Dokumen dengan tombol hapus
        if uploaded_file:
            c1, c2 = st.columns([1, 4])
            with c1:
                st.markdown("📄")
            with c2:
                st.info(f"File: {uploaded_file.name}")
                if st.button("❌ Hapus Dokumen", key="del_doc", use_container_width=True):
                    st.rerun()

        # Input Chat Utama
        user_input = st.chat_input("Ketik perintahmu di sini... (Contoh: 'Analisis gambar ini' atau 'Rangkum dokumen')")

    # --- 4. PROSES LOGIKA CHAT & API ---
    if user_input:
        teks_perintah = str(user_input).strip()

        # 1. Proses Ekstraksi Dokumen (jika ada)
        konteks_dokumen = ""
        if uploaded_file:
            with st.spinner("📡 Membaca dan menganalisis dokumen..."):
                isi_teks = ekstrak_teks_dari_dokumen(uploaded_file)
                if isi_teks:
                    konteks_dokumen = (
                        f"Konteks Dokumen ({uploaded_file.name}):\n"
                        f"\"\"\"\n{isi_teks}\n\"\"\"\n\n"
                    )
                else:
                    st.warning("Dokumen tidak memiliki teks yang bisa dibaca.")

        # 2. Siapkan Payload untuk API
        payload_konten = []

        # Logika Multimodal: Jika ada gambar, format khusus untuk API
        if uploaded_image:
            base64_img = konversi_gambar_ke_base64(uploaded_image)
            # Gabungkan konteks dokumen (jika ada) dengan perintah teks
            final_prompt = f"{konteks_dokumen}{teks_perintah}"

            payload_konten = [
                {"type": "text", "text": final_prompt},
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}
                }
            ]
        else:
            # Jika hanya teks + dokumen (tanpa gambar)
            final_prompt = f"{konteks_dokumen}{teks_perintah}"
            payload_konten = final_prompt

        # 3. Tampilkan Pesan User di UI
        with st.chat_message("user"):
            st.markdown(teks_perintah)
            if uploaded_image:
                st.image(uploaded_image, width=250)
            if uploaded_file:
                st.caption(f"Dokumen terlampir: {uploaded_file.name}")

        # 4. Simpan ke sesi chat
        st.session_state.messages.append({"role": "user", "content": payload_konten})

        # 5. Panggil API Qwen 3.5 & Tampilkan Respons Streaming
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
                            # Tampilkan dengan efek kursor berkedip
                            placeholder.markdown(full_response + "▌")

                # Update akhir tanpa kursor
                placeholder.markdown(full_response)

                # Simpan respons AI ke sesi chat
                st.session_state.messages.append({"role": "assistant", "content": full_response})

            except Exception as e:
                st.error(f"⚠️ Terjadi kesalahan koneksi ke Qwen 3.5: {e}")
                st.info("Silakan coba lagi atau periksa koneksi internet.")
