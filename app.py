import streamlit as st
from openai import OpenAI
import io
import re
import base64
import requests
from docx import Document
from audio_recorder_streamlit import audio_recorder
import speech_recognition as sr
import sqlite3
import json
import uuid
import hashlib
from datetime import datetime, timedelta
import streamlit.components.v1 as components
import extra_streamlit_components as stx

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Premium",
    page_icon="🔮",
    layout="centered", 
    initial_sidebar_state="expanded"
)

# --- INIT COOKIE MANAGER ---
# Parameter experimental_allow_widgets dihapus agar tidak error di versi Streamlit terbaru
@st.cache_resource
def get_cookie_manager():
    return stx.CookieManager()

cookie_manager = get_cookie_manager()

# --- 2. CUSTOM CSS (MINIMALIS & PROFESIONAL) ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        
        #MainMenu, footer, header {visibility: hidden;}

        /* Tipografi Header */
        .brand-title {
            text-align: center;
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(135deg, #7d4eff 0%, #00d2ff 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.2rem;
            padding-top: 1rem;
        }
        .brand-subtitle {
            text-align: center;
            color: var(--text-color);
            opacity: 0.6;
            font-size: 1rem;
            font-weight: 400;
            margin-bottom: 2.5rem;
        }

        /* Tampilan Bubble Chat */
        .stChatMessage {
            padding: 1.5rem !important;
            border-radius: 12px;
            margin-bottom: 1rem;
        }
        .stChatMessage:nth-child(even) {
            background-color: var(--secondary-background-color) !important;
            border: 1px solid var(--border-color);
        }

        /* Tampilan Pill (Indikator File) */
        .file-indicator {
            display: inline-flex;
            align-items: center;
            background: rgba(125, 78, 255, 0.1);
            color: #7d4eff;
            padding: 6px 16px;
            border-radius: 50px;
            font-size: 0.85rem;
            font-weight: 500;
            margin-right: 10px;
            margin-bottom: 15px;
            border: 1px solid rgba(125, 78, 255, 0.2);
        }

        /* Tombol Bulat (Popover Attachment) */
        [data-testid="stPopover"] button {
            border-radius: 50% !important;
            height: 52px !important;
            width: 52px !important;
            padding: 0 !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            border: 1.5px solid var(--border-color) !important;
            background-color: var(--secondary-background-color) !important;
            transition: all 0.2s ease !important;
        }
        [data-testid="stPopover"] button:hover {
            border-color: #7d4eff !important;
            color: #7d4eff !important;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(125, 78, 255, 0.15);
        }
        
        /* Merapikan tulisan pada tombol sidebar agar tidak menabrak batas */
        [data-testid="stSidebar"] button p {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
    </style>
""", unsafe_allow_html=True)

# --- FUNGSI DATABASE MULTI-USER ---
DB_NAME = 'lagos_multiuser.db'

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS sessions (session_id TEXT PRIMARY KEY, username TEXT, title TEXT, updated_at TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS messages (id INTEGER PRIMARY KEY AUTOINCREMENT, session_id TEXT, role TEXT, content TEXT)''')
    conn.commit()
    conn.close()

def register_user(username, password):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, hash_password(password)))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False 
    conn.close()
    return success

def authenticate_user(username, password):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT password FROM users WHERE username=?", (username,))
    row = c.fetchone()
    conn.close()
    if row and row[0] == hash_password(password):
        return True
    return False

def get_user_sessions(username):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT session_id, title FROM sessions WHERE username=? ORDER BY updated_at DESC", (username,))
    rows = c.fetchall()
    conn.close()
    return rows

def load_session_messages(session_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT role, content FROM messages WHERE session_id=? ORDER BY id ASC", (session_id,))
    rows = c.fetchall()
    conn.close()
    msgs = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
    for r, c in rows:
        try:
            msgs.append({"role": r, "content": json.loads(c)})
        except:
            msgs.append({"role": r, "content": c})
    return msgs

def save_session_db(session_id, username, title, messages):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO sessions (session_id, username, title, updated_at) VALUES (?, ?, ?, ?)", 
              (session_id, username, title, datetime.now()))
    c.execute("DELETE FROM messages WHERE session_id=?", (session_id,))
    for msg in messages:
        if msg["role"] != "system":
            c.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)", 
                      (session_id, msg["role"], json.dumps(msg["content"])))
    conn.commit()
    conn.close()

def delete_session_db(session_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("DELETE FROM sessions WHERE session_id=?", (session_id,))
    c.execute("DELETE FROM messages WHERE session_id=?", (session_id,))
    conn.commit()
    conn.close()

init_db()

# --- 3. SISTEM AUTENTIKASI DENGAN COOKIES ---
# Cek apakah ada cookie yang tersimpan dari sesi sebelumnya
cached_user = cookie_manager.get(cookie="lagos_username")

if "logged_in" not in st.session_state:
    if cached_user:
        st.session_state.logged_in = True
        st.session_state.username = cached_user
    else:
        st.session_state.logged_in = False
        st.session_state.username = ""

if not st.session_state.logged_in:
    st.markdown('<div class="brand-title">Lagos AI 9.1</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Platform Asisten Analitik Multimodal</div>', unsafe_allow_html=True)
    
    # Layout pembagian untuk meletakkan form di tengah
    _, col_form, _ = st.columns([1, 1.2, 1])
    
    with col_form:
        auth_mode = st.radio("Mode Akses", ["Masuk ke Akun", "Daftar Akun Baru"], horizontal=True, label_visibility="collapsed")
        
        if auth_mode == "Masuk ke Akun":
            with st.form("login_form"):
                st.markdown("#### Akses Portal")
                log_user = st.text_input("Username")
                log_pass = st.text_input("Password", type="password")
                keep_login = st.checkbox("Biarkan saya tetap masuk")
                submit_login = st.form_submit_button("Masuk", use_container_width=True, type="primary")
                
                if submit_login:
                    if authenticate_user(log_user, log_pass):
                        st.session_state.logged_in = True
                        st.session_state.username = log_user
                        
                        # Set cookie jika user menceklis "Tetap Masuk" (aktif 30 hari)
                        if keep_login:
                            cookie_manager.set("lagos_username", log_user, expires_at=datetime.now() + timedelta(days=30))
                        
                        st.rerun()
                    else:
                        st.error("Kredensial tidak valid. Silakan periksa kembali.")
                        
        else:
            with st.form("register_form"):
                st.markdown("#### Pendaftaran Akun")
                reg_user = st.text_input("Pilih Username")
                reg_pass = st.text_input("Buat Password", type="password")
                submit_register = st.form_submit_button("Daftar Sekarang", use_container_width=True)
                
                if submit_register:
                    if reg_user and reg_pass:
                        if register_user(reg_user, reg_pass):
                            st.success("Registrasi berhasil! Silakan pilih 'Masuk ke Akun'.")
                        else:
                            st.error("Username ini sudah digunakan.")
                    else:
                        st.warning("Username dan Password wajib diisi.")
    
    st.stop() # Blok eksekusi berlanjut jika belum login

# ==========================================
# KODE BERJALAN SETELAH LOGIN SUKSES
# ==========================================

# --- KONFIGURASI API ---
API_KEY = st.secrets["NVIDIA_API_KEY"] 
BASE_URL = "https://integrate.api.nvidia.com/v1"

# --- FUNGSI PEMBANTU MULTIMEDIA ---
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

def generate_title_from_messages(messages):
    for msg in messages:
        if msg["role"] == "user":
            content = msg["content"]
            text = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
            text = text.split("[AKHIR KONTEN]\n\n")[-1]
            return text[:25] + "..." if len(text) > 25 else (text if text else "Obrolan Gambar/File")
    return "Obrolan Baru"

# --- 4. INISIALISASI SESSION STATE OBROLAN ---
if "current_session_id" not in st.session_state:
    st.session_state.current_session_id = None
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
if "temp_image" not in st.session_state:
    st.session_state.temp_image = None
if "temp_doc" not in st.session_state:
    st.session_state.temp_doc = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# --- SIDEBAR (UI MANAJEMEN YANG BERSIH) ---
with st.sidebar:
    st.markdown(f"Masuk sebagai: **{st.session_state.username}**")
    st.divider()
    
    # Tombol utama untuk aksi prioritas
    if st.button("➕ Obrolan Baru", use_container_width=True, type="primary"):
        st.session_state.current_session_id = None
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        st.rerun()
        
    # Fitur Hapus hanya muncul jika sedang membuka riwayat lama
    if st.session_state.current_session_id is not None:
        if st.button("🗑️ Hapus Obrolan Ini", use_container_width=True):
            delete_session_db(st.session_state.current_session_id)
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
            st.rerun()

    st.markdown("### Riwayat Anda")
    sessions = get_user_sessions(st.session_state.username)
    
    if not sessions:
        st.caption("Belum ada obrolan.")
    else:
        with st.container(height=350, border=False):
            for sess_id, title in sessions:
                btn_type = "secondary"
                if st.session_state.current_session_id == sess_id:
                    btn_type = "primary"
                
                if st.button(title, key=f"btn_{sess_id}", use_container_width=True, type=btn_type):
                    st.session_state.current_session_id = sess_id
                    st.session_state.messages = load_session_messages(sess_id)
                    st.rerun()

    st.divider()
    
    st.markdown("### Pengaturan Engine")
    MODEL_MAPPING = {
        "openai/gpt-oss-120b": "Kecepatan Tinggi (Teks)",
        "thinkingmachines/inkling": "Optimal Standar (Teks)",
        "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning": "Analisis Mendalam",
        "google/diffusiongemma-26b-a4b-it": "Stabil & Presisi",
        "mistralai/mistral-large-3-675b-instruct-2512": "Projek Khusus"
    }
    
    MODEL_NAME = st.selectbox(
        label="Pilih model aktif:",
        options=list(MODEL_MAPPING.keys()),
        index=0,
        format_func=lambda x: MODEL_MAPPING[x],
        label_visibility="collapsed"
    )
    
    if len(st.session_state.messages) > 1:
        file_word = buat_file_word(st.session_state.messages)
        st.download_button(
            label="📥 Unduh Laporan DOCX",
            data=file_word,
            file_name="Lagos_AI_9.1_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    st.divider()
    col_out, col_adm = st.columns(2)
    with col_out:
        if st.button("🚪 Keluar", use_container_width=True):
            # Hapus cookie agar user benar-benar logout
            cookie_manager.delete("lagos_username")
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
            st.rerun()
            
    with col_adm:
        try:
            with open(DB_NAME, "rb") as db_file:
                st.download_button("⚙️ DB", data=db_file, file_name="lagos_multiuser.db", mime="application/octet-stream", use_container_width=True, help="Download Database (Admin)")
        except FileNotFoundError:
            st.button("⚙️ DB", disabled=True, use_container_width=True)

# --- BRANDING UTAMA DI HALAMAN CHAT ---
st.markdown('<div class="brand-title">Lagos AI 9.1</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-subtitle">Siap membantu menganalisis dokumen dan gambar.</div>', unsafe_allow_html=True)

# --- 5. AREA OBROLAN UTAMA ---
if len(st.session_state.messages) == 1:
    st.info("Ketik pertanyaan, rekam suara, atau lampirkan dokumen/gambar di bawah ini untuk memulai analisis.")

for message in st.session_state.messages:
    if message["role"] == "system": continue
    with st.chat_message(message["role"]):
        content = message["content"]
        text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        st.markdown(text_disp)

st.markdown("<div style='height: 90px'></div>", unsafe_allow_html=True)

# --- SKRIP AUTO-SCROLL KE BAWAH ---
st.markdown("<div id='bottom-marker'></div>", unsafe_allow_html=True)
components.html(
    """
    <script>
        setTimeout(function() {
            var parentDoc = window.parent.document;
            var marker = parentDoc.getElementById('bottom-marker');
            if (marker) {
                marker.scrollIntoView({behavior: 'auto', block: 'end'});
            } else {
                var scrollNode = parentDoc.querySelector('.stMainBlockContainer') || parentDoc.querySelector('.main');
                if(scrollNode) scrollNode.scrollTo(0, scrollNode.scrollHeight);
            }
        }, 300);
    </script>
    """,
    height=0
)

# --- 6. AREA INPUT TERPADU (UI GEMINI-STYLE) ---
input_container = st.container()

with input_container:
    current_img = st.session_state.get(f"img_{st.session_state.uploader_key}")
    current_doc = st.session_state.get(f"doc_{st.session_state.uploader_key}")

    if current_img:
        st.markdown(f"<div class='file-indicator'>📸 Gambar telah dilampirkan</div>", unsafe_allow_html=True)
    if current_doc:
        st.markdown(f"<div class='file-indicator'>📄 Dokumen telah dilampirkan</div>", unsafe_allow_html=True)

    col_attach, col_input, col_mic = st.columns([1, 8, 1])
    
    with col_attach:
        with st.popover("📎"): 
            st.markdown("**Lampirkan File**")
            up_img = st.file_uploader("Upload Image", type=["jpg", "png", "jpeg"], label_visibility="collapsed", key=f"img_{st.session_state.uploader_key}")
            up_doc = st.file_uploader("Upload Doc", type=["pdf", "txt"], label_visibility="collapsed", key=f"doc_{st.session_state.uploader_key}")
            st.session_state.temp_image = up_img
            st.session_state.temp_doc = up_doc

    with col_input:
        prompt_text = st.chat_input("Tanyakan sesuatu pada Lagos AI...")

    with col_mic:
        audio_bytes = audio_recorder(
            text="", 
            recording_color="#ff4b4b",
            neutral_color="#888888", 
            icon_name="microphone", 
            icon_size="1.8x",
            key=f"mic_{st.session_state.uploader_key}"
        )

# --- 7. LOGIKA PEMROSESAN & TRANSLASI SUARA ---
prompt = prompt_text

if audio_bytes and not prompt_text:
    with st.spinner("Memproses suara..."):
        r = sr.Recognizer()
        try:
            with io.BytesIO(audio_bytes) as source_bytes:
                with sr.AudioFile(source_bytes) as source:
                    audio_data = r.record(source)
                    prompt = r.recognize_google(audio_data, language="id-ID")
        except sr.UnknownValueError:
            st.warning("Suara tidak terdengar jelas. Silakan coba lagi.")
            prompt = None
        except Exception as e:
            st.error(f"Gagal memproses suara: {e}")
            prompt = None

if prompt:
    teks_dokumen = ""
    if st.session_state.temp_doc:
        with st.spinner("Mengekstrak teks dokumen..."):
            teks_dokumen = ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
        if teks_dokumen:
            teks_dokumen = f"[KONTEN DOKUMEN: {st.session_state.temp_doc.name}]\n{teks_dokumen}\n[AKHIR KONTEN]\n\n"

    final_prompt = teks_dokumen + prompt

    if st.session_state.temp_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.temp_image)
        konten_payload = [
            {"type": "text", "text": final_prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
        ]
    else:
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
                temperature=0.3,
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
            
            st.session_state.messages[-1] = {"role": "user", "content": f"[User Query] {prompt}"}
            st.session_state.messages.append({"role": "assistant", "content": full_response})
            
            if st.session_state.current_session_id is None:
                st.session_state.current_session_id = str(uuid.uuid4())
            
            judul_chat = generate_title_from_messages(st.session_state.messages)
            
            # SIMPAN KE DATABASE
            save_session_db(st.session_state.current_session_id, st.session_state.username, judul_chat, st.session_state.messages)

            st.session_state.temp_image = None
            st.session_state.temp_doc = None
            st.session_state.uploader_key += 1 
            
            st.rerun()

        except Exception as e:
            st.error(f"Kesalahan teknis pada engine Lagos AI: {str(e)}")
            if st.session_state.messages[-1]["role"] == "user":
                st.session_state.messages.pop()
