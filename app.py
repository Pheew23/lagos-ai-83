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
from datetime import datetime
import streamlit.components.v1 as components

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Premium Chat",
    page_icon="🔮",
    layout="centered", 
    initial_sidebar_state="expanded"
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

        [data-testid="stHorizontalBlock"] {
            align-items: center !important;
        }

        [data-testid="stPopover"] button {
            border-radius: 50% !important;
            height: 48px !important;
            width: 48px !important;
            padding: 0 !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
            background-color: transparent !important;
            transition: all 0.3s ease !important;
        }
        
        [data-testid="stPopover"] button:hover {
            border-color: #7d4eff !important;
            background-color: rgba(125, 78, 255, 0.1) !important;
            color: #7d4eff !important;
            transform: scale(1.05) !important;
        }
        
        .history-btn p {
            margin: 0;
            font-size: 0.9rem;
            text-align: left;
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
    # Tabel Pengguna
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (username TEXT PRIMARY KEY, password TEXT)''')
    # Tabel Sesi
    c.execute('''CREATE TABLE IF NOT EXISTS sessions
                 (session_id TEXT PRIMARY KEY, username TEXT, title TEXT, updated_at TIMESTAMP)''')
    # Tabel Pesan
    c.execute('''CREATE TABLE IF NOT EXISTS messages
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, session_id TEXT, role TEXT, content TEXT)''')
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

# --- 3. SISTEM AUTENTIKASI (LOGIN/REGISTER) ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.markdown('<div class="header-title">🔮 Lagos AI 9.1</div>', unsafe_allow_html=True)
    st.markdown('<div class="header-subtitle">Silakan Masuk untuk Mengakses Asisten</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        tab_login, tab_register = st.tabs(["🔑 Masuk", "📝 Daftar Baru"])
        
        with tab_login:
            log_user = st.text_input("Username", key="log_user")
            log_pass = st.text_input("Password", type="password", key="log_pass")
            if st.button("Masuk", use_container_width=True, type="primary"):
                if authenticate_user(log_user, log_pass):
                    st.session_state.logged_in = True
                    st.session_state.username = log_user
                    st.rerun()
                else:
                    st.error("Username atau password salah!")
                    
        with tab_register:
            reg_user = st.text_input("Username Baru", key="reg_user")
            reg_pass = st.text_input("Password Baru", type="password", key="reg_pass")
            if st.button("Daftar & Buat Akun", use_container_width=True):
                if reg_user and reg_pass:
                    if register_user(reg_user, reg_pass):
                        st.success("✅ Berhasil mendaftar! Silakan buka tab 'Masuk'.")
                    else:
                        st.error("❌ Username sudah dipakai, silakan pilih yang lain.")
                else:
                    st.warning("⚠️ Harap isi username dan password!")
    
    st.stop()

# ==========================================
# KODE DI BAWAH INI HANYA JALAN JIKA SUDAH LOGIN
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

# --- BRANDING UTAMA ---
st.markdown('<div class="header-title">🔮 Lagos AI 9.1</div>', unsafe_allow_html=True)
st.markdown('<div class="header-subtitle">Premium Multimodal Assistant</div>', unsafe_allow_html=True)

# --- SIDEBAR (FITUR HISTORY SPESIFIK USER) ---
with st.sidebar:
    st.success(f"👤 Login sebagai: **{st.session_state.username}**")
    
    if st.button("➕ Mulai Obrolan Baru", use_container_width=True, type="primary"):
        st.session_state.current_session_id = None
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        st.rerun()

    st.markdown("### 🗂️ Riwayat Obrolan Anda")
    sessions = get_user_sessions(st.session_state.username)
    
    if not sessions:
        st.caption("Belum ada riwayat obrolan.")
    else:
        for sess_id, title in sessions:
            col_btn, col_del = st.columns([8, 2])
            with col_btn:
                btn_type = "primary" if st.session_state.current_session_id == sess_id else "secondary"
                if st.button(title, key=f"btn_{sess_id}", use_container_width=True, type=btn_type):
                    st.session_state.current_session_id = sess_id
                    st.session_state.messages = load_session_messages(sess_id)
                    st.rerun()
            with col_del:
                if st.button("🗑️", key=f"del_{sess_id}"):
                    delete_session_db(sess_id)
                    if st.session_state.current_session_id == sess_id:
                        st.session_state.current_session_id = None
                        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
                    st.rerun()

    st.divider()
    
    st.markdown("### 🧠 Pilih Model AI")
    MODEL_MAPPING = {
        "openai/gpt-oss-120b": "1. Sangat Cepat (text only)",
        "thinkingmachines/inkling": "2. Cepat(text only)",
        "mistralai/mistral-medium-3.5-128b": "3. Analisis Mendalam",
        "deepseek-ai/deepseek-v4-pro": "4. Stabil",
        "nvidia/nemotron-3-ultra-550b-a55b": "5. Projek Khusus"
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
            label="📥 Unduh Laporan (.DOCX)",
            data=file_word,
            file_name="Lagos_AI_9.1_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    st.divider()
    if st.button("🚪 Keluar (Logout)", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.current_session_id = None
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        st.rerun()
        
    st.markdown("### 🛠️ Admin Panel")
    try:
        with open(DB_NAME, "rb") as db_file:
            st.download_button(
                label="📦 Download Database (Admin)",
                data=db_file,
                file_name="lagos_multiuser.db",
                mime="application/octet-stream",
                use_container_width=True
            )
    except FileNotFoundError:
        pass

# --- 5. AREA OBROLAN UTAMA ---
if len(st.session_state.messages) == 1:
    st.markdown("<p style='text-align: center; margin-top: 5vh; color: #666;'>Sistem siap. Lampirkan file, bicara, atau ketik <b>/gambar [teks]</b> untuk buat gambar.</p>", unsafe_allow_html=True)

for message in st.session_state.messages:
    if message["role"] == "system": continue
    with st.chat_message(message["role"]):
        content = message["content"]
        text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        st.markdown(text_disp, unsafe_allow_html=True)

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
        st.markdown(f"<div class='file-pill'>📷 Gambar telah dilampirkan</div>", unsafe_allow_html=True)
    if current_doc:
        st.markdown(f"<div class='file-pill'>📄 Dokumen telah dilampirkan</div>", unsafe_allow_html=True)

    col_attach, col_input, col_mic = st.columns([1, 8, 1])
    
    with col_attach:
        with st.popover("➕"): 
            st.markdown("**Lampirkan File**")
            up_img = st.file_uploader("Upload Image", type=["jpg", "png", "jpeg"], label_visibility="collapsed", key=f"img_{st.session_state.uploader_key}")
            up_doc = st.file_uploader("Upload Doc", type=["pdf", "txt"], label_visibility="collapsed", key=f"doc_{st.session_state.uploader_key}")
            st.session_state.temp_image = up_img
            st.session_state.temp_doc = up_doc

    with col_input:
        prompt_text = st.chat_input("Tanyakan AI... (Gunakan /gambar [deskripsi] untuk visual)")

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
    with st.spinner("Menerjemahkan suara..."):
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
            st.error(f"Sistem gagal memproses suara: {e}")
            prompt = None

if prompt:
    # --- FITUR BARU: GENERATE GAMBAR ---
    if prompt.strip().lower().startswith("/gambar "):
        image_prompt = prompt.strip()[8:].strip()
        
        with st.chat_message("user"):
            st.markdown(prompt)
            
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        with st.chat_message("assistant"):
            with st.spinner("🎨 Lagos AI sedang merender gambar..."):
                try:
                    import urllib.parse
                    base64_img = ""
                    image_url_final = ""
                    
                    # 1. Coba gunakan Hugging Face terlebih dahulu
                    HF_TOKEN = st.secrets.get("HF_TOKEN", "") 
                    if HF_TOKEN:
                        try:
                            api_url = "https://api-inference.huggingface.co/models/black-forest-labs/FLUX.1-schnell"
                            headers = {"Authorization": f"Bearer {HF_TOKEN}"}
                            payload = {"inputs": image_prompt, "options": {"wait_for_model": True}}
                            
                            response = requests.post(api_url, headers=headers, json=payload, timeout=15)
                            if response.status_code == 200:
                                image_bytes = response.content
                                base64_img = base64.b64encode(image_bytes).decode('utf-8')
                        except:
                            pass # Jika gagal terhubung/DNS error, otomatis lanjut ke cadangan
                            
                    # 2. Jika Hugging Face gagal (termasuk karena masalah NameResolution), gunakan Pollinations sebagai cadangan stabil
                    if base64_img:
                        img_markdown = f"Berikut adalah hasil render untuk: **{image_prompt}**\n\n![Generated Image](data:image/jpeg;base64,{base64_img})"
                    else:
                        encoded_prompt = urllib.parse.quote(image_prompt)
                        image_url_final = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=1024&height=1024&nologo=true"
                        img_markdown = f"Berikut adalah hasil render untuk: **{image_prompt}**\n\n![Generated Image]({image_url_final})"
                        
                    st.markdown(img_markdown, unsafe_allow_html=True)
                    st.session_state.messages.append({"role": "assistant", "content": img_markdown})
                    
                    # Simpan ke Database
                    if st.session_state.current_session_id is None:
                        st.session_state.current_session_id = str(uuid.uuid4())
                    
                    judul_chat = generate_title_from_messages(st.session_state.messages)
                    save_session_db(st.session_state.current_session_id, st.session_state.username, judul_chat, st.session_state.messages)
                    
                    st.session_state.temp_image = None
                    st.session_state.temp_doc = None
                    st.session_state.uploader_key += 1 
                            
                except Exception as e:
                    st.error(f"Kesalahan internal: {str(e)}")
                    st.session_state.messages.pop()
                    
    # --- FITUR STANDAR: LLM CHAT (TIDAK DIUBAH) ---
    else:
        teks_dokumen = ""
        if st.session_state.temp_doc:
            with st.spinner("Membaca referensi dokumen..."):
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
                
                save_session_db(st.session_state.current_session_id, st.session_state.username, judul_chat, st.session_state.messages)

                st.session_state.temp_image = None
                st.session_state.temp_doc = None
                st.session_state.uploader_key += 1 
                
                st.rerun()

            except Exception as e:
                st.error(f"Kesalahan teknis pada engine Lagos AI: {str(e)}")
                if st.session_state.messages[-1]["role"] == "user":
                    st.session_state.messages.pop()
