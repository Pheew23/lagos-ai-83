import streamlit as st
import extra_streamlit_components as stx
from openai import OpenAI
import io
import re
import base64
import requests
from docx import Document
from audio_recorder_streamlit import audio_recorder
import speech_recognition as sr
import json
import uuid
import hashlib
import datetime
import streamlit.components.v1 as components
from bs4 import BeautifulSoup 

# --- IMPORT FIREBASE ---
import firebase_admin
from firebase_admin import credentials, firestore

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Lagos AI 9.1 | Premium Chat",
    page_icon="🔮",
    layout="centered", 
    initial_sidebar_state="expanded"
)

# --- 2. CUSTOM CSS ---
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
            color: var(--text-color);
            opacity: 0.7;
            font-size: 0.95rem;
            font-weight: 300;
            margin-bottom: 30px;
        }
        
        .stChatMessage:nth-child(even) {
            background-color: var(--secondary-background-color) !important;
            border-radius: 12px;
            padding: 1rem;
        }
        
        .file-pill {
            display: inline-block;
            background: var(--secondary-background-color);
            color: var(--text-color);
            padding: 4px 14px;
            border-radius: 20px;
            font-size: 0.8rem;
            margin-right: 8px;
            margin-bottom: 12px;
            border: 1px solid var(--border-color);
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
            border: 1px solid var(--border-color) !important;
            background-color: transparent !important;
            transition: all 0.3s ease !important;
        }
        
        [data-testid="stPopover"] button:hover {
            border-color: #7d4eff !important;
            background-color: rgba(125, 78, 255, 0.1) !important;
            color: #7d4eff !important;
            transform: scale(1.05) !important;
        }
        
        [data-testid="stSidebar"] .stButton > button {
            border-radius: 8px !important;
            padding: 0.5rem 1rem !important;
            text-align: left !important;
            justify-content: flex-start !important;
        }
        
        [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button {
            background-color: transparent !important;
            border: 1px solid transparent !important;
            justify-content: center !important;
            padding: 0.5rem 0 !important;
            color: #888888 !important;
            transition: all 0.2s ease;
        }
        
        [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button:hover {
            border: 1px solid #ff4b4b !important;
            color: #ff4b4b !important;
            background-color: rgba(255, 75, 75, 0.1) !important;
        }
        
        [data-testid="stSidebar"] .stButton > button p {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            width: 100%;
            margin: 0;
        }
    </style>
""", unsafe_allow_html=True)

# --- PENGELOLA COOKIE ---
cookie_manager = stx.CookieManager(key="cookie_manager")

# --- INISIALISASI FIREBASE ---
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        # Mengambil kredensial dari st.secrets
        cred_dict = dict(st.secrets["firebase"])
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_firebase()

# --- FUNGSI DATABASE FIREBASE ---
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def register_user(username, password):
    user_ref = db.collection('users').document(username)
    if user_ref.get().exists:
        return False # Username sudah ada
    user_ref.set({'password': hash_password(password)})
    return True

def authenticate_user(username, password):
    user_doc = db.collection('users').document(username).get()
    if user_doc.exists:
        if user_doc.to_dict().get('password') == hash_password(password):
            return True
    return False

def get_user_sessions(username):
    # Membutuhkan Index Composite Firestore di konsol nantinya
    docs = db.collection('sessions').where('username', '==', username).order_by('updated_at', direction=firestore.Query.DESCENDING).stream()
    return [(doc.id, doc.to_dict().get('title', 'Obrolan Baru')) for doc in docs]

def load_session_messages(session_id):
    session_doc = db.collection('sessions').document(session_id).get()
    if session_doc.exists:
        data = session_doc.to_dict()
        if 'messages' in data:
            return data['messages']
    
    return [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]

def save_session_db(session_id, username, title, messages):
    db.collection('sessions').document(session_id).set({
        'username': username,
        'title': title,
        'updated_at': firestore.SERVER_TIMESTAMP,
        'messages': messages
    })

def delete_session_db(session_id):
    db.collection('sessions').document(session_id).delete()


# --- 3. SISTEM AUTENTIKASI (LOGIN/REGISTER DENGAN COOKIES) ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

cookie_logged_in = cookie_manager.get("is_logged_in")
cookie_username = cookie_manager.get("saved_username")

if st.session_state.get("del_cookie") == True:
    cookie_manager.delete("is_logged_in", key="del_login_cookie")
    cookie_manager.delete("saved_username", key="del_user_cookie")
    st.session_state.del_cookie = False 
    cookie_logged_in = None 
    cookie_username = None

if cookie_logged_in == "True" and not st.session_state.logged_in:
    st.session_state.logged_in = True
    st.session_state.username = cookie_username

if st.session_state.get("set_cookie") == True:
    expire_date = datetime.datetime.now() + datetime.timedelta(days=7)
    cookie_manager.set("is_logged_in", "True", expires_at=expire_date, key="set_login_cookie")
    cookie_manager.set("saved_username", st.session_state.username, expires_at=expire_date, key="set_user_cookie")
    st.session_state.set_cookie = False

if not st.session_state.logged_in:
    st.markdown('<div class="header-title">🔮 Lagos AI 9.1</div>', unsafe_allow_html=True)
    st.markdown('<div class="header-subtitle">Silakan Masuk untuk Mengakses Asisten</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.container(border=True):
            tab_login, tab_register = st.tabs(["🔑 Masuk", "📝 Daftar Baru"])
            
            with tab_login:
                st.markdown("<h4 style='text-align: center; margin-bottom: 20px;'>Selamat Datang Kembali</h4>", unsafe_allow_html=True)
                log_user = st.text_input("Username", key="log_user", placeholder="Masukkan username Anda...")
                log_pass = st.text_input("Password", type="password", key="log_pass", placeholder="Masukkan password Anda...")
                st.markdown("<br>", unsafe_allow_html=True)
                
                if st.button("Masuk", use_container_width=True, type="primary"):
                    if authenticate_user(log_user, log_pass):
                        st.session_state.logged_in = True
                        st.session_state.username = log_user
                        st.session_state.set_cookie = True 
                        st.rerun()
                    else:
                        st.error("Username atau password salah!")
                        
            with tab_register:
                st.markdown("<h4 style='text-align: center; margin-bottom: 20px;'>Buat Akun Baru</h4>", unsafe_allow_html=True)
                reg_user = st.text_input("Username Baru", key="reg_user", placeholder="Pilih username unik...")
                reg_pass = st.text_input("Password Baru", type="password", key="reg_pass", placeholder="Buat password yang kuat...")
                st.markdown("<br>", unsafe_allow_html=True)
                
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

def ambil_teks_dari_link(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'} 
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        paragraphs = soup.find_all('p')
        text = ' '.join([p.get_text() for p in paragraphs])
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    except Exception as e:
        return f"Error saat membaca link: {str(e)}"

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

# --- SIDEBAR ---
with st.sidebar:
    st.success(f"👤 Login sebagai: **{st.session_state.username}**")
    
    if st.button("➕ Mulai Obrolan Baru", use_container_width=True, type="primary"):
        st.session_state.current_session_id = None
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        st.rerun()

    st.markdown("### 🗂️ Riwayat Obrolan")
    sessions = get_user_sessions(st.session_state.username)
    
    if not sessions:
        st.caption("Belum ada riwayat obrolan.")
    else:
        with st.container(height=300, border=False):
            for sess_id, title in sessions:
                col_btn, col_del = st.columns([6, 1], gap="small") 
                
                with col_btn:
                    btn_type = "primary" if st.session_state.current_session_id == sess_id else "secondary"
                    if st.button(title, key=f"btn_{sess_id}", use_container_width=True, type=btn_type):
                        st.session_state.current_session_id = sess_id
                        st.session_state.messages = load_session_messages(sess_id)
                        st.rerun()
                with col_del:
                    if st.button("🗑️", key=f"del_{sess_id}", help="Hapus obrolan ini"):
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
        "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning": "3. Analisis Mendalam",
        "google/diffusiongemma-26b-a4b-it": "4. Stabil",
        "deepseek-ai/deepseek-v4-flash": "5. Projek Khusus",
        "google/veo-3.1-fast-generate-preview": "6. Generator Gambar (Veo)"
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
    
    # TOMBOL LOGOUT 
    if st.button("🚪 Keluar (Logout)", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.current_session_id = None
        st.session_state.messages = [{"role": "system", "content": "Anda adalah Lagos AI 9.1 (Rian Dev), asisten analitik tingkat tinggi."}]
        
        st.session_state.del_cookie = True 
        st.rerun()
        
    st.markdown("### 🛠️ Admin Panel")
    # Karena tidak menggunakan SQLite lagi, download database tidak berlaku.
    st.caption("Penyimpanan menggunakan Firebase (Cloud).")

    # Ruang Placeholder untuk Integrasi Iklan
    st.markdown(
        '''
        <div style="background-color: var(--secondary-background-color); border: 1px dashed var(--border-color); padding: 15px; border-radius: 10px; text-align: center; margin-top: 20px;">
            <span style="color: #888; font-size: 0.75rem;">Advertisement Space</span><br>
            <span style="font-size: 0.9rem;">Integrasi Iklan Akan Ditampilkan Di Sini</span>
        </div>
        ''', unsafe_allow_html=True
    )

# --- 5. AREA OBROLAN UTAMA ---
if len(st.session_state.messages) == 1:
    st.markdown("<p style='text-align: center; margin-top: 5vh; color: #666;'>Sistem siap. Lampirkan gambar/dokumen atau bicara melalui mikrofon.</p>", unsafe_allow_html=True)

for message in st.session_state.messages:
    if message["role"] == "system": continue
    with st.chat_message(message["role"]):
        content = message["content"]
        text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
        st.markdown(text_disp)

st.markdown("<div style='height: 90px'></div>", unsafe_allow_html=True)

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

# --- 6. AREA INPUT TERPADU ---
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
        prompt_text = st.chat_input("Tanyakan sesuatu atau tempelkan link (http://...) pada Lagos AI 9.1...")

    with col_mic:
        audio_bytes = audio_recorder(
            text="", 
            recording_color="#ff4b4b",
            neutral_color="#888888", 
            icon_name="microphone", 
            icon_size="1.8x",
            key=f"mic_{st.session_state.uploader_key}"
        )

# --- 7. LOGIKA PEMROSESAN ---
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
    with st.chat_message("user"):
        st.markdown(prompt)

    teks_tambahan = ""

    if st.session_state.temp_doc:
        with st.spinner("Membaca referensi dokumen..."):
            teks_dok = ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
            if teks_dok:
                teks_tambahan += f"\n[KONTEN DOKUMEN: {st.session_state.temp_doc.name}]\n{teks_dok}\n[AKHIR KONTEN DOKUMEN]\n"

    url_pattern = re.compile(r'https?://\S+')
    urls_found = url_pattern.findall(prompt)
    
    if urls_found:
        with st.spinner("Mengekstrak informasi dari Link..."):
            for url in urls_found:
                teks_web = ambil_teks_dari_link(url)
                teks_web_singkat = teks_web[:4000]
                teks_tambahan += f"\n[ISI WEBSITE: {url}]\n{teks_web_singkat}\n[AKHIR ISI WEBSITE]\n"

    if teks_tambahan:
        final_prompt = f"{teks_tambahan}\n\nPertanyaan/Instruksi Pengguna:\n{prompt}"
    else:
        final_prompt = prompt

    if st.session_state.temp_image:
        base64_img = konversi_gambar_ke_base64(st.session_state.temp_image)
        konten_payload = [
            {"type": "text", "text": final_prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
        ]
    else:
        konten_payload = final_prompt 

    st.session_state.messages.append({"role": "user", "content": konten_payload})

    with st.chat_message("assistant"):
        client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
        placeholder = st.empty()
        full_response = ""

        try:
            if MODEL_NAME == "google/veo-3.1-fast-generate-preview":
                with st.spinner("Menghasilkan gambar..."):
                    img_response = client.images.generate(
                        model=MODEL_NAME,
                        prompt=prompt,
                        n=1
                    )
                    image_url = img_response.data[0].url
                    full_response = f"![Gambar yang Dihasilkan]({image_url})"
                    placeholder.markdown(full_response)
            else:
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
            
            st.session_state.messages[-1] = {"role": "user", "content": prompt}
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
