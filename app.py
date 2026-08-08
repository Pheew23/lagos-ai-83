import io
import os
import re
import json
import uuid
import base64
import sqlite3
import hashlib
import datetime
from contextlib import contextmanager
from typing import List, Dict, Any, Optional

import requests
from bs4 import BeautifulSoup
from docx import Document
from pptx import Presentation
import speech_recognition as sr
import streamlit as st
import streamlit.components.v1 as components
import extra_streamlit_components as stx
from openai import OpenAI
from audio_recorder_streamlit import audio_recorder

# ==========================================
# 1. KONSTANTA & KONFIGURASI PENGATURAN
# ==========================================
DB_NAME = 'lagos_multiuser.db'
API_KEY = st.secrets["NVIDIA_API_KEY"]
BASE_URL = "https://integrate.api.nvidia.com/v1"

MODEL_MAPPING = {
    "openai/gpt-oss-120b": "1. Flash (text only)",
    "thinkingmachines/inkling": "2. Pro(text only)",
    "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning": "3. Pro (Analisis)",
    "google/diffusiongemma-26b-a4b-it": "4. Stable",
    "deepseek-ai/deepseek-v4-flash": "5. limited",
    "google/veo-3.1-fast-generate-preview": "6. Generator Gambar (coming soon)"
}

SYSTEM_PROMPT = """Anda adalah Lagøs AI 9.1, asisten analitik tingkat tinggi yang dikembangkan oleh Rian Dev.

ATURAN KETAT UNTUK MERESPONS:
1. JANGAN PERNAH memperkenalkan diri, menyebutkan nama, atau menjelaskan kemampuan Anda, KECUALI pengguna secara spesifik bertanya tentang identitas Anda (contoh: "Siapa kamu?", "Buatan siapa kamu?").
2. Jika tidak ditanya tentang identitas, jawab langsung ke inti pertanyaan pengguna tanpa basa-basi pengenalan diri.
3. Dilarang keras menyebutkan identitas model AI dasar Anda. Anda hanya Lagøs AI 9.1.

ATURAN PEMBUATAN APLIKASI WEB (HTML):
Jika pengguna meminta Anda untuk membuat web app atau aplikasi, Anda HARUS menuliskan seluruh kodenya dalam SATU file HTML lengkap (gabungkan CSS dan JS di dalamnya) dan bungkus dengan blok kode html.

ATURAN PEMBUATAN PRESENTASI (PPT):
Jika pengguna meminta Anda untuk merangkum teks/file menjadi presentasi atau membuat PPT, Anda HARUS bertindak sebagai Art Director. Analisis topiknya dan pilih TEMA yang paling cocok dari daftar ini: "bisnis", "kreatif", "akademik", atau "gelap".
Rangkum materi menjadi poin-poin presentasi dan kembalikan datanya MURNI dalam format JSON di dalam blok kode json berikut (tanpa tambahan teks lain):
```json
{
  "judul_presentasi": "Judul Utama PPT",
  "rekomendasi_tema": "bisnis",
  "slides": [
    {
      "slide_type": "title",
      "title": "Judul Presentasi Utama",
      "content": "Sub-judul / Nama Penulis"
    },
    {
      "slide_type": "content",
      "title": "Judul Slide 1",
      "content": ["Poin penting 1", "Poin penting 2", "Poin penting 3"]
    }
  ]
}
```"""

# ==========================================
# 2. MANAJER DATABASE
# ==========================================
class DatabaseManager:
    @staticmethod
    @contextmanager
    def get_connection():
        conn = sqlite3.connect(DB_NAME)
        try:
            yield conn
        finally:
            conn.close()

    @classmethod
    def init_db(cls):
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS sessions (session_id TEXT PRIMARY KEY, username TEXT, title TEXT, updated_at TIMESTAMP)''')
            c.execute('''CREATE TABLE IF NOT EXISTS messages (id INTEGER PRIMARY KEY AUTOINCREMENT, session_id TEXT, role TEXT, content TEXT)''')
            conn.commit()

    @staticmethod
    def hash_password(password: str) -> str:
        return hashlib.sha256(password.encode()).hexdigest()

    @classmethod
    def register_user(cls, username: str, password: str) -> bool:
        with cls.get_connection() as conn:
            c = conn.cursor()
            try:
                c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, cls.hash_password(password)))
                conn.commit()
                return True
            except sqlite3.IntegrityError:
                return False

    @classmethod
    def authenticate_user(cls, username: str, password: str) -> bool:
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT password FROM users WHERE username=?", (username,))
            row = c.fetchone()
            if row and row[0] == cls.hash_password(password):
                return True
            return False

    @classmethod
    def get_user_sessions(cls, username: str) -> List[tuple]:
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT session_id, title FROM sessions WHERE username=? ORDER BY updated_at DESC", (username,))
            return c.fetchall()

    @classmethod
    def load_session_messages(cls, session_id: str) -> List[Dict[str, Any]]:
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT role, content FROM messages WHERE session_id=? ORDER BY id ASC", (session_id,))
            rows = c.fetchall()
            
        msgs = [{"role": "system", "content": SYSTEM_PROMPT}]
        for r, content_str in rows:
            try:
                msgs.append({"role": r, "content": json.loads(content_str)})
            except json.JSONDecodeError:
                msgs.append({"role": r, "content": content_str})
        return msgs

    @classmethod
    def save_session(cls, session_id: str, username: str, title: str, messages: List[Dict[str, Any]]):
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute("INSERT OR REPLACE INTO sessions (session_id, username, title, updated_at) VALUES (?, ?, ?, ?)", 
                      (session_id, username, title, datetime.datetime.now()))
            c.execute("DELETE FROM messages WHERE session_id=?", (session_id,))
            
            for msg in messages:
                if msg["role"] != "system":
                    content = json.dumps(msg["content"]) if isinstance(msg["content"], (dict, list)) else json.dumps(msg["content"])
                    c.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)", 
                              (session_id, msg["role"], content))
            conn.commit()

    @classmethod
    def delete_session(cls, session_id: str):
        with cls.get_connection() as conn:
            c = conn.cursor()
            c.execute("DELETE FROM sessions WHERE session_id=?", (session_id,))
            c.execute("DELETE FROM messages WHERE session_id=?", (session_id,))
            conn.commit()

# ==========================================
# 3. UTILITIES & PEMROSESAN MULTIMEDIA
# ==========================================
class MediaUtils:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def konversi_gambar_ke_base64(uploaded_file) -> Optional[str]:
        if uploaded_file is not None:
            return base64.b64encode(uploaded_file.read()).decode('utf-8')
        return None

    @staticmethod
    @st.cache_data(show_spinner=False)
    def ekstrak_teks_dari_dokumen(uploaded_file) -> str:
        teks_hasil = ""
        nama_file = uploaded_file.name.lower()
        try:
            if nama_file.endswith('.pdf'):
                from pypdf import PdfReader
                reader = PdfReader(uploaded_file)
                for page in reader.pages:
                    teks = page.extract_text()
                    if teks: 
                        teks_hasil += teks + "\n"
            elif nama_file.endswith('.txt'):
                teks_hasil = uploaded_file.read().decode("utf-8")
            elif nama_file.endswith('.docx'):
                doc = Document(uploaded_file)
                for para in doc.paragraphs:
                    teks_hasil += para.text + "\n"
            return teks_hasil.strip()
        except Exception as e:
            return f"Gagal membaca dokumen: {str(e)}"

    @staticmethod
    def buat_file_word(riwayat_pesan: List[Dict[str, Any]]) -> io.BytesIO:
        doc = Document()
        doc.add_heading('Lagøs AI 9.1 - Analisis Laporan', 0)
        
        for msg in riwayat_pesan:
            if msg["role"] == "system": continue
            role_title = "User" if msg["role"] == "user" else "Lagøs AI 9.1"
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

    @staticmethod
    def generate_title_from_messages(messages: List[Dict[str, Any]]) -> str:
        for msg in messages:
            if msg["role"] == "user":
                content = msg["content"]
                text = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
                text = text.split("[AKHIR KONTEN]\n\n")[-1]
                return text[:25] + "..." if len(text) > 25 else (text if text else "Obrolan Gambar/File")
        return "Obrolan Baru"

    @staticmethod
    def ambil_teks_dari_link(url: str) -> str:
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

    @staticmethod
    def ekstrak_kode_html(teks: str) -> Optional[str]:
        match = re.search(r'```html\n(.*?)\n```', teks, re.DOTALL | re.IGNORECASE)
        return match.group(1) if match else None

    @staticmethod
    def ekstrak_json_ppt(teks: str) -> Optional[dict]:
        match = re.search(r'```json\n(.*?)\n```', teks, re.DOTALL | re.IGNORECASE)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                return None
        return None

    @staticmethod
    def buat_file_ppt(data_json: dict) -> io.BytesIO:
        tema_pilihan = data_json.get("rekomendasi_tema", "bisnis").lower()
        peta_template = {
            "bisnis": "tema_bisnis.pptx",
            "kreatif": "tema_kreatif.pptx",
            "akademik": "tema_akademik.pptx",
            "gelap": "tema_gelap.pptx"
        }
        
        file_template = peta_template.get(tema_pilihan, "tema_bisnis.pptx")
        
        if os.path.exists(file_template):
            prs = Presentation(file_template)
        else:
            prs = Presentation()
            
        for slide_data in data_json.get("slides", []):
            stype = slide_data.get("slide_type", "content")
            if stype == "title":
                slide_layout = prs.slide_layouts[0]
                slide = prs.slides.add_slide(slide_layout)
                try:
                    slide.shapes.title.text = slide_data.get("title", "")
                    slide.placeholders[1].text = slide_data.get("content", "")
                except Exception:
                    pass
            else:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                try:
                    slide.shapes.title.text = slide_data.get("title", "")
                    body = slide.placeholders[1]
                    tf = body.text_frame
                    content = slide_data.get("content", [])
                    
                    if isinstance(content, list):
                        for i, poin in enumerate(content):
                            if i == 0:
                                tf.text = poin
                            else:
                                p = tf.add_paragraph()
                                p.text = poin
                                p.level = 0
                    else:
                        tf.text = str(content)
                except Exception:
                    pass
                    
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio

# ==========================================
# 4. KOMPONEN UI & TAMPILAN
# ==========================================
def inject_custom_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700&display=swap');
            html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            
            .header-title {
                text-align: center; font-size: 2.2rem; font-weight: 700;
                background: linear-gradient(90deg, #7d4eff, #00d2ff);
                -webkit-background-clip: text; -webkit-text-fill-color: transparent;
                margin-bottom: 0px; padding-top: 10px;
            }
            .header-subtitle {
                text-align: center; color: var(--text-color); opacity: 0.7;
                font-size: 0.95rem; font-weight: 300; margin-bottom: 30px;
            }
            .stChatMessage:nth-child(even) {
                background-color: var(--secondary-background-color) !important;
                border-radius: 12px; padding: 1rem;
            }
            .file-pill {
                display: inline-block; background: var(--secondary-background-color);
                color: var(--text-color); padding: 4px 14px; border-radius: 20px;
                font-size: 0.8rem; margin-right: 8px; margin-bottom: 12px;
                border: 1px solid var(--border-color);
            }
            [data-testid="stHorizontalBlock"] { align-items: center !important; }
            [data-testid="stPopover"] button {
                border-radius: 50% !important; height: 48px !important; width: 48px !important;
                padding: 0 !important; display: flex !important; justify-content: center !important;
                align-items: center !important; border: 1px solid var(--border-color) !important;
                background-color: transparent !important; transition: all 0.3s ease !important;
            }
            [data-testid="stPopover"] button:hover {
                border-color: #7d4eff !important; background-color: rgba(125, 78, 255, 0.1) !important;
                color: #7d4eff !important; transform: scale(1.05) !important;
            }
            [data-testid="stSidebar"] .stButton > button {
                border-radius: 8px !important; padding: 0.5rem 1rem !important;
                text-align: left !important; justify-content: flex-start !important;
            }
            [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button {
                background-color: transparent !important; border: 1px solid transparent !important;
                justify-content: center !important; padding: 0.5rem 0 !important;
                color: #888888 !important; transition: all 0.2s ease;
            }
            [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button:hover {
                border: 1px solid #ff4b4b !important; color: #ff4b4b !important;
                background-color: rgba(255, 75, 75, 0.1) !important;
            }
        </style>
    """, unsafe_allow_html=True)

def render_login_ui():
    st.markdown('<div class="header-title">🔮 Lagøs AI 9.1</div>', unsafe_allow_html=True)
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
                    if DatabaseManager.authenticate_user(log_user, log_pass):
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
                        if DatabaseManager.register_user(reg_user, reg_pass):
                            st.success("✅ Berhasil mendaftar! Silakan buka tab 'Masuk'.")
                        else:
                            st.error("❌ Username sudah dipakai, silakan pilih yang lain.")
                    else:
                        st.warning("⚠️ Harap isi username dan password!")

def inject_auto_scroll():
    components.html("""
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
    """, height=0)

@st.dialog("🌐 Web App Preview", width="large")
def render_webapp_modal(html_code: str):
    st.info("💡 Interaksi dengan Web App di bawah ini. Tekan 'X' di sudut kanan atas untuk menutup.")
    injection = """
    <base target="_blank">
    <script>
        document.addEventListener("click", function(e) {
            var target = e.target.closest("a");
            if(target && target.href) {
                target.setAttribute("target", "_blank");
            }
        });
    </script>
    """
    if re.search(r'<head[^>]*>', html_code, re.IGNORECASE):
        html_code = re.sub(r'(<head[^>]*>)', r'\1\n' + injection, html_code, count=1, flags=re.IGNORECASE)
    else:
        html_code = injection + "\n" + html_code
    components.html(html_code, height=600, scrolling=True)

# ==========================================
# 5. INISIALISASI & SETUP APLIKASI
# ==========================================
def init_session_state():
    defaults = {
        "logged_in": False,
        "username": "",
        "current_session_id": None,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT}],
        "temp_image": None,
        "temp_doc": None,
        "uploader_key": 0
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

# ==========================================
# 6. LOGIKA UTAMA APLIKASI
# ==========================================
def main():
    st.set_page_config(
        page_title="Lagøs AI 9.1 | Premium Chat",
        page_icon="🔮",
        layout="centered", 
        initial_sidebar_state="expanded"
    )
    inject_custom_css()
    DatabaseManager.init_db()
    init_session_state()

    cookie_manager = stx.CookieManager(key="cookie_manager")
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
        render_login_ui()
        st.stop()
    
    st.markdown('<div class="header-title">🔮 Lagøs AI 9.1</div>', unsafe_allow_html=True)
    st.markdown('<div class="header-subtitle">Assistant AI</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.success(f"👤 Login sebagai: **{st.session_state.username}**")
        
        if st.button("➕ Mulai Obrolan Baru", use_container_width=True, type="primary"):
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
            st.rerun()

        st.markdown("### 🗂️ Riwayat Obrolan")
        sessions = DatabaseManager.get_user_sessions(st.session_state.username)
        
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
                            st.session_state.messages = DatabaseManager.load_session_messages(sess_id)
                            st.rerun()
                    with col_del:
                        if st.button("🗑️", key=f"del_{sess_id}", help="Hapus obrolan ini"):
                            DatabaseManager.delete_session(sess_id)
                            if st.session_state.current_session_id == sess_id:
                                st.session_state.current_session_id = None
                                st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
                            st.rerun()

        st.divider()
        st.markdown("### 🧠 Pilih Model AI")
        selected_model = st.selectbox(
            label="Pilih model aktif:",
            options=list(MODEL_MAPPING.keys()),
            index=0,
            format_func=lambda x: MODEL_MAPPING[x],
            label_visibility="collapsed"
        )
        
        if len(st.session_state.messages) > 1:
            file_word = MediaUtils.buat_file_word(st.session_state.messages)
            st.download_button(
                label="📥 Unduh Laporan (.DOCX)",
                data=file_word,
                file_name="Lagøs_AI_9.1_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        st.divider()
        if st.button("🚪 Keluar (Logout)", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
            st.session_state.del_cookie = True 
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

        st.markdown(
            '''
            <div style="background-color: var(--secondary-background-color); border: 1px dashed var(--border-color); padding: 15px; border-radius: 10px; text-align: center; margin-top: 20px;">
                <span style="color: #888; font-size: 0.75rem;">Advertisement Space</span><br>
                <span style="font-size: 0.9rem;">Integrasi Iklan Akan Ditampilkan Di Sini</span>
            </div>
            ''', unsafe_allow_html=True
        )

    if len(st.session_state.messages) == 1:
        st.markdown("<p style='text-align: center; margin-top: 5vh; color: #666;'>Sistem siap.</p>", unsafe_allow_html=True)

    for idx, message in enumerate(st.session_state.messages):
        if message["role"] == "system": continue
        with st.chat_message(message["role"]):
            content = message["content"]
            text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
            
            st.markdown(text_disp)
            
            if message["role"] == "assistant":
                html_code = MediaUtils.ekstrak_kode_html(text_disp)
                if html_code:
                    st.write("") 
                    if st.button("🚀 Tampilkan Web App", key=f"btn_webapp_{idx}", use_container_width=True):
                        render_webapp_modal(html_code)
                
                json_ppt = MediaUtils.ekstrak_json_ppt(text_disp)
                if json_ppt:
                    st.write("")
                    ppt_file = MediaUtils.buat_file_ppt(json_ppt)
                    nama_file = f"{json_ppt.get('judul_presentasi', 'Lagøs_Presentation')}.pptx"
                    
                    st.download_button(
                        label="📊 Unduh Presentasi (.PPTX)",
                        data=ppt_file,
                        file_name=nama_file.replace(" ", "_"),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"btn_ppt_{idx}",
                        use_container_width=True,
                        type="primary"
                    )

    st.markdown("<div style='height: 90px'></div>", unsafe_allow_html=True)
    st.markdown("<div id='bottom-marker'></div>", unsafe_allow_html=True)
    inject_auto_scroll()

    with st.container():
        uploader_idx = st.session_state.uploader_key
        current_img = st.session_state.get(f"img_{uploader_idx}")
        current_doc = st.session_state.get(f"doc_{uploader_idx}")

        if current_img:
            st.markdown(f"<div class='file-pill'>📷 Gambar telah dilampirkan</div>", unsafe_allow_html=True)
        if current_doc:
            st.markdown(f"<div class='file-pill'>📄 Dokumen telah dilampirkan</div>", unsafe_allow_html=True)

        col_attach, col_input, col_mic = st.columns([1, 8, 1])
        
        with col_attach:
            with st.popover("➕"): 
                st.markdown("**Lampirkan File**")
                # DI SINI PERUBAHANNYA: docx ditambahkan ke tipe file yang diterima
                up_img = st.file_uploader("Upload Image", type=["jpg", "png", "jpeg"], label_visibility="collapsed", key=f"img_{uploader_idx}")
                up_doc = st.file_uploader("Upload Doc", type=["pdf", "txt", "docx"], label_visibility="collapsed", key=f"doc_{uploader_idx}")
                st.session_state.temp_image = up_img
                st.session_state.temp_doc = up_doc

        with col_input:
            prompt_text = st.chat_input("Tanyakan sesuatu pada Lagøs AI 9.1...")

        with col_mic:
            audio_bytes = audio_recorder(
                text="", 
                recording_color="#ff4b4b",
                neutral_color="#888888", 
                icon_name="microphone", 
                icon_size="1.8x",
                key=f"mic_{uploader_idx}"
            )

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
                teks_dok = MediaUtils.ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
                if teks_dok:
                    teks_tambahan += f"\n[KONTEN DOKUMEN: {st.session_state.temp_doc.name}]\n{teks_dok}\n[AKHIR KONTEN DOKUMEN]\n"

        urls_found = re.compile(r'https?://\S+').findall(prompt)
        if urls_found:
            with st.spinner("Mengekstrak informasi dari Link..."):
                for url in urls_found:
                    teks_web = MediaUtils.ambil_teks_dari_link(url)
                    teks_web_singkat = teks_web[:4000]
                    teks_tambahan += f"\n[ISI WEBSITE: {url}]\n{teks_web_singkat}\n[AKHIR ISI WEBSITE]\n"

        final_prompt = f"{teks_tambahan}\n\nPertanyaan/Instruksi Pengguna:\n{prompt}" if teks_tambahan else prompt

        if st.session_state.temp_image:
            base64_img = MediaUtils.konversi_gambar_ke_base64(st.session_state.temp_image)
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
                if selected_model == "google/veo-3.1-fast-generate-preview":
                    with st.spinner("Menghasilkan gambar..."):
                        img_response = client.images.generate(
                            model=selected_model,
                            prompt=prompt,
                            n=1
                        )
                        image_url = img_response.data[0].url
                        full_response = f"![Gambar yang Dihasilkan]({image_url})"
                        placeholder.markdown(full_response)
                else:
                    response_stream = client.chat.completions.create(
                        model=selected_model, 
                        messages=st.session_state.messages,
                        temperature=0.7,
                        max_tokens=12096,
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
                
                judul_chat = MediaUtils.generate_title_from_messages(st.session_state.messages)
                DatabaseManager.save_session(
                    st.session_state.current_session_id, 
                    st.session_state.username, 
                    judul_chat, 
                    st.session_state.messages
                )

                st.session_state.temp_image = None
                st.session_state.temp_doc = None
                st.session_state.uploader_key += 1 
                
                st.rerun()

            except Exception as e:
                st.error(f"Kesalahan teknis pada engine Lagøs AI: {str(e)}")
                if st.session_state.messages[-1]["role"] == "user":
                    st.session_state.messages.pop()

if __name__ == "__main__":
    main()
