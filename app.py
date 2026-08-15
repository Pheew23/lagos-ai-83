import io
import os
import re
import json
import uuid
import base64
import sqlite3
import hashlib
import datetime
import copy
from contextlib import contextmanager
from typing import List, Dict, Any, Optional

import requests
import yfinance as yf
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

# Trik untuk menghindari bug render markdown di antarmuka web
B3 = "`" * 3

MODEL_MAPPING = {
    "minimaxai/minimax-m3": "1. Flash (Pro)",
    "google/diffusiongemma-26b-a4b-it": "2. Stable",
    "thinkingmachines/inkling": "3. Pro(text only)",
    "z-ai/glm-5.2": "4. Pro (Analisis)",
    "openai/gpt-oss-120b": "5. limited",
    "google/veo-3.1-fast-generate-preview": "6. Generator Gambar (coming soon)"
}

SYSTEM_PROMPT = """Anda adalah Lagøs AI 9.1, asisten analitik tingkat tinggi yang dikembangkan oleh Rian Dev.

ATURAN KETAT UNTUK MERESPONS UMUM:
1. JANGAN PERNAH memperkenalkan diri, menyebutkan nama, atau menjelaskan kemampuan Anda, KECUALI pengguna secara spesifik bertanya tentang identitas Anda.
2. Jika tidak ditanya tentang identitas, jawab langsung ke inti pertanyaan pengguna tanpa basa-basi.
3. Dilarang keras menyebutkan identitas model AI dasar Anda. Anda hanya Lagøs AI 9.1.
4. Jangan Pernah membagikan informasi sensitif.

ATURAN KONFIRMASI FORMAT OUTPUT:
Jika pengguna memerintahkan Anda untuk membuat sesuatu (misalnya tugas, rencana, rangkuman, dll) namun BELUM menyebutkan format spesifik, Anda WAJIB menahan diri untuk tidak langsung membuatnya dan HARUS bertanya kembali kepada pengguna dengan kalimat: "Dalam bentuk apa hasilnya? aplikasi atau word/pdf?". JANGAN hasilkan kontennya sebelum pengguna memilih.

ATURAN ANALISIS TRADING (LONG & SHORT):
Jika pengguna bertanya tentang prospek pasar, koin, saham, atau kapan harus LONG/SHORT, dan sistem melampirkan [DATA PASAR TERBARU]:
1. Bertindaklah sebagai Master Trader Institusional. 
2. Analisis tren dari data harga yang diberikan.
3. Tentukan probabilitas kecenderungan arah pasar (Bullish/Bearish).
4. Berikan rekomendasi tegas: "🟢 SINYAL LONG", "🔴 SINYAL SHORT", atau "🟡 HOLD (Tunggu/Jangan Masuk)".
5. WAJIB sertakan estimasi level Take Profit (TP) dan Stop Loss (SL) yang logis.

ATURAN PEMBUATAN APLIKASI WEB (HTML):
Perhatikan baik-baik [STATUS SAKLAR] yang dikirimkan bersama pertanyaan. Jika pengguna menyalakan saklar (ON), Anda boleh menulis kode aplikasi dalam SATU file HTML lengkap. Jika saklar MATI (OFF), Anda DILARANG menulis kode HTML/aplikasi sama sekali.

ATURAN PEMBUATAN DOKUMEN (WORD/PDF):
Jika pengguna meminta Anda untuk membuat dokumen, artikel, surat, makalah, atau laporan dalam bentuk Word, DOCX, atau PDF, Anda HARUS merangkum isi kontennya dan menaruhnya MURNI di dalam blok kode `document`.
Contoh:
%sdocument
# Judul Dokumen
## Sub Judul
Ini adalah paragraf dari dokumen...
- Poin 1
- Poin 2
%s

ATURAN PEMBUATAN PRESENTASI (PPT):
Jika diminta merangkum teks menjadi PPT, Anda HARUS bertindak sebagai Art Director. Pilih TEMA ("bisnis", "kreatif", "akademik", atau "gelap"). Kembalikan MURNI dalam JSON:
%sjson
{
  "judul_presentasi": "Judul Utama PPT",
  "rekomendasi_tema": "bisnis",
  "slides": [
    {
      "slide_type": "title",
      "title": "Judul Utama",
      "content": "Sub-judul"
    }
  ]
}
%s""" % (B3, B3, B3, B3)

# ==========================================
# 2. MANAJER DATABASE
# ==========================================
class DatabaseManager:
    @staticmethod
    @contextmanager
    def get_connection():
        conn = sqlite3.connect(DB_NAME)
        try: yield conn
        finally: conn.close()

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
            return bool(row and row[0] == cls.hash_password(password))

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
            try: msgs.append({"role": r, "content": json.loads(content_str)})
            except: msgs.append({"role": r, "content": content_str})
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
        if uploaded_file is not None: return base64.b64encode(uploaded_file.read()).decode('utf-8')
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
                    if teks: teks_hasil += teks + "\n"
            elif nama_file.endswith('.txt'):
                teks_hasil = uploaded_file.read().decode("utf-8")
            elif nama_file.endswith('.docx'):
                doc = Document(uploaded_file)
                for para in doc.paragraphs: teks_hasil += para.text + "\n"
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
    def ekstrak_dokumen(teks: str) -> Optional[str]:
        match = re.search(r'`{3}document\n(.*?)\n`{3}', teks, re.DOTALL | re.IGNORECASE)
        return match.group(1) if match else None

    @staticmethod
    def buat_dokumen_docx(konten: str) -> io.BytesIO:
        doc = Document()
        for line in konten.split('\n'):
            line = line.strip()
            if not line: continue
            if line.startswith('# '): doc.add_heading(line[2:], level=1)
            elif line.startswith('## '): doc.add_heading(line[3:], level=2)
            elif line.startswith('### '): doc.add_heading(line[4:], level=3)
            elif line.startswith('- '): doc.add_paragraph(line[2:], style='List Bullet')
            else: doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    @staticmethod
    def buat_dokumen_pdf(konten: str) -> io.BytesIO:
        try:
            from fpdf import FPDF
        except ImportError:
            raise ImportError("Fitur PDF diblokir karena library fpdf2 belum diinstal. Jalankan di terminal: pip install fpdf2")
            
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("helvetica", size=12)
        
        for line in konten.split('\n'):
            line = line.strip()
            if not line:
                pdf.ln(5)
                continue
                
            if line.startswith('# '):
                pdf.set_font("helvetica", style="B", size=16)
                pdf.multi_cell(0, 10, text=line[2:])
                pdf.set_font("helvetica", size=12)
            elif line.startswith('## '):
                pdf.set_font("helvetica", style="B", size=14)
                pdf.multi_cell(0, 10, text=line[3:])
                pdf.set_font("helvetica", size=12)
            elif line.startswith('### '):
                pdf.set_font("helvetica", style="B", size=12)
                pdf.multi_cell(0, 8, text=line[4:])
                pdf.set_font("helvetica", size=12)
            elif line.startswith('- '):
                pdf.multi_cell(0, 8, text=f"• {line[2:]}")
            else:
                pdf.multi_cell(0, 8, text=line)
                
        bio = io.BytesIO(pdf.output())
        return bio

    @staticmethod
    def generate_title_from_messages(messages: List[Dict[str, Any]]) -> str:
        for msg in messages:
            if msg["role"] == "user":
                content = msg["content"]
                text = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
                text = text.split("[AKHIR KONTEN]\n\n")[-1]
                return text[:25] + "..." if len(text) > 25 else (text if text else "Obrolan Baru")
        return "Obrolan Baru"

    @staticmethod
    def ambil_teks_dari_link(url: str) -> str:
        try:
            headers = {'User-Agent': 'Mozilla/5.0'} 
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            text = ' '.join([p.get_text() for p in soup.find_all('p')])
            return re.sub(r'\s+', ' ', text).strip()
        except Exception as e: return f"Error: {str(e)}"

    @staticmethod
    def ekstrak_kode_html(teks: str) -> Optional[str]:
        match = re.search(r'`{3}html\n(.*?)\n`{3}', teks, re.DOTALL | re.IGNORECASE)
        return match.group(1) if match else None

    @staticmethod
    def ekstrak_json_ppt(teks: str) -> Optional[dict]:
        match = re.search(r'`{3}json\n(.*?)\n`{3}', teks, re.DOTALL | re.IGNORECASE)
        if match:
            try: return json.loads(match.group(1))
            except: pass
        return None

    @staticmethod
    def buat_file_ppt(data_json: dict) -> io.BytesIO:
        tema_pilihan = data_json.get("rekomendasi_tema", "bisnis").lower()
        peta = {"bisnis": "tema_bisnis.pptx", "kreatif": "tema_kreatif.pptx", "akademik": "tema_akademik.pptx", "gelap": "tema_gelap.pptx"}
        file_template = peta.get(tema_pilihan, "tema_bisnis.pptx")
        
        prs = Presentation(file_template) if os.path.exists(file_template) else Presentation()
            
        for slide_data in data_json.get("slides", []):
            stype = slide_data.get("slide_type", "content")
            if stype == "title":
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                try:
                    slide.shapes.title.text = slide_data.get("title", "")
                    slide.placeholders[1].text = slide_data.get("content", "")
                except: pass
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                try:
                    slide.shapes.title.text = slide_data.get("title", "")
                    tf = slide.placeholders[1].text_frame
                    content = slide_data.get("content", [])
                    if isinstance(content, list):
                        for i, poin in enumerate(content):
                            if i == 0: tf.text = poin
                            else: tf.add_paragraph().text, tf.paragraphs[-1].level = poin, 0
                    else: tf.text = str(content)
                except: pass
                    
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio

class MarketUtils:
    @staticmethod
    def ambil_data_pasar(simbol_ticker: str) -> str:
        try:
            if not simbol_ticker.endswith("-USD") and len(simbol_ticker) <= 5:
                if simbol_ticker.upper() in ['BTC', 'ETH', 'SOL', 'BNB', 'XRP', 'ADA', 'DOGE']:
                    simbol_ticker = f"{simbol_ticker.upper()}-USD"
            ticker = yf.Ticker(simbol_ticker)
            hist = ticker.history(period="5d")
            if hist.empty: return f"Data pasar '{simbol_ticker}' tidak ditemukan."
            data_str = hist[['Open', 'High', 'Low', 'Close', 'Volume']].to_string()
            return f"Data 5 Hari Terakhir {simbol_ticker}:\n{data_str}"
        except Exception as e: return f"Gagal: {str(e)}"

# ==========================================
# 4. KOMPONEN UI & TAMPILAN
# ==========================================
def inject_custom_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700&display=swap');
            html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
            #MainMenu {visibility: hidden;} footer {visibility: hidden;}
            .header-title { text-align: center; font-size: 2.2rem; font-weight: 700; background: linear-gradient(90deg, #7d4eff, #00d2ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0px; padding-top: 10px; }
            .header-subtitle { text-align: center; color: var(--text-color); opacity: 0.7; font-size: 0.95rem; font-weight: 300; margin-bottom: 30px; }
            .stChatMessage:nth-child(even) { background-color: var(--secondary-background-color) !important; border-radius: 12px; padding: 1rem; }
            .file-pill { display: inline-block; background: var(--secondary-background-color); color: var(--text-color); padding: 4px 14px; border-radius: 20px; font-size: 0.8rem; margin-right: 8px; margin-bottom: 12px; border: 1px solid var(--border-color); }
        </style>
    """, unsafe_allow_html=True)

def inject_auto_scroll():
    components.html("""
        <script>
            setTimeout(function() {
                var parentDoc = window.parent.document;
                var marker = parentDoc.getElementById('bottom-marker');
                if (marker) marker.scrollIntoView({behavior: 'auto', block: 'end'});
            }, 300);
        </script>
    """, height=0)

@st.dialog("🌐 Web App Preview", width="large")
def render_webapp_modal(html_code: str):
    st.info("💡 Interaksi dengan Web App di bawah ini.")
    injection = "<base target='_blank'><script>document.addEventListener('click', function(e) { var t = e.target.closest('a'); if(t && t.href) { t.setAttribute('target', '_blank'); } });</script>"
    if re.search(r'<head[^>]*>', html_code, re.IGNORECASE):
        html_code = re.sub(r'(<head[^>]*>)', r'\1\n' + injection, html_code, count=1, flags=re.IGNORECASE)
    else: html_code = injection + "\n" + html_code
    components.html(html_code, height=600, scrolling=True)

def init_session_state():
    defaults = {
        "logged_in": False, "username": "", "current_session_id": None,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT}],
        "temp_image": None, "temp_doc": None, "uploader_key": 0,
        "tahap2_pending": False, "trigger_tahap2": False
    }
    for key, val in defaults.items():
        if key not in st.session_state: st.session_state[key] = val

def main():
    st.set_page_config(page_title="Lagøs AI 9.1", page_icon="🔮", layout="centered", initial_sidebar_state="expanded")
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
        st.markdown('<div class="header-title">🔮 Lagøs AI 9.1</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1.5, 1])
        with col2:
            with st.container(border=True):
                tab_login, tab_register = st.tabs(["🔑 Masuk", "📝 Daftar Baru"])
                with tab_login:
                    log_user = st.text_input("Username", key="log_user")
                    log_pass = st.text_input("Password", type="password", key="log_pass")
                    if st.button("Masuk", use_container_width=True, type="primary"):
                        if DatabaseManager.authenticate_user(log_user, log_pass):
                            st.session_state.logged_in = True
                            st.session_state.username = log_user
                            st.session_state.set_cookie = True 
                            st.rerun()
                        else: st.error("Username atau password salah!")
                with tab_register:
                    reg_user = st.text_input("Username Baru", key="reg_user")
                    reg_pass = st.text_input("Password Baru", type="password", key="reg_pass")
                    if st.button("Daftar & Buat Akun", use_container_width=True):
                        if reg_user and reg_pass:
                            if DatabaseManager.register_user(reg_user, reg_pass): st.success("✅ Berhasil mendaftar!")
                            else: st.error("❌ Username sudah dipakai.")
                        else: st.warning("⚠️ Harap isi data!")
        st.stop()
    
    st.markdown('<div class="header-title">🔮 Lagøs AI 9.1</div>', unsafe_allow_html=True)
    st.markdown('<div class="header-subtitle">Assistant AI</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.success(f"👤 Login sebagai: **{st.session_state.username}**")
        
        st.divider()
        
        if st.button("➕ Mulai Obrolan Baru", use_container_width=True, type="primary"):
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
            st.session_state.tahap2_pending = False
            st.session_state.trigger_tahap2 = False
            st.rerun()

        st.markdown("### 🗂️ Riwayat Obrolan")
        sessions = DatabaseManager.get_user_sessions(st.session_state.username)
        if sessions:
            with st.container(height=300, border=False):
                for sess_id, title in sessions:
                    col_btn, col_del = st.columns([6, 1], gap="small") 
                    with col_btn:
                        btn_type = "primary" if st.session_state.current_session_id == sess_id else "secondary"
                        if st.button(title, key=f"btn_{sess_id}", use_container_width=True, type=btn_type):
                            st.session_state.current_session_id = sess_id
                            st.session_state.messages = DatabaseManager.load_session_messages(sess_id)
                            st.session_state.tahap2_pending = False
                            st.session_state.trigger_tahap2 = False
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
        selected_model = st.selectbox("Pilih model aktif:", list(MODEL_MAPPING.keys()), format_func=lambda x: MODEL_MAPPING[x], label_visibility="collapsed")
        
        if len(st.session_state.messages) > 1:
            st.download_button("📥 Unduh Laporan Chat", data=MediaUtils.buat_file_word(st.session_state.messages), file_name="Lagøs_AI_Chat.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        st.divider()
        if st.button("🚪 Keluar (Logout)", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.current_session_id = None
            st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
            st.session_state.tahap2_pending = False
            st.session_state.trigger_tahap2 = False
            st.session_state.del_cookie = True 
            st.rerun()

    for idx, message in enumerate(st.session_state.messages):
        if message["role"] == "system": continue
        with st.chat_message(message["role"]):
            content = message["content"]
            text_disp = next((item["text"] for item in content if item["type"] == "text"), "") if isinstance(content, list) else str(content)
            
            st.markdown(text_disp)
            
            if message["role"] == "assistant":
                is_last_message = (idx == len(st.session_state.messages) - 1)
                is_pending = is_last_message and st.session_state.get("tahap2_pending")
                
                # Cek HTML Web App
                html_code = MediaUtils.ekstrak_kode_html(text_disp)
                if html_code and not is_pending:
                    st.write("") 
                    if st.button("🚀 Tampilkan Web App", key=f"btn_webapp_{idx}", use_container_width=True):
                        render_webapp_modal(html_code)
                
                # Cek PPT
                json_ppt = MediaUtils.ekstrak_json_ppt(text_disp)
                if json_ppt:
                    st.write("")
                    ppt_file = MediaUtils.buat_file_ppt(json_ppt)
                    st.download_button("📊 Unduh Presentasi (.PPTX)", data=ppt_file, file_name=f"{json_ppt.get('judul_presentasi', 'PPT')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"btn_ppt_{idx}", use_container_width=True, type="primary")

                # Cek Dokumen PDF/DOCX baru
                dokumen_teks = MediaUtils.ekstrak_dokumen(text_disp)
                if dokumen_teks and not is_pending:
                    st.write("")
                    col_doc1, col_doc2 = st.columns(2)
                    with col_doc1:
                        docx_file = MediaUtils.buat_dokumen_docx(dokumen_teks)
                        st.download_button(
                            label="📄 Unduh (.DOCX)", 
                            data=docx_file, 
                            file_name="Dokumen_Lagos.docx", 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                            key=f"btn_docx_{idx}", 
                            use_container_width=True, 
                            type="primary"
                        )
                    with col_doc2:
                        try:
                            pdf_file = MediaUtils.buat_dokumen_pdf(dokumen_teks)
                            st.download_button(
                                label="📕 Unduh (.PDF)", 
                                data=pdf_file, 
                                file_name="Dokumen_Lagos.pdf", 
                                mime="application/pdf", 
                                key=f"btn_pdf_{idx}", 
                                use_container_width=True, 
                                type="primary"
                            )
                        except Exception as e:
                            st.error(str(e))

    st.markdown("<div style='height: 40px'></div>", unsafe_allow_html=True)
    st.markdown("<div id='bottom-marker'></div>", unsafe_allow_html=True)
    inject_auto_scroll()

    if st.session_state.get("tahap2_pending"):
        st.info("⚠️ **Tahap 1 Selesai.** Aplikasi belum lengkap. Lanjutkan untuk mencegah timeout.")
        if st.button("⚡ Lanjutkan ke Tahap 2", type="primary", use_container_width=True):
            st.session_state.trigger_tahap2 = True
            st.session_state.tahap2_pending = False
            st.rerun()

    # ==========================================
    # POSISI BARU SAKLAR (Tepat di atas kolom chat)
    # ==========================================
    with st.container():
        app_mode = st.toggle("🚀 Izinkan Buat Aplikasi Web", value=False, help="Nyalakan jika Anda secara eksplisit meminta AI membuat web (HTML). Jika mati, AI hanya akan membalas dengan obrolan biasa.")
        
        uploader_idx = st.session_state.uploader_key
        if st.session_state.get(f"img_{uploader_idx}"): st.markdown(f"<div class='file-pill'>📷 Gambar telah dilampirkan</div>", unsafe_allow_html=True)
        if st.session_state.get(f"doc_{uploader_idx}"): st.markdown(f"<div class='file-pill'>📄 Dokumen telah dilampirkan</div>", unsafe_allow_html=True)

        col_attach, col_input, col_mic = st.columns([1, 8, 1])
        with col_attach:
            with st.popover("➕"): 
                st.session_state.temp_image = st.file_uploader("Upload Image", type=["jpg", "png", "jpeg"], label_visibility="collapsed", key=f"img_{uploader_idx}")
                st.session_state.temp_doc = st.file_uploader("Upload Doc", type=["pdf", "txt", "docx"], label_visibility="collapsed", key=f"doc_{uploader_idx}")
        with col_input:
            prompt_text = st.chat_input("Tanyakan sesuatu..." if not st.session_state.get("tahap2_pending") else "Terkunci (Selesaikan Tahap 2)", disabled=st.session_state.get("tahap2_pending"))
        with col_mic:
            audio_bytes = audio_recorder(text="", recording_color="#ff4b4b", neutral_color="#888888", icon_name="microphone", icon_size="1.8x", key=f"mic_{uploader_idx}")

    prompt = prompt_text
    if audio_bytes and not prompt_text:
        with st.spinner("Menerjemahkan suara..."):
            try: prompt = sr.Recognizer().recognize_google(sr.Recognizer().record(sr.AudioFile(io.BytesIO(audio_bytes))), language="id-ID")
            except: st.warning("Suara tidak terdengar jelas.")

    is_tahap2_exec = st.session_state.get("trigger_tahap2", False)

    if prompt or is_tahap2_exec:
        if is_tahap2_exec:
            st.session_state.trigger_tahap2 = False 
        else:
            with st.chat_message("user"): st.markdown(prompt)
            teks_tambahan = ""
            
            # INJEKSI LOGIKA SAKLAR APLIKASI
            if app_mode:
                teks_tambahan += "\n[STATUS SAKLAR: ON (MODE APLIKASI). Anda DIIZINKAN merender kode HTML/Aplikasi lengkap jika pengguna memintanya.]\n"
            else:
                teks_tambahan += "\n[STATUS SAKLAR: OFF (MODE NORMAL). Anda DILARANG KERAS membuat atau menulis kode aplikasi web/HTML. Jawablah menggunakan teks biasa saja meskipun pengguna meminta dibuatkan aplikasi.]\n"
            
            if st.session_state.temp_doc:
                teks_dok = MediaUtils.ekstrak_teks_dari_dokumen(st.session_state.temp_doc)
                if teks_dok: teks_tambahan += f"\n[KONTEN DOKUMEN]\n{teks_dok}\n"

            urls_found = re.compile(r'https?://\S+').findall(prompt)
            for url in urls_found:
                teks_tambahan += f"\n[ISI WEBSITE: {url}]\n{MediaUtils.ambil_teks_dari_link(url)[:4000]}\n"

            if any(kata in prompt.lower() for kata in ["analisis", "short", "long", "beli", "jual", "prospek"]):
                potensi_ticker = re.findall(r'\b[A-Z]{3,5}(?:-[A-Z]+|\.JK)?\b', prompt.upper())
                if potensi_ticker: teks_tambahan += f"\n[DATA PASAR TERBARU]\n{MarketUtils.ambil_data_pasar(potensi_ticker[0])}\n"

            final_prompt = f"{teks_tambahan}\n\nPertanyaan/Instruksi Pengguna:\n{prompt}" if teks_tambahan else prompt

            if st.session_state.temp_image:
                base64_img = MediaUtils.konversi_gambar_ke_base64(st.session_state.temp_image)
                st.session_state.messages.append({"role": "user", "content": [{"type": "text", "text": final_prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}]})
            else:
                st.session_state.messages.append({"role": "user", "content": final_prompt})

        with st.chat_message("assistant"):
            client = OpenAI(base_url=BASE_URL, api_key=API_KEY)
            placeholder = st.empty()
            full_response = ""

            try:
                if selected_model == "google/veo-3.1-fast-generate-preview":
                    with st.spinner("Menghasilkan gambar..."):
                        image_url = client.images.generate(model=selected_model, prompt=prompt, n=1).data[0].url
                        placeholder.markdown(f"![Gambar yang Dihasilkan]({image_url})")
                        st.session_state.messages.append({"role": "assistant", "content": f"![Gambar yang Dihasilkan]({image_url})"})
                else:
                    if is_tahap2_exec:
                        st.info("⏳ Memproses Tahap 2...")
                        tahap2_msgs = copy.deepcopy(st.session_state.messages)
                        tahap2_msgs.append({
                            "role": "user",
                            "content": "TAHAP 2: Lanjutkan kode sebelumnya. Tuliskan HANYA sisa kode JavaScript-nya (di dalam tag <script>) dan tutup semua sisa tag HTML/BODY-nya. \nPENTING: JANGAN ulangi kode awal dan JANGAN membuka awalan blok markdown baru (jangan tulis " + B3 + "html atau " + B3 + "javascript lagi), sambung langsung saja agar menyatu."
                        })
                        
                        last_msg_content = st.session_state.messages[-1]["content"]
                        if last_msg_content.strip().endswith(B3):
                            last_msg_content = last_msg_content.rstrip("` \n") 
                            
                        response_stream = client.chat.completions.create(model=selected_model, messages=tahap2_msgs, temperature=0.7, max_tokens=12096, stream=True)
                        for chunk in response_stream:
                            if chunk.choices and len(chunk.choices) > 0:
                                delta = chunk.choices[0].delta.content
                                if delta:
                                    full_response += delta
                                    placeholder.markdown(last_msg_content + "\n" + full_response + "▌")
                                    
                        gabungan_kode = last_msg_content + "\n" + full_response
                        if not gabungan_kode.strip().endswith(B3): gabungan_kode += "\n" + B3
                        
                        placeholder.markdown(gabungan_kode)
                        st.session_state.messages[-1]["content"] = gabungan_kode 
                        
                    else:
                        # PROSES TAHAP 1 (Hanya dieksekusi jika SAKLAR ON dan ada kata kunci aplikasi)
                        is_web_app = app_mode and any(kata in prompt.lower() for kata in ["buat", "bikin", "aplikasi", "web", "html", "app"])
                        payload_msgs = copy.deepcopy(st.session_state.messages)
                        
                        if is_web_app:
                            st.info("⏳ Mode Web/App (Saklar ON): Memproses Tahap 1...")
                            instruksi_tahap_1 = "\n\n[INSTRUKSI SISTEM PENTING]: Karena potensi timeout, kerjakan pembuatan web dalam 2 TAHAP. TAHAP 1: Tuliskan kerangka dasar HTML dan CSS-nya saja, bungkus dalam SATU blok " + B3 + "html. PENTING: JANGAN tulis JavaScript, dan JANGAN tutup tag </body> atau </html> pada tahap ini."
                            if isinstance(payload_msgs[-1]["content"], list): payload_msgs[-1]["content"][0]["text"] += instruksi_tahap_1
                            else: payload_msgs[-1]["content"] += instruksi_tahap_1

                        response_stream = client.chat.completions.create(model=selected_model, messages=payload_msgs, temperature=0.7, max_tokens=12096, stream=True)
                        for chunk in response_stream:
                            if chunk.choices and len(chunk.choices) > 0:
                                delta = chunk.choices[0].delta.content
                                if delta:
                                    full_response += delta
                                    placeholder.markdown(full_response + "▌")
                                    
                        placeholder.markdown(full_response)
                        st.session_state.messages.append({"role": "assistant", "content": full_response})
                        
                        if is_web_app:
                            st.session_state.tahap2_pending = True

                if st.session_state.current_session_id is None:
                    st.session_state.current_session_id = str(uuid.uuid4())
                
                DatabaseManager.save_session(st.session_state.current_session_id, st.session_state.username, MediaUtils.generate_title_from_messages(st.session_state.messages), st.session_state.messages)

                if not is_tahap2_exec:
                    st.session_state.temp_image = None
                    st.session_state.temp_doc = None
                    st.session_state.uploader_key += 1 
                
                st.rerun()

            except Exception as e:
                st.error(f"Kesalahan teknis: {str(e)}")
                if st.session_state.messages[-1]["role"] == "user": st.session_state.messages.pop()

if __name__ == "__main__":
    main()
