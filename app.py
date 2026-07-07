import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document
from pypdf import PdfReader
import datetime

# --- 1. KONFIGURASI UTAMA & CSS CUSTOM ---
st.set_page_config(
    page_title="Qwen 3.5 - Professional Workspace",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan lebih profesional
st.markdown("""
    <style>
    /* Mengatur font dan warna dasar */
    .stApp {
        background-color: #f8f9fa;
    }
    .chat-message {
        padding: 1.5rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
        border: 1px solid #e0e0e0;
    }
    .user-msg {
        background-color: #e3f2fd;
        border-left: 5px solid #1976d2;
    }
    .ai-msg {
        background-color: #ffffff;
        border-left: 5px solid #6c5ce7;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    /* Styling Sidebar */
    section[data-testid=stSidebar] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
    }
    /* Styling Widget Input */
    .stTextInput > div > div > input {
        border-radius: 10px;
        border: 1px solid #ccc;
    }
    .stFileUploader > div {
        border-radius: 10px;
        border: 1px dashed #aaa;
    }
    /* Header Login */
    .login-container {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .login-box {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        text-align: center;
        width: 400px;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNGSI PEMBANTU ---

def konversi_gambar_ke_base64(uploaded_file):
    if uploaded_file is not None:
        return base64.b64encode(uploaded_file.read()).decode('utf-8')
    return None

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
        st.error(f"Gagal mengekstrak dokumen: {e}")

    return teks_hasil.strip()

def buat_file_word(riwayat_pesan, nama_user="Pengguna"):
    doc = Document()

    # Header Dokumen
    doc.add_heading(f'Laporan Analisis AI - {nama_user}', level=0)
    doc.add_paragraph(f'Dibuat pada: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}')
    doc.add_paragraph("-" * 50)

    for msg in riwayat_pesan:
        if msg["role"] == "system":
            continue

        if msg["role"] == "user":
            doc.add_heading(f"Pertanyaan dari {nama_user}:", level=2)
            if isinstance(msg["content"], list):
                for item in msg["content"]:
                    if item["type"] == "text":
                        doc.add_paragraph(item["text"])
            elif isinstance(msg["content"], str):
                doc.add_paragraph(msg["content"])

        elif msg["role"] == "assistant":
            doc.add_heading("Jawaban Qwen 3.5:", level=2)

            paragraf_list = msg["content"].split('\n')
            for p_text in paragraf_list:
                teks_bersih = p_text.strip()
                if not teks_bersih:
                    continue

                # Deteksi Heading
                match_heading = re.match(r'^(#{1,6})\s+(.*)$', teks_bersih)
                if match_heading:
                    level_pagar = len(match_heading.group(1))
                    konten_judul = match_heading.group(2).replace('**', '').replace('*', '')
                    doc.add_heading(konten_judul, level=min(level_pagar, 3))
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

                # Format Bold
                parts = re.split(r'(\*\*.*?\*\*)', teks_bersih)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        teks_tebal = part.replace('**', '')
                        p.add_run(teks_tebal).bold = True
                    else:
                        p.add_run(part)

            doc.add_paragraph("_" * 50).italic = True
            doc.add_paragraph() # Spasi antar blok

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 2. LOGIKA LOGIN (TABEL NAMA) ---

def check_login():
    if "user_name" not in st.session_state or not st.session_state.user_name:
        st.markdown("""
        <div class="login-container">
            <div class="login-box">
                <h2 style="color: #667eea;">🔮 Selamat Datang</h2>
                <p>Silakan masukkan nama Anda untuk memulai sesi analisis.</p>
                <form id="loginForm">
                    <input type="text" id="userNameInput" placeholder="Nama Lengkap Anda" style="width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #ddd; border-radius: 5px;">
                    <button type="submit" style="width: 100%; padding: 10px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer;">Masuk ke Workspace</button>
                </form>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Menggunakan form HTML sederhana untuk menangkap input di Streamlit
        # Streamlit tidak mendukung form HTML langsung, jadi kita gunakan st.form standar yang dimodifikasi tampilannya
        with st.form("login_form"):
            st.markdown("### Identifikasi Pengguna")
            nama_input = st.text_input("Masukkan Nama Anda:", label_visibility="collapsed", placeholder="Contoh: Budi Santoso")
            submit_login = st.form_submit_button("Lanjut ke Workspace 🚀", use_container_width=True)

            if submit_login and nama_input:
                st.session_state.user_name = nama_input.strip()
                st.session_state.messages = [
                    {"role": "system", "content": f"Anda adalah Qwen 3.5, asisten analitik tingkat tinggi. Berikan jawaban cerdas dalam Bahasa Indonesia. Nama pengguna saat ini adalah: {nama_input}."}
                ]
                st.rerun()
            elif submit_login and not nama_input:
                st.warning("Mohon isi nama Anda terlebih dahulu.")
        return False
    return True

# --- 3. LOGIKA UTAMA APLIKASI ---

def main():
    # Cek apakah user sudah login
    if not check_login():
        return

    nama_user = st.session_state.user_name

    # --- SIDEBAR ---
    with st.sidebar:
        st.markdown(f"### 👤 {nama_user}")
        st.markdown("---")
        st.markdown("### 🔮 Kontrol AI")
        st.info("⚡ Model: Qwen 3.5 122B (Multimodal)")

        st.divider()
        st.markdown("### 📥 Ekspor Dokumen")
        if "messages" in st.session_state and len(st.session_state.messages) > 1:
            file_word = buat_file_word(st.session_state.messages, nama_user)
            st.download_button(
                label="📥 Download Laporan (.docx)",
                data=file_word,
                file_name=f"Laporan_{nama_user.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.info("Mulai percakapan untuk mengunduh laporan.")

        st.divider()
        if st.button("🗑️ Reset Sesi & Logout", use_container_width=True, type="secondary"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    # --- HEADER UTAMA ---
    st.title(f"🔮 Lagos AI 8.3 - Workspace {nama_user}")
    st.caption("Asisten Analitik Cerdas untuk Dokumen & Gambar")

    # --- CHAT HISTORY ---
    # Tampilkan pesan chat dengan styling custom
    for message in st.session_state.messages:
        if message["role"] == "system":
            continue

        with st.chat_message(message["role"]):
            if message["role"] == "user":
                st.markdown(f"**{nama_user}:**")
                if isinstance(message["content"], list):
                    for item in message["content"]:
                        if item["type"] == "text":
                            st.markdown(item["text"])
                        elif item["type"] == "image_url":
                            st.image(item["image_url"]["url"].split(",")[1], caption="Gambar Terlampir")
                else:
                    st.markdown(message["content"])
            else:
                st.markdown(message["content"])

    st.divider()

    # --- INPUT AREA ---
    input_container = st.container()

    with input_container:
        col_img, col_doc = st.columns(2)

        with col_img:
            uploaded_image = st.file_uploader(
                "📷 Unggah Gambar (JPG/PNG)", 
                type=["jpg", "jpeg", "png"],
                label_visibility="collapsed",
                key="img_uploader"
            )

        with col_doc:
            uploaded_file = st.file_uploader(
                "📄 Unggah Dokumen (PDF/TXT)",
                type=["pdf", "txt"],
                label_visibility="collapsed",
                key="doc_uploader"
            )

        # Preview & Hapus File
        cols_preview = st.columns(2)
        with cols_preview[0]:
            if uploaded_image:
                st.image(uploaded_image, caption="Gambar Terpilih", width=100)
                if st.button("❌ Hapus Gambar", key="del_img_btn"):
                    del st.session_state.img_uploader
                    st.rerun()

        with cols_preview[1]:
            if uploaded_file:
                st.info(f"📄 {uploaded_file.name}")
                if st.button("❌ Hapus Dokumen", key="del_doc_btn"):
                    del st.session_state.doc_uploader
                    st.rerun()

        user_input = st.chat_input("Ketik pertanyaan Anda di sini...")

    # --- PROSES RESPON ---
    if user_input:
        teks_perintah_bersih = str(user_input).strip()
        perintah_final = teks_perintah_bersih

        # Handle Dokumen
        if uploaded_file:
            with st.spinner("📄 Sedang membaca dokumen..."):
                isi_teks_dokumen = ekstrak_teks_dari_dokumen(uploaded_file)

            if isi_teks_dokumen:
                perintah_final = (
                    f"Berikut adalah isi dokumen ({uploaded_file.name}):\n"
                    f"\"\"\"\n{isi_teks_dokumen}\n\"\"\"\n\n"
                    f"Perintah: {teks_perintah_bersih}"
                )

        # Handle Gambar
        if uploaded_image:
            base64_image = konversi_gambar_ke_base64(uploaded_image)
            konten_payload = [
                {"type": "text", "text": perintah_final},
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                }
            ]
        else:
            konten_payload = perintah_final

        # Tampilkan pesan user
        with st.chat_message("user"):
            if isinstance(konten_payload, list):
                for item in konten_payload:
                    if item["type"] == "text":
                        st.write(item["text"])
                    elif item["type"] == "image_url":
                        st.image(base64.b64decode(item["image_url"]["url"].split(",")[1]))
            else:
                st.write(konten_payload)

        st.session_state.messages.append({"role": "user", "content": konten_payload})

        # Panggil API
        with st.chat_message("assistant"):
            try:
                placeholder = st.empty()
                full_response = ""

                # Konfigurasi API (Pastikan API Key aman di production, simpan di secret)
                # Untuk demo ini, kita gunakan variabel langsung atau st.secrets
                api_key = st.secrets.get("NVIDIA_API_KEY", "nvapi-...") # Ganti dengan secret jika deploy

                client = OpenAI(base_url="https://integrate.api.nvidia.com/v1", api_key=api_key)

                response_stream = client.chat.completions.create(
                    model="qwen/qwen3.5-122b-a10b",
                    messages=st.session_state.messages,
                    temperature=0.6,
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
                st.session_state.messages.append({"role": "assistant", "content": full_response})

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

if __name__ == "__main__":
    main()
