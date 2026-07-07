import streamlit as st
from openai import OpenAI
import io
import re
import base64
from docx import Document
# Import pustaka untuk membaca PDF
from pypdf import PdfReader

# --- 1. KONFIGURASI UTAMA STREAMLIT ---
st.set_page_config(
    page_title="Qwen 3.5 Document Workspace",
    page_icon="🔮",
    layout="wide"
)

# --- FUNGSI PEMBANTU: KONVERSI GAMBAR KE BASE64 ---
def konversi_gambar_ke_base64(uploaded_file):
    if uploaded_file is not None:
        return base64.b64encode(uploaded_file.read()).decode('utf-8')
    return None

# --- FUNGSI PEMBANTU: EKSTRAKSI TEKS DARI PDF/TXT ---
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
        st.error(f"Gagal mengekstrak dokumen {uploaded_file.name}: {e}")
        
    return teks_hasil.strip()

# --- 2. FUNGSI UNTUK MEMBUAT FILE WORD (.DOCX) ---
def buat_file_word(riwayat_pesan):
    doc = Document()
    doc.add_heading('Draf Hasil Kerja AI - Qwen Workspace', level=0)
    
    for msg in riwayat_pesan:
        if msg["role"] == "system":
            continue
            
        if msg["role"] == "user":
            doc.add_heading("Pertanyaan / Instruksi Anda:", level=2)
            if isinstance(msg["content"], list):
                for item in msg["content"]:
                    if item["type"] == "text":
                        doc.add_paragraph(item["text"])
            elif isinstance(msg["content"], str):
                doc.add_paragraph(msg["content"])
                    
        elif msg["role"] == "assistant":
            doc.add_heading("Jawaban AI:", level=2)
            paragraf_list = msg["content"].split('\n')
            for p_text in paragraf_list:
                if not p_text.strip():
                    continue
                
                match_heading = re.match(r'^(#{1,6})\s+(.*)$', p_text.strip())
                if match_heading:
                    level_pagar = len(match_heading.group(1))
                    teks_judul = match_heading.group(2).replace('**', '')
                    doc.add_heading(teks_judul, level=min(level_pagar, 3))
                    continue
                
                p = doc.add_paragraph()
                parts = re.split(r'(\*\*.*?\*\*)', p_text)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        p.add_run(part.replace('**', '')).bold = True
                    else:
                        p.add_run(part)
                        
            p_line = doc.add_paragraph()
            p_line.add_run("_" * 40).italic = True
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 3. PANEL CONTROL SIDEBAR ---
with st.sidebar:
    st.title("🔮 Kontrol AI")
    st.info("⚡ Model: Qwen 3.5 122B (Multimodal + Dokumen)")
    
    st.divider()
    st.markdown("### 📥 Ekspor Dokumen")
    if "messages" in st.session_state and len(st.session_state.messages) > 1:
        file_word = buat_file_word(st.session_state.messages)
        st.download_button(
            label="📥 Download Jadi Word (.docx)",
            data=file_word,
            file_name="Analisis_Qwen_Workspace.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    else:
        st.info("Mulai obrolan untuk mengunduh dokumen.")
            
    st.divider()
    if st.button("🗑️ Reset & Bersihkan Memori"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# --- 4. PEMASANGAN API KEY & KONFIGURASI MODEL ---
BASE_URL = "https://integrate.api.nvidia.com/v1"
nvidia_api_key = "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"
MODEL_NAME = "qwen/qwen3.5-122b-a10b"

client = OpenAI(base_url=BASE_URL, api_key=nvidia_api_key)

# --- 5. MANAJEMEN MEMORI CHAT ---
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": "Anda adalah Qwen 3.5, asisten analitik tingkat tinggi. Anda dapat menganalisis gambar dan mengekstrak konteks dari file dokumen (PDF/TXT) yang dikirimkan pengguna. Berikan jawaban yang cerdas, mendalam, dan terstruktur dalam Bahasa Indonesia."}
    ]

# --- 6. TAMPILAN UTAMA INTERFASE CHAT ---
st.title("🔮 Lagos AI 8.3 (Vision + Document)")
st.caption("Bisa Apa Saja AKU?. Tentukan Sendiri")

# Tampilkan riwayat chat
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

# --- 7. AREA INPUT MULTIMODAL & DOKUMEN (GAYA GEMINI) ---
input_container = st.container()
perintah_final = ""

with input_container:
    # Baris pertama: Slot Unggah Gambar dan Slot Unggah Dokumen PDF/TXT
    col_img, col_doc = st.columns(2)
    
    with col_img:
        uploaded_image = st.file_uploader(
            "Unggah Gambar (JPG/PNG)", 
            type=["jpg", "jpeg", "png"],
            label_visibility="collapsed",
            key="pengunggah_gambar"
        )
    
    with col_doc:
        uploaded_file = st.file_uploader(
            "Unggah Dokumen (PDF/TXT)",
            type=["pdf", "txt"],
            label_visibility="collapsed",
            key="pengunggah_dokumen"
        )

    # Preview Interaktif untuk Gambar
    if uploaded_image:
        col_img_preview, _ = st.columns([1, 4])
        with col_img_preview:
            st.image(uploaded_image, caption="📷 Gambar Terlampir", width=120)
            if st.button("❌ Hapus Gambar", use_container_width=True, key="del_img"):
                st.cache_data.clear()
                st.rerun()

    # Preview Interaktif untuk File Dokumen
    if uploaded_file:
        col_doc_preview, _ = st.columns([1, 4])
        with col_doc_preview:
            st.info(f"📄 {uploaded_file.name}")
            if st.button("❌ Hapus Dokumen", use_container_width=True, key="del_doc"):
                st.cache_data.clear()
                st.rerun()

    # Kotak Teks Input Utama
    user_input = st.chat_input("Ketik perintah teks atau tanyakan sesuatu tentang file Anda di sini...")

# --- 8. PROSES EVALUASI RESPONS MULTIMODAL ---
if user_input:
    teks_perintah_bersih = str(user_input).strip()
    
    # 1. Jika ada dokumen PDF/TXT, ekstrak teks isinya dan rekayasa ke dalam perintah
    if uploaded_file:
        with st.spinner("📄 Membaca berkas dokumen..."):
            isi_teks_dokumen = ekstrak_teks_dari_dokumen(uploaded_file)
        
        if isi_teks_dokumen:
            perintah_final = (
                f"Berikut adalah isi teks dari dokumen yang diunggah ({uploaded_file.name}):\n"
                f"\"\"\"\n{isi_teks_dokumen}\n\"\"\"\n\n"
                f"Pertanyaan/Instruksi User: {teks_perintah_bersih}"
            )
        else:
            perintah_final = teks_perintah_bersih
    else:
        perintah_final = teks_perintah_bersih

    # 2. Format struktur payload multimodal (jika ada gambar)
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
    
    # Tampilkan teks asli ketikan user di layar agar bersih
    with st.chat_message("user"):
        st.markdown(teks_perintah_bersih)
    
    # Masukkan ke riwayat memori chat utama
    st.session_state.messages.append({"role": "user", "content": konten_payload})
    
    # Ambil respons dari Qwen 3.5
    with st.chat_message("assistant"):
        try:
            placeholder_teks = st.empty()
            full_response = ""

            response_stream = client.chat.completions.create(
                model=MODEL_NAME,
                messages=st.session_state.messages,
                temperature=0.6,
                max_tokens=4096,
                stream=True
            )
            
            for chunk in response_stream:
                if chunk.choices and len(chunk.choices) > 0:
                    delta_content = chunk.choices[0].delta.content
                    if delta_content:
                        full_response += delta_content
                        placeholder_teks.markdown(full_response + "▌")
            
            placeholder_teks.markdown(full_response)
            st.session_state.messages.append({"role": "assistant", "content": full_response})
            
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis saat memproses permintaan. Detail: {e}")
