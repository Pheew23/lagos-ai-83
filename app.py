/* --- STYLING UNTUK RIWAYAT OBROLAN SIDEBAR --- */
        /* Membuat list riwayat terlihat lebih rapi dan ringkas */
        [data-testid="stSidebar"] .stButton > button {
            border-radius: 8px !important;
            padding: 0.5rem 1rem !important;
            text-align: left !important;
            justify-content: flex-start !important;
        }
        
        /* Modifikasi khusus untuk tombol delete (kolom ke-2) agar transparan dan presisi */
        [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button {
            background-color: transparent !important;
            border: 1px solid transparent !important;
            justify-content: center !important;
            padding: 0.5rem 0 !important;
            color: #888888 !important;
            transition: all 0.2s ease;
        }
        
        /* Efek hover untuk tombol delete */
        [data-testid="stSidebar"] [data-testid="column"]:nth-of-type(2) .stButton > button:hover {
            border: 1px solid #ff4b4b !important;
            color: #ff4b4b !important;
            background-color: rgba(255, 75, 75, 0.1) !important;
        }
        
        /* Memastikan teks pada tombol riwayat terpotong rapi dengan ellipsis (...) jika terlalu panjang */
        [data-testid="stSidebar"] .stButton > button p {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            width: 100%;
        }
