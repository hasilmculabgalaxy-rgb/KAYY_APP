"""
KAY App - FINAL SINGLE PAGE APP (SPA) - VERSI DIPERBAHARUI & FIX SYNTAX ERROR
- Memperbaiki SyntaxError: 'return' outside function.
- Menambahkan FITUR BARU: Batch Rename PDF dan Gambar berdasarkan Excel.
"""

import os
import io
import zipfile
import shutil
import traceback
import tempfile
import time

import streamlit as st
import pandas as pd
from PIL import Image

# PDF libs
PdfReader = PdfWriter = None
try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception:
    pass

pdfplumber = None
try:
    import pdfplumber
except Exception:
    pass

Document = None
try:
    from docx import Document
except Exception:
    pass

# pdf2image: try both convert_from_bytes and convert_from_path availability
PDF2IMAGE_AVAILABLE = False
convert_from_bytes = convert_from_path = None
try:
    from pdf2image import convert_from_bytes, convert_from_path 
    PDF2IMAGE_AVAILABLE = True
except Exception:
    try:
        from pdf2image import convert_from_path
        PDF2IMAGE_AVAILABLE = True
        convert_from_bytes = None
    except Exception:
        pass # convert_from_path and convert_from_bytes remain None

# ----------------- Helpers -----------------
def make_zip_from_map(bytes_map: dict) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        for name, data in bytes_map.items():
            z.writestr(name, data)
    b.seek(0)
    return b.getvalue()

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

def try_encrypt(writer, password: str):
    """Fungsi untuk enkripsi PDF, menampung try/except"""
    try:
        writer.encrypt(password)
    except TypeError:
        try:
            writer.encrypt(user_pwd=password, owner_pwd=None)
        except Exception:
            writer.encrypt(user_pwd=password, owner_pwd=password)

def rotate_page_safe(page, angle):
    """
    Fungsi untuk rotasi halaman PDF. 
    (Ini adalah fungsi yang menampung 'return' pada baris 374/379 di script lama)
    """
    try:
        page.rotate(angle)
    except Exception:
        try:
            from PyPDF2.generic import NameObject, NumberObject
            page.__setitem__(NameObject("/Rotate"), NumberObject(angle))
        except Exception:
            pass
            
def navigate_to(target_menu):
    """Helper global untuk navigasi antar halaman/menu."""
    st.session_state.menu_selection = target_menu
    try:
        st.rerun() 
    except AttributeError:
        st.experimental_rerun()

# ----------------- Streamlit config & CSS (Perapihan Ikon) -----------------
LOGO_PATH = os.path.join("assets", "logo.png")
# Menggunakan ikon palet untuk memastikan ikon terlihat di semua sistem
page_icon = LOGO_PATH if os.path.exists(LOGO_PATH) else "🛠️" 
st.set_page_config(page_title="KAY App – Tools MCU", page_icon=page_icon, layout="wide", initial_sidebar_state="collapsed")

# CSS / Theme
st.markdown("""
<style>
/* 1. HILANGKAN SEMUA UI SIDEBAR */
[data-testid="stSidebarToggleButton"], 
section[data-testid="stSidebar"],      
[data-testid="stDecoration"]            
{
    visibility: hidden !important;
    display: none !important;
    width: 0 !important;
    padding: 0 !important;
}

/* 2. Hilangkan logo GitHub 'Fork' */
.stApp a[href*="github.com/"],
.stApp header > div:last-child {
    display: none !important;
}

/* 3. Background gradient lembut */
.stApp {
    background: linear-gradient(180deg, #e9f2ff 0%, #f4f9ff 100%); 
    color: #002b5b;
    font-family: 'Inter', sans-serif;
}

/* 4. Tombol modern glossy */
div.stButton > button {
    background: linear-gradient(90deg, #5dade2, #3498db);
    color: white; 
    border: none;
    border-radius: 12px;
    padding: 0.5rem 1rem;
    font-weight: 600;
    transition: 0.2s;
    box-shadow: 0 4px 8px rgba(52, 152, 219, 0.4); 
    cursor: pointer;
    width: auto;
}

/* Tombol Dashboard & Kembali dibuat lebar penuh di konteks masing-masing */
div.stButton > button[key*="dash_"], 
div.stButton > button[key*="back_"] {
    width: 100%;
    margin-top: 10px; 
}

div.stButton > button:hover {
    background: linear-gradient(90deg, #3498db, #2e86c1); 
    transform: scale(1.01);
    box-shadow: 0 6px 14px rgba(52, 152, 219, 0.5); 
}

/* 5. Card fitur */
.feature-card {
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06); 
    padding: 18px; 
    transition: 0.2s; 
    border: 1px solid #d0e3ff; 
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    height: 100%; 
}
.feature-card:hover {
    box-shadow: 0 6px 18px rgba(0,0,0,0.10);
    transform: translateY(-2px);
}

/* 6. Small UI tweaks */
h1 { color: #1b4f72; font-weight: 800;
}
.stInfo, .stWarning {
    border-radius: 8px;
    padding: 1rem;
}
.stInfo { background-color: #e3f2fd; border-left: 5px solid #2196f3; }
.stWarning { background-color: #fff3e0; border-left: 5px solid #ff9800;
}

</style>
""", unsafe_allow_html=True)

# ----------------- Global Header & Navigation Setup -----------------

# Inisialisasi session state
if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "Dashboard"
    
menu = st.session_state.menu_selection

# ----------------- Fungsi Tombol Kembali (Perapihan Ikon) -----------------
def add_back_to_dashboard_button():
    """Menambahkan tombol 'Kembali ke Dashboard' di halaman fitur dengan ikon 🏠."""
    if st.button("🏠 Kembali ke Dashboard", key="back_to_dash"):
        navigate_to("Dashboard")
    st.markdown("---")

# ----------------- Halaman Header (Selalu ditampilkan) -----------------
if os.path.exists(LOGO_PATH):
    try:
        st.image(LOGO_PATH, width=110)
    except Exception:
        st.write("KAY App")
st.title("KAY App – Tools MCU")
st.markdown("Aplikasi serbaguna untuk pengolahan dokumen, PDF, gambar, dan data MCU — UI elegan + fungsi lengkap.")
st.markdown("---")


# -----------------------------------------------------------------------------
# ----------------- FUNGSI UTAMA (Diperbarui untuk fitur baru) -----------------
# -----------------------------------------------------------------------------

# -------------- Dashboard (Menu Utama) --------------
if menu == "Dashboard":
    st.markdown("### Pilih Fitur Utama")
    
    # ------------------ FITUR UTAMA ------------------
    cols1 = st.columns(3)

    # Kompres Foto
    with cols1[0]:
        # Tambahkan notifikasi "Baru" untuk rename
        st.markdown('<div class="feature-card"><b>Kompres Foto / Gambar Tools</b><br>Perkecil ukuran, ubah format, **Batch Rename Sesuai Excel (Baru)**.</div>', unsafe_allow_html=True)
        if st.button("Buka Kompres Foto", key="dash_foto"):
            navigate_to("Kompres Foto")

    # PDF Tools
    with cols1[1]:
        # Tambahkan notifikasi "Baru" untuk rename
        st.markdown('<div class="feature-card"><b>PDF Tools</b><br>Gabung, pisah, encrypt, Reorder & **Batch Rename Sesuai Excel (Baru)**.</div>', unsafe_allow_html=True)
        if st.button("Buka PDF Tools", key="dash_pdf"):
            navigate_to("PDF Tools")

    # MCU Tools
    with cols1[2]:
        st.markdown('<div class="feature-card"><b>MCU Tools</b><br>Proses Excel + PDF untuk hasil MCU / Analisis Data.</div>', unsafe_allow_html=True)
        if st.button("Buka MCU Tools", key="dash_mcu"):
            navigate_to("MCU Tools")
            
    # ------------------ FITUR LAINNYA ------------------
    st.markdown("### Fitur Lainnya")
    cols2 = st.columns(3)
    
    # File Tools
    with cols2[0]:
        # Tambahkan notifikasi "Baru" untuk rename
        st.markdown('<div class="feature-card"><b>File Tools</b><br>Zip/unzip, konversi dasar, Batch Rename Gambar & PDF.</div>', unsafe_allow_html=True)
        if st.button("Buka File Tools", key="dash_file"):
            navigate_to("File Tools")

    # Tentang
    with cols2[1]:
        st.markdown('<div class="feature-card"><b>Tentang Aplikasi</b><br>Informasi dan kebutuhan library.</div>', unsafe_allow_html=True)
        if st.button("Lihat Tentang", key="dash_about"):
            navigate_to("Tentang")

    # Kolom kosong (dibuat 3 kolom agar layout tetap rapi)
    with cols2[2]:
        st.markdown('<div class="feature-card" style="visibility:hidden; height: 100%;">.</div>', unsafe_allow_html=True)
        

    st.markdown("---")
    st.info("Semua proses berlangsung lokal di perangkat server tempat Streamlit dijalankan.")

# -------------- Kompres Foto / Image Tools --------------
if menu == "Kompres Foto":
    add_back_to_dashboard_button() 
    st.subheader("Kompres & Kelola Foto/Gambar")
    
    # Sub-menu untuk gambar
    img_tool = st.selectbox("Pilih Fitur Gambar", [
        "Kompres Foto (Batch)", 
        "🔢 Batch Rename/Format Gambar (Sequential)", 
        "📄 Batch Rename Gambar Sesuai Excel (Fitur Baru)" # FITUR BARU: Batch Rename Gambar by Excel
        ])

    if img_tool == "Kompres Foto (Batch)":
        st.markdown("---")
        uploaded = st.file_uploader("Unggah gambar (jpg/png) — bisa banyak", type=["jpg","jpeg","png"], accept_multiple_files=True)
        quality = st.slider("Kualitas JPEG", 10, 95, 75)
        max_side = st.number_input("Max side (px)", min_value=100, max_value=4000, value=1200)
        if uploaded and st.button("Kompres Semua"):
            out_map = {}
            total = len(uploaded)
            prog = st.progress(0)
            with st.spinner("Mengompres..."):
                for i, f in enumerate(uploaded):
                    try:
                        raw = f.read()
                        im = Image.open(io.BytesIO(raw))
                        im.thumbnail((max_side, max_side))
                        buf = io.BytesIO()
                        im.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
                        out_map[f"compressed_{f.name}"] = buf.getvalue()
                    except Exception as e:
                        st.warning(f"Gagal: {f.name} — {e}")
                    prog.progress(int((i+1)/total*100))
            if out_map:
                zipb = make_zip_from_map(out_map)
                st.success(f"{len(out_map)} file berhasil dikompres")
                st.download_button("Unduh Hasil (ZIP)", zipb, file_name="foto_kompres.zip", mime="application/zip")
            else:
                st.warning("Tidak ada file berhasil dikompres.")

    # --- FITUR Batch Rename Gambar (Sequential) ---
    elif img_tool == "🔢 Batch Rename/Format Gambar (Sequential)":
        st.markdown("---")
        st.subheader("🔢 Ganti Nama & Ubah Format Gambar Massal (Sequential)")
        uploaded_files = st.file_uploader(
            "Unggah file Gambar (JPG, PNG, dll.):", 
            type=["jpg", "jpeg", "png", "webp"], 
            accept_multiple_files=True,
            key="batch_rename_uploader"
        )
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="KAY_File", help="Contoh: KAY_File_001.jpg", key="prefix_img_seq")
            new_format = col2.selectbox("Format Output Baru:", ["Sama seperti Asli", "JPG", "PNG", "WEBP"], index=0, key="format_img_seq")

            if st.button("Proses Batch File", key="process_batch_rename_seq"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, 1):
                                _, original_ext = os.path.splitext(file.name)
                                img = Image.open(file)
                                img_io = io.BytesIO()
                                if new_format == "Sama seperti Asli":
                                    output_format_pil = img.format if img.format else 'JPEG'
                                    output_ext = original_ext
                                else:
                                    output_ext = "." + new_format.lower()
                                    output_format_pil = new_format.upper()
                                new_filename = f"{new_prefix}_{i:03d}{output_ext}"
                                if output_format_pil in ('JPEG', 'JPG'):
                                    img.convert("RGB").save(img_io, format='JPEG', quality=95) 
                                elif output_format_pil == 'PNG':
                                    img.save(img_io, format='PNG')
                                elif output_format_pil == 'WEBP':
                                    img.save(img_io, format='WEBP')
                                else:
                                    img.save(img_io, format=output_format_pil) 
                                img_io.seek(0)
                                zf.writestr(new_filename, img_io.read())
                        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Batch", data=output_zip.getvalue(), file_name="hasil_batch_gambar.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

    # --- FITUR BARU 1: Batch Rename Gambar Sesuai Excel ---
    elif img_tool == "📄 Batch Rename Gambar Sesuai Excel (Fitur Baru)":
        st.markdown("---")
        st.subheader("📄 Ganti Nama Gambar (PNG/JPEG) Berdasarkan Excel")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (termasuk ekstensi, misal: `foto_123.jpg`) dan **`nama_baru`** (termasuk ekstensi, misal: `ID_001.png`).")
        
        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_img_excel_up")
        files = st.file_uploader("Unggah Gambar (JPG/PNG/JPEG, multiple):", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="rename_img_files_up")
        
        if excel_up and files and st.button("Proses Ganti Nama Gambar (ZIP)", key="process_img_rename_excel"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    # 1. Baca Excel
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    # 2. Validasi Kolom
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        # 3. Map File dan Proses Rename
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []
                        df['nama_lama_str'] = df['nama_lama'].astype(str).str.strip() # Gunakan str.strip() untuk membersihkan spasi

                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            # Cek di file yang diupload (case-sensitive)
                            if old_name in file_map:
                                # Pastikan nama baru memiliki ekstensi
                                if not os.path.splitext(new_name)[1]:
                                    # Coba ambil ekstensi dari nama lama jika nama baru tidak ada
                                    _, old_ext = os.path.splitext(old_name)
                                    new_name = new_name + old_ext
                                    
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        # 4. Buat ZIP
                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ {len(out_map)} file berhasil diganti namanya dan dikemas.")
                            st.download_button("Unduh Hasil (ZIP)", zipb, file_name="gambar_renamed_by_excel.zip", mime="application/zip")
                        else:
                            st.warning("Tidak ada file yang cocok ditemukan atau diproses.")
                        
                        if not_found:
                            st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan di file yang diunggah. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()

# -------------- PDF Tools (Diperbarui Menu dan Fitur) --------------
if menu == "PDF Tools":
    add_back_to_dashboard_button() 
    st.subheader("PDF Tools")

    # Menu yang lebih terstruktur dan ditambahkan fitur baru
    pdf_options = [
        "--- Pilih Tools ---",
        "📂 Gabung PDF",
        "✂️ Pisah PDF", 
        "🔀 Reorder/Hapus Halaman PDF", 
        "📑 Batch Rename PDF (Sequential)", 
        "📄 Batch Rename PDF Sesuai Excel (Fitur Baru)", # FITUR BARU: Batch Rename PDF by Excel
        "➡️ Image -> PDF",
        "⬅️ PDF -> Image", 
        "📝 Ekstraksi Teks/Tabel",
        "🔄 Konversi PDF",
        "🔒 Proteksi PDF",
        "🛠️ Utility PDF",
    ]
    
    tool_select = st.selectbox("Pilih fitur PDF", pdf_options)

    # Mapping
    if tool_select == "--- Pilih Tools ---": tool = None
    elif tool_select == "📝 Ekstraksi Teks/Tabel": tool = st.selectbox("Pilih mode ekstraksi", ["Extract Text", "Extract Tables -> Excel"])
    elif tool_select == "🔄 Konversi PDF": tool = st.selectbox("Pilih mode konversi", ["PDF -> Word", "PDF -> Excel (text)"])
    elif tool_select == "🔒 Proteksi PDF": tool = st.selectbox("Pilih mode proteksi", ["Encrypt PDF", "Decrypt PDF", "Batch Lock (Excel)"])
    elif tool_select == "🛠️ Utility PDF": tool = st.selectbox("Pilih mode utilitas", ["Hapus Halaman", "Rotate PDF", "Kompres PDF", "Watermark PDF", "Preview PDF"])
    elif tool_select == "📂 Gabung PDF": tool = "Gabung PDF"
    elif tool_select == "✂️ Pisah PDF": tool = "Pisah PDF"
    elif tool_select == "🔀 Reorder/Hapus Halaman PDF": tool = "Reorder PDF" 
    elif tool_select == "📑 Batch Rename PDF (Sequential)": tool = "Batch Rename PDF Seq" 
    elif tool_select == "📄 Batch Rename PDF Sesuai Excel (Fitur Baru)": tool = "Batch Rename PDF Excel" # FITUR BARU: Batch Rename PDF by Excel
    elif tool_select == "⬅️ PDF -> Image": tool = "PDF -> Image"
    elif tool_select == "➡️ Image -> PDF": tool = "Image -> PDF"
    else: tool = None
    

    # --- FITUR BARU 2: Batch Rename PDF Sesuai Excel ---
    if tool == "Batch Rename PDF Excel":
        st.markdown("---")
        st.subheader("📄 Ganti Nama File PDF Berdasarkan Excel")
        st.markdown("Unggah banyak file PDF dan ganti namanya sesuai daftar di Excel/CSV.")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (misal: `ID_123.pdf`) dan **`nama_baru`** (misal: `Hasil_123.pdf`).")

        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_pdf_excel_up")
        files = st.file_uploader("Unggah File PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="rename_pdf_files_up")
        
        if excel_up and files and st.button("Proses Ganti Nama PDF (ZIP)", key="process_pdf_rename_excel"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    # 1. Baca Excel
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    # 2. Validasi Kolom
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        # 3. Map File dan Proses Rename
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []
                        
                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            if old_name in file_map:
                                # Tambahkan ekstensi .pdf jika belum ada di nama baru
                                if not new_name.lower().endswith('.pdf'):
                                    new_name += '.pdf'
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        # 4. Buat ZIP
                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ {len(out_map)} file berhasil diganti namanya dan dikemas.")
                            st.download_button("Unduh Hasil (ZIP)", zipb, file_name="pdf_renamed_by_excel.zip", mime="application/zip")
                        else:
                            st.warning("Tidak ada file yang cocok ditemukan atau diproses.")
                        
                        if not_found:
                            st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan di file yang diunggah. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()

    # --- FITUR Batch Rename PDF (Sequential) ---
    if tool == "Batch Rename PDF Seq":
        st.markdown("---")
        st.subheader("📑 Ganti Nama File PDF Massal (Sequential)")
        uploaded_files = st.file_uploader("Unggah file PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="batch_rename_pdf_uploader_seq")
        
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="Hasil_PDF", help="Contoh: Hasil_PDF_001.pdf", key="prefix_pdf_seq")
            start_num = col2.number_input("Mulai dari Angka (Counter):", min_value=1, value=1, step=1, key="start_num_pdf_seq")

            if st.button("Proses Ganti Nama (ZIP)", key="process_batch_rename_pdf_seq"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, start_num):
                                new_filename = f"{new_prefix}_{i:03d}.pdf"
                                zf.writestr(new_filename, file.read())
                        st.success(f"✅ Berhasil mengganti nama {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Rename", data=output_zip.getvalue(), file_name="pdf_renamed.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

    # --- LOGIKA FITUR PDF LAINNYA (tidak diubah) ---
    if tool == "Reorder PDF":
        st.markdown("---")
        st.subheader("🔀 Reorder atau Hapus Halaman PDF")
        st.markdown("Unggah file PDF Anda dan tentukan urutan halaman baru (contoh: `2, 1, 3` untuk membalik, atau `1, 3` untuk menghapus halaman 2).")

        f = st.file_uploader("Unggah 1 file PDF:", type="pdf", key="reorder_pdf_uploader")
        
        if f:
            try:
                raw = f.read()
                reader = PdfReader(io.BytesIO(raw))
                num_pages = len(reader.pages)
                st.info(f"PDF berhasil dimuat. Jumlah total halaman: **{num_pages}**.")
                
                default_order = ", ".join(map(str, range(1, num_pages + 1)))

                new_order_str = st.text_input(
                    f"Masukkan urutan halaman baru (1-{num_pages}) dipisahkan koma:",
                    value=default_order,
                    help="Contoh: '3, 1, 2' untuk mengubah urutan. '1, 3, 5' untuk menghapus halaman genap."
                )

                if st.button("Proses Reorder/Hapus Halaman", key="process_reorder"):
                    new_order_indices = []
                    
                    try:
                        input_list = [int(x.strip()) for x in new_order_str.split(',') if x.strip().isdigit()]
                        
                        if any(n < 1 or n > num_pages for n in input_list):
                            st.error(f"Nomor halaman harus antara 1 sampai {num_pages}.")
                            raise ValueError("Invalid page number in input.")

                        new_order_indices = [n - 1 for n in input_list]
                        
                        writer = PdfWriter()
                        for index in new_order_indices:
                            writer.add_page(reader.pages[index])
                        
                        pdf_buffer = io.BytesIO()
                        writer.write(pdf_buffer)
                        pdf_buffer.seek(0)

                        st.download_button(
                            "✅ Unduh Hasil PDF (Reordered)",
                            data=pdf_buffer,
                            file_name="pdf_reordered.pdf",
                            mime="application/pdf"
                        )
                        st.success(f"Pemrosesan selesai. Total halaman baru: {len(new_order_indices)}.")

                    except ValueError:
                        pass
                    except Exception as e:
                        st.error(f"Format urutan halaman tidak valid atau terjadi kesalahan pemrosesan: {e}")

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses PDF: {e}")
                st.info("Pastikan file yang diunggah adalah PDF yang valid.")

    if tool == "Gabung PDF":
        st.markdown("---")
        files = st.file_uploader("Upload PDFs (multiple):", type="pdf", accept_multiple_files=True)
        if files and st.button("Gabung"):
            try:
                with st.spinner("Menggabungkan..."):
                    writer = PdfWriter()
                    for f in files:
                        reader = PdfReader(io.BytesIO(f.read()))
                        for p in reader.pages:
                            writer.add_page(p)
                    out = io.BytesIO(); writer.write(out); out.seek(0)
                st.download_button("Download merged.pdf", out.getvalue(), file_name="merged.pdf", mime="application/pdf")
                st.success("Selesai")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Pisah PDF":
        st.markdown("---")
        f = st.file_uploader("Upload single PDF:", type="pdf")
        if f and st.button("Split to pages (ZIP)"):
            try:
                with st.spinner("Memisahkan..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    out_map = {}
                    for i, p in enumerate(reader.pages):
                        w = PdfWriter(); w.add_page(p)
                        buf = io.BytesIO(); w.write(buf); buf.seek(0)
                        out_map[f"page_{i+1}.pdf"] = buf.getvalue()
                    zipb = make_zip_from_map(out_map)
                st.download_button("Download pages.zip", zipb, file_name="pages.zip", mime="application/zip")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Hapus Halaman":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        page_no = st.number_input("Halaman yang dihapus (1-based)", min_value=1, value=1)
        if f and st.button("Hapus Halaman"):
            try:
                with st.spinner("Menghapus..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    writer = PdfWriter()
                    for i, p in enumerate(reader.pages):
                        if i+1 != page_no:
                            writer.add_page(p)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download result", buf.getvalue(), file_name="removed_page.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Rotate PDF":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        angle = st.selectbox("Rotate degrees", [90, 180, 270])
        if f and st.button("Rotate"):
            try:
                with st.spinner("Memutar..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    writer = PdfWriter()
                    for p in reader.pages:
                        rotate_page_safe(p, angle)
                        writer.add_page(p)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download rotated.pdf", buf.getvalue(), file_name="rotated.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Kompres PDF":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        if f and st.button("Compress (rewrite)"):
            try:
                with st.spinner("Mengompres (rewrite)..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    writer = PdfWriter()
                    for p in reader.pages:
                        writer.add_page(p)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download compressed.pdf", buf.getvalue(), file_name="compressed.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Watermark PDF":
        st.markdown("---")
        base = st.file_uploader("Base PDF", type="pdf")
        watermark = st.file_uploader("Watermark PDF (single page)", type="pdf")
        if base and watermark and st.button("Apply watermark"):
            try:
                with st.spinner("Menerapkan watermark..."):
                    rb = PdfReader(io.BytesIO(base.read()))
                    rm = PdfReader(io.BytesIO(watermark.read()))
                    wm = rm.pages[0]
                    writer = PdfWriter()
                    for p in rb.pages:
                        try:
                            p.merge_page(wm)
                        except Exception:
                            try:
                                p.mergeTranslatedPage(wm, 0, 0)
                            except Exception:
                                pass
                        writer.add_page(p)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download watermarked.pdf", buf.getvalue(), file_name="watermarked.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "PDF -> Image":
        st.markdown("---")
        st.info("Requires pdf2image + poppler (server).")
        f = st.file_uploader("Upload PDF", type="pdf")
        dpi = st.slider("DPI", 100, 300, 150)
        fmt = st.radio("Format", ["PNG", "JPEG"])
        if f and st.button("Convert to images"):
            try:
                if not PDF2IMAGE_AVAILABLE:
                    st.error("pdf2image not installed or poppler missing.")
                else:
                    with st.spinner("Converting..."):
                        pdf_bytes = f.read()
                        images = None
                        if convert_from_bytes is not None:
                            images = convert_from_bytes(pdf_bytes, dpi=dpi)
                        else:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                tmp.write(pdf_bytes)
                                tmp_path = tmp.name
                            images = convert_from_path(tmp_path, dpi=dpi)
                            try:
                                os.unlink(tmp_path)
                            except Exception:
                                pass
                        out_map = {}
                        for i, img in enumerate(images):
                            b = io.BytesIO(); img.save(b, format=fmt); out_map[f"page_{i+1}.{fmt.lower()}"] = b.getvalue()
                        zipb = make_zip_from_map(out_map)
                    st.download_button("Download images.zip", zipb, file_name="pdf_images.zip", mime="application/zip")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Image -> PDF":
        st.markdown("---")
        imgs = st.file_uploader("Upload images", type=["jpg","png","jpeg"], accept_multiple_files=True)
        if imgs and st.button("Images -> PDF"):
            try:
                with st.spinner("Membuat PDF dari gambar..."):
                    pil = [Image.open(io.BytesIO(i.read())).convert("RGB") for i in imgs]
                    buf = io.BytesIO()
                    if len(pil) == 1:
                        pil[0].save(buf, format="PDF")
                    else:
                        pil[0].save(buf, save_all=True, append_images=pil[1:], format="PDF")
                    buf.seek(0)
                st.download_button("Download images_as_pdf.pdf", buf.getvalue(), file_name="images_as_pdf.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Extract Text":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        if f and st.button("Extract text"):
            try:
                with st.spinner("Mengekstrak teks..."):
                    text_blocks = []
                    raw = f.read()
                    if pdfplumber:
                        with pdfplumber.open(io.BytesIO(raw)) as doc:
                            for i, p in enumerate(doc.pages):
                                text_blocks.append(f"--- Page {i+1} ---\n" + (p.extract_text() or ""))
                    else:
                        reader = PdfReader(io.BytesIO(raw))
                        for i, p in enumerate(reader.pages):
                            text_blocks.append(f"--- Page {i+1} ---\n" + (p.extract_text() or ""))
                    full = "\n".join(text_blocks)
                st.text_area("Extracted text (preview)", full[:10000], height=300)
                st.download_button("Download .txt", full, file_name="extracted_text.txt", mime="text/plain")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Extract Tables -> Excel":
        st.markdown("---")
        if pdfplumber is None:
            st.error("pdfplumber is required for table extraction (pip install pdfplumber)")
        else:
            f = st.file_uploader("Upload PDF", type="pdf")
            if f and st.button("Extract tables"):
                try:
                    with st.spinner("Mengekstrak tabel..."):
                        tables = []
                        with pdfplumber.open(io.BytesIO(f.read())) as doc:
                            for page in doc.pages:
                                for tbl in page.extract_tables():
                                    if tbl and len(tbl) > 1:
                                        # Menghilangkan baris header yang mungkin diduplikasi
                                        df = pd.DataFrame(tbl[1:], columns=tbl[0]) 
                                        tables.append(df)
                        if tables:
                            df_all = pd.concat(tables, ignore_index=True)
                            st.dataframe(df_all.head())
                            excel_bytes = df_to_excel_bytes(df_all)
                            st.download_button("Download Excel", data=excel_bytes, file_name="extracted_tables.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            st.info("No tables found.")
                except Exception:
                    st.error(traceback.format_exc())

    if tool == "PDF -> Word":
        st.markdown("---")
        if Document is None:
            st.error("python-docx is required for PDF->Word (pip install python-docx)")
        else:
            f = st.file_uploader("Upload PDF", type="pdf")
            if f and st.button("Convert to Word"):
                try:
                    with st.spinner("Converting..."):
                        reader = PdfReader(io.BytesIO(f.read()))
                        doc = Document()
                        for p in reader.pages:
                            txt = p.extract_text() or ""
                            doc.add_paragraph(txt)
                        out = io.BytesIO(); doc.save(out); out.seek(0)
                    st.download_button("Download .docx", out.getvalue(), file_name="converted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception:
                    st.error(traceback.format_exc())

    if tool == "PDF -> Excel (text)":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        if f and st.button("Convert to Excel (text)"):
            try:
                with st.spinner("Converting..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    rows = []
                    for i, p in enumerate(reader.pages):
                        rows.append({"page": i+1, "text": p.extract_text() or ""})
                    df = pd.DataFrame(rows)
                    excel_bytes = df_to_excel_bytes(df)
                    st.download_button("Download Excel", excel_bytes, file_name="pdf_text.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Encrypt PDF":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        pw = st.text_input("Password", type="password")
        if f and pw and st.button("Encrypt"):
            try:
                with st.spinner("Mengunci PDF..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    writer = PdfWriter()
                    for p in reader.pages:
                        writer.add_page(p)
                    try_encrypt(writer, pw)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download encrypted.pdf", buf.getvalue(), file_name="encrypted.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Decrypt PDF":
        st.markdown("---")
        f = st.file_uploader("Upload encrypted PDF", type="pdf")
        pw = st.text_input("Password for decryption", type="password")
        if f and pw and st.button("Decrypt"):
            try:
                with st.spinner("Membuka PDF..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    if getattr(reader, "is_encrypted", False):
                        reader.decrypt(pw)
                    writer = PdfWriter()
                    for p in reader.pages:
                        writer.add_page(p)
                    buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                st.download_button("Download decrypted.pdf", buf.getvalue(), file_name="decrypted.pdf", mime="application/pdf")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Batch Lock (Excel)":
        st.markdown("---")
        excel_file = st.file_uploader("Upload Excel (filename,password) or CSV", type=["xlsx","csv"])
        pdfs = st.file_uploader("Upload PDFs (multiple)", type="pdf", accept_multiple_files=True)
        if excel_file and pdfs and st.button("Batch Lock"):
            try:
                with st.spinner("Batch locking PDFs..."):
                    if excel_file.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_file.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_file.read()))
                    pdf_map = {p.name: p.read() for p in pdfs}
                    out_map = {}
                    not_found = []
                    total = len(df)
                    prog = st.progress(0)
                    for idx, (_, row) in enumerate(df.iterrows()):
                        cols = [c.lower() for c in df.columns]
                        try:
                            # Safely extract column names
                            target_col = df.columns[cols.index('filename')]
                            pwd_col = df.columns[cols.index('password')]
                            target = str(row[target_col]).strip()
                            pwd = str(row[pwd_col]).strip()
                        except Exception:
                            target = None; pwd = None 
                        
                        if target and pwd:
                            # Cek yang exact match dulu
                            matches = [k for k in pdf_map.keys() if k == target]
                            if matches:
                                key = matches[0]
                                reader = PdfReader(io.BytesIO(pdf_map[key]))
                                writer = PdfWriter()
                                for p in reader.pages: writer.add_page(p)
                                try_encrypt(writer, pwd)
                                b = io.BytesIO(); writer.write(b);
                                out_map[f"locked_{key}"] = b.getvalue()
                            else:
                                not_found.append(target)
                        prog.progress(int((idx+1)/total*100))
                if out_map:
                    st.download_button("Download locked_pdfs.zip", make_zip_from_map(out_map), file_name="locked_pdfs.zip", mime="application/zip")
                if not_found:
                    st.warning(f"{len(not_found)} files not found sample: {not_found[:10]}")
            except Exception:
                st.error(traceback.format_exc())

    if tool == "Preview PDF":
        st.markdown("---")
        f = st.file_uploader("Upload PDF", type="pdf")
        mode = st.radio("Preview mode", ["First page (fast)", "All pages (slow)"])
        if f and st.button("Show Preview"):
            try:
                with st.spinner("Preparing preview..."):
                    pdf_bytes = f.read()
                    if PDF2IMAGE_AVAILABLE:
                        if mode.startswith("First"):
                            if convert_from_bytes is not None:
                                imgs = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)
                            else:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                    tmp.write(pdf_bytes)
                                    tmp_path = tmp.name
                                imgs = convert_from_path(tmp_path, first_page=1, last_page=1)
                                try:
                                    os.unlink(tmp_path)
                                except Exception:
                                    pass
                            buf = io.BytesIO(); imgs[0].save(buf, format="PNG"); buf.seek(0)
                            st.image(buf.getvalue(), caption="Page 1")
                        else:
                            if convert_from_bytes is not None:
                                imgs = convert_from_bytes(pdf_bytes)
                            else:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                    tmp.write(pdf_bytes)
                                    tmp_path = tmp.name
                                imgs = convert_from_path(tmp_path)
                                try: os.unlink(tmp_path)
                                except: pass
                            for i, img in enumerate(imgs):
                                buf = io.BytesIO(); img.save(buf, format="PNG"); st.image(buf.getvalue(), caption=f"Page {i+1}")
                    else:
                        reader = PdfReader(io.BytesIO(pdf_bytes))
                        if mode.startswith("First"):
                            st.text(reader.pages[0].extract_text() or "[no text]")
                        else:
                            for i, p in enumerate(reader.pages):
                                st.write(f"--- Page {i+1} ---"); st.text(p.extract_text() or "[no text]")
            except Exception:
                st.error(traceback.format_exc())


# -------------- MCU Tools (Diperbarui untuk Fitur Baru) --------------
if menu == "MCU Tools":
    add_back_to_dashboard_button() 
    st.subheader("MCU Tools - Organise & Analyze Data")

    mcu_mode = st.selectbox("Pilih Mode MCU", ["Organise by Excel (Original)", "📊 Analisis Data MCU Massal (Dashboard Baru)"])
    
    # --- Analisis Data MCU Massal (Logika Asli) ---
    if mcu_mode == "📊 Analisis Data MCU Massal (Dashboard Baru)":
        st.subheader("📊 Dashboard Analisis Hasil MCU Massal")
        uploaded_file = st.file_uploader(
            "Unggah file Data MCU:",
            type=["xlsx", "csv"],
            key="mcu_data_uploader"
        )

        if uploaded_file:
            try:
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)

                st.success(f"File **{uploaded_file.name}** berhasil dimuat. Total Baris: {len(df)}")
                # Normalisasi kolom
                df.columns = df.columns.str.replace('[^A-Za-z0-9_]+', '', regex=True).str.lower()
                st.dataframe(df.head(), use_container_width=True)

                st.markdown("---")
                st.markdown("#### 📈 Hasil Analisis Agregat")
                status_cols = [col for col in df.columns if 'status' in col or 'fit' in col]
                
                if status_cols:
                    status_col = status_cols[0]
                    st.write(f"##### 1. Distribusi Status Kesehatan (Menggunakan kolom: `{status_col}`)")
                    df[status_col] = df[status_col].fillna("TIDAK DIKETAHUI") 
                    status_counts = df[status_col].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Jumlah']
                    
                    if len(status_counts) > 0:
                        st.bar_chart(status_counts.set_index('Status'), color="#4CAF50")
                    else:
                        st.info("Tidak ada data unik yang valid dalam kolom status.")
                else:
                    st.warning("Kolom yang mengandung kata 'status' atau 'fit' tidak ditemukan untuk Analisis Cepat.")
                
                st.markdown("##### 2. Data Hasil Terfilter")
                filter_cols = [col for col in df.columns if df[col].dtype == 'object' and df[col].nunique() <= 20]
                
                if filter_cols:
                    col_to_filter = st.selectbox("Pilih Kolom untuk Filter:", filter_cols)
                    filter_values = ["Semua"] + list(df[col_to_filter].dropna().unique())
                    selected_value = st.selectbox(f"Pilih Nilai {col_to_filter}:", filter_values)
                    
                    if selected_value != "Semua":
                        df_filtered = df[df[col_to_filter] == selected_value]
                    else:
                        df_filtered = df
                    
                    st.info(f"Menampilkan {len(df_filtered)} baris data.")
                    st.dataframe(df_filtered.head(10), use_container_width=True)
                else:
                    st.info("Tidak ada kolom kategorikal yang cocok untuk filter cepat.")
            except Exception as e:
                st.error(f"Gagal memuat atau memproses file: {e}")
                traceback.print_exc()

    # Organise by Excel (Original Logic)
    if mcu_mode == "Organise by Excel (Original)":
        st.subheader("MCU Tools - Organise by Excel")
        excel_up = st.file_uploader("Upload Excel (No_MCU, Nama, Departemen, JABATAN) or (filename,target_folder)", type=["xlsx","csv"], key="mcu_organize_excel")
        pdfs = st.file_uploader("Upload PDF files (multiple)", type="pdf", accept_multiple_files=True, key="mcu_organize_pdf")
        if excel_up and pdfs and st.button("Process MCU"):
            try:
                with st.spinner("Memproses MCU..."):
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                
                    pdf_map = {p.name: p.read() for p in pdfs}
                    out_map = {}
                    not_found = []
                    
                    if all(c in df.columns for c in ["No_MCU","Nama","Departemen","JABATAN"]):
                        total = len(df)
                        prog = st.progress(0)
                        for idx, r in df.iterrows():
                            no = str(r["No_MCU"]).strip()
                            dept = str(r["Departemen"]) if not pd.isna(r["Departemen"]) else "Unknown"
                            jab = str(r["JABATAN"]) if not pd.isna(r["JABATAN"]) else "Unknown"
                            matches = [k for k in pdf_map.keys() if k.startswith(no)]
                            if matches:
                                out_map[f"{dept}/{jab}/{matches[0]}"] = pdf_map[matches[0]]
                            else:
                                not_found.append(no)
                            prog.progress(int((idx+1)/total*100))
                    elif "filename" in df.columns and "target_folder" in df.columns:
                        for _, r in df.iterrows():
                            fn = str(r["filename"]).strip(); tgt = str(r["target_folder"]).strip()
                            if fn in pdf_map:
                                out_map[f"{tgt}/{fn}"] = pdf_map[fn]
                            else:
                                not_found.append(fn)
                if out_map:
                    st.download_button("Download MCU zip", make_zip_from_map(out_map), file_name="mcu_structured.zip", mime="application/zip")
                if not_found:
                    st.warning(f"{len(not_found)} not found sample: {not_found[:10]}")
            except Exception:
                st.error(traceback.format_exc())

# -------------- File Tools (Diperbarui untuk Fitur Baru) --------------
if menu == "File Tools":
    add_back_to_dashboard_button() 
    st.subheader("File Tools - zip / unzip / conversions / Rename")
    mode = st.selectbox("Mode", [
        "Zip files", 
        "Unzip file", 
        "Excel -> CSV", 
        "Word -> PDF (text)", 
        "🔢 Batch Rename/Format Gambar (Sequential)", 
        "📄 Batch Rename Gambar Sesuai Excel", # Batch Rename Gambar by Excel
        "📑 Batch Rename PDF (Sequential)", 
        "📄 Batch Rename PDF Sesuai Excel" # Batch Rename PDF by Excel
        ])
    
    # --- LOGIKA Batch Rename PDF Sesuai Excel (Duplikasi dari PDF Tools) ---
    if mode == "📄 Batch Rename PDF Sesuai Excel":
        st.markdown("---")
        st.subheader("📄 Ganti Nama File PDF Berdasarkan Excel")
        st.markdown("Unggah banyak file PDF dan ganti namanya sesuai daftar di Excel/CSV.")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (misal: `ID_123.pdf`) dan **`nama_baru`** (misal: `Hasil_123.pdf`).")

        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_pdf_excel_up_2")
        files = st.file_uploader("Unggah File PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="rename_pdf_files_up_2")
        
        if excel_up and files and st.button("Proses Ganti Nama PDF (ZIP)", key="process_pdf_rename_excel_2"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []
                        
                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            if old_name in file_map:
                                if not new_name.lower().endswith('.pdf'): new_name += '.pdf'
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ {len(out_map)} file berhasil diganti namanya.")
                            st.download_button("Unduh File ZIP Hasil Rename", data=zipb, file_name="pdf_renamed_2.zip", mime="application/zip")
                        else:
                            st.warning("Tidak ada file yang cocok ditemukan.")
                        
                        if not_found: st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()

    # --- LOGIKA Batch Rename Gambar Sesuai Excel (Duplikasi dari Kompres Foto) ---
    if mode == "📄 Batch Rename Gambar Sesuai Excel":
        st.markdown("---")
        st.subheader("📄 Ganti Nama Gambar (PNG/JPEG) Berdasarkan Excel")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (termasuk ekstensi) dan **`nama_baru`** (termasuk ekstensi).")
        
        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_img_excel_up_2")
        files = st.file_uploader("Unggah Gambar (JPG/PNG/JPEG, multiple):", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="rename_img_files_up_2")
        
        if excel_up and files and st.button("Proses Ganti Nama Gambar (ZIP)", key="process_img_rename_excel_2"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []
                        
                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            if old_name in file_map:
                                # Tambahkan ekstensi jika nama baru tidak memilikinya
                                if not os.path.splitext(new_name)[1]:
                                    _, old_ext = os.path.splitext(old_name)
                                    new_name = new_name + old_ext
                                    
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ {len(out_map)} file berhasil diganti namanya.")
                            st.download_button("Unduh Hasil (ZIP)", zipb, file_name="gambar_renamed_by_excel_2.zip", mime="application/zip")
                        else:
                            st.warning("Tidak ada file yang cocok ditemukan.")
                        
                        if not_found: st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()


    # --- LOGIKA FITUR FILE LAINNYA (tidak diubah) ---
    if mode == "📑 Batch Rename PDF (Sequential)":
        st.markdown("---")
        st.subheader("📑 Ganti Nama File PDF Massal (Sequential)")
        uploaded_files = st.file_uploader("Unggah file PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="batch_rename_pdf_uploader_file_tool")
        
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="Hasil_PDF", help="Contoh: Hasil_PDF_001.pdf", key="prefix_pdf_file_tool")
            start_num = col2.number_input("Mulai dari Angka (Counter):", min_value=1, value=1, step=1, key="start_num_pdf_file_tool")

            if st.button("Proses Ganti Nama (ZIP)", key="process_batch_rename_pdf_file_tool"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, start_num):
                                new_filename = f"{new_prefix}_{i:03d}.pdf"
                                zf.writestr(new_filename, file.read())
                        st.success(f"✅ Berhasil mengganti nama {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Rename", data=output_zip.getvalue(), file_name="pdf_renamed_file_tool.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()
            
    if mode == "🔢 Batch Rename/Format Gambar (Sequential)":
        st.markdown("---")
        st.subheader("🔢 Ganti Nama & Ubah Format Gambar Massal (Sequential)")
        uploaded_files = st.file_uploader(
            "Unggah file Gambar (JPG, PNG, dll.):", 
            type=["jpg", "jpeg", "png", "webp"], 
            accept_multiple_files=True,
            key="batch_rename_uploader_file_tool"
        )

        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="KAY_File", help="Contoh: KAY_File_001.jpg", key="prefix_img_file_tool")
            new_format = col2.selectbox("Format Output Baru:", ["Sama seperti Asli", "JPG", "PNG", "WEBP"], index=0, key="format_img_file_tool")

            if st.button("Proses Batch File", key="process_batch_rename_file_tool"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, 1):
                                
                                _, original_ext = os.path.splitext(file.name)
                                img = Image.open(file)
                                img_io = io.BytesIO()
                                
                                if new_format == "Sama seperti Asli":
                                    output_format_pil = img.format if img.format else 'JPEG'
                                    output_ext = original_ext
                                else:
                                    output_ext = "." + new_format.lower()
                                    output_format_pil = new_format.upper()

                                new_filename = f"{new_prefix}_{i:03d}{output_ext}"
                                
                                if output_format_pil in ('JPEG', 'JPG'):
                                    img.convert("RGB").save(img_io, format='JPEG', quality=95) 
                                elif output_format_pil == 'PNG':
                                    img.save(img_io, format='PNG')
                                elif output_format_pil == 'WEBP':
                                    img.save(img_io, format='WEBP')
                                else:
                                    img.save(img_io, format=output_format_pil) 

                                img_io.seek(0)
                                zf.writestr(new_filename, img_io.read())

                        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Batch", data=output_zip.getvalue(), file_name="hasil_batch_gambar_file_tool.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

    if mode == "Zip files":
        st.markdown("---")
        ups = st.file_uploader("Select files to zip", accept_multiple_files=True)
        if ups and st.button("Create ZIP"):
            try:
                with st.spinner("Membuat ZIP..."):
                    out = io.BytesIO()
                    with zipfile.ZipFile(out, "w") as z:
                        total = len(ups)
                        prog = st.progress(0)
                        for i, f in enumerate(ups):
                            z.writestr(f.name, f.read())
                            prog.progress(int((i+1)/total*100))
                        out.seek(0)
                st.download_button("Download ZIP", out.getvalue(), file_name="files.zip", mime="application/zip")
            except Exception:
                st.error(traceback.format_exc())
    
    elif mode == "Unzip file":
        st.markdown("---")
        zf = st.file_uploader("Upload zip file", type="zip")
        if zf and st.button("Extract"):
            try:
                with st.spinner("Mengekstrak ZIP..."):
                    with zipfile.ZipFile(io.BytesIO(zf.read()), "r") as z:
                        members = z.namelist()
                        st.write("Contains:", members)
                        tmpdir = tempfile.mkdtemp()
                        z.extractall(tmpdir)
                        shutil.make_archive(tmpdir, "zip", tmpdir)
                        with open(tmpdir + ".zip", "rb") as fh:
                            st.download_button("Download extracted as zip", fh.read(), file_name="extracted.zip", mime="application/zip")
                        shutil.rmtree(tmpdir)
            except Exception:
                st.error(traceback.format_exc())
    
    elif mode == "Excel -> CSV":
        st.markdown("---")
        file = st.file_uploader("Unggah file Excel:", type=["xlsx"])
        if file and st.button("Konversi ke CSV"):
            try:
                df = pd.read_excel(file)
                csv = df.to_csv(index=False).encode("utf-8")
                st.download_button("Unduh CSV", csv, "konversi.csv", "text/csv")
                st.success("Konversi berhasil")
            except Exception:
                st.error(traceback.format_exc())
    
    elif mode == "Word -> PDF (text)":
        st.markdown("---")
        file = st.file_uploader("Unggah file Word (.docx):", type=["docx"])
        if file and st.button("Konversi ke PDF"):
            if Document is None:
                st.error("python-docx is required for Word->PDF (pip install python-docx)")
            else:
                try:
                    doc = Document(io.BytesIO(file.read()))
                    text = "\n".join([p.text for p in doc.paragraphs])
                    pdf_buffer = io.BytesIO()
                    pdf_buffer.write(text.encode("utf-8"))
                    st.download_button("Unduh Hasil PDF (raw text)", pdf_buffer.getvalue(), "konversi.pdf", "application/pdf")
                    st.success("Konversi selesai (simple text dump). For accurate conversions, use LibreOffice headless or other converters.")
                except Exception:
                    st.error(traceback.format_exc())

# -------------- Tentang (Diperbarui) --------------
if menu == "Tentang":
    add_back_to_dashboard_button() 
    st.subheader("Tentang KAY App – Tools MCU")
    st.markdown("""
    **KAY App** adalah aplikasi serbaguna berbasis Streamlit untuk membantu:
    - Kompres foto & gambar
    - Pengelolaan dokumen PDF (gabung, pisah, proteksi, ekstraksi, **Reorder/Hapus Halaman**, **Batch Rename PDF**)
    - Analisis & pengolahan hasil Medical Check Up (MCU) (**Dashboard Analisis Data**)
    - Manajemen file & konversi dasar (**Batch Rename/Format Gambar**, **Batch Rename PDF**)

    Beberapa fitur memerlukan library tambahan (instal di environment Anda):
    - `PyPDF2` (Dasar PDF)
    - `pdfplumber` untuk ekstraksi tabel teks: `pip install pdfplumber`
    - `python-docx` untuk menghasilkan .docx: `pip install python-docx`
    - `pdf2image` + poppler untuk konversi PDF->Gambar / Preview gambar: `pip install pdf2image`
    - `pandas` & `openpyxl` untuk Analisis MCU dan **Batch Rename by Excel**: `pip install pandas openpyxl`
    """)
    st.info("Data diproses di server tempat Streamlit dijalankan. Untuk mengaktifkan semua fitur, pasang dependensi yang diperlukan.")

# ----------------- Footer -----------------
st.markdown("""
<hr style="border: none; border-top: 1px solid #cfe2ff; margin-top: 1.5rem; margin-bottom: 0.5rem;">
<div style="text-align:center; color:#5d6d7e; font-size:0.9rem;">
    © 2025 KAY App – Tools MCU | Built with 🛠️
</div>

""", unsafe_allow_html=True)
