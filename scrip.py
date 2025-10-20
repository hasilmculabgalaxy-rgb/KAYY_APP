# ===============================================================
# KAY App – FINAL FULL VERSION (Gabungan SCRIP 1 + SCRIP 2)
# Bagian 1 : Setup, Import, UI, Dashboard dasar
# ===============================================================

import os, io, zipfile, shutil, traceback, tempfile, time
import streamlit as st
import pandas as pd
from PIL import Image

# === PDF & DOCX Libraries ===
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

# === pdf2image (optional) ===
PDF2IMAGE_AVAILABLE = False
convert_from_bytes = convert_from_path = None
try:
    from pdf2image import convert_from_bytes, convert_from_path
    PDF2IMAGE_AVAILABLE = True
except Exception:
    try:
        from pdf2image import convert_from_path
        PDF2IMAGE_AVAILABLE = True
    except Exception:
        pass

# === Translator (deep_translator) ===
Translator = None
try:
    from deep_translator import GoogleTranslator as Translator
except Exception:
    pass

# === QR CODE (baru dari SCRIP 2) ===
import qrcode
from PIL import Image as PILImage
from io import BytesIO

# ===============================================================
# HELPERS
# ===============================================================

def make_zip_from_map(bytes_map: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, data in bytes_map.items():
            z.writestr(name, data)
    buf.seek(0)
    return buf.getvalue()

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

def try_encrypt(writer, password: str):
    """Encrypt PDF safely"""
    try:
        writer.encrypt(password)
    except TypeError:
        try:
            writer.encrypt(user_pwd=password, owner_pwd=None)
        except Exception:
            writer.encrypt(user_pwd=password, owner_pwd=password)

def rotate_page_safe(page, angle):
    try:
        page.rotate(angle)
    except Exception:
        try:
            from PyPDF2.generic import NameObject, NumberObject
            page.__setitem__(NameObject("/Rotate"), NumberObject(angle))
        except Exception:
            pass

def navigate_to(target_menu):
    st.session_state.menu_selection = target_menu
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

# ===============================================================
# STREAMLIT CONFIG & CSS
# ===============================================================

st.set_page_config(
    page_title="KAY App – Tools Serbaguna",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
[data-testid="stSidebarToggleButton"],
section[data-testid="stSidebar"],
[data-testid="stDecoration"] {display:none!important;}
.stApp {
    background: linear-gradient(180deg,#e9f2ff 0%,#f4f9ff 100%);
    color:#002b5b; font-family:'Inter',sans-serif;
}
div.stButton>button {
    background:linear-gradient(90deg,#5dade2,#3498db);
    color:white;border:none;border-radius:12px;
    padding:.7rem 1.2rem;font-weight:600;
    box-shadow:0 4px 10px rgba(52,152,219,.4);
    transition:.2s;cursor:pointer;
}
div.stButton>button:hover {
    transform:scale(1.03);
    box-shadow:0 8px 18px rgba(52,152,219,.6);
}
.feature-card {
    background:white;border-radius:16px;
    box-shadow:0 4px 15px rgba(0,0,0,.08);
    padding:24px;transition:.3s;border:1px solid #d0e3ff;
    height:100%;display:flex;flex-direction:column;justify-content:space-between;
}
.feature-card:hover {
    box-shadow:0 10px 30px rgba(0,0,0,.15);
    transform:translateY(-5px);border-color:#3498db;
}
h1{color:#1b4f72;font-weight:800;}
h2,h3{color:#2e86c1;}
.dashboard-container{padding:20px 0;margin-bottom:20px;}
</style>
""", unsafe_allow_html=True)

# ===============================================================
# HEADER
# ===============================================================

if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "Dashboard"
menu = st.session_state.menu_selection

header_col1, header_col2 = st.columns([1, 4])
with header_col1:
    st.markdown("<h1 style='font-size:3rem;'>🧩</h1>", unsafe_allow_html=True)
with header_col2:
    st.title("KAY App – Tools Serbaguna MCU + PDF + QR")
    st.markdown("Solusi lengkap untuk PDF, gambar, MCU dan kode QR dalam satu aplikasi Streamlit elegan.")
st.markdown("---")

# ===============================================================
# DASHBOARD UTAMA
# ===============================================================

if menu == "Dashboard":
    st.markdown('<div class="dashboard-container">', unsafe_allow_html=True)
    st.markdown("## ✳️ Pilih Fitur Utama")

    cols1 = st.columns(3)
    # Kompres Foto
    with cols1[0]:
        st.markdown('<div class="feature-card"><b>🖼️ Kompres Foto / Gambar</b><br>Perkecil ukuran file dan rename massal.</div>', unsafe_allow_html=True)
        if st.button("Buka Kompres Foto", key="dash_foto"):
            navigate_to("Kompres Foto")
    # PDF Tools
    with cols1[1]:
        st.markdown('<div class="feature-card"><b>📄 PDF Tools</b><br>Gabung, pisah, proteksi, rename, translate PDF.</div>', unsafe_allow_html=True)
        if st.button("Buka PDF Tools", key="dash_pdf"):
            navigate_to("PDF Tools")
    # MCU Tools
    with cols1[2]:
        st.markdown('<div class="feature-card"><b>🩺 MCU Tools</b><br>Analisis dan organisasi hasil MCU berbasis Excel.</div>', unsafe_allow_html=True)
        if st.button("Buka MCU Tools", key="dash_mcu"):
            navigate_to("MCU Tools")

    st.markdown("## 🧰 Fitur Lainnya")
    cols2 = st.columns(3)
    with cols2[0]:
        st.markdown('<div class="feature-card"><b>📂 File Tools</b><br>Zip/unzip, konversi dan rename batch.</div>', unsafe_allow_html=True)
        if st.button("Buka File Tools", key="dash_file"):
            navigate_to("File Tools")
    with cols2[1]:
        st.markdown('<div class="feature-card"><b>🔳 QR Code Generator ✨</b><br>Buat QR Code dengan logo khusus anda.</div>', unsafe_allow_html=True)
        if st.button("Buka QR Code Generator", key="dash_qr"):
            navigate_to("QR Code")
    with cols2[2]:
        st.markdown('<div class="feature-card"><b>ℹ️ Tentang</b><br>Informasi dan library yang dibutuhkan.</div>', unsafe_allow_html=True)
        if st.button("Lihat Tentang", key="dash_about"):
            navigate_to("Tentang")

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")
    st.info("Semua proses berlangsung lokal pada server Streamlit anda.")

# ---------------------------------------------------------------
# Bagian 2 akan berisi semua fungsi fitur lengkap
# ---------------------------------------------------------------
# ===============================================================
# BAGIAN 2 – Semua Fitur Utama: Foto, PDF, MCU, File
# ===============================================================

# ---------------------------------------------------------------
# 🔙 Tombol kembali ke dashboard
# ---------------------------------------------------------------
def add_back_to_dashboard_button():
    if st.button("🏠 Kembali ke Dashboard", key="back_dash"):
        navigate_to("Dashboard")
    st.markdown("---")

# ===============================================================
# 📸 FITUR: KOMPRES & KELOLA FOTO
# ===============================================================
if menu == "Kompres Foto":
    add_back_to_dashboard_button()
    st.subheader("🖼️ Kompres & Kelola Foto / Gambar")

    img_tool = st.selectbox("Pilih Fitur:", [
        "📉 Kompres Foto (Batch)",
        "🧾 Batch Rename Gambar (Sequential)",
        "📑 Batch Rename Gambar Sesuai Excel"
    ])

    # Kompres Foto
    if img_tool == "📉 Kompres Foto (Batch)":
        uploaded = st.file_uploader("Unggah gambar (jpg/png) – bisa banyak:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        quality = st.slider("Kualitas JPEG", 10, 95, 75)
        max_side = st.number_input("Ukuran maksimum (px)", 100, 4000, 1200)
        if uploaded and st.button("Kompres Semua"):
            out_map = {}
            total = len(uploaded)
            prog = st.progress(0)
            with st.spinner("Mengompres..."):
                for i, f in enumerate(uploaded):
                    try:
                        im = Image.open(io.BytesIO(f.read()))
                        im.thumbnail((max_side, max_side))
                        buf = io.BytesIO()
                        im.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
                        out_map[f"compressed_{f.name}"] = buf.getvalue()
                    except Exception as e:
                        st.warning(f"Gagal: {f.name} — {e}")
                    prog.progress(int((i + 1) / total * 100))
            if out_map:
                st.success(f"✅ {len(out_map)} file berhasil dikompres")
                st.download_button("Unduh Hasil (ZIP)", make_zip_from_map(out_map),
                                   file_name="foto_kompres.zip", mime="application/zip")

    # Batch Rename Sequential
    elif img_tool == "🧾 Batch Rename Gambar (Sequential)":
        uploaded_files = st.file_uploader("Unggah file gambar:", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
        prefix = st.text_input("Prefix nama file:", "KAY_File")
        fmt = st.selectbox("Format output:", ["Sama seperti asli", "JPG", "PNG", "WEBP"], index=0)
        if uploaded_files and st.button("Proses Batch Rename"):
            output = io.BytesIO()
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, f in enumerate(uploaded_files, 1):
                    im = Image.open(f)
                    ext = os.path.splitext(f.name)[1]
                    if fmt != "Sama seperti asli":
                        ext = "." + fmt.lower()
                    name = f"{prefix}_{i:03d}{ext}"
                    buf = io.BytesIO()
                    im.convert("RGB").save(buf, format=fmt if fmt != "Sama seperti asli" else im.format)
                    buf.seek(0)
                    zf.writestr(name, buf.read())
            st.download_button("Unduh ZIP", output.getvalue(), file_name="gambar_batch_rename.zip", mime="application/zip")
            st.success("✅ Berhasil mengganti nama file")

    # Batch Rename Gambar via Excel
    elif img_tool == "📑 Batch Rename Gambar Sesuai Excel":
        st.info("Excel harus memiliki kolom `nama_lama` dan `nama_baru`")
        excel = st.file_uploader("Unggah Excel/CSV:", type=["xlsx", "csv"])
        files = st.file_uploader("Unggah gambar:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        if excel and files and st.button("Proses Rename"):
            df = pd.read_excel(excel) if excel.name.endswith(".xlsx") else pd.read_csv(excel)
            required_cols = ['nama_lama', 'nama_baru']
            if not all(c in df.columns for c in required_cols):
                st.error(f"Kolom wajib: {', '.join(required_cols)}")
            else:
                file_map = {f.name: f.read() for f in files}
                out_map, not_found = {}, []
                for _, row in df.iterrows():
                    old, new = str(row['nama_lama']).strip(), str(row['nama_baru']).strip()
                    if old in file_map:
                        if not os.path.splitext(new)[1]:
                            new += os.path.splitext(old)[1]
                        out_map[new] = file_map[old]
                    else:
                        not_found.append(old)
                if out_map:
                    st.success(f"✅ {len(out_map)} file berhasil diganti nama")
                    st.download_button("Unduh ZIP", make_zip_from_map(out_map),
                                       file_name="gambar_renamed.zip", mime="application/zip")
                if not_found:
                    st.warning(f"Tidak ditemukan: {not_found[:5]} ...")

# ===============================================================
# 📄 FITUR: PDF TOOLS
# ===============================================================
if menu == "PDF Tools":
    add_back_to_dashboard_button()
    st.subheader("📄 PDF Tools")

    pdf_option = st.selectbox("Pilih Fitur:", [
        "--- Pilih ---", "📎 Gabung PDF", "✂️ Pisah PDF", "🔄 Reorder / Hapus Halaman",
        "🔐 Proteksi / Encrypt PDF", "🔓 Decrypt PDF",
        "🧾 Batch Rename PDF (Sequential)", "📑 Batch Rename PDF Sesuai Excel",
        "🌍 Terjemahkan PDF ke Bahasa Lain",
        "🖼️ PDF ➜ Image", "📃 Image ➜ PDF"
    ])

    # Gabung PDF
    if pdf_option == "📎 Gabung PDF":
        files = st.file_uploader("Unggah PDF (multiple):", type="pdf", accept_multiple_files=True)
        if files and st.button("Gabung"):
            writer = PdfWriter()
            for f in files:
                reader = PdfReader(io.BytesIO(f.read()))
                for p in reader.pages:
                    writer.add_page(p)
            buf = io.BytesIO(); writer.write(buf); buf.seek(0)
            st.download_button("Unduh PDF Gabungan", buf.getvalue(), file_name="merged.pdf", mime="application/pdf")
            st.success("✅ PDF berhasil digabung")

    # Pisah PDF
    if pdf_option == "✂️ Pisah PDF":
        f = st.file_uploader("Unggah 1 PDF:", type="pdf")
        if f and st.button("Pisahkan"):
            reader = PdfReader(io.BytesIO(f.read()))
            out_map = {}
            for i, p in enumerate(reader.pages):
                w = PdfWriter(); w.add_page(p)
                buf = io.BytesIO(); w.write(buf); buf.seek(0)
                out_map[f"page_{i+1}.pdf"] = buf.getvalue()
            st.download_button("Unduh ZIP", make_zip_from_map(out_map), file_name="pages.zip", mime="application/zip")
            st.success("✅ PDF berhasil dipisah")

    # Reorder / Hapus Halaman
    if pdf_option == "🔄 Reorder / Hapus Halaman":
        f = st.file_uploader("Unggah PDF:", type="pdf")
        order = st.text_input("Urutan halaman (misal: 2,1,3)", "")
        if f and st.button("Proses"):
            reader = PdfReader(io.BytesIO(f.read()))
            pages = [int(x.strip()) - 1 for x in order.split(",") if x.strip().isdigit()]
            writer = PdfWriter()
            for i in pages:
                writer.add_page(reader.pages[i])
            buf = io.BytesIO(); writer.write(buf); buf.seek(0)
            st.download_button("Unduh PDF Baru", buf.getvalue(), file_name="reordered.pdf", mime="application/pdf")
            st.success("✅ Halaman berhasil diatur ulang")

    # Batch Rename PDF Sequential
    if pdf_option == "🧾 Batch Rename PDF (Sequential)":
        files = st.file_uploader("Unggah PDF (multiple):", type="pdf", accept_multiple_files=True)
        prefix = st.text_input("Prefix:", "Hasil_PDF")
        if files and st.button("Proses Rename"):
            output = io.BytesIO()
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, f in enumerate(files, 1):
                    zf.writestr(f"{prefix}_{i:03d}.pdf", f.read())
            st.download_button("Unduh ZIP", output.getvalue(), file_name="pdf_renamed.zip", mime="application/zip")
            st.success("✅ Berhasil ganti nama")

    # Batch Rename PDF by Excel
    if pdf_option == "📑 Batch Rename PDF Sesuai Excel":
        st.info("Excel wajib punya kolom `nama_lama` dan `nama_baru`")
        excel = st.file_uploader("Unggah Excel:", type=["xlsx", "csv"])
        files = st.file_uploader("Unggah PDF:", type=["pdf"], accept_multiple_files=True)
        if excel and files and st.button("Proses"):
            df = pd.read_excel(excel) if excel.name.endswith(".xlsx") else pd.read_csv(excel)
            file_map = {f.name: f.read() for f in files}
            out_map, not_found = {}, []
            for _, row in df.iterrows():
                old, new = str(row['nama_lama']).strip(), str(row['nama_baru']).strip()
                if old in file_map:
                    if not new.lower().endswith(".pdf"):
                        new += ".pdf"
                    out_map[new] = file_map[old]
                else:
                    not_found.append(old)
            if out_map:
                st.download_button("Unduh ZIP", make_zip_from_map(out_map), file_name="pdf_renamed_by_excel.zip", mime="application/zip")
                st.success(f"✅ {len(out_map)} file berhasil diganti nama")
            if not_found:
                st.warning(f"Tidak ditemukan: {not_found[:5]}")

# (lanjutan PDF ➜ Translate, Encrypt, MCU Tools, File Tools akan dilanjut di Bagian 3)
# ===============================================================
# BAGIAN 3 – Lanjutan: Translate PDF, MCU, File Tools, QR, Footer
# ===============================================================

# ---------------------------------------------------------------
# 🌍 Terjemahan PDF
# ---------------------------------------------------------------
if menu == "PDF Tools" and pdf_option == "🌍 Terjemahkan PDF ke Bahasa Lain":
    st.markdown("---")
    st.subheader("🌍 Terjemahkan PDF ke Bahasa Lain")

    f = st.file_uploader("Unggah PDF:", type="pdf")
    col1, col2 = st.columns(2)
    src = col1.text_input("Bahasa sumber (ISO, contoh: auto)", "auto")
    tgt = col2.text_input("Bahasa tujuan (contoh: en, id, fr)", "en")

    if f and st.button("Proses Terjemahan"):
        try:
            if Translator is None or Document is None:
                st.error("Pastikan library `deep_translator` dan `python-docx` sudah terinstall.")
                st.stop()

            raw = f.read()
            text_blocks = []
            if pdfplumber:
                with pdfplumber.open(io.BytesIO(raw)) as doc:
                    for p in doc.pages:
                        text_blocks.append(p.extract_text() or "")
            else:
                reader = PdfReader(io.BytesIO(raw))
                for p in reader.pages:
                    text_blocks.append(p.extract_text() or "")

            full_text = "\n\n".join(text_blocks)
            translator = Translator(source=src, target=tgt)
            chunks = [full_text[i:i+4000] for i in range(0, len(full_text), 4000)]
            translated = ""
            prog = st.progress(0)
            for i, chunk in enumerate(chunks):
                translated += translator.translate(chunk) + "\n"
                prog.progress(int((i + 1) / len(chunks) * 100))
            prog.empty()

            doc = Document()
            for para in translated.split("\n\n"):
                doc.add_paragraph(para.strip())
            out = io.BytesIO()
            doc.save(out)
            out.seek(0)
            st.success("✅ Terjemahan selesai")
            st.download_button("Unduh File Word", out.getvalue(), file_name="translated.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

# ---------------------------------------------------------------
# 🔐 Encrypt / Decrypt PDF
# ---------------------------------------------------------------
if menu == "PDF Tools" and pdf_option in ["🔐 Proteksi / Encrypt PDF", "🔓 Decrypt PDF"]:
    f = st.file_uploader("Unggah PDF:", type="pdf")
    pw = st.text_input("Password:", type="password")
    if f and pw and st.button("Proses"):
        reader = PdfReader(io.BytesIO(f.read()))
        writer = PdfWriter()
        for p in reader.pages:
            writer.add_page(p)
        if pdf_option == "🔐 Proteksi / Encrypt PDF":
            try_encrypt(writer, pw)
            out_name = "encrypted.pdf"
        else:
            reader.decrypt(pw)
            out_name = "decrypted.pdf"
        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
        st.download_button("Unduh PDF", buf.getvalue(), file_name=out_name, mime="application/pdf")
        st.success("✅ Proses selesai")

# ===============================================================
# 🩺 MCU TOOLS
# ===============================================================
if menu == "MCU Tools":
    add_back_to_dashboard_button()
    st.subheader("🩺 MCU Tools – Analisis Data & Organisasi File")

    tool = st.selectbox("Pilih Fitur:", [
        "📊 Dashboard Analisis Data MCU (Excel)",
        "🗂️ Organise File MCU by Excel"
    ])

    if tool == "📊 Dashboard Analisis Data MCU (Excel)":
        excel = st.file_uploader("Unggah Excel hasil MCU:", type=["xlsx", "csv"])
        if excel:
            df = pd.read_excel(excel) if excel.name.endswith(".xlsx") else pd.read_csv(excel)
            st.dataframe(df.head())
            st.success(f"✅ Data terbaca: {len(df)} baris")
            if st.checkbox("Tampilkan statistik ringkas"):
                st.dataframe(df.describe(include='all'))

    elif tool == "🗂️ Organise File MCU by Excel":
        st.info("Excel harus memiliki kolom `nama_lama` dan `folder_tujuan`.")
        excel = st.file_uploader("Unggah Excel mapping:", type=["xlsx", "csv"])
        files = st.file_uploader("Unggah file MCU:", accept_multiple_files=True)
        if excel and files and st.button("Proses Organisasi"):
            df = pd.read_excel(excel) if excel.name.endswith(".xlsx") else pd.read_csv(excel)
            file_map = {f.name: f.read() for f in files}
            out_zip = io.BytesIO()
            with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for _, r in df.iterrows():
                    old = str(r['nama_lama']).strip()
                    folder = str(r['folder_tujuan']).strip()
                    if old in file_map:
                        zf.writestr(f"{folder}/{old}", file_map[old])
            out_zip.seek(0)
            st.download_button("Unduh ZIP Terorganisir", out_zip.getvalue(), file_name="mcu_organised.zip", mime="application/zip")
            st.success("✅ File terorganisir sesuai Excel")

# ===============================================================
# 📂 FILE TOOLS
# ===============================================================
if menu == "File Tools":
    add_back_to_dashboard_button()
    st.subheader("📂 File Tools")

    tool = st.selectbox("Pilih Fitur:", [
        "📦 Zip / Unzip File",
        "📑 Konversi TXT/CSV/JSON ke Excel"
    ])

    if tool == "📦 Zip / Unzip File":
        mode = st.radio("Mode:", ["Zip Files", "Extract ZIP"])
        if mode == "Zip Files":
            files = st.file_uploader("Unggah file:", accept_multiple_files=True)
            if files and st.button("Buat ZIP"):
                out = {f.name: f.read() for f in files}
                st.download_button("Unduh ZIP", make_zip_from_map(out), file_name="files.zip", mime="application/zip")
        else:
            f = st.file_uploader("Unggah ZIP:", type=["zip"])
            if f and st.button("Ekstrak ZIP"):
                z = zipfile.ZipFile(io.BytesIO(f.read()))
                out = {name: z.read(name) for name in z.namelist() if not name.endswith("/")}
                st.download_button("Unduh ZIP Hasil", make_zip_from_map(out), file_name="unzipped.zip", mime="application/zip")

    elif tool == "📑 Konversi TXT/CSV/JSON ke Excel":
        f = st.file_uploader("Unggah File:", type=["txt", "csv", "json"])
        if f and st.button("Konversi"):
            df = None
            if f.name.endswith(".csv"):
                df = pd.read_csv(f)
            elif f.name.endswith(".json"):
                df = pd.read_json(f)
            elif f.name.endswith(".txt"):
                df = pd.read_csv(f)
            if df is not None:
                st.dataframe(df.head())
                st.download_button("Unduh Excel", df_to_excel_bytes(df), file_name="converted.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success("✅ Konversi berhasil")

# ===============================================================
# 🔳 QR CODE GENERATOR DENGAN LOGO
# ===============================================================
if menu == "QR Code":
    add_back_to_dashboard_button()
    st.subheader("🔳 QR Code Generator dengan Logo")

    data = st.text_input("Masukkan teks atau URL:", "")
    logo_file = st.file_uploader("Unggah logo (opsional):", type=["png", "jpg", "jpeg", "webp"])
    col1, col2 = st.columns(2)
    box_size = col1.slider("Ukuran kotak QR", 5, 20, 10)
    border = col2.slider("Tebal border", 2, 10, 4)

    if st.button("🚀 Buat QR Code"):
        if not data:
            st.warning("Masukkan teks terlebih dahulu.")
        else:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=box_size,
                border=border,
            )
            qr.add_data(data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

            if logo_file is not None:
                logo = PILImage.open(logo_file)
                w, h = img.size
                logo_size = int(w * 0.2)
                logo = logo.resize((logo_size, logo_size))
                pos = ((w - logo_size) // 2, (h - logo_size) // 2)
                img.paste(logo, pos, mask=logo if logo.mode == "RGBA" else None)

            buf = BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            st.image(img, caption="Hasil QR Code", use_container_width=False)
            st.download_button("💾 Unduh QR Code", buf.getvalue(), file_name="qrcode.png", mime="image/png")
            st.success("✅ QR Code berhasil dibuat!")

# ===============================================================
# ℹ️ TENTANG
# ===============================================================
if menu == "Tentang":
    add_back_to_dashboard_button()
    st.subheader("ℹ️ Tentang Aplikasi KAY App")
    st.markdown("""
    **KAY App – Tools Serbaguna**
    - Gabungan fitur: PDF, Gambar, MCU, QR Code.
    - Dibangun dengan Streamlit.
    - Dibuat dengan ❤️ oleh tim pengembang KAY.

    **Library Utama yang digunakan:**
    - streamlit, PyPDF2, pdfplumber, python-docx
    - pillow, pdf2image, qrcode, pandas, openpyxl, deep_translator
    """)

# ===============================================================
# FOOTER
# ===============================================================
def footer():
    st.markdown("""
    <hr style="border:none;height:1px;background:#ccc;">
    <div style="text-align:center;color:gray;font-size:0.8em;">
        © 2025 KAY App – Dibuat dengan ❤️ dan Streamlit.
    </div>
    """, unsafe_allow_html=True)

footer()

