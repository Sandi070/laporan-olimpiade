import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
from google import genai 

# --- FUNGSI TAMBAHAN: FORMAT TANGGAL INDONESIA ---
def tanggal_indo(tgl):
    """Mengubah format tanggal menjadi Bahasa Indonesia (Contoh: Senin, 16 September 2024)"""
    # Kamus Hari
    hari_dict = {
        'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
        'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu',
        'Sunday': 'Minggu'
    }
    # Kamus Bulan
    bulan_dict = {
        'January': 'Januari', 'February': 'Februari', 'March': 'Maret',
        'April': 'April', 'May': 'Mei', 'June': 'Juni',
        'July': 'Juli', 'August': 'Agustus', 'September': 'September',
        'October': 'Oktober', 'November': 'November', 'December': 'Desember'
    }
    
    # Ambil nama hari dan bulan dalam bahasa Inggris dulu
    nama_hari_eng = tgl.strftime("%A")
    nama_bulan_eng = tgl.strftime("%B")
    
    # Terjemahkan
    nama_hari_indo = hari_dict[nama_hari_eng]
    nama_bulan_indo = bulan_dict[nama_bulan_eng]
    tahun = tgl.strftime("%Y")
    tanggal = tgl.strftime("%d")
    
    # Gabungkan
    return f"{nama_hari_indo}, {tanggal} {nama_bulan_indo} {tahun}"

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Generator Laporan Olimpiade", layout="wide")

# --- Judul Aplikasi ---
st.title("🏆 Generator Laporan Pembimbingan Olimpiade")
st.markdown("Status: **Siap Digunakan** (Support Tanggal Indo & Tahun Otomatis)")

# --- Sidebar: Konfigurasi API ---
with st.sidebar:
    st.header("⚙️ Pengaturan")
    api_key = st.text_input("Masukkan Google Gemini API Key", type="password")
    
    # Model default stabil
    model_pilihan = "gemini-flash-latest" 
    
    if api_key:
        st.success("API Key terdeteksi.")
    else:
        st.warning("Masukkan API Key untuk fitur Otomatis.")
        st.markdown("[Buat API Key di sini](https://aistudio.google.com/app/apikey)")

# --- Input Data Kegiatan ---
st.header("📝 Input Data Laporan")

col1, col2 = st.columns(2)

with col1:
    bidang_osn = ["Kimia", "Fisika", "Biologi", "Matematika", "Ekonomi", "Geografi", "Kebumian", "Astronomi", "Informatika"]
    mapel = st.selectbox("Mata Pelajaran", bidang_osn)
    nama_pembahas = st.text_input("Nama Tutor", "Sandi Saputra, S.Pd.")
    
with col2:
    # Input tanggal standar
    tanggal = st.date_input("Tanggal Kegiatan", datetime.date.today())
    waktu_mulai = st.time_input("Waktu Mulai", datetime.time(9, 0))
    waktu_selesai = st.time_input("Waktu Selesai", datetime.time(11, 0))

materi = st.text_input("Materi Pembahasan", "Contoh: Stoikiometri / Konsep Mol")
jumlah_peserta = st.number_input("Jumlah Peserta", min_value=1, value=8)

# --- Upload Dokumentasi ---
st.subheader("📷 Dokumentasi")
uploaded_file = st.file_uploader("Upload Foto Kegiatan", type=['png', 'jpg', 'jpeg'])
if uploaded_file:
    st.image(uploaded_file, width=400, caption="Preview Foto")

# --- Logika AI ---
def generate_description_ai(api_key, model, mapel, materi, tanggal_obj, peserta, mulai, selesai):
    try:
        client = genai.Client(api_key=api_key)
        
        # Format tanggal ke Indonesia dulu sebelum dikirim ke AI
        tgl_teks = tanggal_indo(tanggal_obj)
        
        prompt = f"""
        Bertindaklah sebagai guru pembimbing olimpiade di sekolah unggulan (MAN Insan Cendekia).
        Buatkan 1 paragraf laporan deskriptif formal dalam Bahasa Indonesia baku.
        
        Data Kegiatan:
        - Mapel: {mapel}
        - Materi: {materi}
        - Hari/Tanggal: {tgl_teks}
        - Waktu: {mulai} - {selesai}
        - Peserta: {peserta} orang
        
        Instruksi Isi Laporan:
        1. Sebutkan hari dan tanggal kegiatan dengan lengkap (Bahasa Indonesia).
        2. Tekankan bahwa materi ini adalah fondasi penting.
        3. Gambarkan suasana kelas yang kondusif dan siswa antusias.
        
        Output: Hanya paragraf isinya saja. Jangan pakai judul.
        """

        response = client.models.generate_content(
            model=model, 
            contents=prompt
        )
        return response.text
        
    except Exception as e:
        return f"Gagal Generate. Error: {str(e)}"

def generate_description_manual(mapel, materi, tanggal_obj, peserta, mulai, selesai):
    # Menggunakan fungsi tanggal_indo agar manual pun tetap Bahasa Indonesia
    tgl_teks = tanggal_indo(tanggal_obj)
    
    return f"Pada hari ini, {tgl_teks}, dilakukan pembimbingan olimpiade bidang {mapel} dengan peserta sebanyak {peserta} orang. Bimbingan dimulai jam {mulai} sampai {selesai} WITA. Materi yang diajarkan adalah {materi}. Materi ini merupakan dasar yang sangat penting. Fokus pertemuan kali ini adalah pendalaman konsep dan latihan soal. Hasilnya, siswa terlihat antusias dan mampu mengerjakan soal dengan baik."

# Tombol Generate
st.divider()
col_btn, col_res = st.columns([1, 2])

with col_btn:
    st.subheader("Langkah 1: Generate Teks")
    if st.button("✨ Buat Deskripsi Otomatis"):
        if api_key:
            with st.spinner("AI sedang menyusun laporan..."):
                res = generate_description_ai(api_key, model_pilihan, mapel, materi, tanggal, jumlah_peserta, waktu_mulai, waktu_selesai)
                st.session_state['deskripsi'] = res
        else:
            st.session_state['deskripsi'] = generate_description_manual(mapel, materi, tanggal, jumlah_peserta, waktu_mulai, waktu_selesai)

with col_res:
    deskripsi_final = st.text_area("Hasil Deskripsi (Silakan edit jika perlu):", value=st.session_state.get('deskripsi', ''), height=180)

# --- Fungsi Membuat Dokumen Word ---
def create_docx(nama, tgl_obj, mat, deskripsi, gambar):
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    # 1. Judul Header (TAHUN OTOMATIS)
    tahun_laporan = tgl_obj.year 
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"LAPORAN KEGIATAN\nPROGRAM OLIMPIADE\nMAN INSAN CENDEKIA KOTA KENDARI TAHUN {tahun_laporan}")
    run.bold = True
    run.font.size = Pt(14)
    
    doc.add_paragraph() # Spasi
    
    # 2. Tabel Metadata (TANGGAL INDONESIA)
    # Kita panggil fungsi tanggal_indo di sini dan di-uppercase
    tgl_indo_lengkap = tanggal_indo(tgl_obj).upper()
    
    table = doc.add_table(rows=3, cols=3)
    metadata = [
        ("NAMA PEMBAHAS", ":", nama),
        ("HARI/TANGGAL", ":", tgl_indo_lengkap), # Hasil: SENIN, 16 SEPTEMBER 2024
        ("MATERI", ":", mat.upper())
    ]
    
    for i, data in enumerate(metadata):
        row = table.rows[i]
        row.cells[0].text = data[0]
        row.cells[1].text = data[1]
        row.cells[2].text = data[2]
        row.cells[0].paragraphs[0].runs[0].bold = True 
        
    doc.add_paragraph() 

    # 3. Deskripsi
    title_desc = doc.add_paragraph("DESKRIPSI SINGKAT KEGIATAN")
    title_desc.runs[0].bold = True
    doc.add_paragraph(deskripsi)
    
    doc.add_paragraph() 
    
    # 4. Dokumentasi
    title_doc = doc.add_paragraph("DOKUMENTASI")
    title_doc.runs[0].bold = True
    
    if gambar:
        doc.add_picture(gambar, width=Inches(6.0))
        last_p = doc.paragraphs[-1]
        last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Tombol Download
st.divider()
st.subheader("Langkah 2: Download File")

if st.button("💾 Download Laporan (.docx)"):
    if deskripsi_final and "Gagal Generate" not in deskripsi_final:
        file_docx = create_docx(nama_pembahas, tanggal, materi, deskripsi_final, uploaded_file)
        file_name = f"Laporan_{mapel}_{tanggal}.docx"
        
        st.download_button(
            label="Klik untuk Mengunduh",
            data=file_docx,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success(f"File laporan Tahun {tanggal.year} berhasil dibuat!")
    else:
        st.error("Pastikan deskripsi sudah terisi dan tidak error sebelum download.")