import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
from google import genai 

# --- FUNGSI FORMAT TANGGAL INDONESIA ---
def tanggal_indo(tgl):
    hari_dict = {'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu', 'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'}
    bulan_dict = {'January': 'Januari', 'February': 'Februari', 'March': 'Maret', 'April': 'April', 'May': 'Mei', 'June': 'Juni', 'July': 'Juli', 'August': 'Agustus', 'September': 'September', 'October': 'Oktober', 'November': 'November', 'December': 'Desember'}
    return f"{hari_dict[tgl.strftime('%A')]}, {tgl.strftime('%d')} {bulan_dict[tgl.strftime('%B')]} {tgl.strftime('%Y')}"

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Sistem Laporan MAN IC Kendari", layout="wide")

# --- JUDUL APLIKASI ---
st.title("🏫 Generator Laporan Kegiatan MAN IC Kendari")
st.markdown("Satu aplikasi untuk semua laporan program unggulan (Olimpiade, TKA, UTBK, Pengasuhan, dll).")

# --- SIDEBAR: MENU UTAMA & API ---
with st.sidebar:
    st.header("⚙️ Pengaturan")
    api_key = st.text_input("Google Gemini API Key", type="password")
    if not api_key:
        st.warning("Masukkan API Key agar fitur AI aktif.")
    
    st.divider()
    
    st.header("📂 Pilih Program")
    # Menu Dropdown untuk memilih jenis kegiatan
    jenis_program = st.selectbox(
        "Jenis Kegiatan Laporan:",
        [
            "Bimbingan Olimpiade",
            "Bimbingan TKA (Kompetensi Akademik)",
            "Bimbingan UTBK/SNBT",
            "Klinik Mata Pelajaran (Remedial)",
            "Karya Ilmiah Remaja (KIR)",
            "Pendampingan Belajar Malam",
            "Kegiatan Pengasuhan (Guru Asuh)"
        ]
    )

# --- INPUT DATA UTAMA (DINAMIS SESUAI PROGRAM) ---
st.header(f"📝 Input Data: {jenis_program}")

col1, col2 = st.columns(2)

# Variabel penampung data (default kosong)
mapel = "-"
materi = "-"
kategori_malam = "-"
topik_pengasuhan = "-"

with col1:
    nama_pembahas = st.text_input("Nama Guru/Pembina", "Sandi Saputra, S.Pd.")
    
    # LOGIKA INPUT BERDASARKAN PROGRAM
    if jenis_program in ["Bimbingan Olimpiade", "Bimbingan TKA (Kompetensi Akademik)", "Bimbingan UTBK/SNBT", "Klinik Mata Pelajaran (Remedial)"]:
        bidang_mapel = ["Kimia", "Fisika", "Biologi", "Matematika", "Ekonomi", "Geografi", "Kebumian", "Astronomi", "Informatika", "Bahasa Inggris", "Bahasa Indonesia", "TPS/TPA"]
        mapel = st.selectbox("Mata Pelajaran", bidang_mapel)
        materi = st.text_input("Materi / Topik Bahasan", "Contoh: Stoikiometri / Latihan Soal Paket 1")
        
    elif jenis_program == "Karya Ilmiah Remaja (KIR)":
        mapel = st.selectbox("Bidang Penelitian", ["IPA (Saintek)", "IPS (Soshum)", "Keagamaan", "Teknologi"])
        materi = st.text_input("Judul/Topik Penelitian Siswa", "Contoh: Pengaruh Limbah Sagu terhadap...")

    elif jenis_program == "Pendampingan Belajar Malam":
        kategori_malam = st.selectbox("Jenis Kegiatan Malam", ["Belajar Mandiri (KBM)", "Latihan Upacara", "Latihan Seni/Pentas", "Kegiatan Asrama Lainnya"])
        materi = st.text_input("Detail Kegiatan", "Contoh: Persiapan Petugas Upacara Hari Senin")
        mapel = "Pendampingan Asrama" # Default label

    elif jenis_program == "Kegiatan Pengasuhan (Guru Asuh)":
        mapel = "Pengasuhan & Konseling"
        topik_pengasuhan = st.text_area("Topik/Masalah yang Dibahas", "Contoh: Keluhan air asrama mati, diskusi menu kantin, motivasi belajar, atau penegakan disiplin kebersihan.")

with col2:
    tanggal = st.date_input("Tanggal Kegiatan", datetime.date.today())
    waktu_mulai = st.time_input("Waktu Mulai", datetime.time(9, 0))
    waktu_selesai = st.time_input("Waktu Selesai", datetime.time(11, 0))
    jumlah_peserta = st.number_input("Jumlah Siswa Hadir", min_value=1, value=10)

# --- UPLOAD FOTO ---
st.subheader("📷 Dokumentasi")
uploaded_file = st.file_uploader("Upload Foto Kegiatan", type=['png', 'jpg', 'jpeg'])
if uploaded_file:
    st.image(uploaded_file, width=400, caption="Preview Foto")

# --- LOGIKA AI (PROMPT DINAMIS) ---
def generate_description_ai(api_key, program, mapel, materi, topik, kategori_malam, tanggal_obj, peserta, mulai, selesai):
    try:
        client = genai.Client(api_key=api_key)
        tgl_teks = tanggal_indo(tanggal_obj)
        
        # PROMPT ENGINEER: Menyesuaikan instruksi berdasarkan jenis program
        konteks_khusus = ""
        
        if "Olimpiade" in program:
            konteks_khusus = f"Fokus laporan: Pendalaman materi olimpiade {mapel} topik {materi}. Tekankan materi ini fundamental untuk kompetisi."
        elif "TKA" in program:
            konteks_khusus = f"Fokus laporan: Persiapan Tes Kompetensi Akademik (TKA) Kemendikdasmen mapel {mapel}. Laporan berisi drill soal dan asesmen kompetensi siswa."
        elif "UTBK" in program:
            konteks_khusus = f"Fokus laporan: Persiapan seleksi masuk PTN (SNBT/UTBK). Bahas strategi pengerjaan soal {materi} agar siswa siap tes."
        elif "Klinik" in program:
            konteks_khusus = f"Fokus laporan: Program remedial/klinik bagi siswa yang butuh tambahan jam. Tekankan pendekatan personal agar siswa paham materi {materi}."
        elif "KIR" in program:
            konteks_khusus = f"Fokus laporan: Bimbingan riset/karya ilmiah bidang {mapel}. Diskusikan progres penelitian berjudul '{materi}'."
        elif "Belajar Malam" in program:
            if "Belajar" in kategori_malam:
                konteks_khusus = "Fokus laporan: Mendampingi siswa belajar mandiri di asrama/kelas. Suasana tenang dan kondusif."
            else:
                konteks_khusus = f"Fokus laporan: Mendampingi kegiatan non-akademik malam hari yaitu {materi}. Pastikan kegiatan berjalan tertib."
        elif "Pengasuhan" in program:
            konteks_khusus = f"Fokus laporan: Sesi 'Jumat Curhat' atau pembinaan guru asuh. Masalah yang dibahas/diselesaikan: {topik}. Tekankan peran guru sebagai orang tua asuh yang memberi solusi dan penguatan mental."

        prompt = f"""
        Bertindaklah sebagai Guru di MAN Insan Cendekia Kendari. Buatkan 1 paragraf laporan kegiatan formal Bahasa Indonesia.
        
        Jenis Kegiatan: {program}
        Waktu: {tgl_teks}, Pukul {mulai}-{selesai}.
        Peserta: {peserta} siswa.
        
        Instruksi Khusus:
        {konteks_khusus}
        
        Output: Hanya isi paragraf laporannya saja. Jangan pakai judul. Gunakan bahasa baku yang santun.
        """

        # Menggunakan model flash-latest (paling aman & support API Key anda)
        response = client.models.generate_content(
            model="gemini-flash-latest", 
            contents=prompt
        )
        return response.text
        
    except Exception as e:
        return f"Gagal Generate. Error: {str(e)}"

# --- TOMBOL GENERATE ---
st.divider()
col_btn, col_res = st.columns([1, 2])

with col_btn:
    st.subheader("Langkah 1: Generate Teks")
    if st.button("✨ Buat Laporan Otomatis"):
        if api_key:
            with st.spinner("AI sedang menyusun laporan..."):
                # Kirim semua parameter ke fungsi AI (termasuk yang kosong tidak masalah)
                res = generate_description_ai(api_key, jenis_program, mapel, materi, topik_pengasuhan, kategori_malam, tanggal, jumlah_peserta, waktu_mulai, waktu_selesai)
                st.session_state['deskripsi'] = res
        else:
            st.error("Harap masukkan API Key terlebih dahulu.")

with col_res:
    deskripsi_final = st.text_area("Hasil Laporan:", value=st.session_state.get('deskripsi', ''), height=200)

# --- FUNGSI WORD (HEADER DINAMIS) ---
def create_docx(nama, program, tgl_obj, mapel_display, deskripsi, gambar):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # 1. Judul Header (NAMA PROGRAM & TAHUN OTOMATIS)
    tahun = tgl_obj.year
    # Membersihkan nama program agar rapi di judul (misal menghapus dalam kurung)
    judul_program = program.upper().split('(')[0].strip()
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"LAPORAN KEGIATAN\nPROGRAM {judul_program}\nMAN INSAN CENDEKIA KOTA KENDARI TAHUN {tahun}")
    run.bold = True
    run.font.size = Pt(14)
    
    doc.add_paragraph()
    
    # 2. Tabel Metadata
    tgl_indo_lengkap = tanggal_indo(tgl_obj).upper()
    
    table = doc.add_table(rows=3, cols=3)
    # Sesuaikan label baris ke-3 berdasarkan jenis program
    label_baris_3 = "MATERI/KEGIATAN"
    
    metadata = [
        ("NAMA PEMBINA", ":", nama),
        ("HARI/TANGGAL", ":", tgl_indo_lengkap),
        (label_baris_3, ":", mapel_display.upper()) # Mapel Display bisa berisi Mapel atau Topik
    ]
    
    for i, data in enumerate(metadata):
        r = table.rows[i]
        r.cells[0].text = data[0]
        r.cells[1].text = data[1]
        r.cells[2].text = data[2]
        r.cells[0].paragraphs[0].runs[0].bold = True 
        
    doc.add_paragraph() 

    # 3. Isi Laporan
    doc.add_paragraph("DESKRIPSI KEGIATAN").runs[0].bold = True
    doc.add_paragraph(deskripsi)
    doc.add_paragraph() 
    
    # 4. Dokumentasi
    doc.add_paragraph("DOKUMENTASI").runs[0].bold = True
    if gambar:
        doc.add_picture(gambar, width=Inches(6.0))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- TOMBOL DOWNLOAD ---
st.subheader("Langkah 2: Download File")

if st.button("💾 Download Dokumen (.docx)"):
    if deskripsi_final:
        # Menentukan apa yang ditampilkan di tabel metadata baris ke-3
        if "Pengasuhan" in jenis_program:
            isi_metadata = "PENGASUHAN SISWA"
        elif "Malam" in jenis_program:
            isi_metadata = f"{kategori_malam} ({materi})"
        elif "KIR" in jenis_program:
            isi_metadata = f"KIR {mapel}: {materi}"
        else:
            isi_metadata = f"{mapel}: {materi}"

        file_docx = create_docx(nama_pembahas, jenis_program, tanggal, isi_metadata, deskripsi_final, uploaded_file)
        
        file_name = f"Laporan_{jenis_program}_{tanggal}.docx"
        st.download_button("Klik untuk Unduh", file_docx, file_name, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.success(f"Laporan {jenis_program} berhasil dibuat!")
    else:
        st.error("Laporan belum digenerate.")
