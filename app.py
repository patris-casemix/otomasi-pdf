import streamlit as st
import os
import shutil
import pandas as pd
import re
from PyPDF2 import PdfMerger 
import fitz # PyMuPDF
import qrcode
from pathlib import Path
import io
import zipfile # <-- NEW IMPORT

# Tentukan lokasi kerja sementara
TEMP_DIR = "data"
os.makedirs(TEMP_DIR, exist_ok=True)

# ----------------------------------------------------------------------
# UTILITY FUNCTIONS
# ----------------------------------------------------------------------

def save_uploaded_file(uploaded_file, name=None):
    """Menyimpan objek file yang diunggah ke disk dan mengembalikan path."""
    file_name = name if name else uploaded_file.name
    # Pastikan nama file adalah string aman
    file_name = "".join([c if c.isalnum() or c in (' ', '.', '_', '-') else '_' for c in file_name])
    path = os.path.join(TEMP_DIR, file_name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

def save_excel_df(df, output_filename, output_dir):
    """Menyimpan DataFrame sebagai Excel di output folder dan mengembalikan path."""
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, output_filename)
    df.to_excel(path, index=False)
    return path

def zip_folder_and_download(folder_path, zip_filename):
    """Membuat file ZIP dari folder dan menyediakan tombol download Streamlit."""
    zip_path = os.path.join(TEMP_DIR, zip_filename)
    
    # Pastikan folder_path ada dan tidak kosong sebelum zipping
    if not os.path.exists(folder_path) or not os.listdir(folder_path):
        st.warning(f"⚠️ Folder hasil **'{os.path.basename(folder_path)}'** kosong. Tidak ada file PDF yang berhasil diproses/dihasilkan.")
        return
        
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Menambahkan semua file di dalam folder_path ke ZIP
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Tentukan nama file di dalam ZIP (menghapus nama folder_path agar tidak ada duplikasi)
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))

        # Tombol Unduh ZIP
        with open(zip_path, "rb") as file:
            st.download_button(
                label=f"⬇️ Unduh Semua File Hasil ({os.path.basename(folder_path)}.zip)",
                data=file,
                file_name=zip_filename,
                mime="application/zip"
            )
    except Exception as e:
        st.error(f"Gagal membuat file ZIP: {e}")

# ----------------------------------------------------------------------
# FUNGSI LOGIKA (PENGGABUNGAN)
# ----------------------------------------------------------------------

def run_pdf_merge_3_way_giant(uploaded_individu, uploaded_rawat, uploaded_billing):
    """Logika Gabung Raksasa 3-Arah: INDIVIDU + RAWAT JALAN + BILLING (gabung_pdf_giant_3.py)"""
    st.info("Memproses Penggabungan Raksasa 3-Arah (INDIVIDU + RJ + BILLING)...")
    
    individu_paths = {f.name: save_uploaded_file(f, name=f"I_{f.name}") for f in uploaded_individu}
    rawat_paths = {f.name: save_uploaded_file(f, name=f"RJ_{f.name}") for f in uploaded_rawat}
    billing_paths = {f.name: save_uploaded_file(f, name=f"BIL_{f.name}") for f in uploaded_billing}

    output_dir = os.path.join(TEMP_DIR, "Gabungan_3_Arah")
    os.makedirs(output_dir, exist_ok=True)
    
    gagal_gabung = []
    tidak_lengkap = []
    
    # Acuan adalah semua file yang diunggah
    file_set = set(individu_paths.keys()) | set(rawat_paths.keys()) | set(billing_paths.keys())
    
    for file_name in file_set:
        if file_name.lower().endswith('.pdf'):
            file_individu = individu_paths.get(file_name)
            file_rawat = rawat_paths.get(file_name)
            file_billing = billing_paths.get(file_name)
            file_output = os.path.join(output_dir, file_name)

            # Cek kelengkapan (HARUS ADA di KETIGA folder)
            if not (file_individu and file_rawat and file_billing):
                status = ""
                if not file_individu: status += "Tidak ada di INDIVIDU; "
                if not file_rawat: status += "Tidak ada di RAWAT JALAN; "
                if not file_billing: status += "Tidak ada di BILLING;"
                tidak_lengkap.append([file_name, status])
                continue

            # Gabungkan PDF sesuai urutan
            merger = PdfMerger()
            try:
                merger.append(file_individu) # Urutan 1
                merger.append(file_rawat)    # Urutan 2
                merger.append(file_billing)  # Urutan 3
                merger.write(file_output)
            except Exception as e:
                gagal_gabung.append([file_name, str(e)])
            finally:
                merger.close()

    st.success("Proses Penggabungan Selesai. Mohon cek log di bawah.")
    
    # Logik menyimpan log Excel
    if gagal_gabung:
        df_gagal = pd.DataFrame(gagal_gabung, columns=["Nama File", "Alasan Gagal Gabung"])
        log_gagal_path = save_excel_df(df_gagal, "log_gagal_gabung_3way.xlsx", output_dir)
        st.warning(f"❌ Ada {len(gagal_gabung)} file gagal digabung.")
        with open(log_gagal_path, "rb") as file:
            st.download_button("Unduh Log Gagal Gabung (Excel)", file, "log_gagal_gabung_3way.xlsx")
        
    if tidak_lengkap:
        df_tidak_lengkap = pd.DataFrame(tidak_lengkap, columns=["Nama File", "Keterangan Tidak Lengkap"])
        log_lengkap_path = save_excel_df(df_tidak_lengkap, "log_tidak_lengkap_3way.xlsx", output_dir)
        st.info(f"⚠️ Ada {len(tidak_lengkap)} file tidak lengkap.")
        with open(log_lengkap_path, "rb") as file:
            st.download_button("Unduh Log Tidak Lengkap (Excel)", file, "log_tidak_lengkap_3way.xlsx")
    
    st.info(f"Total Berhasil: {len(file_set) - len(gagal_gabung) - len(tidak_lengkap)} file.")
    zip_folder_and_download(output_dir, "Gabungan_3_Arah_Hasil.zip") # <-- PANGGILAN BARU

def run_pdf_merge_2_way_giant(uploaded_folder1, uploaded_folder2):
    """Logika Gabung Raksasa 2-Arah: Folder 1 + Folder 2 (gabung_pdf_giant_2.py)"""
    st.info("Memproses Penggabungan Raksasa 2-Arah (Folder 1 + Folder 2)...")

    folder1_paths = {f.name: save_uploaded_file(f, name=f"F1_{f.name}") for f in uploaded_folder1}
    folder2_paths = {f.name: save_uploaded_file(f, name=f"F2_{f.name}") for f in uploaded_folder2}

    output_dir = os.path.join(TEMP_DIR, "Gabungan_2_Arah_Giant")
    os.makedirs(output_dir, exist_ok=True)
    
    gagal_gabung = []
    tidak_lengkap = []
    
    file_set = set(folder1_paths.keys()) | set(folder2_paths.keys())
    
    for file_name in file_set:
        if file_name.lower().endswith('.pdf'):
            path1 = folder1_paths.get(file_name)
            path2 = folder2_paths.get(file_name)
            file_output = os.path.join(output_dir, file_name)

            # Cek kelengkapan
            if not (path1 and path2):
                status = f"Tidak ada di {'Folder 1' if not path1 else 'Folder 2'}"
                tidak_lengkap.append([file_name, status])
                continue

            # Gabungkan PDF sesuai urutan
            merger = PdfMerger()
            try:
                merger.append(path1)
                merger.append(path2)
                merger.write(file_output)
            except Exception as e:
                gagal_gabung.append([file_name, str(e)])
            finally:
                merger.close()

    st.success("Proses Penggabungan Selesai. Mohon cek log di bawah.")
    
    # Logik menyimpan log Excel
    if gagal_gabung:
        df_gagal = pd.DataFrame(gagal_gabung, columns=["Nama File", "Alasan Gagal Gabung"])
        log_gagal_path = save_excel_df(df_gagal, "log_gagal_gabung_2way_giant.xlsx", output_dir)
        st.warning(f"❌ Ada {len(gagal_gabung)} file gagal digabung.")
        with open(log_gagal_path, "rb") as file:
            st.download_button("Unduh Log Gagal Gabung (Excel)", file, "log_gagal_gabung_2way_giant.xlsx")
        
    if tidak_lengkap:
        df_tidak_lengkap = pd.DataFrame(tidak_lengkap, columns=["Nama File", "Keterangan Tidak Lengkap"])
        log_lengkap_path = save_excel_df(df_tidak_lengkap, "log_tidak_lengkap_2way_giant.xlsx", output_dir)
        st.info(f"⚠️ Ada {len(tidak_lengkap)} file tidak lengkap.")
        with open(log_lengkap_path, "rb") as file:
            st.download_button("Unduh Log Tidak Lengkap (Excel)", file, "log_tidak_lengkap_2way_giant.xlsx")
    
    st.info(f"Total Berhasil: {len(file_set) - len(gagal_gabung) - len(tidak_lengkap)} file.")
    zip_folder_and_download(output_dir, "Gabungan_2_Arah_Giant_Hasil.zip") # <-- PANGGILAN BARU


def run_pdf_merge_simple(uploaded_folder1, uploaded_folder2):
    """Logika Gabung Sederhana 2-Arah: Folder1 + Folder2 (merge_pdf.py)"""
    st.info("Memproses Penggabungan Sederhana (Folder 1 + Folder 2)...")

    folder1_paths = {f.name: save_uploaded_file(f, name=f"F1_{f.name}") for f in uploaded_folder1}
    folder2_paths = {f.name: save_uploaded_file(f, name=f"F2_{f.name}") for f in uploaded_folder2}

    output_dir = os.path.join(TEMP_DIR, "Gabungan_2_Arah_Simple")
    os.makedirs(output_dir, exist_ok=True)
    
    berhasil_gabung = 0
    gagal_gabung = []
    
    file_list_acuan = set(folder1_paths.keys())
    
    for file_name in file_list_acuan:
        if file_name.lower().endswith('.pdf'):
            path1 = folder1_paths.get(file_name)
            path2 = folder2_paths.get(file_name)

            if path1 and path2:
                output_path = os.path.join(output_dir, file_name)
                merger = PdfMerger()
                try:
                    merger.append(path1)
                    merger.append(path2)
                    merger.write(output_path)
                    merger.close()
                    berhasil_gabung += 1
                except Exception as e:
                    gagal_gabung.append([file_name, str(e)])
            else:
                gagal_gabung.append([file_name, "File tidak ditemukan di Folder 2"])

    st.success(f"🎉 Proses Selesai. Total file yang digabungkan: {berhasil_gabung}. Gagal: {len(gagal_gabung)}.")
    if gagal_gabung:
        df_gagal = pd.DataFrame(gagal_gabung, columns=["Nama File", "Alasan Gagal"])
        log_gagal_path = save_excel_df(df_gagal, "log_gagal_gabung_simple.xlsx", output_dir)
        with open(log_gagal_path, "rb") as file:
            st.download_button("Unduh Log Gagal Gabung (Excel)", file, "log_gagal_gabung_simple.xlsx")
            
    zip_folder_and_download(output_dir, "Gabungan_2_Arah_Simple_Hasil.zip") # <-- PANGGILAN BARU


# ----------------------------------------------------------------------
# FUNGSI LOGIKA (RENAME)
# ----------------------------------------------------------------------

def parse_sep_from_excel_list(df):
    """Logika pembersihan SEP dari skrip rename_by_Excel*"""
    sep_mapping = {}
    for val in df.iloc[:, 0].astype(str):
        parts = val.split(maxsplit=1)
        if len(parts) > 1:
            # Jika ada spasi, ambil bagian kedua (SEP)
            nomor_sep_clean = re.sub(r'\W+', '', parts[1])
        else:
            # Jika tidak ada spasi, ambil seluruhnya
            nomor_sep_clean = re.sub(r'\W+', '', parts[0])
        
        # Simpan mapping {Nomor SEP bersih: Nama Lengkap dari Excel}
        sep_mapping[nomor_sep_clean] = val.strip() 
    return sep_mapping

def run_pdf_rename_excel(uploaded_pdfs, uploaded_excel, search_mode="Page 1"):
    """Logika Rename PDF (SEP Halaman 1 atau Semua Halaman)"""
    st.info(f"Mulai Ganti Nama (SEP Mode: {search_mode})...")
    
    excel_path = save_uploaded_file(uploaded_excel)
    df = pd.read_excel(excel_path)
    sep_mapping = parse_sep_from_excel_list(df)
    
    output_dir = os.path.join(TEMP_DIR, f"Renamed_SEP_{search_mode.replace(' ', '_')}")
    os.makedirs(output_dir, exist_ok=True)
    
    log_ok = []
    log_fail = []
    
    for file_obj in uploaded_pdfs:
        pdf_path = save_uploaded_file(file_obj)
        file_name = file_obj.name
        
        try:
            doc = fitz.open(pdf_path)
            
            # --- Perbedaan Logika Inti (Halaman 1 vs Semua Halaman) ---
            if search_mode == "Page 1":
                text_search = doc[0].get_text() # Hanya halaman pertama
            else: # search_mode == "All Pages"
                text_search = ""
                for page in doc:
                    text_search += page.get_text() # Semua halaman
            # ---------------------------------------------------------

            match = re.search(r"Nomor\s*SEP\s*[:\-]?\s*([A-Za-z0-9]+)", text_search, re.IGNORECASE)
            
            if match:
                nomor_sep_raw = match.group(1).strip()
                nomor_sep_clean = re.sub(r'\W+', '', nomor_sep_raw)

                if nomor_sep_clean in sep_mapping:
                    new_name = f"{sep_mapping[nomor_sep_clean]}.pdf"
                    new_path = os.path.join(output_dir, new_name)
                    doc.save(new_path)
                    log_ok.append([file_name, new_name])
                else:
                    log_fail.append([file_name, f"SEP '{nomor_sep_clean}' tidak ada di Excel"])
            else:
                log_fail.append([file_name, "Nomor SEP tidak ditemukan"])

            doc.close()
        except Exception as e:
            log_fail.append([file_name, f"ERROR PDF: {e}"])
        
    st.success("Proses Penggantian Nama Selesai.")
    
    # Buat Log Excel
    df_ok = pd.DataFrame(log_ok, columns=["Nama File Asal", "Nama File Baru"])
    df_fail = pd.DataFrame(log_fail, columns=["Nama File Asal", "Alasan Gagal"])
    
    log_ok_path = save_excel_df(df_ok, f"log_rename_berhasil_SEP_{search_mode}.xlsx", output_dir)
    log_fail_path = save_excel_df(df_fail, f"log_rename_gagal_SEP_{search_mode}.xlsx", output_dir)
    
    st.info(f"Berhasil: {len(log_ok)} file. Gagal: {len(log_fail)} file.")
    with open(log_ok_path, "rb") as file: st.download_button("Unduh Log Berhasil Rename (Excel)", file, f"log_rename_berhasil_SEP_{search_mode}.xlsx")
    with open(log_fail_path, "rb") as file: st.download_button("Unduh Log Gagal Rename (Excel)", file, f"log_rename_gagal_SEP_{search_mode}.xlsx")

    zip_folder_and_download(output_dir, f"Renamed_SEP_{search_mode.replace(' ', '_')}_Hasil.zip") # <-- PANGGILAN BARU


def run_pdf_rename_strip_tail(uploaded_pdfs):
    """Logika Rename: Hapus Bagian Belakang Setelah Spasi Terakhir (Rename_Belakang_All.py)"""
    st.info("Mulai Ganti Nama (Hapus Bagian Belakang)...")
    
    output_dir = os.path.join(TEMP_DIR, "Renamed_Strip_Tail")
    os.makedirs(output_dir, exist_ok=True)
    
    log_ok = []

    for file_obj in uploaded_pdfs:
        file_name = file_obj.name
        nama_lama = os.path.splitext(file_name)[0]
        ekstensi = os.path.splitext(file_name)[1]
        
        if ' ' in nama_lama:
            nama_baru_base = nama_lama.rsplit(' ', 1)[0]
            nama_baru = nama_baru_base + ekstensi
        else:
            nama_baru = file_name
        
        new_path = os.path.join(output_dir, nama_baru)
        with open(new_path, "wb") as f:
            f.write(file_obj.getbuffer())
            
        log_ok.append([file_name, nama_baru])

    st.success("Proses Penggantian Nama Selesai.")
    
    df_ok = pd.DataFrame(log_ok, columns=["Nama File Asal", "Nama File Baru"])
    log_ok_path = save_excel_df(df_ok, "log_rename_berhasil_StripTail.xlsx", output_dir)
    
    st.info(f"Berhasil mengganti nama {len(log_ok)} file.")
    with open(log_ok_path, "rb") as file: st.download_button("Unduh Log Hasil Rename (Excel)", file, "log_rename_berhasil_StripTail.xlsx")
    
    zip_folder_and_download(output_dir, "Renamed_Strip_Tail_Hasil.zip") # <-- PANGGILAN BARU


# ----------------------------------------------------------------------
# FUNGSI LOGIKA (COPY/MOVE/FILTER)
# ----------------------------------------------------------------------

def run_pdf_copy_excel_list(uploaded_pdfs, uploaded_excel, target_folder_name):
    """Logika Salin/Filter: Mencocokkan file yang diunggah dengan daftar nama di Excel (copy_pdf_gui_baca_SubFolder.py)"""
    st.info("Mulai Penyalinan/Filter Berdasarkan Daftar Excel...")
    
    excel_path = save_uploaded_file(uploaded_excel)
    df = pd.read_excel(excel_path, sheet_name=None)
    
    # Cari sheet "Pending RJ" atau "Pending RI", jika tidak ada ambil sheet pertama
    sheet_name_found = [s for s in df.keys() if "Pending" in s]
    df_data = df.get(sheet_name_found[0], df[list(df.keys())[0]]) if sheet_name_found else df[list(df.keys())[0]]
    
    # Ambil kolom 0, tambahkan ".pdf"
    nama_file_list_excel = set(
        df_data.iloc[:, 0].astype(str).apply(lambda x: x.strip() + ".pdf").tolist()
    )

    output_dir = os.path.join(TEMP_DIR, target_folder_name)
    os.makedirs(output_dir, exist_ok=True)
    
    log_data = []
    berhasil = 0
    uploaded_pdf_names = {f.name: f for f in uploaded_pdfs}
    
    for nama_pdf_dari_excel in nama_file_list_excel:
        if nama_pdf_dari_excel in uploaded_pdf_names:
            file_obj = uploaded_pdf_names[nama_pdf_dari_excel]
            path_tujuan = os.path.join(output_dir, nama_pdf_dari_excel)
            
            with open(path_tujuan, "wb") as f:
                f.write(file_obj.getbuffer())
                
            berhasil += 1
            log_data.append([nama_pdf_dari_excel, "Berhasil", f"Disalin ke {target_folder_name}"])
        else:
            log_data.append([nama_pdf_dari_excel, "Gagal", "Tidak ditemukan di file yang diunggah"])

    st.success(f"Proses Selesai. Total file dicari: {len(nama_file_list_excel)}. Berhasil disalin: {berhasil}.")
    
    df_log = pd.DataFrame(log_data, columns=["Nama File", "Status", "Keterangan"])
    log_path = save_excel_df(df_log, "log_penyalinan.xlsx", output_dir)
    with open(log_path, "rb") as file: st.download_button("Unduh Log Penyalinan (Excel)", file, "log_penyalinan.xlsx")
    
    zip_folder_and_download(output_dir, "Hasil_Penyalinan_Excel.zip") # <-- PANGGILAN BARU


def run_pdf_move_list_excel(uploaded_pdfs, uploaded_excel, target_folder_name):
    """Logika Pindah: Memindahkan file yang terdaftar di Excel ke folder tujuan (pindah_pdf_gagal_purif_final.py)"""
    st.info("Mulai Pemindahan File Berdasarkan Daftar Excel...")

    excel_path = save_uploaded_file(uploaded_excel)
    df = pd.read_excel(excel_path, header=None) # Asumsi: tidak ada header
    daftar_dipindahkan = set(df.iloc[:, 0].astype(str).tolist())
    
    output_dir = os.path.join(TEMP_DIR, target_folder_name)
    os.makedirs(output_dir, exist_ok=True)
    
    log = []
    total_berhasil = 0

    # PENTING: Unggah file PDF harus disimpan terlebih dahulu agar bisa dipindahkan/dihapus
    uploaded_pdf_paths = {f.name: save_uploaded_file(f) for f in uploaded_pdfs}
    
    for nama_file in daftar_dipindahkan:
        path_asal_temp = uploaded_pdf_paths.get(nama_file)
        
        if path_asal_temp and os.path.exists(path_asal_temp):
            path_tujuan = os.path.join(output_dir, nama_file)
            
            try:
                # Pindahkan file dari folder TEMP_DIR ke folder output_dir
                shutil.move(path_asal_temp, path_tujuan)
                log.append([nama_file, "Berhasil", path_tujuan])
                total_berhasil += 1
            except Exception as e:
                log.append([nama_file, "Gagal", str(e)])
        else:
            log.append([nama_file, "Gagal", "File tidak ditemukan di daftar unggahan"])

    st.success(f"Proses Selesai. Total file dalam daftar: {len(daftar_dipindahkan)}. Berhasil dipindahkan: {total_berhasil}.")
    
    df_log = pd.DataFrame(log, columns=["Nama File", "Status", "Keterangan"])
    log_path = save_excel_df(df_log, "log_pemindahan_list.xlsx", output_dir)
    with open(log_path, "rb") as file: st.download_button("Unduh Log Pemindahan (Excel)", file, "log_pemindahan_list.xlsx")
    
    zip_folder_and_download(output_dir, "File_Pindah_Gagal.zip") # <-- PANGGILAN BARU

# ----------------------------------------------------------------------
# FUNGSI LOGIKA (QR CODE & TANDA TANGAN)
# ----------------------------------------------------------------------

def run_qr_code_generator(data_ttd, filename):
    """Logika Generator QR Code (qr_code_tandatangan.py)"""
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4)
    qr.add_data(data_ttd)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    
    st.success("QR Code berhasil dibuat.")
    st.image(img, caption="Hasil QR Code")
    st.download_button(
        label="Unduh Gambar QR Code (.png)",
        data=buf.getvalue(),
        file_name=filename,
        mime="image/png"
    )

def run_qr_code_inserter(uploaded_pdfs, uploaded_qr_img, teks_jangkar, offset_y, lebar_ttd):
    """Logika Sisip QR Code ke PDF (qr_code_ttd_dpjp.py / qr_code_ttd_kasir.py)"""
    st.info(f"Mulai menyisipkan QR Code di bawah teks: **'{teks_jangkar}'**...")

    qr_path = save_uploaded_file(uploaded_qr_img, name="input_qr_code.png")
    
    output_dir = os.path.join(TEMP_DIR, "PDF_dengan_TTD")
    os.makedirs(output_dir, exist_ok=True)
    
    log_data = []

    for file_obj in uploaded_pdfs:
        pdf_path_input = save_uploaded_file(file_obj)
        pdf_path_output = os.path.join(output_dir, file_obj.name)
        ditemukan = False
        
        try:
            doc = fitz.open(pdf_path_input)
            
            # Hanya sisipkan di halaman pertama
            page = doc[0] 
            text_instances = page.search_for(teks_jangkar)
            
            if text_instances:
                rect = text_instances[0]
                
                # Posisi di bawah teks jangkar
                posisi_x = rect.x0
                posisi_y = rect.y1 + offset_y
                
                rect_ttd = fitz.Rect(posisi_x, posisi_y, posisi_x + lebar_ttd, posisi_y + lebar_ttd)
                
                page.insert_image(rect_ttd, filename=qr_path)
                ditemukan = True
                log_data.append([file_obj.name, "Berhasil", f"Disisipkan di halaman 1"])
            
            if not ditemukan:
                log_data.append([file_obj.name, "Gagal", f"Teks '{teks_jangkar}' tidak ditemukan di Halaman 1"])

            if ditemukan:
                 doc.save(pdf_path_output)

            doc.close()

        except Exception as e:
            log_data.append([file_obj.name, "ERROR", str(e)])

    st.success("Proses Sisip QR Code Selesai.")
    
    df_log = pd.DataFrame(log_data, columns=["Nama File", "Status", "Keterangan"])
    log_path = save_excel_df(df_log, "log_sisip_qr.xlsx", output_dir)
    with open(log_path, "rb") as file: st.download_button("Unduh Log Sisip QR (Excel)", file, "log_sisip_qr.xlsx")
    st.info(f"Berhasil: {len([r for r in log_data if r[1] == 'Berhasil'])} file.")
    
    zip_folder_and_download(output_dir, "PDF_dengan_TTD_Hasil.zip") # <-- PANGGILAN BARU


# ----------------------------------------------------------------------
# TAMPILAN APLIKASI STREAMLIT (UI)
# ----------------------------------------------------------------------

st.set_page_config(page_title="Alat Otomasi PDF Kustom", layout="wide")
st.title("🗃️ Alat Otomasi PDF Kustom (Versi Web) - 10 Fungsi")
st.caption("Integrasi dari semua Skrip Python Anda.")

st.sidebar.header("Pilih Tugas Kustom")
selected_task = st.sidebar.radio(
    "Fungsi Skrip Anda:",
    [
        "1. Gabung PDF Raksasa (3 Folder)",
        "2. Gabung PDF Raksasa (2 Folder)",
        "3. Gabung PDF Sederhana (2 Folder)",
        "4. Ganti Nama (SEP Halaman 1 + Excel)",
        "5. Ganti Nama (SEP Semua Halaman + Excel)",
        "6. Ganti Nama (Hapus Kata Belakang)",
        "7. Filter/Salin PDF (Daftar Excel)",
        "8. Pindahkan PDF (Daftar Gagal/Tidak Dibutuhkan)",
        "9. Buat QR Code Tanda Tangan",
        "10. Sisipkan QR Code ke PDF (Teks Jangkar)",
    ]
)

st.divider()

if selected_task == "1. Gabung PDF Raksasa (3 Folder)":
    st.header("1. Gabungkan PDF Raksasa (INDIVIDU + RJ + BILLING)")
    st.warning("Unggah SEMUA file dari setiap folder. File akan dicocokkan berdasarkan **Nama File** yang sama.")

    col1, col2, col3 = st.columns(3)
    
    with col1: uploaded_individu = st.file_uploader("Unggah PDF dari Folder INDIVIDU:", type="pdf", accept_multiple_files=True, key="i")
    with col2: uploaded_rawat = st.file_uploader("Unggah PDF dari Folder RAWAT JALAN:", type="pdf", accept_multiple_files=True, key="rj")
    with col3: uploaded_billing = st.file_uploader("Unggah PDF dari Folder BILLING:", type="pdf", accept_multiple_files=True, key="bil")

    if st.button("Jalankan Penggabungan 3-Arah") and uploaded_individu and uploaded_rawat and uploaded_billing:
        run_pdf_merge_3_way_giant(uploaded_individu, uploaded_rawat, uploaded_billing)

elif selected_task == "2. Gabung PDF Raksasa (2 Folder)":
    st.header("2. Gabungkan PDF Raksasa (Folder 1 + Folder 2) dengan Log")
    st.warning("Unggah SEMUA file dari setiap folder. File yang tidak lengkap di salah satu folder akan tercatat di Log.")

    col1, col2 = st.columns(2)
    
    with col1: uploaded_folder1 = st.file_uploader("Unggah PDF dari Folder 1 (Urutan Awal):", type="pdf", accept_multiple_files=True, key="g2_f1")
    with col2: uploaded_folder2 = st.file_uploader("Unggah PDF dari Folder 2 (Urutan Kedua):", type="pdf", accept_multiple_files=True, key="g2_f2")

    if st.button("Jalankan Penggabungan Raksasa 2-Arah") and uploaded_folder1 and uploaded_folder2:
        run_pdf_merge_2_way_giant(uploaded_folder1, uploaded_folder2)

elif selected_task == "3. Gabung PDF Sederhana (2 Folder)":
    st.header("3. Gabungkan PDF Sederhana (Folder 1 + Folder 2)")
    st.info("File akan dicocokkan berdasarkan Nama File, dan hanya file yang ada di **Folder 1** yang menjadi acuan.")

    col1, col2 = st.columns(2)
    
    with col1: uploaded_folder1 = st.file_uploader("Unggah PDF dari Folder 1 (Acuan & Urutan Awal):", type="pdf", accept_multiple_files=True, key="s2_f1")
    with col2: uploaded_folder2 = st.file_uploader("Unggah PDF dari Folder 2 (Urutan Kedua):", type="pdf", accept_multiple_files=True, key="s2_f2")

    if st.button("Jalankan Penggabungan Sederhana 2-Arah") and uploaded_folder1 and uploaded_folder2:
        run_pdf_merge_simple(uploaded_folder1, uploaded_folder2)

elif selected_task == "4. Ganti Nama (SEP Halaman 1 + Excel)":
    st.header("4. Ganti Nama PDF (Berdasarkan Nomor SEP di **Halaman 1** & Mapping Excel)")
    st.info("Mode cepat: hanya memeriksa Nomor SEP di halaman pertama.")
    
    uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF yang akan diganti namanya:", type="pdf", accept_multiple_files=True, key="r1_pdfs")
    uploaded_excel = st.file_uploader("Unggah File Excel Daftar Mapping (Kolom 1 berisi Nama File Baru):", type=["xlsx"], accept_multiple_files=False, key="r1_excel")

    if st.button("Jalankan Ganti Nama SEP Halaman 1") and uploaded_pdfs and uploaded_excel:
        run_pdf_rename_excel(uploaded_pdfs, uploaded_excel, search_mode="Page 1")

elif selected_task == "5. Ganti Nama (SEP Semua Halaman + Excel)":
    st.header("5. Ganti Nama PDF (Berdasarkan Nomor SEP di **Semua Halaman** & Mapping Excel)")
    st.warning("Mode lambat: mencari Nomor SEP di seluruh dokumen. Gunakan jika Nomor SEP tidak selalu di halaman pertama.")
    
    uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF yang akan diganti namanya:", type="pdf", accept_multiple_files=True, key="r2_pdfs")
    uploaded_excel = st.file_uploader("Unggah File Excel Daftar Mapping (Kolom 1 berisi Nama File Baru):", type=["xlsx"], accept_multiple_files=False, key="r2_excel")

    if st.button("Jalankan Ganti Nama SEP Semua Halaman") and uploaded_pdfs and uploaded_excel:
        run_pdf_rename_excel(uploaded_pdfs, uploaded_excel, search_mode="All Pages")

elif selected_task == "6. Ganti Nama (Hapus Kata Belakang)":
    st.header("6. Ganti Nama PDF (Hapus Bagian Belakang Setelah Spasi Terakhir)")
    st.markdown("Contoh: `Nama File Belakang 123.pdf` akan diubah menjadi `Nama File.pdf`")
    
    uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF yang ingin dibersihkan namanya:", type="pdf", accept_multiple_files=True, key="r3_pdfs")

    if st.button("Jalankan Ganti Nama Belakang") and uploaded_pdfs:
        run_pdf_rename_strip_tail(uploaded_pdfs)

elif selected_task == "7. Filter/Salin PDF (Daftar Excel)":
    st.header("7. Filter/Salin PDF Berdasarkan Daftar Nama File di Excel")
    st.info("Hanya PDF yang ada di daftar Excel yang akan disalin.")

    uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF yang akan difilter/disalin:", type="pdf", accept_multiple_files=True, key="c1_pdfs")
    uploaded_excel = st.file_uploader("Unggah File Excel Daftar Nama File yang Dicari:", type=["xlsx"], accept_multiple_files=False, key="c1_excel")
    target_folder_name = st.text_input("Nama Folder Tujuan untuk Hasil Salin:", "Hasil_Penyalinan_Excel")

    if st.button("Jalankan Penyalinan/Filter") and uploaded_pdfs and uploaded_excel:
        run_pdf_copy_excel_list(uploaded_pdfs, uploaded_excel, target_folder_name)

elif selected_task == "8. Pindahkan PDF (Daftar Gagal/Tidak Dibutuhkan)":
    st.header("8. Pindahkan PDF Berdasarkan Daftar Nama File di Excel")
    st.info("PDF yang namanya persis tercantum di **Kolom A** Excel akan dipindahkan.")

    uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF yang akan dipindahkan (sumber):", type="pdf", accept_multiple_files=True, key="m1_pdfs")
    uploaded_excel = st.file_uploader("Unggah File Excel Daftar Nama File yang Akan Dipindahkan:", type=["xlsx"], accept_multiple_files=False, key="m1_excel")
    target_folder_name = st.text_input("Nama Folder Tujuan untuk File yang Dipindahkan:", "File_Pindah_Gagal")

    if st.button("Jalankan Pemindahan List") and uploaded_pdfs and uploaded_excel:
        run_pdf_move_list_excel(uploaded_pdfs, uploaded_excel, target_folder_name)

elif selected_task == "9. Buat QR Code Tanda Tangan":
    st.header("9. Generator QR Code Tanda Tangan")

    data_ttd = st.text_area("Masukkan Data Tanda Tangan (Nama, Jabatan, Instansi, dll.):", 
                            "Nama: [Nama Anda]\nJabatan: [Jabatan Anda]\nInstansi: [Nama Instansi]")
    
    filename_qr = st.text_input("Nama File Gambar Output (.png):", "qr_tandatangan_saya.png")

    if st.button("Buat QR Code") and data_ttd:
        run_qr_code_generator(data_ttd, filename_qr)
        
elif selected_task == "10. Sisipkan QR Code ke PDF (Teks Jangkar)":
    st.header("10. Sisipkan QR Code ke PDF (Cari Teks Jangkar)")
    st.warning("Mode ini hanya akan mencari teks jangkar di **Halaman 1** untuk penyisipan.")

    col1, col2 = st.columns(2)
    
    with col1:
        uploaded_pdfs = st.file_uploader("Unggah SEMUA file PDF Target:", type="pdf", accept_multiple_files=True, key="q1_pdfs")
        uploaded_qr_img = st.file_uploader("Unggah Gambar QR Code Tanda Tangan (.png):", type=["png"], accept_multiple_files=False, key="q1_qr")
    
    with col2:
        teks_jangkar = st.text_input("Teks Jangkar (Teks di atas lokasi tanda tangan, misal: Kolektor/Kasir):", "Dokter Penanggung jawab Pelayanan")
        offset_y = st.number_input("Offset Y (Jarak vertikal dari Teks Jangkar, dalam pixel):", value=0, step=10)
        lebar_ttd = st.number_input("Lebar Tanda Tangan (dalam pixel, misal: 42 atau 50):", value=50, min_value=1)

    if st.button("Jalankan Sisip QR Code") and uploaded_pdfs and uploaded_qr_img:
        run_qr_code_inserter(uploaded_pdfs, uploaded_qr_img, teks_jangkar, offset_y, lebar_ttd)