import os
import csv
import re
import fitz  # PyMuPDF
import pytesseract
import cv2
import numpy as np
from settings import TESSERACT_CMD, ROOT_FOLDER, DAFTAR_PT

# =====================================================================
# KONFIGURASI PATH & DAFTAR PT
# =====================================================================
# Path & daftar perusahaan dibaca dari settings.py (.env + companies.local.txt)
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
root_folder = ROOT_FOLDER

current_dir = os.path.dirname(os.path.abspath(__file__))
output_csv = os.path.join(current_dir, "hasil_ekstraksi_akta_cepat.csv")
log_gagal_csv = os.path.join(current_dir, "log_gagal_ekstraksi.csv")

# =====================================================================

filter_file_regex = re.compile(r"(akta|\bap\b)", re.IGNORECASE)
kbli_regex = re.compile(r"(\(\d{5}\)\s*[A-Za-z\s\-,\/]+|[A-Za-z\s\-,\/]+\(\d{5}\))", re.IGNORECASE)

total_ditemukan = 0
total_berhasil = 0
total_gagal = 0

print(f"Memulai pemindaian terfokus pada folder PT di: {root_folder}\n" + "="*70)

def tulis_log_gagal(nama_pt, file_name, alasan):
    with open(log_gagal_csv, mode='a', newline='', encoding='utf-8') as f_gagal:
        writer_gagal = csv.writer(f_gagal, delimiter=';')
        writer_gagal.writerow([nama_pt, file_name, alasan])

# Reset file CSV
with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter=';')
    writer.writerow(["NAMA PT", "MAKSUD DAN TUJUAN PERUSAHAAN"])

with open(log_gagal_csv, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter=';')
    writer.writerow(["NAMA PT", "NAMA FILE", "ALASAN GAGAL/LEWAT"])

try:
    semua_subfolder = [d for d in os.listdir(root_folder) if os.path.isdir(os.path.join(root_folder, d))]
except Exception as e:
    print(f"Gagal mengakses root folder: {e}")
    semua_subfolder = []

for pt_target in DAFTAR_PT:
    print(f"\nMencari Dokumen untuk: {pt_target}")
    
    pt_normalized = re.sub(r'\s+', '', pt_target.upper().replace('\xa0', ''))
    pt_clean = re.sub(r'[^A-Z0-9]', '', pt_normalized)
    
    # Antisipasi jika ada perbedaan ejaan kata "INVESMENTS" vs "INVESTMENTS"
    pt_fallback_clean = pt_clean.replace("INVESMENTS", "INVESTMENTS").replace("INVESTMENTS", "INVESMENTS")
    
    folder_pt_ditemukan = None
    for folder in semua_subfolder:
        folder_normalized = re.sub(r'\s+', '', folder.upper().replace('\xa0', ''))
        folder_clean = re.sub(r'[^A-Z0-9]', '', folder_normalized)
        if (pt_clean in folder_clean) or (pt_fallback_clean in folder_clean) or (folder_clean in pt_clean):
            folder_pt_ditemukan = os.path.join(root_folder, folder)
            break
            
    if not folder_pt_ditemukan:
        pesan = "Folder PT tidak ditemukan di root directory"
        print(f"  -> [LEWATKAN] {pesan}")
        tulis_log_gagal(pt_target, "-", pesan)
        total_gagal += 1
        continue

    file_target = []
    for root, dirs, files in os.walk(folder_pt_ditemukan):
        for f in files:
            if f.lower().endswith('.pdf') and filter_file_regex.search(f):
                file_target.append(os.path.join(root, f))
                
    if not file_target:
        pesan = "Tidak ada file PDF mengandung nama 'AKTA' atau 'AP'"
        print(f"  -> [LEWATKAN] {pesan}")
        tulis_log_gagal(pt_target, "-", pesan)
        total_gagal += 1
        continue

    berhasil_ekstrak_pt_ini = False
    alasan_log_terakhir = "Gagal memproses seluruh berkas PDF yang ada"
    file_terakhir_diproses = "-"

    for file_path in file_target:
        file_name = os.path.basename(file_path)
        file_terakhir_diproses = file_name
        print(f"  -> Memproses File: {file_name}")
        total_ditemukan += 1
        
        raw_text = ""
        maksud_tujuan_list = []
        
        try:
            doc = fitz.open(file_path)
            for page in doc:
                page_text = page.get_text()
                if page_text:
                    raw_text += page_text + "\n"
            
            # Jika PDF berupa pindaian/scan, jalankan OpenCV + Tesseract
            if len(raw_text.strip()) < 100:
                print(f"     -> Terdeteksi PDF SCAN. Menjalankan Pemuatan Gambar OpenCV + Tesseract...")
                raw_text = ""  
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    
                    # 1. Mengubah objek bytes PyMuPDF menjadi array numpy
                    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.height, pix.width, pix.n))
                    
                    # 2. Mengubah format warna karena PyMuPDF menghasilkan RGB/RGBA sedangkan OpenCV menggunakan BGR
                    if pix.n == 4:
                        img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
                    else:
                        img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
                    
                    # 3. [Opsional] Mengubah gambar menjadi grayscale agar pembacaan OCR Tesseract jauh lebih enteng & akurat
                    img_gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
                    
                    # Jalankan OCR langsung dari array memori gambar OpenCV
                    raw_text += pytesseract.image_to_string(img_gray, lang='ind') + "\n"
            
            def ekstrak_kbli(teks):
                hasil = []
                matches = kbli_regex.findall(teks)
                for match in matches:
                    bidang_usaha = " ".join(match.split())
                    bidang_usaha = re.sub(r'-{2,}', '', bidang_usaha)
                    bidang_usaha = re.sub(r'.*?(?:dalam bidang usaha|dalam bidang|menjalankan usaha)\s+', '', bidang_usaha, flags=re.IGNORECASE)
                    bidang_usaha = re.sub(r'^adalah\s+', '', bidang_usaha, flags=re.IGNORECASE)
                    bidang_usaha = re.sub(r'\s*,\s*', ', ', bidang_usaha).strip()
                    if bidang_usaha and bidang_usaha not in hasil:
                        hasil.append(bidang_usaha)
                return hasil

            start_match = re.search(r"(Pasal\s*3|MAKSUD\s*DAN\s*TUJUAN)", raw_text, re.IGNORECASE)
            end_match = re.search(r"(Pasal\s*4|MODAL)", raw_text, re.IGNORECASE)
            
            if start_match and end_match and start_match.start() < end_match.start():
                teks_target = raw_text[start_match.start():end_match.start()]
                maksud_tujuan_list = ekstrak_kbli(teks_target)
            
            if not maksud_tujuan_list:
                maksud_tujuan_list = ekstrak_kbli(raw_text)
            
            maksud_tujuan_final = ", ".join(maksud_tujuan_list)
            
            if intent_clean := maksud_tujuan_final.strip():
                with open(output_csv, mode='a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file, delimiter=';')
                    writer.writerow([pt_target, intent_clean])
                print(f"     -> [BERHASIL] Ditulis ke CSV: {intent_clean[:60]}...")
                total_berhasil += 1
                berhasil_ekstrak_pt_ini = True
                break  
            else:
                alasan_log_terakhir = "Data kode KBLI 5 digit tidak ditemukan di dalam teks (Kosong)"
                print(f"     -> {alasan_log_terakhir}")
                
        except Exception as e:
            alasan_log_terakhir = f"Error pembacaan: {str(e)}"
            print(f"     -> {alasan_log_terakhir}")
        finally:
            if 'doc' in locals():
                doc.close()

    if not berhasil_ekstrak_pt_ini:
        tulis_log_gagal(pt_target, file_terakhir_diproses, alasan_log_terakhir)
        total_gagal += 1

# =====================================================================
# LAPORAN STATISTIK AKHIR
# =====================================================================
print("\n" + "="*70)
print("PROSES SELESAI - RE-ORDERED SYSTEM OPENCV MEMORY FIX")
print("="*70)
print(f"Total Folder PT Sukses   : {total_berhasil}")
print(f"Total Folder PT Gagal    : {total_gagal}")
print(f"Total Keseluruhan (54)   : {total_berhasil + total_gagal}")
print(f"Hasil Sukses disimpan di : {output_csv}")
print(f"Log Gagal disimpan di    : {log_gagal_csv}")
print("="*70)