import os
import csv
import re
import pypdf
import fitz  # PyMuPDF
import pytesseract
from settings import TESSERACT_CMD, ROOT_FOLDER, DAFTAR_PT

# =====================================================================
# KONFIGURASI PATH & DAFTAR PT
# =====================================================================
# Path & daftar perusahaan dibaca dari settings.py (.env + companies.local.txt)
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
root_folder = ROOT_FOLDER

current_dir = os.path.dirname(os.path.abspath(__file__))
output_csv = os.path.join(current_dir, "hasil_ekstraksi_akta_legal.csv")

# =====================================================================

# Regex Patterns
filter_file_regex = re.compile(r"(akta\s*pendirian|\bap\b)", re.IGNORECASE)
kbli_regex = re.compile(r"(?:dalam bidang usaha|dalam bidang)\s+(.+?\(\d{5}\)(?:\s*,\s*.+?\(\d{5}\))*)", re.IGNORECASE)

# Inisialisasi variabel statistik
total_ditemukan = 0
total_berhasil = 0
total_gagal = 0

print(f"Memulai pemindaian total pada: {root_folder}")
print("Mode: Ekstrak Teks sampai Kode KBLI 5 Digit -> Hanya Hasil Sukses\n" + "="*70)

# Buat file baru dengan header kolom
with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter=';')
    writer.writerow(["NAMA PT", "MAKSUD DAN TUJUAN PERUSAHAAN"])

for dirpath, dirnames, filenames in os.walk(root_folder):
    for file_name in filenames:
        if file_name.lower().endswith(".pdf") and filter_file_regex.search(file_name):
            
            # Strategi Pencocokan Nama PT dari Daftar Resmi
            nama_pt = None
            full_path_string = os.path.join(dirpath, file_name).upper()
            
            for pt in DAFTAR_PT:
                # Menghapus variasi titik/spasi ("PT." atau "PT ") saat membandingkan
                pt_clean = re.sub(r'[^A-Z0-9]', '', pt.upper())
                path_clean = re.sub(r'[^A-Z0-9]', '', full_path_string)
                
                if pt_clean in path_clean:
                    nama_pt = pt
                    break
            
            # Jika tidak ada di daftar, pakai fallback bawaan awal Anda
            if not nama_pt:
                nama_pt = os.path.splitext(file_name)[0]
                nama_pt = re.sub(r'^SALINAN\s+AKTA\s+', '', nama_pt, flags=re.IGNORECASE)
            
            total_ditemukan += 1
            file_path = os.path.join(dirpath, file_name)
            rel_path = os.path.relpath(dirpath, root_folder)
            print(f"Memproses: {file_name} -> Terdeteksi: {nama_pt} [{rel_path}]")
            
            raw_text = ""
            maksud_tujuan_list = []
            
            try:
                # STRATEGI 1: BACA TEKS DIGITAL DULU
                with open(file_path, 'rb') as f:
                    reader = pypdf.PdfReader(f)
                    for page in reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            raw_text += page_text + "\n"
                
                # STRATEGI 2: JIKA SCAN, PAKAI OCR
                if len(raw_text.strip()) < 100:
                    print(f"  -> Terdeteksi sebagai PDF SCAN. Menjalankan OCR...")
                    doc = fitz.open(file_path)
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_data = pix.tobytes("png")
                        raw_text += pytesseract.image_to_string(img_data, lang='ind') + "\n"
                
                # PROSES SLICING (PASAL 3 SAMPAI PASAL 4)
                start_match = re.search(r"(Pasal\s*3|MAKSUD\s*DAN\s*TUJUAN)", raw_text, re.IGNORECASE)
                end_match = re.search(r"(Pasal\s*4|MODAL)", raw_text, re.IGNORECASE)
                
                if start_match and end_match and start_match.start() < end_match.start():
                    teks_target = raw_text[start_match.start():end_match.start()]
                else:
                    teks_target = raw_text
                
                matches = kbli_regex.findall(teks_target)
                
                for match in matches:
                    bidang_usaha = " ".join(match.split())
                    bidang_usaha = re.sub(r'\s*,\s*', ', ', bidang_usaha)
                    
                    if bidang_usaha and bidang_usaha not in maksud_tujuan_list:
                        maksud_tujuan_list.append(bidang_usaha)
                
                maksud_tujuan_final = ", ".join(maksud_tujuan_list)
                
                # HANYA TULIS JIKA BERHASIL
                if intent_clean := maksud_tujuan_final.strip():
                    with open(output_csv, mode='a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file, delimiter=';')
                        writer.writerow([nama_pt, intent_clean])
                    print(f"  -> [BERHASIL] Ditulis ke CSV: {intent_clean[:60]}...")
                    total_berhasil += 1
                else:
                    print(f"  -> [LEWATKAN] Data bidang usaha kosong.")
                    total_gagal += 1
                
            except Exception as e:
                print(f"  -> [LEWATKAN] Error: {e}")
                total_gagal += 1

# =====================================================================
# BLOK STATISTIK AKHIR
# =====================================================================
print("\n" + "="*70)
print("PROSES SELESAI - LAPORAN STATISTIK PEMINDAIAN")
print("="*70)
print(f"Total File PDF Ditemukan  : {total_ditemukan}")
print(f"Berhasil Diekstraksi      : {total_berhasil}")
print(f"Gagal / Dilewatkan        : {total_gagal}")
print(f"Hasil disimpan di         : {output_csv}")
print("="*70)