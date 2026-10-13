import os
import csv
import re

# Nama file CSV input dan output
input_csv = "hasil_ekstraksi_akta_cepat.csv"
output_csv = "hasil_ekstraksi_bersih.csv"

def clean_pt_name(name):
    """Menghilangkan kata kunci akta pendirian dan merapikan format menjadi PT. NAMA"""
    name = str(name).strip()
    
    # Hapus prefix/awalan yang tidak diperlukan
    prefix_patt = r"^(AKTA\s+PENDIRIAN\s+PERSEROAN\s+TERBATAS|AKTA\s+PENDIRIANN|AKTA\s+PENDIRIAN|PENDIRIAN|[A-Z]{2,5}-AP\[-|[A-Z]{2,5}-AP\[|4\.\s*AKTA\s+PENDIRIAN)\s*"
    name = re.sub(prefix_patt, "", name, flags=re.IGNORECASE)
    
    # Bersihkan tanda kurung atau bracket di ujung teks
    name = name.strip("()[] ")
    
    # Normalkan spasi berlebih
    name = " ".join(name.split())
    
    # Bersihkan tanda baca di ujung yang tersisa
    name = name.strip(") ")
    
    # Ubah format awal agar seragam menjadi "PT. NAMA"
    if name.upper().startswith("PT"):
        name = re.sub(r"^PT[\s\.]*", "PT. ", name, flags=re.IGNORECASE)
    else:
        name = "PT. " + name
        
    # Merapikan spasi ganda setelah tanda titik PT.
    name = re.sub(r"^PT\.\s*", "PT. ", name)
    return name.strip()

def clean_maksud_tujuan(text):
    """Mengekstrak hanya nama bidang usaha dan kode KBLI (e.g., Nama KBLI (55110))"""
    text = str(text).strip()
    
    # Cari semua pola yang diakhiri oleh kode angka 4-5 digit di dalam kurung
    kbli_matches = re.findall(r"([^,;\-_–—]+?\(\d{4,5}\))", text)
    
    if not kbli_matches:
        return "Kata kunci bidang usaha tidak ditemukan"
        
    cleaned_items = []
    for item in kbli_matches:
        # Bersihkan noise kata awalan struktural akta
        item = re.sub(r"^(usaha\s+Penyediaan|Penyediaan\s+usaha|usaha\s+|Penyediaan\s+)", "", item, flags=re.IGNORECASE)
        
        # Bersihkan noise jika ada pecahan teks pasal masuk ke regex
        item = re.sub(r"^\d+\.\s*Maksud\s*dan\s*tujuan.*?(bidang|usaha)\s+", "", item, flags=re.IGNORECASE)
        
        # Rapikan spasi
        item = " ".join(item.split())
        item = item.strip(";:, ")
        
        if item and item not in cleaned_items:
            cleaned_items.append(item)
            
    return ", ".join(cleaned_items)

print("Memulai proses pembersihan data CSV...")

if not os.path.exists(input_csv):
    print(f"Error: File '{input_csv}' tidak ditemukan. Pastikan file berada di folder yang sama.")
else:
    data_bersih = []
    
    # 1. Baca data mentah
    with open(input_csv, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file, delimiter=';')
        
        for row in reader:
            nama_mentah = row.get("NAMA PT", "")
            tujuan_mentah = row.get("MAKSUD DAN TUJUAN PERUSAHAAN", "")
            
            # Jalankan fungsi pembersihan
            nama_pt_bersih = clean_pt_name(nama_mentah)
            tujuan_bersih = clean_maksud_tujuan(tujuan_mentah)
            
            data_bersih.append([nama_pt_bersih, tujuan_bersih])
            
    # 2. Tulis ke file CSV baru yang rapi
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(["NAMA PT", "MAKSUD DAN TUJUAN PERUSAHAAN"])
        writer.writerows(data_bersih)
        
    print(f"Selesai! File bersih dan rapi sukses dibuat: {os.path.abspath(output_csv)}")