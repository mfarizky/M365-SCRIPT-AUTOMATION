import os
import csv
import re
from collections import Counter

# =====================================================================
# KONFIGURASI PATH
# =====================================================================
current_dir = os.path.dirname(os.path.abspath(__file__))

# File input (Hasil ekstraksi bersih sebelumnya)
input_csv = os.path.join(current_dir, "hasil_ekstraksi_bersih.csv")

# File output hasil perhitungan statistik KBLI
output_stats_csv = os.path.join(current_dir, "Statistik_KBLI.csv")
# =====================================================================

list_semua_kbli = []

print(f"Membaca data dari: {input_csv}...")

try:
    with open(input_csv, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file, delimiter=';')
        
        for row in reader:
            maksud_tujuan = row.get("MAKSUD DAN TUJUAN PERUSAHAAN", "").strip()
            
            if maksud_tujuan and "tidak ditemukan" not in maksud_tujuan.lower():
                # Pisahkan KBLI jika dalam satu perusahaan terdaftar beberapa kode sekaligus
                kbli_split = [k.strip() for k in maksud_tujuan.split(',')]
                
                for kbli in kbli_split:
                    if kbli:
                        # Normalisasi spasi ganda agar teks seragam saat dihitung
                        kbli_bersih = " ".join(kbli.split())
                        list_semua_kbli.append(kbli_bersih)

    # Hitung total kemunculan tiap KBLI
    counter_kbli = Counter(list_semua_kbli)
    
    # Urutkan dari yang terbanyak (most_common otomatis mengurutkan secara descending)
    kbli_terurut = counter_kbli.most_common()

    # 1. Cetak hasil rapi di Terminal
    print("\n" + "="*80)
    print(f"{'KBLI / BIDANG USAHA':<65} | {'JUMLAH':<6}")
    print("="*80)
    for kbli, jumlah in kbli_terurut:
        print(f"{kbli:<65} | {jumlah:<6}")
    print("="*80)

    # 2. Simpan hasil perhitungan ke file CSV Baru (KBLI; JUMLAH)
    with open(output_stats_csv, mode='w', newline='', encoding='utf-8') as out_file:
        writer = csv.writer(out_file, delimiter=';')
        writer.writerow(["KBLI", "JUMLAH"])
        writer.writerows(kbli_terurut)
        
    print(f"\nSukses! File rangkuman statistik KBLI berhasil disimpan di:\n{output_stats_csv}")

except FileNotFoundError:
    print(f"\n[ERROR] File '{input_csv}' tidak ditemukan. Pastikan sudah menjalankan skrip ekstraksi utama.")
except Exception as e:
    print(f"\n[ERROR] Terjadi kendala saat menghitung data: {e}")