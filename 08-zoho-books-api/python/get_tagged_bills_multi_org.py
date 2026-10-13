import json
import os
import requests

from auth_helper import get_new_access_token

# === KONFIGURASI MULTI ORGANISASI (INPUT ARRAY/LIST) ===
# Daftar Organization ID dibaca dari Auth/.env (ZOHO_ORG_IDS=111,222,333)
TARGET_ORG_IDS = [x.strip() for x in os.getenv("ZOHO_ORG_IDS", "").split(",") if x.strip()]
# ========================================================

current_working_dir = os.getcwd()
output_dir = os.path.join(current_working_dir, "Output")

def fetch_and_filter_bills(access_token, org_id):
    # Pastikan folder output sudah terbuat
    os.makedirs(output_dir, exist_ok=True)
    
    # Nama file dinamis berdasarkan masing-masing ORG ID
    output_json_file = os.path.join(output_dir, f"filtered_bills_{org_id}.json")
    
    url = "https://www.zohoapis.com/books/v3/bills"
    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
        "Content-Type": "application/json"
    }
    
    filtered_bills = []
    current_page = 1
    has_more_page = True
    total_raw_bills = 0
    
    print(f"\n→ [ORG {org_id}] Mulai menarik data dan menyaring Bills yang memiliki Tags...")
    
    while has_more_page:
        params = {
            "organization_id": org_id,
            "date_start": "2026-01-01",
            "date_end": "2026-12-31",
            "per_page": 200, 
            "page": current_page  
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            res_json = response.json()
            page_bills = res_json.get("bills", [])
            total_raw_bills += len(page_bills)
            
            # --- PROSES FILTER TAGS NOT EMPTY ---
            for bill in page_bills:
                tags = bill.get("tags", [])
                
                # Cek apakah tags berbentuk list dan memiliki isi (tidak kosong)
                if isinstance(tags, list) and len(tags) > 0:
                    filtered_bills.append(bill)
            # -------------------------------------
            
            page_context = res_json.get("page_context", {})
            has_more_page = page_context.get("has_more_page", False)
            current_page += 1
        else:
            print(f"❌ [ORG {org_id}] Gagal pada halaman {current_page}: {response.status_code} - {response.text}")
            break
            
    # Simpan hasil filter ke file JSON per Org ID
    if filtered_bills:
        output_data = {
            "code": 0,
            "message": "success",
            "organization_id": org_id,
            "total_scanned": total_raw_bills,
            "total_filtered": len(filtered_bills),
            "bills": filtered_bills
        }
        
        with open(output_json_file, mode="w", encoding="utf-8") as file:
            json.dump(output_data, file, indent=4, ensure_ascii=False)
            
        print(f"✔ [ORG {org_id}] Selesai!")
        print(f"   Total Bills discan      : {total_raw_bills}")
        print(f"   Total Bills dengan Tags : {len(filtered_bills)}")
        print(f"   File disimpan ke        : '{output_json_file}'")
    else:
        print(f"⚠ [ORG {org_id}] Selesai scan {total_raw_bills} data, tetapi TIDAK ADA Bills dengan properti 'tags'.")

if __name__ == "__main__":
    try:
        token = get_new_access_token()
        print("✔ Access Token diperoleh.")
        
        # Looping menjalankan fungsi untuk tiap-tiap ORG ID yang ada di dalam list/array
        for current_org_id in TARGET_ORG_IDS:
            fetch_and_filter_bills(token, current_org_id)
            
        print("\n⚡ Seluruh proses multi-organisasi selesai dijalankan!")
        
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")