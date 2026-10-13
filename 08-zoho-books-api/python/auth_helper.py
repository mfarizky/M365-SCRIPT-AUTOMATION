import os
import requests
from pathlib import Path
from dotenv import load_dotenv

# Path menuju Auth/.env
current_file_dir = Path(__file__).parent.resolve()
env_path = current_file_dir / "Auth" / ".env"

if not env_path.exists():
    print(f"❌ ERROR: File .env tidak ditemukan di jalur: {env_path}")
else:
    load_dotenv(dotenv_path=env_path)

def get_new_access_token():
    refresh_token = os.getenv("ZOHO_REFRESH_TOKEN")
    client_id = os.getenv("ZOHO_CLIENT_ID")
    client_secret = os.getenv("ZOHO_CLIENT_SECRET")
    
    if not all([refresh_token, client_id, client_secret]):
        raise Exception("❌ Gagal membaca kredensial dari .env. Pastikan isi file .env sudah benar.")
        
    auth_url = "https://accounts.zoho.com/oauth/v2/token"
    payload = {
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token"
    }
    
    response = requests.post(auth_url, data=payload)
    if response.status_code == 200:
        token_data = response.json()
        access_token = token_data.get("access_token")
        
        # Jangan cetak token utuh ke terminal (hindari bocor di log/screenshot)
        print(f"🔑 Access Token diperoleh: {access_token[:10]}...")
        
        return access_token
    else:
        raise Exception(f"Gagal generate Access Token: {response.text}")

# === TAMBAHKAN INI DI PALING BAWAH ===
if __name__ == "__main__":
    print("⚡ Menjalankan auth_helper langsung...")
    try:
        get_new_access_token()
    except Exception as e:
        print(e)