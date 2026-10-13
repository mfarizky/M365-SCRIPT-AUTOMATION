"""
Konfigurasi lokal untuk KBLI extractor.

Nilai sensitif TIDAK disimpan di repo:
  - .env                  -> path folder dokumen & path Tesseract
  - companies.local.txt   -> daftar nama perusahaan (data klien, 1 nama per baris)

Salin `.env.example` -> `.env` dan `companies.example.txt` -> `companies.local.txt`, lalu isi.
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    from dotenv import load_dotenv  # pip install python-dotenv
    load_dotenv(os.path.join(BASE_DIR, ".env"))
except ImportError:
    pass  # fallback: pakai environment variable sistem

TESSERACT_CMD = os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
ROOT_FOLDER = os.getenv("KBLI_ROOT_FOLDER", os.path.join(BASE_DIR, "sample_docs"))


def _load_companies():
    for name in ("companies.local.txt", "companies.example.txt"):
        path = os.path.join(BASE_DIR, name)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return [line.strip() for line in f if line.strip() and not line.startswith("#")]
    return []


DAFTAR_PT = _load_companies()
