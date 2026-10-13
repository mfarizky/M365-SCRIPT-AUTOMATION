# 07 · KBLI Extractor (Python + OCR)

Indonesian companies list their business activities as **KBLI** codes (5-digit classification) in Article 3 of their deed of establishment (*Akta Pendirian*). The firm needed a table of KBLI codes for every client company, but the information was locked in hundreds of PDF deeds — many of them scans.

## Pipeline

```
extract_kbli_per_company.py   (or extract_kbli_folder.py)
        │  PDF text (PyMuPDF / pypdf) → OCR fallback (Tesseract, Indonesian)
        ▼
hasil_ekstraksi_akta_cepat.csv
        │
clean_results.py   → normalise company names, keep "Activity (KBLI)" items
        ▼
hasil_ekstraksi_bersih.csv
        │
summarize_kbli.py  → frequency of each KBLI across all companies
        ▼
Statistik_KBLI.csv
```

| Script | What it does |
|---|---|
| `extract_kbli_per_company.py` | For each company, fuzzy-matches its folder, scans deed PDFs, falls back to OCR (OpenCV grayscale + Tesseract) for scanned pages, stops at first success, logs failures to a separate CSV |
| `extract_kbli_folder.py` | Alternative: walks the whole folder, slices text between *Pasal 3* and *Pasal 4*, regex-extracts KBLI codes |
| `clean_results.py` | Cleans and normalises the raw output |
| `summarize_kbli.py` | Counts KBLI occurrences (`collections.Counter`) |
| `settings.py` | Loads paths from `.env` and the company list from `companies.local.txt` |

## Setup

```bash
pip install pymupdf pypdf pytesseract opencv-python numpy python-dotenv
# Install Tesseract-OCR with the Indonesian language pack (ind)
cp .env.example .env                              # set KBLI_ROOT_FOLDER & TESSERACT_CMD
cp companies.example.txt companies.local.txt      # one company name per line
python extract_kbli_per_company.py
```

`companies.local.txt`, `.env` and all CSV output are git-ignored (they contain client data).
