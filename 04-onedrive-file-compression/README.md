# 04 · OneDrive / SharePoint File Compression

SharePoint storage was filling up with scanned PDFs and phone photos of 5–30 MB each. This pipeline finds them through the local OneDrive sync client, compresses them in batches, and records progress in a CSV so it can resume at any time.

```
Scan (inventory CSV)  ──►  Compress in batches  ──►  CSV status updated
                                   │
                                   └─► backup / temp file / validation before overwrite
```

## `image/`

| Script | What it does |
|---|---|
| `Scan-LargeImages.ps1` | Recursively finds JPG/PNG ≥ 5 MB, categorises (5–8 MB, > 8 MB), maps local path → SharePoint URL, **appends** new files to a tracker CSV without duplicates |
| `Compress-ImagesInPlace.ps1` | Processes the next batch of pending rows with ImageMagick: compress to temp → validate with `magick identify` → skip if not smaller → overwrite in place → update `Status` and new size in the CSV |

**Why overwrite in place instead of delete + rename?** OneDrive tracks a file by its item ID. Writing new content to the same path creates a *new version* of the same file, so the original (large) version stays available in Version History. Delete + rename would create a new item and cut the history.

The script also skips cloud-only placeholders (`Offline` / `RecallOnDataAccess` attributes) and locked files instead of failing.

## `pdf/`

| Script | What it does |
|---|---|
| `Scan-LargePdfs.ps1` | Inventories PDFs ≥ 5 MB and detects whether each is **scanned (image-only)** or has text, using Ghostscript `txtwrite`; appends row-by-row so an interrupted run keeps its progress |
| `Compress-PdfGhostscript.ps1` | Batch compression with Ghostscript `/ebook`: backup copy → compress to temp → replace original. "Already processed" = backup exists, so runs are resumable. Pending + execution CSV logs |
| `Compress-PdfAdobeApi.ps1` | Same flow using the **Adobe PDF Services REST API** (OAuth token → asset upload → `compresspdf` job → poll `Location` header → download) |

## `recovery/`

Built after one compression run produced unwanted results — used to find and roll back every affected file.

| Script | What it does |
|---|---|
| `Get-UserFileActivity.ps1` | Tenant-wide KQL search for files modified by one user in a time window (UTC → local) → CSV |
| `Restore-PreviousFileVersion.ps1` | Reads that CSV, connects to each file's site (reconnecting only when the site changes) and runs `Restore-PnPFileVersion`, logging each result |

## Requirements

- [ImageMagick](https://imagemagick.org) (`magick` in PATH) – images
- [Ghostscript](https://ghostscript.com) (`gswin64c`) – PDFs
- Adobe PDF Services credentials (optional)
- OneDrive sync client with the target libraries synced locally
- Paths are set in `$Config.Compression` (see [`config`](../config)). Reports are written to `REPORT/` / `RESULT/` next to the scripts (git-ignored).
