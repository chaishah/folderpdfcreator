# Folder PDF Merger

Merges all supported files in a folder into a single PDF, sorted numerically by filename.

## Supported file types

| Category | Extensions |
|----------|-----------|
| Images | `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`, `.tif`, `.gif`, `.webp` |
| PDF | `.pdf` |
| Word | `.docx`, `.doc` |
| Email (MIME) | `.eml` |
| Outlook email | `.msg` |

## Installation

```bash
pip install -r requirements.txt
```

> **Word documents (.doc / .docx):** conversion uses Microsoft Word via COM automation (`docx2pdf`). If Word is not installed, the tool falls back to plain-text extraction.

## Usage

```bash
python merge_to_pdf.py FOLDER [OPTIONS]
```

### Arguments

| Argument | Description |
|----------|-------------|
| `FOLDER` | Path to the folder containing the files to merge |

### Options

| Option | Description |
|--------|-------------|
| `-o, --output PATH` | Output PDF path. Defaults to `<folder>/<folder_name>.pdf` |
| `--compress` | Lossless compression — deflates streams and deduplicates objects. Best for text/vector PDFs. |
| `--image-quality 1-95` | Re-encodes all embedded images as JPEG at the given quality (implies `--compress`). Best for scans and photos. |
| `--skip-errors` | Skip files that fail to convert instead of aborting the whole run. |
| `--help` | Show help and exit. |

## Examples

**Basic merge** — output saved to `MyFolder/MyFolder.pdf`:
```bash
python merge_to_pdf.py "C:\path\to\MyFolder"
```

**Custom output path:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" -o "C:\output\merged.pdf"
```

**Lossless compression** (good for text-heavy PDFs):
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --compress
```

**Image recompression** (best for scans and photos):
```bash
# High quality — moderate savings
python merge_to_pdf.py "C:\path\to\MyFolder" --image-quality 85

# Medium quality — good savings
python merge_to_pdf.py "C:\path\to\MyFolder" --image-quality 60

# Small file — visible quality loss
python merge_to_pdf.py "C:\path\to\MyFolder" --image-quality 40
```

**Skip broken files instead of aborting:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --image-quality 85 --skip-errors
```

## How files are sorted

Files are sorted **numerically** by filename, so `10.pdf` comes after `9.pdf` (not after `1.pdf`). Mixed names like `1.png`, `2.docx`, `3.jpg`, `10.pdf` are ordered correctly.

The output file is automatically excluded from the input list, so running the command twice on the same folder won't merge the previous output into itself.

## Compression guide

| Flag | How it works | When to use |
|------|-------------|-------------|
| *(none)* | No compression | Files are already small / optimised |
| `--compress` | Lossless: deflate + deduplication via `pikepdf` | Text, vectors, mixed documents |
| `--image-quality 85` | Lossy JPEG re-encoding of all embedded images | Scanned documents, photos |
| `--image-quality 60` | Same, more aggressive | Large scan archives |

- Images that end up **larger** after recompression are kept at their original encoding automatically.
- Tiny images (icons, signatures < 100×100 px) are skipped to preserve quality.
- `--image-quality` always implies `--compress`.

## Output

The tool prints a progress line for every file and a final size report:

```
Found 5 file(s):
  1.jpg
  2.jpg
  3.pdf
  4.docx
  5.eml

  OK  [1/5] 1.jpg
  OK  [2/5] 2.jpg
  OK  [3/5] 3.pdf
  OK  [4/5] 4.docx
  OK  [5/5] 5.eml

Merging 5 PDF segment(s) ...
Compressing + recompressing images at quality=85 ...

Done! -> C:\path\to\MyFolder\MyFolder.pdf  [12.4 MB -> 4.9 MB  (saved 7.6 MB, 60.9%)]
```
