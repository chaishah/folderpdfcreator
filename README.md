# 📂 Folder PDF Merger

> A powerful, interactive command-line tool that merges all supported files in a folder into a single, beautifully formatted PDF. Automatically handles images, office documents, emails, and PDFs, sorting them perfectly.

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## ✨ Features
- 🗂️ **Format Flexibility:** Merges Images (`.jpg`, `.png`, etc.), Word (`.docx`), Emails (`.eml`, `.msg`), and PDFs.
- 📑 **Smart Formatting:** Automatically generates a **Table of Contents**, sets up **Bookmarks**, and stamps **Page Numbers** (`N / Total`).
- 🗜️ **Deep Compression:** Lossless object deflation configures right alongside smart image re-compression (`JPEG` with custom quality) for huge space savings.
- 🔎 **Metadata Extraction:** Rips EXIF data and GPS coordinates from images to a separate `.metadata.txt` file.
- 🪄 **Interactive Mode:** A beautiful, terminal-based selection menu means you don't have to memorize a single command flag!

---

## 🎥 Watch it in Action

*(Upload a screen recording via terminal tools like [Asciinema](https://asciinema.org/) or VHS, and replace the placeholder GIF link below)*

![Demo of Interactive Mode](https://raw.githubusercontent.com/your-username/your-repo/main/demo.gif)

```text
? Folder to merge: ./Quarterly_Reports
? Output PDF path: ./Quarterly_Reports.pdf
? Compression: Image recompression  —  re-encode as JPEG (best for scans)
? Image quality (1-95): 85
? Output options:  (Space to toggle, Enter to confirm)
 ◯ Add PDF bookmarks  (one per source file)
 ◉ Add table of contents page  (implies bookmarks)
 ◉ Stamp page numbers  (N / Total at page bottom)
 ◯ Extract image EXIF metadata  (writes .txt file)
 ◯ Skip files that fail to convert  (default: abort)
```

---

## 🛠️ Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

> 💡 **Word documents (`.doc` / `.docx`):** Conversion uses Microsoft Word via COM automation (`docx2pdf`). If Word is not installed, the tool will fall back to plain-text extraction automatically.

---

## 🚀 Usage

### 🪄 Interactive Mode (Highly Recommended)

Run with no arguments to launch the interactive prompt. 

```bash
python merge_to_pdf.py
```

Use your arrow keys to navigate and `Enter` to confirm. The prompt smoothly guides you through:
1. **Target Folder** (with tab-completion)
2. **Output Path** (pre-filled smartly)
3. **Compression Level** (None, Lossless, or Lossy Image Re-compression)
4. **Navigational Add-ons** (TOC, Bookmarks, Page Numbers)
5. **Metadata & Error Handling**

### 💻 Command-line Mode

Perfect for scripts, aliases, or CI pipelines.

```bash
python merge_to_pdf.py FOLDER [OPTIONS]
```

| Argument / Option | Description |
|-------------------|-------------|
| `FOLDER` | Target folder containing the files to merge. |
| `-o, --output PATH` | Output PDF path (Defaults to `<folder>/<folder_name>.pdf`) |
| `--bookmarks` | Add a PDF bookmark (outline entry) for each source file. 🔖|
| `--toc` | Prepend a beautifully formatted Table of Contents page (implies `--bookmarks`). 📑 |
| `--page-numbers` | Stamp '`N / Total`' footer page numbers at the bottom of every page. 🔢 |
| `--compress` | Lossless compression (deflate streams, deduplicate objects). Best for Text/Vector PDFs. |
| `--image-quality [1-95]` | Re-encode embedded images as JPEG at specific quality (implies `--compress`). Ideal for scans. |
| `--extract-metadata` | Extract EXIF/image metadata from all images and save to `.metadata.txt`. |
| `--skip-errors` | Skip broken files instead of aborting the entire run. |
| `-h, --help` | Show help and exit |

---

## 📖 Examples

**Complete document assembly with TOC and Page Numbers:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --toc --page-numbers
```

**Heavy compression for scanned archives:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --image-quality 60
```

**Extract metadata to a text file:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --extract-metadata
```

**Skip broken/corrupted files:**
```bash
python merge_to_pdf.py "C:\path\to\MyFolder" --skip-errors
```

---

## 🧠 Behind the Scenes

### 📂 How files are sorted
Files are sorted strictly **numerically**, ensuring `10.png` comes after `9.jpg` (not after `1.pdf`). Mixed extensions sequence correctly! The tool automatically excludes the newly created output PDF so running it multiple times in the same directory won't cause infinite recursion errors.

### 🗜️ Compression Guide
- Images that would end up **larger** after re-compression will automatically be kept in their original encoding.
- Tiny graphics (like icon files or signatures < 100×100 px) are automatically skipped to preserve edge quality.

### 📝 Document Navigation & Formatting
- **Table of Contents (`--toc`)**: Scans all output items and automatically calculates the right page mapping to prepend a styled Title layout + TOC matching your file's pagination layout.
- **Page Numbering (`--page-numbers`)**: Dynamically overlays a transparent canvas layer placing `N / Total` dead-center at the bottom of each page without disturbing existing margins.

---
*Happy Merging! 🎉*
