#!/usr/bin/env python3
"""
Folder PDF Merger
Merges all supported files in a folder into a single PDF, sorted numerically by filename.

Supported formats:
  Images  : .png, .jpg, .jpeg, .bmp, .tiff, .tif, .gif, .webp
  PDF     : .pdf
  Word    : .docx, .doc
  Email   : .eml, .msg
"""

import html
import io
import re
import shutil
import sys
import tempfile
from pathlib import Path

import click


# ---------------------------------------------------------------------------
# Supported file extensions
# ---------------------------------------------------------------------------
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp"}
PDF_EXTS   = {".pdf"}
WORD_EXTS  = {".docx", ".doc"}
EML_EXTS   = {".eml"}
MSG_EXTS   = {".msg"}
ALL_EXTS   = IMAGE_EXTS | PDF_EXTS | WORD_EXTS | EML_EXTS | MSG_EXTS


# ---------------------------------------------------------------------------
# Sorting
# ---------------------------------------------------------------------------
def _natural_sort_key(path: Path):
    """Sort filenames naturally so '2' comes before '10'."""
    parts = re.split(r"(\d+)", path.stem)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


# ---------------------------------------------------------------------------
# Converters
# ---------------------------------------------------------------------------
def _rl_escape(text: str) -> str:
    """Escape characters that break ReportLab XML parser."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )


def _text_to_pdf(lines: list[str], out: Path, title_lines: list[str] | None = None) -> Path:
    """Render a list of text lines to a PDF page using ReportLab."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer

    styles = getSampleStyleSheet()
    story: list = []

    if title_lines:
        for line in title_lines:
            story.append(Paragraph(_rl_escape(line), styles["Normal"]))
        story.append(HRFlowable(width="100%"))
        story.append(Spacer(1, 6 * mm))

    for line in lines:
        line = line.rstrip()
        if line:
            story.append(Paragraph(_rl_escape(line), styles["Normal"]))
        else:
            story.append(Spacer(1, 3 * mm))

    if not story:
        story.append(Paragraph("(empty)", styles["Normal"]))

    doc = SimpleDocTemplate(
        str(out),
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )
    doc.build(story)
    return out


def image_to_pdf(src: Path, tmp_dir: str) -> Path:
    from PIL import Image

    out = Path(tmp_dir) / f"{src.stem}_img.pdf"
    img = Image.open(src)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.save(out, "PDF", resolution=150)
    return out


def _docx_fallback(src: Path, out: Path) -> Path:
    """Convert .docx to PDF via ReportLab when docx2pdf is unavailable."""
    from docx import Document  # python-docx

    doc = Document(str(src))
    lines = [p.text for p in doc.paragraphs]
    return _text_to_pdf(lines, out)


def docx_to_pdf(src: Path, tmp_dir: str) -> Path:
    out = Path(tmp_dir) / f"{src.stem}_word.pdf"
    try:
        from docx2pdf import convert  # requires Microsoft Word on the system
        convert(str(src), str(out))
        return out
    except Exception:
        return _docx_fallback(src, out)


def eml_to_pdf(src: Path, tmp_dir: str) -> Path:
    import email
    from email import policy

    with open(src, "rb") as fh:
        msg = email.message_from_bytes(fh.read(), policy=policy.default)

    headers = []
    for field in ("From", "To", "CC", "Subject", "Date"):
        val = msg.get(field, "")
        if val:
            headers.append(f"{field}: {val}")

    body_text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                body_text = part.get_content()
                break
        if not body_text:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    raw = part.get_content()
                    body_text = html.unescape(re.sub(r"<[^>]+>", "", raw))
                    break
    else:
        ct = msg.get_content_type()
        if ct == "text/html":
            raw = msg.get_content()
            body_text = html.unescape(re.sub(r"<[^>]+>", "", raw))
        else:
            body_text = msg.get_content()

    out = Path(tmp_dir) / f"{src.stem}_eml.pdf"
    return _text_to_pdf(body_text.splitlines(), out, title_lines=headers)


def msg_to_pdf(src: Path, tmp_dir: str) -> Path:
    import extract_msg

    msg = extract_msg.Message(str(src))
    headers = []
    for label, val in [
        ("From", msg.sender),
        ("To", msg.to),
        ("CC", msg.cc),
        ("Subject", msg.subject),
        ("Date", str(msg.date) if msg.date else ""),
    ]:
        if val:
            headers.append(f"{label}: {val}")

    body = (msg.body or "").splitlines()
    msg.close()

    out = Path(tmp_dir) / f"{src.stem}_msg.pdf"
    return _text_to_pdf(body, out, title_lines=headers)


# ---------------------------------------------------------------------------
# Image metadata extraction
# ---------------------------------------------------------------------------
_EXIF_FRIENDLY: dict[str, str] = {
    "DateTime":          "Date modified",
    "DateTimeOriginal":  "Date taken",
    "DateTimeDigitized": "Date digitised",
    "Make":              "Camera make",
    "Model":             "Camera model",
    "Software":          "Software",
    "Artist":            "Artist",
    "Copyright":         "Copyright",
    "ImageDescription":  "Description",
    "ExposureTime":      "Exposure time",
    "FNumber":           "F-number",
    "ISOSpeedRatings":   "ISO",
    "FocalLength":       "Focal length",
    "Flash":             "Flash",
    "Orientation":       "Orientation",
    "XResolution":       "X resolution",
    "YResolution":       "Y resolution",
    "ResolutionUnit":    "Resolution unit",
    "ColorSpace":        "Colour space",
    "ExifImageWidth":    "Exif width",
    "ExifImageHeight":   "Exif height",
    "LensMake":          "Lens make",
    "LensModel":         "Lens model",
}


def _dms_to_decimal(dms: tuple, ref: str) -> float:
    d, m, s = (float(x) for x in dms)
    decimal = d + m / 60 + s / 3600
    return -decimal if ref in ("S", "W") else decimal


def _fmt_gps(gps_ifd: dict) -> str | None:
    try:
        lat = _dms_to_decimal(gps_ifd[2], gps_ifd[1])
        lon = _dms_to_decimal(gps_ifd[4], gps_ifd[3])
        lat_dir = "N" if lat >= 0 else "S"
        lon_dir = "E" if lon >= 0 else "W"
        return f"{abs(lat):.6f} {lat_dir}, {abs(lon):.6f} {lon_dir}"
    except Exception:
        return None


def extract_image_metadata(src: Path) -> dict | None:
    """Return a dict of metadata fields for an image file, or None on failure."""
    from PIL import ExifTags, Image

    try:
        img = Image.open(src)
    except Exception:
        return None

    fields: dict[str, str] = {
        "Filename":   src.name,
        "Format":     img.format or "unknown",
        "Mode":       img.mode,
        "Dimensions": f"{img.width} x {img.height} px",
        "File size":  _fmt_size(src.stat().st_size),
    }

    try:
        exif = img.getexif()
    except Exception:
        exif = {}

    if exif:
        tag_map = {v: k for k, v in ExifTags.TAGS.items()}
        all_tags: dict = dict(exif)
        try:
            all_tags.update(exif.get_ifd(ExifTags.IFD.Exif))
        except Exception:
            pass

        for tag_name, friendly_label in _EXIF_FRIENDLY.items():
            tag_id = tag_map.get(tag_name)
            if tag_id is None:
                continue
            value = all_tags.get(tag_id)
            if value is None:
                continue
            if hasattr(value, "numerator"):
                value = f"{value.numerator}/{value.denominator}" if value.denominator != 1 else str(value.numerator)
            fields[friendly_label] = str(value).strip()

        try:
            gps_ifd = exif.get_ifd(ExifTags.IFD.GPSInfo)
            if gps_ifd:
                gps_str = _fmt_gps(gps_ifd)
                if gps_str:
                    fields["GPS coordinates"] = gps_str
        except Exception:
            pass

    return fields


def write_metadata_file(entries: list[tuple[Path, dict | None]], output: Path) -> int:
    """Write metadata for all images to a text file. Returns count of files with EXIF."""
    sep = "=" * 72
    has_exif_count = 0

    with open(output, "w", encoding="utf-8") as fh:
        fh.write("Image Metadata Report\n")
        fh.write(f"Generated from: {output.parent}\n")
        fh.write(f"{sep}\n\n")

        for src, fields in entries:
            fh.write(f"{sep}\n")
            fh.write(f"  {src.name}\n")
            fh.write(f"{sep}\n")

            if fields is None:
                fh.write("  (could not read file)\n\n")
                continue

            exif_keys = set(fields) - {"Filename", "Format", "Mode", "Dimensions", "File size"}
            if exif_keys:
                has_exif_count += 1

            label_width = max(len(k) for k in fields) + 2
            for label, value in fields.items():
                fh.write(f"  {label:<{label_width}}: {value}\n")
            fh.write("\n")

    return has_exif_count


# ---------------------------------------------------------------------------
# PDF utilities: page count, TOC, page numbers, bookmarks
# ---------------------------------------------------------------------------
def _count_pdf_pages(path: Path) -> int:
    """Return the number of pages in a PDF file."""
    from pypdf import PdfReader
    return len(PdfReader(str(path)).pages)


def _generate_toc_pdf(entries: list[tuple[str, int]], output: Path) -> int:
    """
    Render a Table of Contents to *output*.

    entries  : list of (filename, 1-indexed page number in the final merged PDF)
    Returns  : number of PDF pages written (usually 1; more for very long lists).
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas

    W, H   = A4
    left   = 20 * mm
    right  = W - 20 * mm
    top    = H - 20 * mm
    bottom = 20 * mm
    line_h = 6.5 * mm

    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    page_count = 1

    # ── Title ──────────────────────────────────────────────────────────────
    c.setFont("Helvetica-Bold", 15)
    c.drawString(left, top - 15, "Table of Contents")

    c.setLineWidth(0.6)
    c.setStrokeColorRGB(0.55, 0.55, 0.55)
    rule_y = top - 15 - 5 * mm
    c.line(left, rule_y, right, rule_y)

    y = rule_y - 8 * mm

    # ── Entries ────────────────────────────────────────────────────────────
    for name, page_num in entries:
        if y < bottom + line_h:
            c.showPage()
            page_count += 1
            y = top
            c.setFont("Helvetica", 10)

        # Truncate long filenames
        display  = name if len(name) <= 68 else name[:65] + "..."
        page_str = str(page_num)

        c.setFont("Helvetica", 10)
        name_w = c.stringWidth(display,  "Helvetica", 10)
        pg_w   = c.stringWidth(page_str, "Helvetica", 10)
        gap    = 3 * mm
        dot_w  = c.stringWidth(".", "Helvetica", 10)
        n_dots = max(3, int(((right - left) - name_w - pg_w - 2 * gap) / dot_w))

        c.setFillColorRGB(0, 0, 0)
        c.drawString(left, y, display)

        c.setFillColorRGB(0.6, 0.6, 0.6)
        c.drawString(left + name_w + gap, y, "." * n_dots)

        c.setFillColorRGB(0, 0, 0)
        c.drawRightString(right, y, page_str)

        y -= line_h

    c.save()
    buf.seek(0)
    output.write_bytes(buf.getvalue())
    return page_count


def _stamp_page_numbers(src: Path, output: Path) -> None:
    """
    Overlay a centred "N / Total" footer on every page of *src* and write
    the result to *output*.
    """
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(src))
    total  = len(reader.pages)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        c.setFont("Helvetica", 9)
        c.setFillColorRGB(0.45, 0.45, 0.45)
        c.drawCentredString(w / 2, 7 * mm, f"{i + 1}  /  {total}")
        c.save()

        buf.seek(0)
        overlay = PdfReader(buf).pages[0]
        page.merge_page(overlay)
        writer.add_page(page)

    with open(output, "wb") as fh:
        writer.write(fh)


# ---------------------------------------------------------------------------
# Merger + compression
# ---------------------------------------------------------------------------
def _fmt_size(n_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n_bytes < 1024:
            return f"{n_bytes:.1f} {unit}"
        n_bytes /= 1024
    return f"{n_bytes:.1f} TB"


def merge_pdfs(
    pdf_paths: list[Path],
    output: Path,
    bookmarks: list[tuple[str, int]] | None = None,
) -> None:
    """
    Merge *pdf_paths* into *output*.

    bookmarks : optional list of (title, 0-indexed page number) added as
                top-level PDF outline entries.
    """
    from pypdf import PdfWriter

    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(str(p))

    if bookmarks:
        for title, page_idx in bookmarks:
            writer.add_outline_item(title, page_idx)

    with open(output, "wb") as fh:
        writer.write(fh)


def _recompress_images(pdf: "pikepdf.Pdf", quality: int) -> None:  # noqa: F821
    """Re-encode every large image in the PDF as JPEG at the given quality."""
    import pikepdf
    from PIL import Image

    processed: set = set()

    for page in pdf.pages:
        resources = page.get("/Resources")
        if not resources:
            continue
        xobjects = resources.get("/XObject")
        if not xobjects:
            continue

        for key in list(xobjects.keys()):
            xobj = xobjects[key]

            try:
                objgen = xobj.objgen
                if objgen in processed:
                    continue
                processed.add(objgen)
            except Exception:
                pass

            if xobj.get("/Subtype") != "/Image":
                continue

            try:
                pdfimage = pikepdf.PdfImage(xobj)
                pil_img  = pdfimage.as_pil_image()

                w, h = pil_img.size
                if w * h < 10_000:
                    continue

                if pil_img.mode == "RGBA":
                    bg = Image.new("RGB", pil_img.size, (255, 255, 255))
                    bg.paste(pil_img, mask=pil_img.split()[3])
                    pil_img = bg
                elif pil_img.mode not in ("RGB", "L"):
                    pil_img = pil_img.convert("RGB")

                buf = io.BytesIO()
                pil_img.save(buf, format="JPEG", quality=quality, optimize=True)
                jpeg_bytes = buf.getvalue()

                try:
                    if len(jpeg_bytes) >= len(xobj.read_raw_bytes()):
                        continue
                except Exception:
                    pass

                cs = "/DeviceGray" if pil_img.mode == "L" else "/DeviceRGB"
                xobj.write(jpeg_bytes, filter=pikepdf.Name("/DCTDecode"))
                xobj["/ColorSpace"] = pikepdf.Name(cs)
                xobj["/Width"]  = w
                xobj["/Height"] = h
                xobj["/BitsPerComponent"] = 8
                for remove_key in ("/DecodeParms", "/SMask", "/Mask", "/Intent"):
                    try:
                        del xobj[remove_key]
                    except Exception:
                        pass

            except Exception:
                pass  # leave images that can't be processed untouched


def compress_pdf(src: Path, output: Path, image_quality: int | None = None) -> None:
    """Compress a PDF using pikepdf. Optionally recompress embedded images."""
    import pikepdf

    with pikepdf.open(str(src)) as pdf:
        if image_quality is not None:
            _recompress_images(pdf, image_quality)
        pdf.save(
            str(output),
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            recompress_flate=True,
        )


# ---------------------------------------------------------------------------
# Core execution (shared by CLI and interactive mode)
# ---------------------------------------------------------------------------
def _execute_merge(
    folder_path: Path,
    output_path: Path,
    compress: bool,
    image_quality: int | None,
    extract_metadata: bool,
    skip_errors: bool,
    add_bookmarks: bool = False,
    add_toc: bool = False,
    add_page_numbers: bool = False,
) -> None:
    do_compress = compress or (image_quality is not None)
    if add_toc:
        add_bookmarks = True  # TOC always pairs with bookmarks

    files = sorted(
        [
            f for f in folder_path.iterdir()
            if f.is_file()
            and f.suffix.lower() in ALL_EXTS
            and f.resolve() != output_path.resolve()
        ],
        key=_natural_sort_key,
    )

    if not files:
        click.echo(
            click.style(
                f"No supported files found in '{folder_path}'.\n"
                f"Supported: {', '.join(sorted(ALL_EXTS))}",
                fg="red",
            ),
            err=True,
        )
        sys.exit(1)

    click.echo(click.style(f"Found {len(files)} file(s):", bold=True))
    for f in files:
        click.echo(f"  {f.name}")
    click.echo()

    # ── Metadata extraction ────────────────────────────────────────────────
    if extract_metadata:
        image_files = [f for f in files if f.suffix.lower() in IMAGE_EXTS]
        if not image_files:
            click.echo(click.style("No image files found — skipping metadata extraction.", fg="yellow"))
        else:
            meta_path = output_path.with_suffix(".metadata.txt")
            entries_meta = [(f, extract_image_metadata(f)) for f in image_files]
            exif_count = write_metadata_file(entries_meta, meta_path)
            click.echo(
                click.style(
                    f"Metadata written: {meta_path}"
                    f"  ({exif_count}/{len(image_files)} image(s) had EXIF data)",
                    fg="cyan",
                )
            )
            click.echo()

    with tempfile.TemporaryDirectory() as tmp_dir:

        # ── Convert files to PDF parts ─────────────────────────────────────
        converted: list[tuple[Path, Path]] = []  # (original, pdf_part)

        for i, f in enumerate(files, 1):
            ext   = f.suffix.lower()
            label = f"[{i}/{len(files)}] {f.name}"
            try:
                if ext in IMAGE_EXTS:
                    part = image_to_pdf(f, tmp_dir)
                elif ext in PDF_EXTS:
                    part = f
                elif ext in WORD_EXTS:
                    part = docx_to_pdf(f, tmp_dir)
                elif ext in EML_EXTS:
                    part = eml_to_pdf(f, tmp_dir)
                elif ext in MSG_EXTS:
                    part = msg_to_pdf(f, tmp_dir)
                else:
                    continue
                converted.append((f, part))
                click.echo(click.style(f"  OK  {label}", fg="green"))
            except Exception as exc:
                err_msg = f"  ERR {label}: {exc}"
                if skip_errors:
                    click.echo(click.style(err_msg, fg="yellow"), err=True)
                else:
                    click.echo(click.style(err_msg, fg="red"), err=True)
                    sys.exit(1)

        if not converted:
            click.echo(click.style("No files were successfully converted.", fg="red"), err=True)
            sys.exit(1)

        pdf_parts = [part for _, part in converted]

        # ── Page counts / TOC / bookmarks ──────────────────────────────────
        bookmarks_list: list[tuple[str, int]] | None = None

        if add_bookmarks or add_toc:
            click.echo()
            click.echo("Counting pages ...")
            page_counts = [_count_pdf_pages(p) for p in pdf_parts]

        if add_toc:
            click.echo("Generating table of contents ...")

            # First pass: measure how many pages the TOC itself will need.
            # (Page numbers don't affect page count, so any numbers work here.)
            draft_toc = Path(tmp_dir) / "_toc_draft.pdf"
            draft_entries = [
                (orig.name, 1 + sum(page_counts[:i]))
                for i, (orig, _) in enumerate(converted)
            ]
            toc_n_pages = _generate_toc_pdf(draft_entries, draft_toc)

            # Second pass: correct page numbers now that we know toc_n_pages.
            toc_path = Path(tmp_dir) / "_toc_final.pdf"
            toc_entries = [
                (orig.name, toc_n_pages + 1 + sum(page_counts[:i]))
                for i, (orig, _) in enumerate(converted)
            ]
            _generate_toc_pdf(toc_entries, toc_path)

            # Prepend TOC; update page_counts so bookmark offsets stay correct.
            pdf_parts   = [toc_path] + pdf_parts
            page_counts = [toc_n_pages] + page_counts

            # Bookmarks point to the first page of each source file (0-indexed).
            bookmarks_list = [
                (orig.name, toc_n_pages + sum(page_counts[1 : 1 + i]))
                for i, (orig, _) in enumerate(converted)
            ]

        elif add_bookmarks:
            cumsum = 0
            bookmarks_list = []
            for (orig, _), count in zip(converted, page_counts):
                bookmarks_list.append((orig.name, cumsum))
                cumsum += count

        # ── Merge ──────────────────────────────────────────────────────────
        click.echo()
        click.echo(f"Merging {len(pdf_parts)} PDF segment(s) ...")

        need_tmp = do_compress or add_page_numbers
        raw_path = Path(tmp_dir) / "_merged_raw.pdf" if need_tmp else output_path
        merge_pdfs(pdf_parts, raw_path, bookmarks=bookmarks_list)
        raw_size = raw_path.stat().st_size
        current  = raw_path

        # ── Page numbers ───────────────────────────────────────────────────
        if add_page_numbers:
            click.echo("Stamping page numbers ...")
            numbered_path = Path(tmp_dir) / "_merged_numbered.pdf"
            _stamp_page_numbers(current, numbered_path)
            current = numbered_path

        # ── Compression ────────────────────────────────────────────────────
        if do_compress:
            if image_quality is not None:
                click.echo(f"Compressing + recompressing images at quality={image_quality} ...")
            else:
                click.echo("Compressing (lossless) ...")
            compress_pdf(current, output_path, image_quality=image_quality)
            final_size = output_path.stat().st_size
            saved = raw_size - final_size
            pct   = (saved / raw_size * 100) if raw_size else 0
            size_info = (
                f"{_fmt_size(raw_size)} -> {_fmt_size(final_size)}"
                f"  (saved {_fmt_size(saved)}, {pct:.1f}%)"
            )
        elif current != output_path:
            shutil.copy2(str(current), str(output_path))
            size_info = _fmt_size(output_path.stat().st_size)
        else:
            size_info = _fmt_size(output_path.stat().st_size)

    click.echo()
    click.echo(click.style(f"Done! -> {output_path}  [{size_info}]", fg="green", bold=True))


# ---------------------------------------------------------------------------
# Interactive mode
# ---------------------------------------------------------------------------
_Q_STYLE = None  # lazily initialised after questionary import


def _q_style():
    global _Q_STYLE
    if _Q_STYLE is None:
        from questionary import Style
        _Q_STYLE = Style([
            ("qmark",       "fg:cyan bold"),
            ("question",    "bold"),
            ("answer",      "fg:cyan bold"),
            ("pointer",     "fg:cyan bold"),
            ("highlighted", "fg:cyan bold"),
            ("selected",    "fg:green"),
            ("instruction", "fg:grey italic"),
            ("text",        ""),
            ("disabled",    "fg:grey italic"),
        ])
    return _Q_STYLE


def _ask(prompt_fn, *args, **kwargs):
    """Call a questionary prompt and exit cleanly on Ctrl-C / Ctrl-D."""
    result = prompt_fn(*args, style=_q_style(), **kwargs).ask()
    if result is None:
        click.echo("\nCancelled.")
        sys.exit(0)
    return result


def interactive_mode() -> None:
    import questionary

    click.echo(click.style("\n  Folder PDF Merger\n", bold=True))

    # ── Folder ────────────────────────────────────────────────────────────
    folder_str  = _ask(questionary.path, "Folder to merge:", only_directories=True)
    folder_path = Path(folder_str).expanduser().resolve()
    if not folder_path.is_dir():
        click.echo(click.style(f"Error: '{folder_path}' is not a directory.", fg="red"))
        sys.exit(1)

    # ── Output path ───────────────────────────────────────────────────────
    default_out = str(folder_path / f"{folder_path.name}.pdf")
    output_str  = _ask(questionary.text, "Output PDF path:", default=default_out)
    output_path = Path(output_str).expanduser()

    # ── Compression ───────────────────────────────────────────────────────
    compress_choice = _ask(
        questionary.select,
        "Compression:",
        choices=[
            questionary.Choice("None  —  merge only",                                        value="none"),
            questionary.Choice("Lossless  —  deflate streams & deduplicate objects",          value="lossless"),
            questionary.Choice("Image recompression  —  re-encode as JPEG (best for scans)",  value="images"),
        ],
    )

    image_quality: int | None = None
    if compress_choice == "images":
        quality_str = _ask(
            questionary.text,
            "Image quality (1-95):",
            default="85",
            validate=lambda v: (v.isdigit() and 1 <= int(v) <= 95) or "Enter a whole number between 1 and 95",
        )
        image_quality = int(quality_str)

    # ── Output options (multi-select checkbox) ────────────────────────────
    extra_opts: list[str] = _ask(
        questionary.checkbox,
        "Output options:  (Space to toggle, Enter to confirm)",
        choices=[
            questionary.Choice("Add PDF bookmarks  (one per source file)",            value="bookmarks"),
            questionary.Choice("Add table of contents page  (implies bookmarks)",     value="toc"),
            questionary.Choice("Stamp page numbers  (N / Total at page bottom)",      value="page_numbers"),
            questionary.Choice("Extract image EXIF metadata  (writes .txt file)",     value="metadata"),
            questionary.Choice("Skip files that fail to convert  (default: abort)",   value="skip_errors"),
        ],
    )

    add_toc          = "toc"          in extra_opts
    add_bookmarks    = "bookmarks"    in extra_opts or add_toc
    add_page_numbers = "page_numbers" in extra_opts
    extract_metadata = "metadata"     in extra_opts
    skip_errors      = "skip_errors"  in extra_opts

    # ── Summary + confirm ─────────────────────────────────────────────────
    click.echo()
    click.echo(click.style("  Summary", bold=True))
    click.echo(f"  Folder     : {folder_path}")
    click.echo(f"  Output     : {output_path}")
    if compress_choice == "none":
        click.echo( "  Compression: None")
    elif compress_choice == "lossless":
        click.echo( "  Compression: Lossless")
    else:
        click.echo(f"  Compression: Images at quality {image_quality}")

    flags = []
    if add_bookmarks:
        flags.append("bookmarks")
    if add_toc:
        flags.append("TOC")
    if add_page_numbers:
        flags.append("page numbers")
    if extract_metadata:
        flags.append("metadata")
    if skip_errors:
        flags.append("skip errors")
    click.echo(f"  Options    : {', '.join(flags) if flags else 'none'}")
    click.echo()

    if not _ask(questionary.confirm, "Proceed?", default=True):
        click.echo("Cancelled.")
        sys.exit(0)

    click.echo()
    _execute_merge(
        folder_path=folder_path,
        output_path=output_path,
        compress=(compress_choice == "lossless"),
        image_quality=image_quality,
        extract_metadata=extract_metadata,
        skip_errors=skip_errors,
        add_bookmarks=add_bookmarks,
        add_toc=add_toc,
        add_page_numbers=add_page_numbers,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.argument("folder", required=False, default=None,
                type=click.Path(exists=False, file_okay=False))
@click.option("-o", "--output",        default=None,
              help="Output PDF path. Defaults to <folder>/<folder_name>.pdf")
@click.option("--compress",            is_flag=True, default=False,
              help="Lossless compression: deflate streams + deduplicate objects.")
@click.option("--image-quality",       type=click.IntRange(1, 95), default=None, metavar="1-95",
              help="Re-encode images as JPEG at this quality (implies --compress). 85=high, 60=medium, 40=small.")
@click.option("--bookmarks",           is_flag=True, default=False,
              help="Add a PDF bookmark (outline entry) for each source file.")
@click.option("--toc",                 is_flag=True, default=False,
              help="Prepend a Table of Contents page (implies --bookmarks).")
@click.option("--page-numbers",        is_flag=True, default=False,
              help="Stamp 'N / Total' page numbers at the bottom of every page.")
@click.option("--extract-metadata",    is_flag=True, default=False,
              help="Save EXIF/image metadata to a .txt file alongside the output PDF.")
@click.option("--skip-errors",         is_flag=True, default=False,
              help="Skip files that fail to convert instead of aborting.")
def main(
    folder: str | None,
    output: str | None,
    compress: bool,
    image_quality: int | None,
    bookmarks: bool,
    toc: bool,
    page_numbers: bool,
    extract_metadata: bool,
    skip_errors: bool,
) -> None:
    """Merge all supported files in FOLDER into a single PDF.

    Run with no arguments to launch interactive mode.

    \b
    Files are sorted numerically (1.png, 2.docx, 10.pdf …).
    Compression options:
      --compress              lossless (structure only)
      --image-quality 85      lossy image recompression — best for scans/photos
    Navigation options:
      --bookmarks             PDF outline entry per source file
      --toc                   Table of contents page (implies --bookmarks)
      --page-numbers          Stamp N / Total footer on every page
    """
    if folder is None:
        interactive_mode()
        return

    folder_path = Path(folder).resolve()
    if not folder_path.is_dir():
        click.echo(click.style(f"Error: '{folder_path}' is not a directory.", fg="red"), err=True)
        sys.exit(1)

    output_path = Path(output) if output else folder_path / f"{folder_path.name}.pdf"
    _execute_merge(
        folder_path,
        output_path,
        compress,
        image_quality,
        extract_metadata,
        skip_errors,
        add_bookmarks=bookmarks or toc,
        add_toc=toc,
        add_page_numbers=page_numbers,
    )


if __name__ == "__main__":
    main()
