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
import os
import re
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
        # Fall back to text extraction if Word is not installed
        return _docx_fallback(src, out)


def eml_to_pdf(src: Path, tmp_dir: str) -> Path:
    import email
    from email import policy

    with open(src, "rb") as fh:
        msg = email.message_from_bytes(fh.read(), policy=policy.default)

    # Build header block
    headers = []
    for field in ("From", "To", "CC", "Subject", "Date"):
        val = msg.get(field, "")
        if val:
            headers.append(f"{field}: {val}")

    # Extract plain-text body (prefer text/plain, fall back to stripped HTML)
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
# Merger + compression
# ---------------------------------------------------------------------------
def _fmt_size(n_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n_bytes < 1024:
            return f"{n_bytes:.1f} {unit}"
        n_bytes /= 1024
    return f"{n_bytes:.1f} TB"


def merge_pdfs(pdf_paths: list[Path], output: Path) -> None:
    """Merge PDF files into one with no compression applied."""
    from pypdf import PdfWriter

    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(str(p))
    with open(output, "wb") as fh:
        writer.write(fh)


def _recompress_images(pdf: "pikepdf.Pdf", quality: int) -> None:  # noqa: F821
    """Re-encode every large image in the PDF as JPEG at the given quality."""
    import io

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

            # Track indirect objects so shared images are only processed once
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
                pil_img = pdfimage.as_pil_image()

                # Skip tiny images (icons, logos, etc.)
                w, h = pil_img.size
                if w * h < 10_000:
                    continue

                # Flatten transparency for JPEG
                if pil_img.mode == "RGBA":
                    bg = Image.new("RGB", pil_img.size, (255, 255, 255))
                    bg.paste(pil_img, mask=pil_img.split()[3])
                    pil_img = bg
                elif pil_img.mode not in ("RGB", "L"):
                    pil_img = pil_img.convert("RGB")

                buf = io.BytesIO()
                pil_img.save(buf, format="JPEG", quality=quality, optimize=True)
                jpeg_bytes = buf.getvalue()

                # Only replace if actually smaller than the original stream
                try:
                    if len(jpeg_bytes) >= len(xobj.read_raw_bytes()):
                        continue
                except Exception:
                    pass

                cs = "/DeviceGray" if pil_img.mode == "L" else "/DeviceRGB"
                xobj.write(jpeg_bytes, filter=pikepdf.Name("/DCTDecode"))
                xobj["/ColorSpace"] = pikepdf.Name(cs)
                xobj["/Width"] = w
                xobj["/Height"] = h
                xobj["/BitsPerComponent"] = 8
                for remove_key in ("/DecodeParms", "/SMask", "/Mask", "/Intent"):
                    try:
                        del xobj[remove_key]
                    except Exception:
                        pass

            except Exception:
                pass  # Leave images that can't be processed untouched


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
# CLI
# ---------------------------------------------------------------------------
@click.command()
@click.argument("folder", type=click.Path(exists=True, file_okay=False, resolve_path=True))
@click.option(
    "-o", "--output",
    default=None,
    help="Output PDF path. Defaults to <folder>/<folder_name>.pdf",
)
@click.option(
    "--compress",
    is_flag=True,
    default=False,
    help="Lossless compression: deflate streams + deduplicate objects.",
)
@click.option(
    "--image-quality",
    type=click.IntRange(1, 95),
    default=None,
    metavar="1-95",
    help=(
        "Re-encode embedded images as JPEG at this quality (implies --compress). "
        "Good values: 85 = high quality, 60 = medium, 40 = small file."
    ),
)
@click.option(
    "--skip-errors",
    is_flag=True,
    default=False,
    help="Skip files that fail to convert instead of aborting.",
)
def main(
    folder: str,
    output: str | None,
    compress: bool,
    image_quality: int | None,
    skip_errors: bool,
) -> None:
    """Merge all supported files in FOLDER into a single PDF.

    Files are sorted numerically by filename (e.g. 1.png, 2.docx, 10.pdf).

    \b
    Compression options:
      --compress              lossless (structure only)
      --image-quality 85      lossy image recompression — use this for scans/photos
    """
    folder_path = Path(folder)
    output_path = Path(output) if output else folder_path / f"{folder_path.name}.pdf"

    # image-quality implies compress
    do_compress = compress or (image_quality is not None)

    # Collect supported files, excluding the output file itself
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

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_parts: list[Path] = []

        for i, f in enumerate(files, 1):
            ext = f.suffix.lower()
            label = f"[{i}/{len(files)}] {f.name}"
            try:
                if ext in IMAGE_EXTS:
                    pdf_parts.append(image_to_pdf(f, tmp_dir))
                elif ext in PDF_EXTS:
                    pdf_parts.append(f)
                elif ext in WORD_EXTS:
                    pdf_parts.append(docx_to_pdf(f, tmp_dir))
                elif ext in EML_EXTS:
                    pdf_parts.append(eml_to_pdf(f, tmp_dir))
                elif ext in MSG_EXTS:
                    pdf_parts.append(msg_to_pdf(f, tmp_dir))
                click.echo(click.style(f"  OK  {label}", fg="green"))
            except Exception as exc:
                err_msg = f"  ERR {label}: {exc}"
                if skip_errors:
                    click.echo(click.style(err_msg, fg="yellow"), err=True)
                else:
                    click.echo(click.style(err_msg, fg="red"), err=True)
                    sys.exit(1)

        if not pdf_parts:
            click.echo(click.style("No files were successfully converted.", fg="red"), err=True)
            sys.exit(1)

        click.echo()
        click.echo(f"Merging {len(pdf_parts)} PDF segment(s) ...")

        if do_compress:
            raw_path = Path(tmp_dir) / "_merged_raw.pdf"
            merge_pdfs(pdf_parts, raw_path)
            raw_size = raw_path.stat().st_size

            if image_quality is not None:
                click.echo(f"Compressing + recompressing images at quality={image_quality} ...")
            else:
                click.echo("Compressing (lossless) ...")

            compress_pdf(raw_path, output_path, image_quality=image_quality)
            final_size = output_path.stat().st_size

            saved = raw_size - final_size
            pct = (saved / raw_size * 100) if raw_size else 0
            size_info = (
                f"{_fmt_size(raw_size)} -> {_fmt_size(final_size)}"
                f"  (saved {_fmt_size(saved)}, {pct:.1f}%)"
            )
        else:
            merge_pdfs(pdf_parts, output_path)
            size_info = _fmt_size(output_path.stat().st_size)

    click.echo()
    click.echo(click.style(f"Done! -> {output_path}  [{size_info}]", fg="green", bold=True))


if __name__ == "__main__":
    main()
