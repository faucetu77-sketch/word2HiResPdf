#!/usr/bin/env python3
"""DOCX -> PDF for prepress: Word's own pagination, images restored to their DOCX originals.

    python docx2pdf.py report.docx                 # -> report.pdf beside source
    python docx2pdf.py a.docx b.docx -o out/       # batch into out/
    python docx2pdf.py report.docx --audit         # list image codecs inside the PDF
    python docx2pdf.py report.docx --pdfa          # PDF/A-1b
    python docx2pdf.py report.docx --no-repair     # keep Word's raster encoding as is
    python docx2pdf.py report.docx -v              # list every restored image

Why a repair pass: Word's PDF export re-encodes photos to JPEG, which destroys prepress
quality. Layout (pages, headers, footers, columns, breaks, fonts, links, bookmarks)
remains 100% Word. Each raster placed in the PDF is then swapped for the original from
word/media/*, cropped losslessly to the same visible area (a:srcRect) and stored as Flate.
"""
import argparse
import io
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

EXTS = {".docx", ".doc", ".rtf", ".odt"}
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def word_export(sources: list[Path], outdir: Path, pdfa: bool) -> list[Path]:
    import win32com.client as win32

    outdir.mkdir(parents=True, exist_ok=True)
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    made = []
    try:
        for src in sources:
            doc = app.Documents.Open(str(src.resolve()), ReadOnly=True, AddToRecentFiles=False)
            try:
                dst = outdir / (src.stem + ".pdf")
                # wdExportFormatPDF=17, wdExportOptimizeForPrint=0, wdExportAllDocument=0,
                # wdExportDocumentContent=0, wdExportCreateHeadingBookmarks=1
                doc.ExportAsFixedFormat(
                    str(dst.resolve()), 17, False, 0, 0, 0, 0, 0,
                    True,    # IncludeDocProps
                    True,    # KeepIRM
                    1,       # bookmarks from headings -> PDF outline
                    True,    # DocStructureTags -> tagged, selectable text
                    True,    # BitmapMissingFonts
                    pdfa,    # UseISO19005_1
                )
                made.append(dst)
            finally:
                doc.Close(False)
    finally:
        app.Quit()
    return made


# ---------------------------------------------------------------- DOCX side

def docx_images(docx: Path) -> list[dict]:
    """Raster figures in document order: PIL image, its a:srcRect crop, a:lum correction."""
    from PIL import Image

    out: list[dict] = []
    with zipfile.ZipFile(docx) as z:
        rels = {}
        try:
            for rel in ET.fromstring(z.read("word/_rels/document.xml.rels")):
                if rel.tag != f"{{{REL}}}Relationship":
                    continue
                target = rel.get("Target", "").replace("\\", "/").lstrip("/")
                rels[rel.get("Id")] = target if target.startswith("word/") else "word/" + target
        except KeyError:
            pass
        try:
            dom = ET.fromstring(z.read("word/document.xml"))
        except KeyError:
            return out

        for pic in dom.iter(f"{{{PIC}}}pic"):
            blip = pic.find(f".//{{{A}}}blip")
            if blip is None:
                continue
            part = rels.get(blip.get(f"{{{R}}}embed"))
            if not part:
                continue
            # a14:imgEffect is already baked into this image by Word (the untouched original
            # lives in a14:imgLayer only for re-editing), so effects must NOT be re-applied.
            im, raw = load_image(z, part)
            if im is None:
                continue
            rect = pic.find(f".//{{{A}}}srcRect")
            crop = tuple(float(rect.get(k, 0) or 0) / 100000.0 for k in ("l", "t", "r", "b")) \
                if rect is not None else (0.0, 0.0, 0.0, 0.0)
            lum = blip.find(f"{{{A}}}lum")
            out.append({"part": part.rsplit("/", 1)[-1], "image": im, "raw": raw, "crop": crop,
                        "lum": {k: int(v) / 100000.0 for k, v in lum.attrib.items()} if lum is not None else {}})
    return out


def load_image(z: zipfile.ZipFile, part: str):
    """Decode a media part. JPEG/PNG/GIF via Pillow, JPEG XR (.wdp) via Windows WIC.

    Returns (PIL image, raw bytes) or (None, None)."""
    from pathlib import Path

    from PIL import Image

    try:
        data = z.read(part)
    except KeyError:
        return None, None
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
        return im, data
    except Exception:
        pass
    if Path(part).suffix.lower() not in (".wdp", ".jxr"):
        return None, None  # EMF/WMF/SVG: vector, Word keeps them vector
    png = _wic_to_png(data)
    if png is None:
        return None, None
    try:
        im = Image.open(io.BytesIO(png))
        im.load()
        return im, png
    except Exception:
        return None, None


_WIC_PS1 = """Add-Type -AssemblyName PresentationCore
$dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create((New-Object System.Uri($env:W2P_IN)), 'None', 'OnLoad')
$conv = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap($dec.Frames[0], [System.Windows.Media.PixelFormats]::Bgr24, $null, 0)
$enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
$enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($conv))
$fs = [IO.File]::Create($env:W2P_OUT); $enc.Save($fs); $fs.Close()
"""


def _wic_to_png(data: bytes) -> bytes | None:
    """Windows-only JPEG XR decode with no third-party dependency (PowerShell + WIC)."""
    import os
    import subprocess
    import tempfile

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        script = Path(tmp, "w2p-decode.ps1")
        script.write_text(_WIC_PS1, encoding="utf-8")
        src, dst = Path(tmp, "in.wdp"), Path(tmp, "out.png")
        src.write_bytes(data)
        env = {**os.environ, "W2P_IN": str(src), "W2P_OUT": str(dst)}
        try:
            subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
                            "Bypass", "-File", str(script)],
                           capture_output=True, timeout=120, env=env)
            return dst.read_bytes() if dst.exists() else None
        except Exception:
            return None


def apply_effects(im, lum: dict):
    """Reproduce Word's a:lum brightness/contrast at full resolution (verified vs Word output)."""
    import numpy as np
    from PIL import Image

    if not lum:
        return im
    arr = np.asarray(im.convert("RGB"), dtype=np.float64)
    if "bright" in lum:
        arr = arr + lum["bright"] * 255.0
    if "contrast" in lum:
        arr = (arr - 128.0) * (1.0 + lum["contrast"]) + 128.0
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8))


def lossless_crop(im, crop):
    l, t, r, b = crop
    if not any(crop):
        return im
    w, h = im.size
    box = (round(l * w), round(t * h), round(w - r * w), round(h - b * h))
    box = (max(box[0], 0), max(box[1], 0), min(box[2], w), min(box[3], h))
    if box[2] <= box[0] or box[3] <= box[1]:
        return im
    return im.crop(box)


def thumb(im, size=16):
    return im.convert("RGB").resize((size, size))


def tdiff(a, b) -> float:
    from PIL import ImageChops, ImageStat
    return sum(ImageStat.Stat(ImageChops.difference(a, b.resize(a.size))).mean) / 3.0


# ---------------------------------------------------------------- PDF side

def repair(pdf_path: Path, docx: Path, out_path: Path | None = None, match_threshold: float = 24.0) -> list[str]:
    """Swap every re-encoded raster in the PDF for the DOCX original. Returns notes."""
    import os
    import tempfile

    import pymupdf as fitz
    from PIL import Image

    out_path = out_path or pdf_path
    cands = docx_images(docx)
    if not cands:
        return ["no raster images in DOCX; nothing to restore"]
    for c in cands:
        # a:lum is applied by Word at render time; crop happens after it.
        c["corrected"] = apply_effects(c["image"], c["lum"])
        c["cropped"] = lossless_crop(c["corrected"], c["crop"])
        c["thumb"] = thumb(c["cropped"])
        # Pass the original file through byte-for-byte when nothing has to change; the
        # PDF then carries the very same JPEG the author inserted, zero re-compression.
        untouched = not any(c["crop"]) and not c["lum"] and c["cropped"] is c["image"]
        c["stream"] = c["raw"] if untouched else None
        c["png"] = None

    doc = fitz.open(pdf_path)
    seen, notes, restored = set(), [], 0
    for page in doc:
        for info in page.get_images(full=True):
            xref = info[0]
            if xref in seen:
                continue
            seen.add(xref)
            pm = fitz.Pixmap(doc, xref)
            if pm.n != 3 or pm.alpha:
                continue  # greyscale/CMYK/masked: Word stores those losslessly already
            pdfim = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)

            best, bestd = None, None
            for c in cands:
                d = tdiff(thumb(pdfim), c["thumb"])
                if bestd is None or d < bestd:
                    best, bestd = c, d
            if best is None or bestd > match_threshold:
                notes.append(f"xref {xref} {pm.width}x{pm.height}: no DOCX original matched "
                             f"(best diff {bestd:.1f})")
                continue

            target = best["cropped"]
            if target.size == pdfim.size and target.tobytes() == pdfim.tobytes():
                continue  # Word left it lossless
            stream = best["stream"]
            if stream is None:
                if best["png"] is None:
                    buf = io.BytesIO()
                    target.save(buf, "PNG", compress_level=9)
                    best["png"] = buf.getvalue()
                stream = best["png"]
            try:
                page.replace_image(xref, stream=stream)
            except Exception as e:  # pragma: no cover - defensive
                notes.append(f"xref {xref}: replace failed: {e}")
                continue
            restored += 1
            notes.append(f"xref {xref}: {pm.width}x{pm.height} -> {target.width}x{target.height} "
                         f"from {best['part'].rsplit('/', 1)[-1]}")
    # pymupdf cannot overwrite the file it has open: write aside, then swap.
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    os.replace(tmp, out_path)
    return [f"restored {restored} image(s) from DOCX originals"] + notes


def audit(pdf: Path, docx: Path | None = None) -> str:
    """Per-object report. With docx: how many image streams are byte-exact originals."""
    import hashlib

    import pymupdf as fitz

    originals = {}
    if docx is not None and docx.suffix.lower() == ".docx":
        with zipfile.ZipFile(docx) as z:
            for n in z.namelist():
                if n.startswith("word/media/"):
                    originals[hashlib.sha256(z.read(n)).hexdigest()] = n

    doc = fitz.open(pdf)
    seen, exact, recompressed, lossless, other = set(), 0, [], 0, 0
    for page in doc:
        for info in page.get_images(full=True):
            x = info[0]
            if x in seen:
                continue
            seen.add(x)
            filt = doc.xref_get_key(x, "Filter")[1]
            try:
                digest = hashlib.sha256(doc.xref_stream_raw(x)).hexdigest()
            except Exception:
                other += 1
                continue
            if "DCTDecode" in filt or "JPXDecode" in filt:
                if digest in originals:
                    exact += 1
                else:
                    recompressed.append(f"{info[2]}x{info[3]}")
            elif "FlateDecode" in filt:
                lossless += 1
            else:
                other += 1
    doc.close()

    head = f"{pdf.name}: {len(seen)} images"
    if originals:
        verdict = ("OK: no image was re-encoded"
                   if not recompressed else f"WARNING: {len(recompressed)} re-encoded image(s)")
        return (f"{head} | byte-exact JPEG from DOCX: {exact} | lossless (PNG/Flate): {lossless}"
                f" | re-encoded: {len(recompressed)} -> {verdict}")
    return f"{head} | Flate: {lossless} | other: {other}"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("inputs", nargs="+", type=Path)
    ap.add_argument("-o", "--outdir", type=Path, help="output directory (default: beside each source)")
    ap.add_argument("--audit", action="store_true", help="report image codecs for each PDF")
    ap.add_argument("--pdfa", action="store_true", help="export PDF/A-1b")
    ap.add_argument("--no-repair", action="store_true", help="keep Word's raster encoding as is")
    ap.add_argument("-v", "--verbose", action="store_true", help="list every restored image")
    a = ap.parse_args()

    missing = [p for p in a.inputs if not p.is_file()]
    if missing:
        sys.exit("not found: " + ", ".join(map(str, missing)))
    bad = [p for p in a.inputs if p.suffix.lower() not in EXTS]
    if bad:
        sys.exit("unsupported extension: " + ", ".join(map(str, bad)))

    groups: dict[Path, list[Path]] = {}
    for p in a.inputs:
        groups.setdefault(a.outdir or p.parent, []).append(p)

    for outdir, batch in groups.items():
        for src, pdf in zip(batch, word_export(batch, outdir, a.pdfa)):
            if not a.no_repair and src.suffix.lower() == ".docx":
                notes = repair(pdf, src, pdf)
                shown = [n for n in notes if "no DOCX" in n or "failed" in n]
                print(notes[0])
                for n in (notes[1:] if a.verbose else shown):
                    print("  " + n)
            print(pdf)
            if a.audit:
                print("  " + audit(pdf, src))


if __name__ == "__main__":
    main()
