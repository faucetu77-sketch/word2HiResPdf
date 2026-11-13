#!/usr/bin/env python3
"""Self-check: build a DOCX containing one uncompressed PNG, convert it, prove the PDF
carries the same pixels and no lossy JPEG re-encode.

    python selftest.py        # needs Word + Pillow + PyMuPDF
"""
import hashlib
import io
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx2pdf import audit, repair, word_export

W, H = 900, 600

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"""

EMU = 9525  # EMU per pixel at 96 dpi
EXT = int(W * EMU), int(H * EMU)
DOC = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body><w:p><w:r><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{EXT[0]}" cy="{EXT[1]}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="1" name="Picture 1"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{EXT[0]}" cy="{EXT[1]}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
</a:graphicData></a:graphic></wp:inline>
</w:drawing></w:r></w:p></w:body></w:document>"""


def make_docx(path: Path, png: bytes) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", DOC)
        z.writestr("word/media/image1.png", png)


def main() -> None:
    import pymupdf

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        tmp = Path(tmp)
        # Noise: worst case for lossy codecs, easy to spot a re-encode.
        rnd = __import__("random").Random(7)
        img = Image.frombytes("RGB", (W, H), bytes(rnd.randrange(256) for _ in range(W * H * 3)))
        buf = io.BytesIO()
        img.save(buf, "PNG", compress_level=0, optimize=False)
        png = buf.getvalue()
        src = tmp / "selftest.docx"
        make_docx(src, png)
        print(f"source PNG: {len(png)} bytes, sha256 {hashlib.sha256(img.tobytes()).hexdigest()[:16]}")

        pdf = word_export([src], tmp, pdfa=False)[0]
        raw_pdf = tmp / "word-raw.pdf"
        pdf.rename(raw_pdf)
        print("Word export:     " + audit(raw_pdf, src))
        notes = repair(raw_pdf, src, pdf)
        raw_pdf.unlink()
        data = pdf.read_bytes()   # keep the file free for cleanup
        print("After repair:    " + audit(pdf, src))
        print(notes[0])

        doc = pymupdf.open(stream=data)
        assert doc.page_count == 1, f"expected 1 page, got {doc.page_count}"
        found = False
        for info in doc[0].get_images(full=True):
            pm = pymupdf.Pixmap(doc, info[0])
            if pm.n == 3 and pm.width == W and pm.height == H:
                found = pm.samples == img.tobytes()
                break
        doc.close()
        assert found, "PDF image is not pixel-identical to the source PNG (re-compressed or rescaled)"
        assert "re-encoded: 0" in audit(pdf, src), "an image was re-encoded"
    print("SELFTEST PASS: 1 page, pixels exact, no re-encoding")


if __name__ == "__main__":
    main()
