#!/usr/bin/env python3
"""Check the repaired PDF against Word's own export: layout, text and vector art untouched.

    python verify.py book.docx raw-word.pdf final.pdf      # compare against untouched export
    python verify.py book.docx final.pdf                   # audit images only

With two PDFs it proves the repair changed pixels only inside image rectangles and that
page geometry and text are identical. With one it checks no image was re-encoded.
"""
import pathlib
import sys

import numpy as np
import pymupdf

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from docx2pdf import audit


def compare(raw: pathlib.Path, fixed: pathlib.Path) -> int:
    a, b = pymupdf.open(raw), pymupdf.open(fixed)
    if a.page_count != b.page_count:
        print(f"FAIL: page count {a.page_count} -> {b.page_count}")
        return 1
    bad = []
    text_changed = rect_changed = strayed = 0
    for i in range(a.page_count):
        pa, pb = a[i], b[i]
        if pa.rect != pb.rect:
            rect_changed += 1
        if pa.get_text() != pb.get_text():
            text_changed += 1
        ra, rb = pa.get_pixmap(dpi=72), pb.get_pixmap(dpi=72)
        if (ra.width, ra.height) != (rb.width, rb.height):
            bad.append(i + 1)
            continue
        xa = np.frombuffer(ra.samples, np.uint8).reshape(ra.height, ra.width, 3).astype(np.int16)
        xb = np.frombuffer(rb.samples, np.uint8).reshape(rb.height, rb.width, 3).astype(np.int16)
        diff = np.abs(xa - xb).max(axis=2) > 8
        if not diff.any():
            continue
        mask = np.zeros(diff.shape, bool)
        scale = ra.width / pa.rect.width
        for img in pa.get_image_info():
            x0, y0, x1, y1 = img["bbox"]
            mask[max(int(y0 * scale) - 2, 0):int(y1 * scale) + 3,
                 max(int(x0 * scale) - 2, 0):int(x1 * scale) + 3] = True
        if int((diff & ~mask).sum()) > 200:
            strayed += 1
            bad.append(i + 1)

    print(f"pages: {a.page_count}")
    print(f"pages whose page box changed: {rect_changed}")
    print(f"pages whose extracted text changed: {text_changed}")
    print(f"pages with pixel changes outside image areas: {strayed}")
    ok = not (rect_changed or text_changed or strayed)
    print("OK: layout, text and vector content untouched" if ok else f"FAIL on pages {bad}")
    return 0 if ok else 1


def main() -> None:
    args = [pathlib.Path(p) for p in sys.argv[1:]]
    if len(args) == 3:
        print(audit(args[2], args[0]))
        sys.exit(compare(args[1], args[2]))
    if len(args) == 2:
        print(audit(args[1], args[0]))
        return
    sys.exit(__doc__)


if __name__ == "__main__":
    main()
