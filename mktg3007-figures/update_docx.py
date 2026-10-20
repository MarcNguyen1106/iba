#!/usr/bin/env python3
"""Put the redesigned figures into a copy of the report without touching anything else.

Only two things change:
  * word/media/image1.png (Figure 1) and word/media/image2.png (Figure C1) are
    replaced under the same names;
  * the matching <wp:extent> and <a:ext> cy values in word/document.xml are set
    so the displayed aspect ratio equals the PNG's (width cx is kept).
Every other zip entry is copied byte for byte, in the original order.

Usage: python update_docx.py original.docx figure1.png figureC1.png output.docx
"""
import re
import sys
import zipfile

from PIL import Image

src, fig1, figc1, dst = sys.argv[1:5]
if src == dst:
    sys.exit("refusing to overwrite the original; give a new output name")

# relationship id -> (media part, new png)
FIGURES = {"rId7": ("word/media/image1.png", fig1), "rId37": ("word/media/image2.png", figc1)}

with zipfile.ZipFile(src) as zin:
    doc = zin.read("word/document.xml").decode("utf8")
    rels = zin.read("word/_rels/document.xml.rels").decode("utf8")
    for rid, (part, _) in FIGURES.items():
        target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1)
        assert "word/" + target == part, (rid, target)

    new_doc = doc
    for rid, (part, png) in FIGURES.items():
        w, h = Image.open(png).size
        # the <w:drawing> that embeds this relationship
        m = re.search(rf'<w:drawing>(?:(?!</w:drawing>).)*r:embed="{rid}"(?:(?!</w:drawing>).)*</w:drawing>',
                      new_doc, re.S)
        drawing = m.group(0)
        cx = int(re.search(r'<wp:extent cx="(\d+)"', drawing).group(1))
        cy = round(cx * h / w)
        updated = re.sub(r'(<wp:extent cx="\d+" cy=")\d+(")', rf"\g<1>{cy}\2", drawing)
        updated = re.sub(r'(<a:ext cx="\d+" cy=")\d+(")', rf"\g<1>{cy}\2", updated)
        new_doc = new_doc[:m.start()] + updated + new_doc[m.end():]
        print(f"{part}: {w} x {h} px -> {cx / 360000:.2f} x {cy / 360000:.2f} cm "
              f"({w / (cx / 360000 / 2.54):.0f} dpi at print size)")

    replacements = {part: open(png, "rb").read() for part, png in FIGURES.values()}
    replacements["word/document.xml"] = new_doc.encode("utf8")

    with zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = replacements.get(info.filename, None)
            zout.writestr(info, data if data is not None else zin.read(info.filename))

# Prove nothing but the two cy values changed in document.xml.
strip = lambda s: re.sub(r'(<(?:wp:extent|a:ext) cx="\d+" cy=")\d+"', r'\1"', s)  # noqa: E731
assert strip(doc) == strip(new_doc), "document.xml changed beyond the image heights"
print("document.xml: only the two image heights changed")
print("wrote", dst)
