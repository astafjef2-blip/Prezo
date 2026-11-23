# -*- coding: utf-8 -*-
"""Приблизительный рендер превью .pptx без LibreOffice.
Рисует реальную геометрию фигур (python-pptx) поверх фирменных фонов.
Цель — визуальная самопроверка раскладки, а не пиксельная точность.
"""
import json
import os
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Emu

EMU_PX = 1.0 / 9525.0  # EMU -> px @96dpi
W_PX, H_PX = 1280, 720
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
BG = os.path.join("assets", "backgrounds")
OUT = "out/preview"
os.makedirs(OUT, exist_ok=True)

with open("spec.json", encoding="utf-8") as f:
    spec = json.load(f)
typemap = {i: s["type"] for i, s in enumerate(spec["slides"], 1)}
bg_for = {"title": "title.jpg", "section": "section.jpg", "final": "final.jpg"}
DARK = {"title", "section", "final"}


def font(size_pt, bold):
    px = max(8, int(round(size_pt * 96.0 / 72.0)))
    return ImageFont.truetype(FONT_B if bold else FONT, px)


def run_props(runs):
    size, color, bold = 12, None, False
    for r in runs:
        if r.font.size:
            size = r.font.size.pt
        if r.font.bold:
            bold = True
        try:
            if r.font.color and r.font.color.type is not None and r.font.color.rgb:
                color = "#" + str(r.font.color.rgb)
        except Exception:
            pass
    return size, color, bold


def wrap(draw, text, fnt, max_w):
    out = []
    for line in text.split("\n"):
        words = line.split(" ")
        cur = ""
        for w in words:
            test = (cur + " " + w).strip()
            if draw.textlength(test, font=fnt) <= max_w or not cur:
                cur = test
            else:
                out.append(cur)
                cur = w
        out.append(cur)
    return out


prs = Presentation("out/НСПД_Воронежская_область.pptx")
paths = []
for i, slide in enumerate(prs.slides, 1):
    t = typemap.get(i, "content")
    bgname = bg_for.get(t, "content.png")
    img = Image.open(os.path.join(BG, bgname)).convert("RGB").resize((W_PX, H_PX))
    d = ImageDraw.Draw(img, "RGBA")
    dark = t in DARK
    for sh in slide.shapes:
        try:
            x, y = sh.left * EMU_PX, sh.top * EMU_PX
            w, h = sh.width * EMU_PX, sh.height * EMU_PX
        except Exception:
            continue
        if w <= 1 or h <= 1:
            continue
        # outline for layout map
        d.rectangle([x, y, x + w, y + h], outline=(150, 150, 150, 90), width=1)
        if not sh.has_text_frame:
            continue
        txt = sh.text_frame.text.strip()
        if not txt:
            continue
        first = sh.text_frame.paragraphs[0].runs
        size, color, bold = run_props(first) if first else (12, None, False)
        box_fill = (255, 255, 255, 150) if not dark else (8, 18, 28, 165)
        d.rectangle([x, y, x + w, y + h], fill=box_fill)
        fnt = font(size, bold)
        col = color or ("#FFFFFF" if dark else "#39505E")
        lines = wrap(d, txt, fnt, max(20, w - 8))
        lh = fnt.size + 3
        cy = y + 3
        for ln in lines[: max(1, int((h - 4) // lh))]:
            d.text((x + 4, cy), ln, fill=col, font=fnt)
            cy += lh
    p = os.path.join(OUT, f"slide-{i:02d}.png")
    img.save(p)
    paths.append((i, p))

# contact sheet 3 cols
cols = 3
rows = (len(paths) + cols - 1) // cols
th_w, th_h = 384, 216
pad = 12
cs_w = cols * th_w + (cols + 1) * pad
cs_h = rows * th_h + (rows + 1) * pad
sheet = Image.new("RGB", (cs_w, cs_h), (235, 238, 242))
sd = ImageDraw.Draw(sheet)
for idx, (num, p) in enumerate(paths):
    r, c = divmod(idx, cols)
    x0 = pad + c * (th_w + pad)
    y0 = pad + r * (th_h + pad)
    thumb = Image.open(p).resize((th_w, th_h))
    sheet.paste(thumb, (x0, y0))
    sd.rectangle([x0, y0, x0 + th_w, y0 + th_h], outline=(180, 186, 194), width=1)
    sd.text((x0 + 6, y0 + 4), f"{num:02d}", fill=(0, 140, 255), font=ImageFont.truetype(FONT_B, 16))
cs = os.path.join(OUT, "contact_sheet.png")
sheet.save(cs)
print("preview saved:", len(paths), "slides ->", cs)
