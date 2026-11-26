#!/usr/bin/env python3
"""Извлекает графические ассеты из PDF-шаблона Росреестра.

Запуск:  python scripts/extract_assets.py [template/rosreestr_template.pdf]

Создаёт:
  assets/backgrounds/  title.jpg, section.jpg, content.png, final.jpg
  assets/brand/        logo_white.png, check.png
  assets/icons/blue/   <категория>_<имя>.png   (синие градиентные, для белого фона)
  assets/icons/white/  <категория>_<имя>.png   (белые, для синего фона)
  assets/illustrations/*.png
Нужен только при обновлении шаблона: готовые ассеты уже лежат в репозитории.
"""
import os
import sys

import pymupdf
from PIL import Image

sys.path.insert(0, os.path.dirname(__file__))
from icon_names import ICONS  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
A = os.path.join(ROOT, 'assets')
Z = 2  # 960x540 pt -> 1920x1080 px


def blocks(page):
    return [b for b in page.get_text('dict')['blocks'] if b['type'] == 1]


def to_pil(b):
    base = pymupdf.Pixmap(b['image'])
    if base.alpha:
        base = pymupdf.Pixmap(base, 0)
    if base.n > 3:
        base = pymupdf.Pixmap(pymupdf.csRGB, base)
    if b.get('mask'):
        m = pymupdf.Pixmap(b['mask'])
        if m.n > 1:
            m = pymupdf.Pixmap(pymupdf.csGRAY, m)
        if (m.width, m.height) == (base.width, base.height):
            base = pymupdf.Pixmap(base, m)
    mode = 'RGBA' if base.alpha else 'RGB'
    return Image.frombytes(mode, (base.width, base.height), base.samples).convert('RGBA')


def render(src, pn, out, redact_text=True, clip=None, zoom=Z):
    d = pymupdf.open()
    d.insert_pdf(src, from_page=pn, to_page=pn)
    p = d[0]
    if redact_text:
        for b in p.get_text('dict')['blocks']:
            if b['type'] == 0:
                p.add_redact_annot(b['bbox'])
        p.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                           graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    pix = p.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip)
    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    save(img, out)
    return img


def save(img, out):
    os.makedirs(os.path.dirname(out), exist_ok=True)
    if out.endswith('.jpg'):
        img.convert('RGB').save(out, quality=92, optimize=True)
    else:
        img.save(out, optimize=True)


def save_icon(img, out):
    img = img.copy()
    img.thumbnail((200, 200), Image.LANCZOS)
    save(img.quantize(colors=128, method=Image.Quantize.FASTOCTREE), out)


def sig(img):
    a = img.split()[-1].resize((24, 24))
    return list(a.tobytes())


def main(pdf):
    src = pymupdf.open(pdf)
    # ---------------------------------------------------------------- backgrounds
    render(src, 0, f'{A}/backgrounds/title.jpg')
    render(src, 1, f'{A}/backgrounds/section.jpg')
    p14 = render(src, 14, '/tmp/_p14.png')
    bg = Image.new('RGB', p14.size, 'white')
    bg.paste(p14.crop((0, 0, 1920, 140)), (0, 0))
    bg.paste(p14.crop((0, 880, 1920, 1080)), (0, 880))
    save(bg, f'{A}/backgrounds/content.png')
    # final: gradient + circuit + green blob + logo (page 3 without the nested thumbnail)
    b3 = blocks(src[3])
    fin = to_pil(b3[0]).resize((1920, 1080))
    fin.alpha_composite(to_pil(b3[3]).resize(((991 - 50) * Z, 540 * Z)), (50 * Z, 0))
    fin.alpha_composite(to_pil(b3[5]).resize(((926 - 544) * Z, (465 - 84) * Z)), (544 * Z, 84 * Z))
    fin.alpha_composite(to_pil(b3[4]).resize(((227 - 31) * Z, (88 - 16) * Z)), (31 * Z, 16 * Z))
    save(fin, f'{A}/backgrounds/final.jpg')
    # ---------------------------------------------------------------- brand elements
    b0 = blocks(src[0])
    save(to_pil(b0[4]), f'{A}/brand/logo_white.png')
    b4 = blocks(src[4])
    save(to_pil(b4[11]), f'{A}/brand/check.png')
    # ---------------------------------------------------------------- icons
    blue, white = {}, []
    for pn in (8, 9, 10):
        for i, b in enumerate(blocks(src[pn]), 1):
            blue[f'{pn}_{i:02d}'] = to_pil(b)
    for pn in (11, 12, 13):
        for b in blocks(src[pn]):
            img = to_pil(b)
            if img.width < 600:
                white.append(img)
    wsig = [sig(w) for w in white]
    for key, (cat, name, _) in ICONS.items():
        img = blue[key]
        save_icon(img, f'{A}/icons/blue/{cat}_{name}.png')
        s = sig(img)
        best = min(range(len(white)), key=lambda j: sum(abs(x - y) for x, y in zip(s, wsig[j])))
        save_icon(white[best], f'{A}/icons/white/{cat}_{name}.png')
    # ---------------------------------------------------------------- illustrations
    M = 3
    ill = {
        'people_house': (14, (470, 135, 870, 425)),
        'surveyor': (14, (60, 120, 395, 425)),
        'analyst_chart': (15, (620, 105, 945, 425)),
        'specialist_map': (15, (85, 72, 470, 400)),
    }
    for name, (pn, r) in ill.items():
        render(src, pn, f'{A}/illustrations/{name}.png', redact_text=False,
               clip=pymupdf.Rect(*r), zoom=M)
    print('Assets extracted to', A)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'template', 'rosreestr_template.pdf'))
