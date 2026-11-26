#!/usr/bin/env python3
"""Рендерит .pptx в PNG для визуальной проверки (нужен LibreOffice).

    python scripts/render_preview.py deck.pptx [outdir]

Создаёт outdir/slide-01.png … и outdir/contact_sheet.png (все слайды на одном листе).
"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile

CANDIDATES = ['soffice', 'libreoffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice']


def find_soffice():
    for c in CANDIDATES:
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    raise SystemExit('LibreOffice не найден. Установите его (apt install libreoffice-impress / '
                     'brew install --cask libreoffice / libreoffice.org) или проверьте файл в PowerPoint вручную.')


def render(pptx, outdir=None, dpi=80):
    import pymupdf
    from PIL import Image
    outdir = outdir or os.path.splitext(pptx)[0] + '_preview'
    os.makedirs(outdir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    subprocess.run([find_soffice(), '--headless', '--convert-to', 'pdf', '--outdir', tmp, pptx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = glob.glob(os.path.join(tmp, '*.pdf'))[0]
    doc = pymupdf.open(pdf)
    ims = []
    for p in doc:
        pix = p.get_pixmap(dpi=dpi)
        fn = os.path.join(outdir, f'slide-{p.number + 1:02d}.png')
        pix.save(fn)
        ims.append(Image.open(fn).convert('RGB'))
    w, h = ims[0].size
    cols = 3
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * w // 2 + (cols + 1) * 8, rows * h // 2 + (rows + 1) * 8), (225, 230, 236))
    for i, im in enumerate(ims):
        sheet.paste(im.resize((w // 2, h // 2)), (8 + (i % cols) * (w // 2 + 8), 8 + (i // cols) * (h // 2 + 8)))
    sheet.save(os.path.join(outdir, 'contact_sheet.png'))
    shutil.copy(pdf, os.path.join(outdir, os.path.basename(pdf)))
    print(f'✔ Превью: {outdir}  ({len(ims)} слайдов)')
    return outdir


if __name__ == '__main__':
    render(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
