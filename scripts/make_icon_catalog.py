#!/usr/bin/env python3
"""Генерирует каталог иконок: docs/icons.md и превью docs/img/icons_*.png"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(__file__))
from icon_names import ICONS, CATEGORIES  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FONT_PATHS = ['/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 'C:/Windows/Fonts/arial.ttf',
              '/Library/Fonts/Arial.ttf']


def font(size):
    for p in FONT_PATHS:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def sheet(variant, bg):
    cols, cw, ch = 8, 150, 132
    items = sorted(ICONS.values(), key=lambda v: (list(CATEGORIES).index(v[0]), v[1]))
    rows = []
    for cat in CATEGORIES:
        its = [i for i in items if i[0] == cat]
        rows.append(('h', cat))
        for k in range(0, len(its), cols):
            rows.append(('i', its[k:k + cols]))
    H = sum(34 if r[0] == 'h' else ch for r in rows) + 20
    img = Image.new('RGB', (cols * cw + 20, H), bg)
    d = ImageDraw.Draw(img)
    fg = (255, 255, 255) if variant == 'white' else (57, 80, 94)
    y = 10
    for kind, val in rows:
        if kind == 'h':
            d.text((14, y + 8), CATEGORIES[val].upper(), font=font(16), fill=fg)
            y += 34
            continue
        for j, (cat, name, _) in enumerate(val):
            ic = Image.open(f'{ROOT}/assets/icons/{variant}/{cat}_{name}.png').convert('RGBA')
            ic.thumbnail((84, 84))
            x = 10 + j * cw
            img.paste(ic, (x + (cw - ic.width) // 2, y + 6 + (84 - ic.height) // 2), ic)
            label = f'{cat}_{name}'
            f = font(11)
            while d.textlength(label, font=f) > cw - 6 and f.size > 7:
                f = font(f.size - 1)
            d.text((x + (cw - d.textlength(label, font=f)) / 2, y + 98), label, font=f, fill=fg)
        y += ch
    os.makedirs(f'{ROOT}/docs/img', exist_ok=True)
    img.save(f'{ROOT}/docs/img/icons_{variant}.png', optimize=True)


def markdown():
    out = ['# Каталог иконок', '',
           'Все иконки извлечены из корпоративного шаблона Росреестра. Каждая есть в двух вариантах:',
           '', '* `assets/icons/blue/<id>.png` — синий градиент, **для белого фона** (основной вариант);',
           '* `assets/icons/white/<id>.png` — белые, **для синих и градиентных плашек**.', '',
           'В спецификации слайда иконка указывается по `id` (без `.png`), например `"icon": "maps_cadastral_map"`.',
           'Белый вариант библиотека подставляет сама, если иконка стоит на синем фоне.', '',
           '![Синие иконки](img/icons_blue.png)', '']
    for cat, title in CATEGORIES.items():
        out += [f'## {title}', '', '| id | Значение / когда использовать |', '|---|---|']
        for key, (c, name, desc) in sorted(ICONS.items(), key=lambda kv: kv[1][1]):
            if c == cat:
                out.append(f'| `{c}_{name}` | {desc} |')
        out.append('')
    out += ['## Белый вариант', '', '![Белые иконки](img/icons_white.png)', '']
    open(f'{ROOT}/docs/icons.md', 'w', encoding='utf-8').write('\n'.join(out))


if __name__ == '__main__':
    sheet('blue', (255, 255, 255))
    sheet('white', (0, 140, 255))
    markdown()
    print('docs/icons.md and docs/img/icons_*.png generated')
