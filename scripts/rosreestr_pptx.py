#!/usr/bin/env python3
"""
rosreestr_pptx — библиотека для сборки презентаций PowerPoint в фирменном стиле Росреестра.

Пример:
    from rosreestr_pptx import Deck
    d = Deck()
    d.title(title="Итоги работы за 2026 год", kicker="Коллегия", author="И.И. Иванов",
            role="руководитель Управления")
    d.cards(title="Ключевые изменения", cards=[{"icon": "documents_scroll_seal",
            "title": "Закон № 123-ФЗ", "text": "Описание"}])
    d.final()
    d.save("out.pptx")

Все координаты — в пунктах (pt) на холсте 960×540 (16:9, 13,33×7,5 дюйма).
Разметка текста: **жирный акцент** (синий, на тёмном фоне — белый жирный), !!предупреждение!! (оранжевый).
"""
import difflib
import math
import os
import re

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
ASSETS = os.path.join(ROOT, 'assets')

# ----------------------------------------------------------------------------- палитра
BLUE = RGBColor(0x00, 0x8C, 0xFF)      # основной фирменный синий: заголовки, акценты
BLUE2 = RGBColor(0x18, 0x90, 0xFA)     # синий плашек и рамок
NAVY = RGBColor(0x00, 0x47, 0x7F)      # тёмно-синий
GREEN = RGBColor(0x00, 0xCC, 0x99)     # бирюзово-зелёный: рост, позитив
TEAL = RGBColor(0x2E, 0xD3, 0xC0)      # конец фирменного градиента
MINT = RGBColor(0x5F, 0xD8, 0xC0)      # рамки нумерованных блоков
LGREEN = RGBColor(0xDF, 0xF8, 0xEA)    # светло-зелёная подложка
DGREEN = RGBColor(0x00, 0x8F, 0x6B)    # текст на светло-зелёном
LBLUE = RGBColor(0xE3, 0xF2, 0xFE)     # светло-голубая подложка карточек
XBLUE = RGBColor(0xF4, 0xF9, 0xFF)     # очень светлый голубой (чередование строк)
PBLUE = RGBColor(0x99, 0xD2, 0xFF)     # бледно-синий: прошлый период, норматив
ORANGE = RGBColor(0xFF, 0x79, 0x00)    # оранжевый: снижение, предупреждение
TEXT = RGBColor(0x39, 0x50, 0x5E)      # основной текст
GRAY = RGBColor(0x59, 0x59, 0x59)      # вторичный текст
BORDER = RGBColor(0xBF, 0xE0, 0xFF)    # рамки карточек
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

COLORS = {'blue': BLUE, 'blue2': BLUE2, 'navy': NAVY, 'green': GREEN, 'teal': TEAL, 'orange': ORANGE,
          'pale': PBLUE, 'light': LBLUE, 'lightgreen': LGREEN, 'text': TEXT, 'gray': GRAY, 'white': WHITE}
SERIES = [PBLUE, BLUE, GREEN, ORANGE, NAVY, TEAL]
GRAD = (BLUE, TEAL)  # фирменный градиент слева направо

FONT = 'Arial'
W, H = 960, 540
X0, X1 = 30, 930          # горизонтальные границы контента
Y0, Y1 = 82, 462          # вертикальные границы контента (под шапкой / над нижней полосой)
GAP = 15

ILLUSTRATIONS = {'people_house', 'surveyor', 'analyst_chart', 'specialist_map'}


def color(c, default=BLUE):
    if c is None:
        return default
    if isinstance(c, RGBColor):
        return c
    if isinstance(c, str) and c.startswith('#'):
        return RGBColor.from_string(c[1:].upper())
    return COLORS.get(c, default)


# ----------------------------------------------------------------------------- разметка
_TOKEN = re.compile(r'(\*\*.+?\*\*|!!.+?!!)')


def parse_markup(s, accent=BLUE):
    """'**жирный** обычный !!оранжевый!!' -> [(text, opts), ...]"""
    runs = []
    for part in _TOKEN.split(str(s)):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            runs.append((part[2:-2], {'bold': True, 'color': accent}))
        elif part.startswith('!!') and part.endswith('!!'):
            runs.append((part[2:-2], {'bold': True, 'color': ORANGE}))
        else:
            runs.append((part, {}))
    return runs


def plain(s):
    return _TOKEN.sub(lambda m: m.group(0)[2:-2], str(s))


def est_height(paras, w, size, line=1.18, spacing=0, bold=False):
    """Грубая оценка высоты текста (pt) для шрифта Arial."""
    k = 0.61 if bold else 0.55
    total = 0
    for i, p in enumerate(paras):
        txt, sz = p
        cpl = max(1, int(w / (sz * k)))
        words = plain(txt).split()
        lines, cur = 1, 0
        for wd in words:
            L = len(wd) + (1 if cur else 0)
            if cur + L > cpl:
                lines += 1
                cur = len(wd)
            else:
                cur += L
        total += lines * sz * line + (spacing if i else 0)
    return total


# ----------------------------------------------------------------------------- колода
class Deck:
    def __init__(self, assets=ASSETS):
        self.assets = assets
        self.prs = Presentation()
        self.prs.slide_width = Pt(W)
        self.prs.slide_height = Pt(H)
        self._blank = self.prs.slide_layouts[6]
        self.warnings = []
        self.n = 0
        self._icons = None

    # ------------------------------------------------------------------ базовые примитивы
    def _slide(self, bg, notes=None, number=True, title=None, title_size=24):
        self.n += 1
        s = self.prs.slides.add_slide(self._blank)
        path = os.path.join(self.assets, 'backgrounds', bg)
        s.shapes.add_picture(path, 0, 0, self.prs.slide_width, self.prs.slide_height)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        if title is not None:
            self.text(s, 72, 8, 860, 52, title, size=title_size, color=BLUE, bold=True,
                      anchor='middle', fit=17, what=f'заголовок слайда {self.n}')
        if number and bg == 'content.png':
            self.text(s, 880, 515, 66, 18, str(self.n), size=11, color=WHITE, bold=True, align='right')
        return s

    def box(self, s, x, y, w, h, fill=WHITE, line=None, lw=1.25, shape='round', radius=10, grad=None,
            shadow=False):
        shp = {'round': MSO_SHAPE.ROUNDED_RECTANGLE, 'rect': MSO_SHAPE.RECTANGLE,
               'arrow': MSO_SHAPE.RIGHT_ARROW, 'oval': MSO_SHAPE.OVAL,
               'chevron': MSO_SHAPE.CHEVRON, 'pentagon': MSO_SHAPE.PENTAGON}[shape]
        sh = s.shapes.add_shape(shp, Pt(x), Pt(y), Pt(w), Pt(h))
        if shape == 'round':
            sh.adjustments[0] = min(0.5, radius / max(1, min(w, h)))
        if grad:
            sh.fill.gradient()
            sh.fill.gradient_angle = 0
            st = sh.fill.gradient_stops
            st[0].color.rgb, st[0].position = color(grad[0]), 0
            st[1].color.rgb, st[1].position = color(grad[1]), 1.0
        elif fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = color(fill)
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = color(line)
            sh.line.width = Pt(lw)
        ef = etree.SubElement(sh._element.spPr, qn('a:effectLst'))
        if shadow:
            sd = etree.SubElement(ef, qn('a:outerShdw'), blurRad='63500', dist='12700', dir='5400000',
                                  algn='t', rotWithShape='0')
            c = etree.SubElement(sd, qn('a:srgbClr'), val='0A4A80')
            etree.SubElement(c, qn('a:alpha'), val='14000')
        st = sh._element.find(qn('p:style'))
        if st is not None:
            st.find(qn('a:effectRef')).set('idx', '0')
        sh.text_frame.text = ''
        return sh

    def text(self, s, x, y, w, h, paras, size=12, color=TEXT, bold=False, align='left', anchor='top',
             spacing=None, line_spacing=None, fit=None, accent=None, what=None):
        """paras: str | list[str | dict(text,size,bold,color)]. fit=минимальный кегль для автоподбора."""
        col = globals()['color'](color, TEXT)
        if accent is None:
            accent = WHITE if col == WHITE else BLUE
        if isinstance(paras, (str, int, float)):
            paras = [str(paras)]
        norm = []
        for p in paras:
            if isinstance(p, dict):
                norm.append(dict(p))
            else:
                norm.append({'text': str(p)})
        scale = 1.0
        if fit:
            sp = spacing or 0
            while True:
                est = est_height([(p['text'], p.get('size', size) * scale) for p in norm], w,
                                 size, spacing=sp * scale, bold=bold or any(p.get('bold') for p in norm))
                if est <= h or size * scale <= fit:
                    break
                scale -= 0.04
            if est > h * 1.08:
                self.warnings.append(f'Слайд {self.n}: текст может не поместиться ({what or plain(norm[0]["text"])[:40]}…). '
                                     f'Сократите формулировку.')
        tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(0)
        tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE,
                              'bottom': MSO_ANCHOR.BOTTOM}[anchor]
        for i, p in enumerate(norm):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[align]
            if spacing and i:
                para.space_before = Pt(spacing * scale)
            if line_spacing:
                para.line_spacing = line_spacing
            pc = globals()['color'](p.get('color'), col) if p.get('color') else col
            for t, o in parse_markup(p['text'], accent):
                r = para.add_run()
                r.text = t
                f = r.font
                f.name = FONT
                f.size = Pt(round(o.get('size', p.get('size', size)) * scale * 2) / 2)
                f.bold = o.get('bold', p.get('bold', bold))
                f.italic = p.get('italic', False)
                f.color.rgb = o.get('color', pc)
        return tb

    def _icon_path(self, name, white=False):
        if not name:
            return None
        if os.path.exists(name):
            return name
        variant = 'white' if white else 'blue'
        p = os.path.join(self.assets, 'icons', variant, name + '.png')
        if os.path.exists(p):
            return p
        if self._icons is None:
            self._icons = [f[:-4] for f in os.listdir(os.path.join(self.assets, 'icons', 'blue'))]
        sug = difflib.get_close_matches(name, self._icons, n=1, cutoff=0.4)
        if sug:
            self.warnings.append(f'Слайд {self.n}: иконка "{name}" не найдена, использована "{sug[0]}"')
            return os.path.join(self.assets, 'icons', variant, sug[0] + '.png')
        self.warnings.append(f'Слайд {self.n}: иконка "{name}" не найдена — пропущена')
        return None

    def icon(self, s, name, x, y, size, white=False):
        p = self._icon_path(name, white)
        if not p:
            return None
        iw, ih = Image.open(p).size
        k = size / max(iw, ih)
        return s.shapes.add_picture(p, Pt(x + (size - iw * k) / 2), Pt(y + (size - ih * k) / 2),
                                    Pt(iw * k), Pt(ih * k))

    def picture(self, s, name, x, y, w, h):
        """Вписывает картинку/иллюстрацию в прямоугольник с сохранением пропорций (по центру)."""
        p = name
        if name in ILLUSTRATIONS:
            p = os.path.join(self.assets, 'illustrations', name + '.png')
        if not os.path.exists(p):
            self.warnings.append(f'Слайд {self.n}: изображение "{name}" не найдено')
            return None
        iw, ih = Image.open(p).size
        k = min(w / iw, h / ih)
        return s.shapes.add_picture(p, Pt(x + (w - iw * k) / 2), Pt(y + (h - ih * k) / 2),
                                    Pt(iw * k), Pt(ih * k))

    # ------------------------------------------------------------------ компоненты
    def pill(self, s, x, y, w, h, t, fill='light', color_='blue', size=12, bold=True, line=None):
        self.box(s, x, y, w, h, fill=fill, line=line, radius=h / 2)
        self.text(s, x + 10, y, w - 20, h, t, size=size, color=color_, bold=bold, align='center',
                  anchor='middle', fit=9)

    def header_tab(self, s, x, y, w, h, t, icon=None, size=13):
        """Синяя плашка-заголовок карточки с белой иконкой (как «ДОСТИГНУТЫЙ РЕЗУЛЬТАТ» в шаблоне)."""
        self.box(s, x, y, w, h, grad=(BLUE, BLUE2), radius=6)
        tx = x + 12
        if icon:
            self.icon(s, icon, x + 8, y + 5, h - 10, white=True)
            tx = x + h + 4
        self.text(s, tx, y, w - (tx - x) - 8, h, plain(t).upper(), size=size, color=WHITE, bold=True,
                  anchor='middle', fit=9)

    def check(self, s, x, y, size=22):
        s.shapes.add_picture(os.path.join(self.assets, 'brand', 'check.png'), Pt(x), Pt(y), Pt(size), Pt(size))

    def arrow(self, s, x, y, w=28, h=22):
        self.box(s, x, y, w, h, shape='arrow', grad=(PBLUE, BLUE))

    def kpi(self, s, x, y, w, h, k, dark=False):
        """k: {value, label, delta?, delta_color?, icon?}"""
        if dark:
            self.box(s, x, y, w, h, grad=GRAD, radius=12)
        else:
            self.box(s, x, y, w, h, fill=LBLUE, radius=12)
        tx = x + 16
        isz = min(46, h - 24)
        if k.get('icon') and w >= 200:
            self.icon(s, k['icon'], x + 14, y + (h - isz) / 2, isz, white=dark)
            tx = x + 14 + isz + 12
        vsz = 28 if h >= 100 else (24 if h >= 80 else 20)
        dcol = k.get('delta_color')
        if not dcol and k.get('delta'):
            dcol = ORANGE if str(k['delta']).strip()[:1] in '−-▼' else GREEN
        paras = [{'text': str(k['value']), 'size': vsz, 'bold': True, 'color': WHITE if dark else BLUE}]
        if k.get('label'):
            paras.append({'text': k['label'], 'size': 11.5, 'color': WHITE if dark else TEXT})
        if k.get('delta'):
            paras.append({'text': k['delta'], 'size': 12, 'bold': True,
                          'color': WHITE if dark else color(dcol)})
        self.text(s, tx, y + 6, w - (tx - x) - 12, h - 12, paras, anchor='middle', spacing=2, fit=8,
                  what='KPI ' + str(k['value']))

    def kpi_column(self, s, kpis, x=X0, y=Y0, w=290, h=None):
        h = h or (Y1 - y)
        n = len(kpis)
        if not n:
            return
        g = 12
        kh = (h - g * (n - 1)) / n
        for i, k in enumerate(kpis):
            self.kpi(s, x, y + i * (kh + g), w, kh, k, dark=k.get('dark', False))

    def card(self, s, x, y, w, h, c, highlight=False, fill=WHITE):
        """c: {icon?, title, text}"""
        hl = highlight or c.get('highlight')
        self.box(s, x, y, w, h, fill=fill, line=BLUE2 if hl else BORDER, lw=1.5 if hl else 1.25, radius=12)
        isz = 46 if h >= 120 else 36
        tx = x + 16
        if c.get('icon'):
            self.icon(s, c['icon'], x + 14, y + 14, isz)
            tx = x + 14 + isz + 10
        self.text(s, tx, y + 12, w - (tx - x) - 12, isz + 4, c.get('title', ''), size=14, color=BLUE, bold=True,
                  anchor='middle', fit=10, what='заголовок карточки')
        if c.get('text'):
            ty = y + 12 + isz + 12
            self.text(s, x + 16, ty, w - 30, y + h - ty - 10, c['text'], size=13, color=TEXT, fit=9,
                      line_spacing=1.05, what='текст карточки «' + plain(c.get('title', ''))[:30] + '»')

    def progress(self, s, x, y, w, it, h=24, label_w=72):
        label, val = it.get('label', ''), float(it['value'])
        mx = float(it.get('max', 100))
        col = color(it.get('color'), BLUE)
        track = LGREEN if col == GREEN else LBLUE
        self.text(s, x, y, label_w - 4, h, label, size=12, color=GRAY, anchor='middle', fit=8)
        bx, bw = x + label_w, w - label_w
        self.box(s, bx, y, bw, h, fill=track, radius=h / 2)
        fw = max(h * 2.2, bw * min(val, mx) / mx)
        self.box(s, bx, y, fw, h, fill=col, radius=h / 2)
        t = it.get('text') or (f'{val:g}%'.replace('.', ','))
        self.text(s, bx + 10, y, fw - 14, h, t, size=11.5, color=WHITE, bold=True, anchor='middle')

    def progress_box(self, s, x, y, w, h, p):
        """p: {title, items:[{label,value,color?,text?,max?}], badge?:{value,caption}}"""
        self.box(s, x, y, w, h, fill=WHITE, line=BORDER, radius=12)
        self.text(s, x + 18, y + 10, w - 36, 22, p.get('title', ''), size=13, color=BLUE, bold=True, fit=10)
        items = p.get('items', [])
        bw = w - 36 - (150 if p.get('badge') else 0)
        top = y + 42
        avail = y + h - 12 - top
        step = avail / max(1, len(items))
        bh = min(26, step - 8)
        for i, it in enumerate(items):
            self.progress(s, x + 18, top + i * step + (step - bh) / 2, bw, it, h=bh)
        if p.get('badge'):
            b = p['badge']
            bh2 = min(110, h - 30)
            bx = x + w - 150
            self.box(s, bx, y + (h - bh2) / 2, 132, bh2, grad=GRAD, radius=12)
            self.text(s, bx, y + (h - bh2) / 2, 132, bh2,
                      [{'text': b['value'], 'size': 30, 'bold': True}, {'text': b.get('caption', ''), 'size': 12}],
                      color=WHITE, align='center', anchor='middle', fit=12)

    def table(self, s, x, y, w, t, max_h=None, max_row_h=36):
        """t: {columns:[...], rows:[[...]], col_widths?:[...], bold_rows?:[i], delta_cols?:[i],
               accent_cols?:[i], highlight_rows?:[i]}  (индексы строк — среди rows, с 0)"""
        cols = t['columns']
        rows = t['rows']
        n = len(rows) + 1
        max_h = max_h or (Y1 - y)
        rh = min(max_row_h, max_h / n)
        fsz = 14 if rh >= 42 else (12.5 if rh >= 30 else (11 if rh >= 24 else 10))
        cw = t.get('col_widths') or ([2] + [1] * (len(cols) - 1))
        tot = sum(cw)
        gt = s.shapes.add_table(n, len(cols), Pt(x), Pt(y), Pt(w), Pt(rh * n)).table
        tblPr = gt._tbl.tblPr
        tblPr.set('bandRow', '0')
        tblPr.set('firstRow', '0')
        for i, c in enumerate(cw):
            gt.columns[i].width = Pt(w * c / tot)
        bold_rows = set(t.get('bold_rows', []))
        hl_rows = set(t.get('highlight_rows', t.get('bold_rows', [])))
        delta_cols = set(t.get('delta_cols', []))
        accent_cols = set(t.get('accent_cols', []))
        for r in range(n):
            gt.rows[r].height = Pt(rh)
            vals = cols if r == 0 else rows[r - 1]
            for c in range(len(cols)):
                val = str(vals[c]) if c < len(vals) else ''
                cell = gt.cell(r, c)
                cell.margin_left = cell.margin_right = Pt(8)
                cell.margin_top = cell.margin_bottom = Pt(2)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.text = ''
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                run = p.add_run()
                run.text = val
                f = run.font
                f.name = FONT
                f.size = Pt(fsz)
                cell.fill.solid()
                ri = r - 1
                if r == 0:
                    cell.fill.fore_color.rgb = BLUE
                    f.bold, f.color.rgb = True, WHITE
                else:
                    cell.fill.fore_color.rgb = LBLUE if ri in hl_rows else (WHITE if ri % 2 == 0 else XBLUE)
                    f.bold = ri in bold_rows or c in accent_cols
                    f.color.rgb = TEXT
                    if c in accent_cols:
                        f.color.rgb = BLUE
                    if c in delta_cols:
                        v = val.strip()
                        f.color.rgb = GREEN if v.startswith('+') else (ORANGE if v[:1] in '−-' else TEXT)
        return rh * n

    def chart(self, s, x, y, w, h, c):
        """c: {type: column|bar|stacked_column|stacked_bar|line|pie|doughnut, categories, series:[{name, values,
               color?}], number_format?, legend?, labels?, max?, colors?}"""
        kind = c.get('type', 'column')
        types = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED,
                 'stacked_column': XL_CHART_TYPE.COLUMN_STACKED, 'stacked_bar': XL_CHART_TYPE.BAR_STACKED,
                 'line': XL_CHART_TYPE.LINE_MARKERS, 'pie': XL_CHART_TYPE.PIE,
                 'doughnut': XL_CHART_TYPE.DOUGHNUT}
        cd = CategoryChartData()
        cd.categories = c['categories']
        for ser in c['series']:
            cd.add_series(ser.get('name', ''), ser['values'])
        gf = s.shapes.add_chart(types[kind], Pt(x), Pt(y), Pt(w), Pt(h), cd)
        ch = gf.chart
        ch.font.name = FONT
        ch.font.size = Pt(11)
        ch.font.color.rgb = TEXT
        pie = kind in ('pie', 'doughnut')
        show_legend = c.get('legend', pie or len(c['series']) > 1)
        ch.has_legend = bool(show_legend)
        if show_legend:
            ch.legend.position = XL_LEGEND_POSITION.RIGHT if pie else XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(11)
            ch.legend.font.color.rgb = GRAY
        plot = ch.plots[0]
        nf = c.get('number_format', '# ##0' if not pie else '0%')
        if c.get('labels', not (kind == 'line' and len(c['series']) > 1)):
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(11)
            dl.font.bold = True
            dl.font.color.rgb = WHITE if (pie or kind.startswith('stacked')) else TEXT
            dl.number_format = nf
            dl.number_format_is_linked = False
            if pie:
                dl.show_percentage = c.get('percent', True)
                dl.show_value = not c.get('percent', True)
            elif not kind.startswith('stacked') and kind != 'line':
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            elif kind == 'line':
                dl.position = XL_LABEL_POSITION.ABOVE
        palette = [color(v) for v in c.get('colors', [])] or SERIES
        if pie:
            ser = plot.series[0]
            for i in range(len(c['categories'])):
                pt = ser.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = [BLUE, GREEN, PBLUE, ORANGE, NAVY, TEAL][i % 6] \
                    if not c.get('colors') else palette[i % len(palette)]
        else:
            va = ch.value_axis
            va.has_major_gridlines = kind == 'line'
            if va.has_major_gridlines:
                va.major_gridlines.format.line.color.rgb = RGBColor(0xE6, 0xEC, 0xF2)
            va.visible = kind == 'line'
            va.tick_labels.font.size = Pt(10)
            va.minimum_scale = c.get('min', 0)
            mx = c.get('max')
            if mx is None and not kind.startswith('stacked'):
                mx = max(max(sr['values']) for sr in c['series']) * 1.15
            if mx:
                va.maximum_scale = mx
            ca = ch.category_axis
            ca.tick_labels.font.size = Pt(11.5)
            ca.tick_labels.font.color.rgb = TEXT
            ca.format.line.color.rgb = RGBColor(0xD0, 0xD8, 0xE0)
            ca.has_major_gridlines = False
            if kind not in ('line',):
                plot.gap_width = 70
                if not kind.startswith('stacked'):
                    plot.overlap = -10
            for i, ser in enumerate(plot.series):
                col = color(c['series'][i].get('color'), palette[i % len(palette)]) \
                    if c['series'][i].get('color') else palette[i % len(palette)]
                if kind == 'line':
                    ser.format.line.color.rgb = col
                    ser.format.line.width = Pt(2.5)
                    ser.smooth = False
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.rgb = col
                    ser.marker.format.line.color.rgb = col
                else:
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = col
        return ch

    def panel(self, s, x, y, w, h, title=None, fill=WHITE):
        self.box(s, x, y, w, h, fill=fill, line=BORDER if fill == WHITE else None, radius=12)
        if title:
            self.text(s, x + 18, y + 10, w - 36, 22, title, size=13, color=BLUE, bold=True, fit=10)
            return y + 36
        return y + 12

    def callout(self, s, x, y, w, h, c, style='filled'):
        """c: {icon?, text}. style: filled (фирменный градиент) | outline | light"""
        if style == 'filled':
            self.box(s, x, y, w, h, grad=GRAD, radius=12)
            col = WHITE
        elif style == 'light':
            self.box(s, x, y, w, h, fill=LBLUE, radius=12)
            col = TEXT
        else:
            self.box(s, x, y, w, h, fill=WHITE, line=BLUE2, lw=1.5, radius=12)
            col = TEXT
        tx = x + 18
        if c.get('icon'):
            isz = min(52, h - 18)
            self.icon(s, c['icon'], x + 14, y + (h - isz) / 2, isz, white=(style == 'filled'))
            tx = x + 14 + isz + 14
        self.text(s, tx, y + 4, x + w - tx - 14, h - 8, c['text'], size=c.get('size', 13.5), color=col,
                  anchor='middle', fit=9, what='выноска')

    def intro(self, s, text, y=74, h=30):
        if text:
            self.text(s, X0, y, X1 - X0, h, text, size=13, color=TEXT, anchor='middle', fit=10, what='вводная строка')
            return y + h + 8
        return Y0

    # ================================================================== СЛАЙДЫ
    def title(self, title, kicker=None, author=None, role=None, notes=None, **_):
        """Титульный слайд. title ≤ 90 символов, kicker — мероприятие/дата."""
        s = self._slide('title.jpg', notes, number=False)
        if kicker:
            self.text(s, 48, 156, 470, 26, kicker, size=14, color=GREEN, bold=True, fit=10)
        self.text(s, 48, 186, 480, 172, title, size=30, color=BLUE, bold=True, anchor='middle', fit=18,
                  line_spacing=1.0, what='заголовок титула')
        paras = []
        if author:
            paras.append({'text': author, 'size': 18, 'bold': True})
        if role:
            paras.append({'text': role, 'size': 13})
        if paras:
            self.text(s, 48, 400, 470, 70, paras, color=WHITE, spacing=4, fit=10)
        return s

    def section(self, title, number=None, subtitle=None, notes=None, **_):
        """Разделитель раздела."""
        s = self._slide('section.jpg', notes, number=False)
        y = 205
        if number is not None:
            self.text(s, 36, 160, 200, 44, f'{int(number):02d}' if str(number).isdigit() else str(number),
                      size=36, color=PBLUE, bold=True)
        self.text(s, 36, y, 500, 70, title, size=28, color=BLUE, bold=True, anchor='middle', fit=18)
        if subtitle:
            self.text(s, 36, y + 76, 500, 40, subtitle, size=14, color=TEXT, fit=10)
        return s

    def final(self, title='Спасибо за внимание!', subtitle=None, footer=None, notes=None, **_):
        s = self._slide('final.jpg', notes, number=False)
        self.text(s, 48, 196, 500, 64, title, size=36, color=WHITE, bold=True, anchor='middle', fit=22)
        if subtitle:
            self.text(s, 48, 266, 480, 60, subtitle, size=18, color=WHITE, fit=12)
        if footer:
            self.text(s, 48, 420, 480, 50, footer, size=13, color=WHITE, bold=True, fit=9)
        return s

    def kpi_table(self, title, table, kpis=(), progress=None, intro=None, notes=None, **_):
        """Слева столбец KPI (1–4), справа таблица; опционально под таблицей — блок прогресс-баров."""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        x = X0
        if kpis:
            self.kpi_column(s, list(kpis), y=y)
            x = X0 + 290 + 20
        w = X1 - x
        avail = Y1 - y
        th_max = avail - (168 if progress else 0)
        th = self.table(s, x, y, w, table, max_h=th_max)
        if progress:
            py = y + th + 16
            self.progress_box(s, x, py, w, Y1 - py, progress)
        return s

    def kpi_chart(self, title, chart, kpis=(), progress=None, chart_title=None, intro=None, notes=None, **_):
        """Слева столбец KPI, справа диаграмма в карточке; опционально — прогресс-бары под ней."""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        x = X0
        if kpis:
            self.kpi_column(s, list(kpis), y=y)
            x = X0 + 290 + 20
        w = X1 - x
        ch_h = (Y1 - y) - (130 if progress else 0)
        top = self.panel(s, x, y, w, ch_h, chart_title)
        self.chart(s, x + 10, top, w - 20, y + ch_h - top - 6, chart)
        if progress:
            py = y + ch_h + 12
            self.progress_box(s, x, py, w, Y1 - py, progress)
        return s

    def cards(self, title, cards, intro=None, columns=None, notes=None, **_):
        """Сетка карточек (2–6, допускается до 8): иконка + заголовок + текст."""
        s = self._slide('content.png', notes, title=title)
        y = Y0
        if intro:
            self.pill(s, X0, 78, min(870, 40 + 7.2 * len(plain(intro))), 28, intro, size=12)
            y = 116
        n = len(cards)
        cols = columns or {1: 1, 2: 2, 3: 3, 4: 2, 5: 3, 6: 3, 7: 4, 8: 4}.get(n, 3)
        rows = math.ceil(n / cols)
        g = 12
        cw = (X1 - X0 - GAP * (cols - 1)) / cols
        ch = (Y1 - y - g * (rows - 1)) / rows
        for i, c in enumerate(cards):
            r, k = divmod(i, cols)
            self.card(s, X0 + k * (cw + GAP), y + r * (ch + g), cw, ch, c)
        return s

    def metric_bars(self, title, groups, intro=None, callout=None, unit='', notes=None, **_):
        """Сравнение «норматив / факт» горизонтальными полосами. groups (1–4):
           {title, icon, bars:[{label, value, kind: plan|fact|fact2|warn}], badge?}"""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro, h=34)
        y += 6
        n = len(groups)
        cw = (X1 - X0 - GAP * (n - 1)) / n
        bottom = Y1 - (70 if callout else 0)
        gh = bottom - y
        vmax = max(b['value'] for g in groups for b in g['bars'])
        dec = 1 if any(float(b['value']) % 1 for g in groups for b in g['bars']) else 0
        kinds = {'plan': PBLUE, 'fact': BLUE, 'fact2': GREEN, 'warn': ORANGE}
        for i, g in enumerate(groups):
            x = X0 + i * (cw + GAP)
            self.box(s, x, y + 8, cw, gh - 8, fill=WHITE, line=BORDER, radius=12)
            self.header_tab(s, x + 12, y, cw - 24, 40, g['title'], icon=g.get('icon'))
            bars = g['bars']
            top = y + 54
            bot = y + gh - (52 if g.get('badge') else 14)
            step = (bot - top) / max(1, len(bars))
            barw = cw - 32 - 78
            for j, b in enumerate(bars):
                by = top + j * step
                col = kinds.get(b.get('kind', 'fact'), BLUE)
                self.text(s, x + 16, by, cw - 32, 16, b['label'], size=11, color=GRAY)
                bw = max(16, barw * b['value'] / vmax)
                bh = min(16, step - 22)
                self.box(s, x + 16, by + 17, bw, bh, fill=col, radius=bh / 2)
                u = b.get('unit', g.get('unit', unit))
                val = b.get('text') or (f"{b['value']:.{dec}f}".replace('.', ',') + (f' {u}' if u else ''))
                self.text(s, x + 22 + bw, by + 13, 90, bh + 8, val, size=12, bold=True,
                          color=NAVY if col == PBLUE else col, anchor='middle')
            if g.get('badge'):
                self.pill(s, x + 16, y + gh - 44, cw - 32, 30, g['badge'], fill=LGREEN, color_=DGREEN, size=12)
        if callout:
            self.callout(s, X0, Y1 - 58, X1 - X0, 58, callout, style=callout.get('style', 'filled'))
        return s

    def numbered(self, title, items, tiles=(), header=None, header_icon=None, lead=None, illustration=None,
                 status=None, notes=None, **_):
        """Нумерованный перечень (2–5 пунктов) в панели; слева — опциональные плитки-цифры (до 4),
           справа — опциональная иллюстрация; внизу — строка статуса."""
        s = self._slide('content.png', notes, title=title)
        x = X0
        if tiles:
            n = len(tiles)
            g = 10
            th = (Y1 - Y0 - g * (n - 1)) / n
            for i, t in enumerate(tiles):
                ty = Y0 + i * (th + g)
                dark = t.get('dark', False)
                if t.get('icon') and not t.get('value'):
                    self.box(s, X0, ty, 290, th, fill=WHITE, line=BORDER, radius=12)
                    isz = min(50, th - 20)
                    self.icon(s, t['icon'], X0 + 12, ty + (th - isz) / 2, isz)
                    self.text(s, X0 + 24 + isz, ty + 6, 290 - 36 - isz, th - 12, t.get('label', ''), size=11.5,
                              anchor='middle', fit=8)
                    continue
                self.box(s, X0, ty, 290, th, fill=None if dark else LBLUE, grad=GRAD if dark else None, radius=12)
                v = str(t['value'])
                vw = 30 + 17 * len(v)
                self.text(s, X0 + 12, ty, vw, th, v, size=30 if len(v) <= 4 else 22, bold=True,
                          color=WHITE if dark else BLUE, anchor='middle', fit=16)
                self.text(s, X0 + 16 + vw, ty + 6, 290 - 28 - vw, th - 12, t.get('label', ''), size=11.5,
                          color=WHITE if dark else TEXT, anchor='middle', fit=8)
            x = X0 + 290 + 20
        w = X1 - x
        y = Y0 + (10 if header else 0)
        self.box(s, x, y, w, Y1 - y, fill=WHITE, line=BORDER, radius=12)
        cy = y + 12
        if header:
            self.header_tab(s, x + 16, Y0, w - 32, 40, header, icon=header_icon)
            cy = Y0 + 48
        if lead:
            self.text(s, x + 20, cy, w - 40, 34, lead, size=12.5, anchor='middle', fit=10, what='лид')
            cy += 40
        bottom = Y1 - (50 if status else 12)
        ill_w = 200 if illustration else 0
        lw = w - 40 - ill_w
        n = len(items)
        step = (bottom - cy) / n
        bh = min(52, step - 8)
        for i, t in enumerate(items):
            iy = cy + i * step + (step - bh) / 2
            self.text(s, x + 18, iy - 4, 44, bh + 8, str(i + 1), size=min(40, bh * 0.9), bold=True, color=PBLUE,
                      align='center', anchor='middle')
            self.box(s, x + 70, iy, lw - 54, bh, fill=WHITE, line=MINT, radius=8)
            self.text(s, x + 82, iy + 2, lw - 78, bh - 4, t, size=12, anchor='middle', fit=9, what=f'пункт {i + 1}')
        if illustration:
            self.picture(s, illustration, x + w - ill_w - 12, cy, ill_w, bottom - cy)
        if status:
            self.pill(s, x + 16, Y1 - 44, w - 32, 32, status, fill=LGREEN, color_=DGREEN, size=12)
        return s

    def checklist(self, title, items, panel=None, notes=None, **_):
        """Слева градиентная панель (закон / тезис), справа список с галочками (2–6).
           panel: {icon, title, subtitle, text:[...]}; items: ["**Лид:** текст", ...]"""
        s = self._slide('content.png', notes, title=title)
        x = X0
        if panel:
            self.box(s, X0, Y0, 290, Y1 - Y0, grad=GRAD, radius=14)
            py = Y0 + 18
            if panel.get('icon'):
                self.icon(s, panel['icon'], X0 + 20, py, 60, white=True)
                py += 74
            paras = []
            if panel.get('title'):
                paras.append({'text': panel['title'], 'size': 22, 'bold': True})
            if panel.get('subtitle'):
                paras.append({'text': panel['subtitle'], 'size': 14})
            for t in panel.get('text', []):
                paras.append({'text': t, 'size': 13})
            self.text(s, X0 + 20, py, 254, Y1 - py - 14, paras, color=WHITE, spacing=12, fit=9, what='панель')
            x = X0 + 290 + 20
        n = len(items)
        g = 8
        ih = (Y1 - Y0 - g * (n - 1)) / n
        for i, t in enumerate(items):
            y = Y0 + i * (ih + g)
            self.box(s, x, y, X1 - x, ih, fill=XBLUE if i % 2 == 0 else WHITE, line=BORDER, lw=1.0, radius=10)
            cs = min(26, ih - 16)
            self.check(s, x + 14, y + (ih - cs) / 2, cs)
            self.text(s, x + 26 + cs, y + 4, X1 - x - 40 - cs, ih - 8, t, size=13, anchor='middle', fit=9,
                      what=f'пункт {i + 1}')
        return s

    def flow(self, title, columns, intro=None, callouts=(), notes=None, **_):
        """Схема «шаг → шаг → шаг» (2–4 колонки) с пунктами-галочками; снизу до 2 выносок.
           columns: [{title, icon, items:[...], highlight?}]; callouts: [{icon, text, style}]"""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro, y=72, h=26)
        y += 6
        n = len(columns)
        aw = 45
        cw = (X1 - X0 - aw * (n - 1)) / n
        bottom = Y1 - (88 if callouts else 0)
        for i, c in enumerate(columns):
            x = X0 + i * (cw + aw)
            hl = c.get('highlight', i == n // 2 and n == 3)
            self.box(s, x, y + 8, cw, bottom - y - 8, fill=LBLUE if hl else WHITE, line=None if hl else BORDER,
                     radius=12)
            self.header_tab(s, x + 12, y, cw - 24, 40, c['title'], icon=c.get('icon'))
            its = c.get('items', [])
            top = y + 52
            step = (bottom - 10 - top) / max(1, len(its))
            for j, t in enumerate(its):
                iy = top + j * step
                self.check(s, x + 14, iy + 3, 22)
                self.text(s, x + 44, iy + 1, cw - 56, step - 6, t, size=13 if cw > 220 else 12, fit=9,
                          what=f'пункт колонки «{plain(c["title"])}»')
            if i < n - 1:
                self.arrow(s, x + cw + (aw - 28) / 2, (y + bottom) / 2 - 11)
        if callouts:
            m = len(callouts)
            cw2 = (X1 - X0 - 20 * (m - 1)) / m
            for i, c in enumerate(callouts):
                self.callout(s, X0 + i * (cw2 + 20), Y1 - 74, cw2, 74, c,
                             style=c.get('style', 'outline' if i == 0 and m > 1 else 'filled'))
        return s

    def statement(self, title, statement, cards=(), intro=None, notes=None, **_):
        """Крупный тезис слева (запрет, главный вывод) + 1–3 карточки справа.
           statement: {icons:[...], text: "!!Запрещено!! ...", badge?}"""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro, y=72, h=38) if intro else Y0
        y = max(y, Y0)
        lw = 470 if cards else X1 - X0
        self.box(s, X0, y, lw, Y1 - y, fill=LBLUE, radius=14)
        iy = y + 22
        for k, ic in enumerate(statement.get('icons', [])[:3]):
            self.icon(s, ic, X0 + 22 + k * 84, iy, 72)
        ty = iy + (90 if statement.get('icons') else 0)
        bh = 50 if statement.get('badge') else 0
        self.text(s, X0 + 22, ty, lw - 44, Y1 - ty - bh - 22, statement['text'], size=22, bold=True, color=BLUE,
                  fit=14, line_spacing=1.05, what='главный тезис')
        if statement.get('badge'):
            self.pill(s, X0 + 22, Y1 - 60, min(lw - 44, 300), 36, statement['badge'], fill=ORANGE, color_=WHITE,
                      size=13)
        if cards:
            x = X0 + lw + 20
            n = len(cards)
            g = 14
            ch = (Y1 - y - g * (n - 1)) / n
            for i, c in enumerate(cards):
                self.card(s, x, y + i * (ch + g), X1 - x, ch, c, highlight=c.get('highlight', i == n - 1 and n > 1))
        return s

    def cards_image(self, title, cards, illustration='people_house', banner=None, chips=None, notes=None, **_):
        """Баннер сверху (опц.), 1–3 карточки слева, иллюстрация справа, строка «чипов» снизу (опц.).
           banner: {icon, text}; chips: {title, items:[...], note}"""
        s = self._slide('content.png', notes, title=title)
        y = Y0
        if banner:
            self.callout(s, X0, y, X1 - X0, 68, banner, style='filled')
            y += 82
        bottom = Y1 - (98 if chips else 0)
        n = len(cards)
        aw = 590 if illustration else X1 - X0
        cw = (aw - 12 * (n - 1)) / n
        for i, c in enumerate(cards):
            x = X0 + i * (cw + 12)
            ch = bottom - y
            self.box(s, x, y, cw, ch, fill=WHITE, line=BORDER, radius=12)
            self.icon(s, c.get('icon'), x + 14, y + 12, 46)
            self.text(s, x + 14, y + 64, cw - 28, 22, c.get('title', ''), size=14, color=BLUE, bold=True, fit=10)
            self.text(s, x + 14, y + 90, cw - 28, ch - 100, c.get('text', ''), size=13, fit=9,
                      what='текст карточки')
        if illustration:
            self.picture(s, illustration, X0 + aw + 12, y, X1 - X0 - aw - 12, bottom - y)
        if chips:
            cy = Y1 - 86
            self.box(s, X0, cy, X1 - X0, 86, fill=LBLUE, radius=12)
            self.text(s, X0 + 16, cy + 8, 500, 22, chips.get('title', ''), size=13, color=BLUE, bold=True)
            nx = X0 + 16
            for t in chips.get('items', []):
                w = 30 + 8.2 * len(plain(t))
                self.pill(s, nx, cy + 38, w, 32, t, fill=WHITE, color_=BLUE, size=12.5, line=BLUE2)
                nx += w + 10
            if chips.get('note') and X1 - nx > 120:
                self.text(s, nx + 8, cy + 28, X1 - nx - 22, 52, chips['note'], size=11.5, anchor='middle', fit=8)
        return s

    def table_slide(self, title, table, intro=None, note=None, notes=None, **_):
        """Полноширинная таблица (до ~10 строк)."""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        th = self.table(s, X0, y, X1 - X0, table, max_h=Y1 - y - (34 if note else 0), max_row_h=46)
        if note:
            self.text(s, X0, y + th + 8, X1 - X0, 26, note, size=10.5, color=GRAY, fit=8)
        return s

    def chart_slide(self, title, chart, intro=None, chart_title=None, insights=(), notes=None, **_):
        """Диаграмма на всю ширину или диаграмма + колонка выводов справа (до 4)."""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        w = X1 - X0 - (310 if insights else 0)
        top = self.panel(s, X0, y, w, Y1 - y, chart_title)
        self.chart(s, X0 + 10, top, w - 20, Y1 - top - 8, chart)
        if insights:
            x = X0 + w + 20
            n = len(insights)
            g = 12
            ih = (Y1 - y - g * (n - 1)) / n
            for i, it in enumerate(insights):
                if isinstance(it, str):
                    it = {'text': it}
                if 'value' in it:
                    self.kpi(s, x, y + i * (ih + g), X1 - x, ih, it, dark=it.get('dark', False))
                else:
                    self.callout(s, x, y + i * (ih + g), X1 - x, ih, it, style=it.get('style', 'light'))
        return s

    def bullets(self, title, bullets, intro=None, illustration=None, notes=None, **_):
        """Простой список тезисов с галочками (3–6) и опциональной иллюстрацией справа."""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        w = X1 - X0 - (330 if illustration else 0)
        n = len(bullets)
        step = (Y1 - y) / n
        for i, t in enumerate(bullets):
            by = y + i * step
            self.check(s, X0, by + 4, 24)
            self.text(s, X0 + 36, by, w - 36, step - 8, t, size=15, fit=10, what=f'тезис {i + 1}')
        if illustration:
            self.picture(s, illustration, X1 - 310, y, 310, Y1 - y)
        return s

    def timeline(self, title, steps, intro=None, notes=None, **_):
        """Горизонтальная шкала этапов (3–6): steps: [{date, title, text, done?}]"""
        s = self._slide('content.png', notes, title=title)
        y = self.intro(s, intro)
        n = len(steps)
        cw = (X1 - X0 - 12 * (n - 1)) / n
        ly = y + 64
        self.box(s, X0 + 10, ly - 3, X1 - X0 - 20, 6, grad=GRAD, shape='rect')
        for i, st in enumerate(steps):
            x = X0 + i * (cw + 12)
            done = st.get('done', True)
            self.pill(s, x + 8, y + 6, cw - 16, 30, st.get('date', ''), fill=BLUE if done else LBLUE,
                      color_=WHITE if done else BLUE, size=12.5)
            self.box(s, x + cw / 2 - 11, ly - 11, 22, 22, shape='oval', fill=WHITE, line=BLUE if done else PBLUE,
                     lw=4)
            self.box(s, x, ly + 26, cw, Y1 - ly - 26, fill=WHITE, line=BORDER, radius=12)
            if st.get('icon'):
                self.icon(s, st['icon'], x + 12, ly + 38, 40)
            ty = ly + 38 + (48 if st.get('icon') else 0)
            self.text(s, x + 12, ty, cw - 24, 44, st.get('title', ''), size=15, color=BLUE, bold=True, fit=10)
            self.text(s, x + 12, ty + 50, cw - 24, Y1 - ty - 60, st.get('text', ''), size=13.5, fit=9,
                      what=f'этап {i + 1}')
        return s

    # ------------------------------------------------------------------ сохранение
    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.prs.save(path)
        return path


LAYOUTS = {
    'title': Deck.title, 'section': Deck.section, 'final': Deck.final,
    'kpi_table': Deck.kpi_table, 'kpi_chart': Deck.kpi_chart, 'cards': Deck.cards,
    'metric_bars': Deck.metric_bars, 'numbered': Deck.numbered, 'checklist': Deck.checklist,
    'flow': Deck.flow, 'statement': Deck.statement, 'cards_image': Deck.cards_image,
    'table': Deck.table_slide, 'chart': Deck.chart_slide, 'bullets': Deck.bullets, 'timeline': Deck.timeline,
}
