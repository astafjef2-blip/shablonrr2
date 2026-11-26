#!/usr/bin/env python3
"""Собирает .pptx в стиле Росреестра из JSON-спецификации.

    python scripts/build_from_spec.py spec.json [-o out.pptx] [--preview]

Формат спецификации — см. docs/spec-format.md и examples/*.json.
"""
import argparse
import inspect
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rosreestr_pptx import Deck, LAYOUTS  # noqa: E402


def build(spec, out=None):
    deck = Deck()
    for i, sl in enumerate(spec['slides'], 1):
        sl = dict(sl)
        kind = sl.pop('type', None)
        if kind not in LAYOUTS:
            raise SystemExit(f'Слайд {i}: неизвестный type="{kind}". Доступны: {", ".join(LAYOUTS)}')
        fn = LAYOUTS[kind]
        params = inspect.signature(fn).parameters
        unknown = [k for k in sl if k not in params]
        if unknown:
            deck.warnings.append(f'Слайд {i} ({kind}): проигнорированы поля {unknown}')
        missing = [k for k, p in params.items()
                   if k not in ('self', '_') and p.default is inspect.Parameter.empty
                   and p.kind == p.POSITIONAL_OR_KEYWORD and k not in sl]
        if missing:
            raise SystemExit(f'Слайд {i} ({kind}): не заданы обязательные поля {missing}')
        fn(deck, **sl)
    meta = spec.get('meta', {})
    cp = deck.prs.core_properties
    cp.title = meta.get('title', '')
    cp.author = meta.get('author', 'Росреестр')
    cp.subject = meta.get('subject', '')
    out = out or spec.get('output') or 'presentation.pptx'
    deck.save(out)
    return out, deck.warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('spec')
    ap.add_argument('-o', '--output')
    ap.add_argument('--preview', action='store_true', help='отрендерить PNG-превью (нужен LibreOffice)')
    a = ap.parse_args()
    with open(a.spec, encoding='utf-8') as f:
        spec = json.load(f)
    out, warns = build(spec, a.output)
    print(f'✔ Сохранено: {out}  (слайдов: {len(spec["slides"])})')
    for w in warns:
        print('⚠', w)
    if a.preview:
        from render_preview import render
        render(out)


if __name__ == '__main__':
    main()
