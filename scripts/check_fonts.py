#!/usr/bin/env python3
"""Audit a generated poster against the approved poster deck (assets/poster-template.pptx).

The group's PPTX is the single source of truth for typography: every text run of the
generated poster must use a (latin, east-asian) pair that already occurs in the same
shape of the template. Nothing else is invented, so a run that shows None (theme font)
in the template must still show None, 宋体 must stay 宋体, and so on.

One deliberate exception is listed in PINNED: the English line of the header is
theme-driven (+mj-lt -> Aptos Display) in the template, and different Office builds
substitute it differently; the approved poster renders it in Calibri, so the builder
writes that explicitly and this audit allows it.
"""
import argparse
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
ROOT = Path(__file__).resolve().parents[1]
# shape id -> {paragraph index: (latin, east-asian)} allowed on top of the template's own pairs
PINNED = {6: {1: ('Calibri', '宋体')}}
# The body must be all Times New Roman (+ 宋体 for Chinese); nothing else is accepted.
# Latin/formula runs get Times New Roman in the east-asian slot as well, so the font
# dialog shows TNR everywhere inside English and formulas.
EXTRA = {12: {('Times New Roman', 'Times New Roman')},
         19: {('Times New Roman', 'Times New Roman')}}


def tag(name):
    prefix, local = name.split(':')
    return '{' + NS[prefix] + '}' + local


def shapes_of(path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    return {int(sp.find('p:nvSpPr/p:cNvPr', NS).get('id')): sp
            for sp in root.findall('.//p:sp', NS)}


def paragraph_fonts(paragraph):
    props = [r.find('a:rPr', NS) for r in paragraph.findall('a:r', NS)]
    props.append(paragraph.find('a:endParaRPr', NS))
    pairs = []
    for rpr in props:
        if rpr is None:
            continue
        latin = rpr.find('a:latin', NS)
        east = rpr.find('a:ea', NS)
        pairs.append((latin.get('typeface') if latin is not None else None,
                      east.get('typeface') if east is not None else None))
    return pairs


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--template', type=Path, default=ROOT / 'assets' / 'poster-template.pptx')
    args = parser.parse_args()

    template = shapes_of(args.template)
    allowed = {}
    for sid, shape in template.items():
        pairs = set()
        for paragraph in shape.findall('p:txBody/a:p', NS):
            pairs.update(paragraph_fonts(paragraph))
        for index, pair in PINNED.get(sid, {}).items():
            pairs.add(pair)
        pairs.update(EXTRA.get(sid, set()))
        allowed[sid] = pairs

    problems = []
    rows = []
    for sid, shape in sorted(shapes_of(args.pptx).items()):
        for index, paragraph in enumerate(shape.findall('p:txBody/a:p', NS)):
            text = ''.join(t.text or '' for t in paragraph.iter(tag('a:t')))
            if not text.strip():
                continue
            for pair in paragraph_fonts(paragraph):
                ok = pair in allowed.get(sid, set())
                rows.append((sid, index, text[:18], pair, ok))
                if not ok:
                    problems.append(f'shape {sid} "{text[:18]}" uses latin={pair[0]!r} ea={pair[1]!r}; '
                                    f'template allows {sorted(allowed.get(sid, set()), key=str)}')
    print(f'{"id":>4} {"para":>4}  {"text":<20} {"(latin, east-asian)":<34} verdict')
    for sid, index, text, pair, ok in rows:
        print(f'{sid:>4} {index:>4}  {text:<20} {str(pair):<34} {"ok" if ok else "DRIFT"}')
    if problems:
        print('\nTypography does not match the approved deck:')
        for line in sorted(set(problems)):
            print('  - ' + line)
        print(f'\nFAIL: {len(set(problems))} run(s) differ from {args.template.name}')
        return 1
    print(f'\nOK: every run matches the approved deck {args.template.name}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
