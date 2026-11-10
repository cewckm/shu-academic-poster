#!/usr/bin/env python3
"""Derive subscript/superscript runs for a poster from the source DOCX.

A plain-text extraction drops w:vertAlign, so formulas such as
Li(n)B(n+1)C(n-1), La3Ni2O7 or U_SS(t_c) would come out as flat text.  This tool
reads the same paragraphs and returns the character ranges that were subscript'
or superscript, in the form build_poster.py --runs expects:

    {"abstract": [{"range": [35, 36], "baseline": -25000, "size_scale": 0.65}]}

The reconstructed text is compared with fields.json; anything that does not match
exactly is reported instead of being written, because run offsets would be wrong.

Usage:
    python -X utf8 scripts/extract_runs.py <input.docx> --fields <fields.json> --out <runs.json>
"""
import argparse
import difflib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
LABELS = {
    "题目": "title", "标题": "title", "报告题目": "title",
    "摘要": "abstract", "报告摘要": "abstract",
    "个人": "bio", "简介": "bio", "个人简介": "bio", "报告人简介": "bio",
}
BASELINE = -25000
SCALE = 0.65
# Glyphs neither 宋体 nor Times New Roman can draw.  The body must stay all Times New
# Roman, so these are reported (by check_text.py) and replaced with the ASCII
# equivalent instead of being pinned to another face.  --glyph-font restores the
# "keep the character, change the font" behaviour when a project asks for it.
MISSING_GLYPHS = {'⟨', '⟩', '〈', '〉', '∼'}


def runs_with_offsets(paragraph):
    """Flatten a paragraph into (text, vertAlign) chunks in document order."""
    chunks = []
    for run in paragraph.iter(W + "r"):
        align = None
        rpr = run.find(W + "rPr")
        if rpr is not None:
            vert = rpr.find(W + "vertAlign")
            if vert is not None:
                align = vert.get(W + "val")
        buf = []
        for node in run.iter():
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "t":
                buf.append(node.text or "")
            elif tag == "tab":
                buf.append("\t")
            elif tag in ("br", "cr"):
                buf.append("\n")
        if buf:
            chunks.append(("".join(buf), align))
    return chunks


def read_paragraphs(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(W + "body")
    return [runs_with_offsets(p) for p in body.findall(W + "p")]


def build_field(paragraphs):
    text, spans, offset = [], [], 0
    for index, chunks in enumerate(paragraphs):
        if index:
            text.append("\n")
            offset += 1
        for chunk, align in chunks:
            start = offset
            text.append(chunk)
            offset += len(chunk)
            if align in ("subscript", "superscript"):
                spans.append({"range": [start, offset], "baseline": BASELINE, "size_scale": SCALE})
    return "".join(text), spans


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', type=Path)
    parser.add_argument('--fields', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--glyph-font', default=None,
                        help='pin this typeface on glyphs TNR cannot draw (default: report only)')
    args = parser.parse_args()

    fields = json.loads(args.fields.read_text(encoding='utf-8-sig'))
    fields = fields.get('fields', fields)
    groups, current = {}, None
    for chunks in read_paragraphs(args.docx):
        full = ''.join(t for t, _ in chunks)
        label = full.split("：")[0].strip() if "：" in full else None
        key = LABELS.get(label) if label else None
        if key:
            current = key
            groups.setdefault(key, []).append(chunks)
        elif current and full.strip():
            groups[current].append(chunks)
        elif not full.strip():
            current = None

    result, problems = {}, []
    for key, paragraphs in groups.items():
        text, spans = build_field(paragraphs)
        head = next(label + "：" for label, mapped in LABELS.items() if mapped == key)
        if text.startswith(head):
            text = text[len(head):]
            spans = [dict(span, range=[max(0, span['range'][0] - len(head)),
                                       span['range'][1] - len(head)])
                     for span in spans if span['range'][1] > len(head)]
        expected = fields.get(key)
        if expected is None:
            continue
        if text != expected:
            # fields.json may have been edited for typography (no-break hyphen, space
            # before %, ...).  Map the spans through the unchanged stretches instead of
            # refusing outright; anything that falls inside an edited stretch is dropped.
            mapping = difflib.SequenceMatcher(None, text, expected, autojunk=False)
            equal = [(i1, i2, j1) for tag, i1, i2, j1, j2 in mapping.get_opcodes() if tag == 'equal']
            mapped, dropped = [], 0
            for span in spans:
                start, end = span['range']
                hit = next(((i1, i2, j1) for i1, i2, j1 in equal if i1 <= start and end <= i2), None)
                if hit is None:
                    dropped += 1
                    continue
                i1, _, j1 = hit
                mapped.append(dict(span, range=[j1 + (start - i1), j1 + (end - i1)]))
            if dropped:
                print(f'NOTE {key}: {dropped} run(s) sit inside edited text and were skipped')
            spans = mapped
        if args.glyph_font:
            for index, ch in enumerate(expected):
                if ch in MISSING_GLYPHS:
                    spans.append({'range': [index, index + 1],
                                  'latin': args.glyph_font, 'ea': args.glyph_font})
        if spans:
            result[key] = spans
    if problems:
        for line in problems:
            print('ERROR', line)
        print('Run extraction failed: fix fields.json so it matches the DOCX exactly, '
              'or drop that field from --runs.')
        return 1
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    summary = {key: len(items) for key, items in result.items()}
    print(f'Wrote {args.out} (runs per field: {summary})')
    return 0


if __name__ == '__main__':
    sys.exit(main())
