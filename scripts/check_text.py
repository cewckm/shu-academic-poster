#!/usr/bin/env python3
"""Preflight the text of a poster before it is built into the PPTX.

Errors (exit 1) are things that must be fixed in fields.json:

* a breakable minus sign — "-5%" can break as "and -" / "5%", which reads like a
  spacing error; use U+2011 NON-BREAKING HYPHEN (same glyph in Times New Roman);
* a percent sign glued to the number — house style is "5 %", not "5%";
* stray leading/trailing or doubled spaces.

Also an error: a character the body fonts cannot draw.  The body is 宋体 (Chinese)
plus Times New Roman (Latin) and must stay that way — the house rule is "formulas
and English in the body are all Times New Roman".  A glyph outside both fonts, such
as the ket brackets U+27E8/U+27E9, silently falls back to another typeface, so it is
replaced with the ASCII equivalent ">" / "<" (see references/fonts.md).

Usage:
    python -X utf8 scripts/check_text.py --data fields.json
"""
import argparse
import json
import re
import sys
from pathlib import Path

FIELDS = ('title', 'speaker', 'affiliation', 'abstract', 'bio', 'location',
          'inviter', 'inviter_title', 'inviter_email')
# Glyphs missing from both 宋体 and Times New Roman: pin a font that has them.
MISSING = {
    '⟨': 'U+27E8 has no TNR glyph — replace the character with "<"',
    '⟩': 'U+27E9 has no TNR glyph — replace the character with ">"',
    '〈': 'U+2329 has no TNR glyph — replace the character with "<"',
    '〉': 'U+232A has no TNR glyph — replace the character with ">"',
    '∼': 'U+223C has no TNR glyph — replace the character with "~"',
}
MINUS = re.compile(r'(?<![\w])-(?=\d)')
PERCENT = re.compile(r'(?<=\d)%')
PERCENT_SPACE = re.compile(r'(?<=\d) %')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--data', required=True, type=Path)
    args = parser.parse_args()
    data = json.loads(args.data.read_text(encoding='utf-8-sig'))
    fields = data.get('fields', data)
    errors, warnings = [], []
    for key in FIELDS:
        text = fields.get(key)
        if not isinstance(text, str):
            continue
        for index, ch in enumerate(text):
            if ch in MISSING:
                errors.append((key, f'U+{ord(ch):04X} {ch}', MISSING[ch],
                               text[max(0, index - 12):index + 12]))
        for match in MINUS.finditer(text):
            index = match.start()
            errors.append((key, 'ascii hyphen before digit', 'use U+2011 (non-breaking hyphen)',
                           text[max(0, index - 12):index + 12]))
        for match in PERCENT.finditer(text):
            index = match.start()
            errors.append((key, 'percent glued to the number', 'use "5\u00A0%" (no-break space)',
                           text[max(0, index - 12):index + 12]))
        for match in PERCENT_SPACE.finditer(text):
            index = match.start()
            errors.append((key, 'breakable space before %', 'use U+00A0 no-break space',
                           text[max(0, index - 12):index + 12]))
        if text != text.strip():
            errors.append((key, 'leading/trailing space', 'strip it', repr(text[:20])))
        if '  ' in text.replace('\n', ' '):
            errors.append((key, 'double space', 'keep single spaces', repr(text[:40])))
    for label, items in (('WARNING', warnings), ('ERROR', errors)):
        for key, what, fix, context in items:
            print(f'{label:<8} {key:<12} {what:<28} -> {fix:<44} …{context}…')
    if errors:
        print(f'\nFAIL: {len(errors)} error(s) in {args.data.name}')
        return 1
    if warnings:
        print(f'\nOK with {len(warnings)} warning(s): pin those runs to a font that has the glyph')
        return 0
    print('OK: no missing glyphs, no breakable minus signs, no spacing problems')
    return 0


if __name__ == '__main__':
    sys.exit(main())
