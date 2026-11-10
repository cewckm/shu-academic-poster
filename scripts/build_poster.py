#!/usr/bin/env python3
"""Fill the retained one-slide poster. Python standard library only."""
import argparse
import copy
import datetime as dt
import hashlib
import io
import json
import os
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
REQ = ('title', 'speaker', 'affiliation', 'abstract', 'bio', 'date', 'start_time', 'location')
# Typography follows assets/poster-template.pptx (the group's approved poster deck):
#   title / 报告人·时间·地点 / footer  keep the template runs (微软雅黑, Helvetica)
#   abstract + biography              中文宋体 + 西文 Times New Roman, as the template runs do
#   section labels + header             keep the template's theme fonts (等线, 华文隶书)
# Only the header's English line is pinned, because its theme font (+mj-lt/Aptos Display)
# differs per machine while the approved poster renders it in Calibri.
BODY_CJK = '宋体'
BODY_LATIN = 'Times New Roman'
PINNED = {6: {1: ('Calibri', '宋体')}}
# Runs must keep DrawingML order: latin/ea precede cs, sym, hlink and extLst.
FONT_LATER = ('a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst')




def split_by_script(shape, latin_font, cjk_font):
    """Give Latin/formula runs their own east-asian slot too.

    A single run with latin=Times New Roman, ea=宋体 renders English in TNR, but
    PowerPoint's font dialog still reports 宋体 for the run, and any ambiguous
    character can fall to the Chinese face.  Splitting by script makes every Latin
    or formula run read Times New Roman in both slots, while Chinese runs keep 宋体.
    """
    body = shape.find('p:txBody', NS)
    for paragraph in body.findall('a:p', NS):
        ppr = paragraph.find('a:pPr', NS)
        runs = paragraph.findall('a:r', NS)
        end = paragraph.find('a:endParaRPr', NS)
        rebuilt = []
        for run in runs:
            rpr = run.find('a:rPr', NS)
            node = run.find('a:t', NS)
            text = (node.text or '') if node is not None else ''
            if rpr is None or not text:
                rebuilt.append(run)
                continue
            for chunk, cjk in split_script(text):
                new_rpr = copy.deepcopy(rpr)
                set_typeface(new_rpr, 'a:ea', cjk_font if cjk else latin_font)
                new_rpr.set('lang', 'zh-CN' if cjk else 'en-US')
                new_rpr.set('altLang', 'en-US' if cjk else 'zh-CN')
                new_run = ET.Element(tag('a:r'))
                new_run.append(new_rpr)
                ET.SubElement(new_run, tag('a:t')).text = chunk
                rebuilt.append(new_run)
        for child in list(paragraph):
            if child is not ppr:
                paragraph.remove(child)
        for child in rebuilt:
            paragraph.append(child)
        if end is not None:
            paragraph.append(end)


def split_script(text):
    """Yield (chunk, is_cjk) runs of the same script class."""
    out = []
    for ch in text:
        cjk = is_cjk(ch)
        if out and out[-1][1] == cjk:
            out[-1][0] += ch
        else:
            out.append([ch, cjk])
    return [(chunk, cjk) for chunk, cjk in out]


def tag(name):
    prefix, local = name.split(':')
    return '{' + NS[prefix] + '}' + local


def parse_xml(data):
    for _, pair in ET.iterparse(io.BytesIO(data), events=('start-ns',)):
        prefix, uri = pair
        if not re.match(r'ns\d+$', prefix):
            ET.register_namespace(prefix, uri)
    return ET.fromstring(data)


def xml_bytes(root):
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def shape_map(root):
    return {int(s.find('p:nvSpPr/p:cNvPr', NS).get('id')): s
            for s in root.findall('.//p:sp', NS)}


def set_box(shape, y=None, height=None):
    xf = shape.find('p:spPr/a:xfrm', NS)
    if y is not None:
        xf.find('a:off', NS).set('y', str(round(y)))
    if height is not None:
        xf.find('a:ext', NS).set('cy', str(round(height)))


def shift(shape, delta):
    off = shape.find('p:spPr/a:xfrm/a:off', NS)
    off.set('y', str(int(off.get('y')) + delta))


def fill(shape, text, size, latin=None, east_asian=None, line_points=None, runs=None):
    """Preserve paragraph/shape styles, replacing all event-specific text runs.

    `runs` holds character ranges of `text` that need their own run properties, e.g.
    subscripts from the source DOCX (baseline -25000, 65 % size) or the ket brackets
    that Times New Roman cannot draw (an explicit typeface).  Ranges are absolute
    offsets into `text`; newlines count as one character, exactly like the built text.
    """
    body = shape.find('p:txBody', NS)
    old = body.find('a:p', NS)
    ppr = old.find('a:pPr', NS)
    rpr = old.find('a:r/a:rPr', NS)
    rpr = copy.deepcopy(rpr) if rpr is not None else ET.Element(tag('a:rPr'))
    rpr.set('sz', str(round(size * 100)))
    for attr in ('err', 'dirty'):
        rpr.attrib.pop(attr, None)
    for font_tag, family in [('a:latin', latin), ('a:ea', east_asian)]:
        if family:
            child = rpr.find(font_tag, NS)
            if child is None:
                child = ET.SubElement(rpr, tag(font_tag))
            child.attrib.clear()
            child.set('typeface', family)
    for p in list(body.findall('a:p', NS)):
        body.remove(p)
    bp = body.find('a:bodyPr', NS)
    for auto in ('a:spAutoFit', 'a:normAutofit', 'a:noAutofit'):
        for el in list(bp.findall(auto, NS)):
            bp.remove(el)
    ET.SubElement(bp, tag('a:noAutofit'))
    offset = 0
    for line in text.split('\n'):
        p = ET.SubElement(body, tag('a:p'))
        pp = copy.deepcopy(ppr) if ppr is not None else ET.Element(tag('a:pPr'))
        if line_points is not None:
            for ln in list(pp.findall('a:lnSpc', NS)):
                pp.remove(ln)
            ln = ET.Element(tag('a:lnSpc'))
            ET.SubElement(ln, tag('a:spcPts'), {'val': str(round(line_points * 100))})
            pp.insert(0, ln)
        p.append(pp)
        start_abs, end_abs = offset, offset + len(line)
        local = []
        for item in runs or []:
            lo, hi = item['range']
            lo, hi = max(lo, start_abs), min(hi, end_abs)
            if lo < hi:
                local.append((lo - start_abs, hi - start_abs,
                              {k: v for k, v in item.items() if k != 'range'}))
        if not local:
            r = ET.SubElement(p, tag('a:r'))
            r.append(copy.deepcopy(rpr))
            ET.SubElement(r, tag('a:t')).text = line
        else:
            cuts = sorted({0, len(line)} | {a for a, _, _ in local} | {b for _, b, _ in local})
            for a, b in zip(cuts, cuts[1:]):
                attrs = {}
                for lo, hi, props in local:
                    if lo <= a and b <= hi:
                        attrs.update(props)
                run_rpr = copy.deepcopy(rpr)
                if 'baseline' in attrs:
                    run_rpr.set('baseline', str(int(round(attrs['baseline']))))
                if 'size_scale' in attrs:
                    run_rpr.set('sz', str(int(round(size * 100 * attrs['size_scale']))))
                for key, font_tag in (('latin', 'a:latin'), ('ea', 'a:ea')):
                    if attrs.get(key):
                        set_typeface(run_rpr, font_tag, attrs[key])
                text_slice = line[a:b]
                run_rpr.set('lang', 'zh-CN' if any(is_cjk(ch) for ch in text_slice) else 'en-US')
                r = ET.SubElement(p, tag('a:r'))
                r.append(run_rpr)
                ET.SubElement(r, tag('a:t')).text = text_slice
        ep = copy.deepcopy(rpr)
        ep.tag = tag('a:endParaRPr')
        p.append(ep)
        offset += len(line) + 1


def is_cjk(ch):
    code = ord(ch)
    return (0x4E00 <= code <= 0x9FFF or 0x3000 <= code <= 0x303F
            or 0xFF00 <= code <= 0xFFEF or 0x2018 <= code <= 0x201D)




def set_leading(shape, points):
    """Replace fixed line spacing on every paragraph of a text shape."""
    for paragraph in shape.findall('p:txBody/a:p', NS):
        ppr = paragraph.find('a:pPr', NS)
        if ppr is None:
            ppr = ET.Element(tag('a:pPr'))
            paragraph.insert(0, ppr)
        for ln in list(ppr.findall('a:lnSpc', NS)):
            ppr.remove(ln)
        ln = ET.Element(tag('a:lnSpc'))
        ET.SubElement(ln, tag('a:spcPts'), {'val': str(round(points * 100))})
        ppr.insert(0, ln)


def set_hanging_indent(shape, marl_emu):
    first = shape.find('p:txBody/a:p', NS)
    ppr = first.find('a:pPr', NS)
    if ppr is None:
        ppr = ET.Element(tag('a:pPr'))
        first.insert(0, ppr)
    ppr.set('marL', str(int(marl_emu)))
    ppr.set('indent', str(-int(marl_emu)))


def set_typeface(rpr, name, family):
    """Replace one font element, keeping DrawingML child order valid."""
    for child in list(rpr.findall(name, NS)):
        rpr.remove(child)
    element = ET.Element(tag(name))
    element.set('typeface', family)
    later = {tag(t) for t in FONT_LATER}
    index = len(rpr)
    for i, child in enumerate(list(rpr)):
        if child.tag in later:
            index = i
            break
    rpr.insert(index, element)


def apply_font_policy(shape, cn, en, skip=()):
    """Force the bilingual font pair on every run of a shape the template left theme-driven."""
    body = shape.find('p:txBody', NS)
    for index, paragraph in enumerate(body.findall('a:p', NS)):
        if index in skip:
            continue
        props = [r.find('a:rPr', NS) for r in paragraph.findall('a:r', NS)]
        props.append(paragraph.find('a:endParaRPr', NS))
        for rpr in props:
            if rpr is None:
                continue
            set_typeface(rpr, 'a:latin', en)
            set_typeface(rpr, 'a:ea', cn)


def apply_pinned_fonts(shape, pins):
    """Pin specific paragraphs to an explicit pair (used for the header English line)."""
    paragraphs = shape.findall('p:txBody/a:p', NS)
    for index, (latin, east) in pins.items():
        if index >= len(paragraphs):
            continue
        paragraph = paragraphs[index]
        props = [r.find('a:rPr', NS) for r in paragraph.findall('a:r', NS)]
        props.append(paragraph.find('a:endParaRPr', NS))
        for rpr in props:
            if rpr is None:
                continue
            set_typeface(rpr, 'a:latin', latin)
            set_typeface(rpr, 'a:ea', east)


def prepare(data, draft):
    if isinstance(data.get('fields'), dict):
        data = dict(data['fields'])
    else:
        data = dict(data)
    for key in set(REQ) | {'end_time', 'inviter', 'inviter_title', 'inviter_email'}:
        val = data.get(key)
        if val is not None and not isinstance(val, str):
            raise ValueError(f'{key} must be a string or null')
        if isinstance(val, str):
            data[key] = val.strip()
    missing = [k for k in REQ if not data.get(k)]
    if missing and not draft:
        raise ValueError('Missing required fields: ' + ', '.join(missing) + '; resolve them or use --draft')
    date_text = '日期待确认'
    if data.get('date'):
        date = dt.date.fromisoformat(data['date'])
        date_text = f'{date.year}年{date.month}月{date.day}日（周{"一二三四五六日"[date.weekday()]}）'
    for key in ('start_time', 'end_time'):
        if data.get(key) and not re.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d', data[key]):
            raise ValueError(f'{key} must use HH:MM')
    if data.get('start_time') and data.get('end_time') and data['end_time'] <= data['start_time']:
        raise ValueError('end_time must be later than start_time (overnight events need a custom layout)')
    start = data.get('start_time') or '时间待确认'
    time = date_text + ' ' + start + ('~' + data['end_time'] if data.get('end_time') else '')
    def value(k):
        return data.get(k) or {'title':'题目待确认', 'speaker':'报告人待确认',
            'affiliation':'单位待确认', 'abstract':'摘要待补充', 'bio':'报告人简介待补充',
            'location':'地点待确认'}[k]
    footer = ''
    if data.get('inviter'):
        footer = '邀请人：' + data['inviter']
        for k in ('inviter_title', 'inviter_email'):
            if data.get(k):
                footer += '  ' + data[k]
    elif data.get('inviter_title') or data.get('inviter_email'):
        raise ValueError('inviter is required when inviter_title or inviter_email is supplied')
    blocks = {14:value('title').rstrip('。'), 12:value('abstract'), 19:value('bio'),
              22:footer, 25:f"报告人：{value('speaker')}（{value('affiliation')}）\n时    间：{time}\n地    点：{value('location')}"}
    return data, missing, blocks


def build(args):
    source = args.template.resolve()
    target = args.out.resolve()
    manifest = target.with_suffix('.build.json')
    if target == source or target == args.data.resolve():
        raise ValueError('Output cannot overwrite an input')
    if target.suffix.lower() != '.pptx':
        raise ValueError('--out must end in .pptx')
    if not args.force and (target.exists() or manifest.exists()):
        raise FileExistsError('Output exists; use a new name or --force')
    if not 14 <= args.body_size <= 16 or not 16 <= args.title_size <= 20:
        raise ValueError('Supported body size 14–16 pt; title size 16–20 pt')
    if not 13 <= args.meta_size <= 16:
        raise ValueError('Supported metadata size 13–16 pt')
    data, missing, blocks = prepare(json.loads(args.data.read_text(encoding='utf-8-sig')), args.draft)
    # Optional per-character run properties: subscripts/superscripts from the source
    # DOCX, or an explicit typeface for a glyph Times New Roman cannot draw.
    runs = data.pop('_runs', None) if isinstance(data.get('_runs'), dict) else {}
    if args.runs:
        runs = json.loads(args.runs.read_text(encoding='utf-8-sig'))
    runs = runs or {}
    for key in runs:
        if key not in ('title', 'abstract', 'bio'):
            raise ValueError(f'--runs supports title/abstract/bio, got {key}')
    for key, items in runs.items():
        length = len(blocks[{ 'title': 14, 'abstract': 12, 'bio': 19 }[key]])
        for item in items:
            start, end = item['range']
            if not 0 <= start < end <= length:
                raise ValueError(f'runs[{key}] range {start}-{end} outside 0-{length}')
    with zipfile.ZipFile(source) as z:
        slide = parse_xml(z.read('ppt/slides/slide1.xml'))
        shapes = shape_map(slide)
        if not {4,6,10,12,14,15,16,19,20,22,25,26,27} <= shapes.keys():
            raise ValueError('Template shape IDs do not match this skill')
        dims = ET.fromstring(z.read('ppt/presentation.xml')).find('p:sldSz', NS)
        if (dims.get('cx'), dims.get('cy')) != ('7199313', '11520488'):
            raise ValueError('Template dimensions do not match the retained poster')
        # Default: keep the template's own typefaces. --cn-font/--en-font are opt-in
        # overrides for groups that want a different pairing than the approved poster.
        cn, en = args.cn_font, args.en_font
        body_cjk, body_latin = (cn or BODY_CJK), (en or BODY_LATIN)
        fill(shapes[14], blocks[14], args.title_size, en, cn, runs=runs.get('title'))
        for sid, key in ((12, 'abstract'), (19, 'bio')):
            fill(shapes[sid], blocks[sid], args.body_size, body_latin, body_cjk,
                 args.body_size * 1.25, runs=runs.get(key))
        fill(shapes[25], blocks[25], args.meta_size, en, cn)
        fill(shapes[22], blocks[22], 14, en, cn)
        if cn or en:
            for sid in (15, 16):
                apply_font_policy(shapes[sid], body_cjk, body_latin)
        for sid, paragraphs in PINNED.items():
            apply_pinned_fonts(shapes[sid], paragraphs)
        # Body only: English and formulas read Times New Roman in both font slots.
        for sid in (12, 19):
            split_by_script(shapes[sid], body_latin, body_cjk)
        # Fine layout adjustments, all within the template's own blocks.
        if args.shift_abstract:
            delta = round(args.shift_abstract * 12700)
            for sid in (26, 15):
                shift(shapes[sid], delta)
            shift(shapes[12], delta)
            # keep the frame's bottom edge on the template position: the divider below is fixed
            set_box(shapes[12], height=int(shapes[12].find('p:spPr/a:xfrm/a:ext', NS).get('cy')) - delta)
        if args.abstract_leading:
            set_leading(shapes[12], args.abstract_leading)
        for item in (args.leading or '').split(','):
            if not item.strip():
                continue
            sid, _, value = item.partition('=')
            set_leading(shapes[int(sid)], float(value))
        if args.meta_hanging_indent:
            set_hanging_indent(shapes[25], round(4 * args.meta_size * 12700))
        for item in (args.box_height or '').split(','):
            if not item.strip():
                continue
            sid, _, value = item.partition('=')
            shapes[int(sid)].find('p:spPr/a:xfrm/a:ext', NS).set('cy', str(round(float(value) * 12700)))
        # Reserve a second title line by moving metadata and the abstract heading down.
        delta = 418833 if args.title_lines == 2 else 0
        if delta:
            set_box(shapes[14], height=418833 + delta)
            for sid in (25,26,15,12):
                shift(shapes[sid], delta)
            set_box(shapes[12], height=3807073-delta)
        # Move the biography divider in a controlled band to rebalance long sections.
        if args.bio_top is not None:
            if not 0.62 <= args.bio_top <= 0.80:
                raise ValueError('--bio-top must be between 0.62 and 0.80 of canvas height')
            move = round(args.bio_top * 11520488) - 8458827
            for sid in (27,16,19):
                shift(shapes[sid], move)
            set_box(shapes[12], height=3807073-delta+move)
            set_box(shapes[19], height=1858201-move)
        replacement = {'ppt/slides/slide1.xml': xml_bytes(slide)}
        # Drop the source thumbnail so the new event cannot display the previous poster preview.
        rel = parse_xml(z.read('_rels/.rels'))
        for item in list(rel):
            if item.get('Type','').endswith('/metadata/thumbnail'):
                rel.remove(item)
        replacement['_rels/.rels'] = xml_bytes(rel)
        ct = parse_xml(z.read('[Content_Types].xml'))
        for item in list(ct):
            if item.get('PartName') == '/docProps/thumbnail.jpeg':
                ct.remove(item)
        replacement['[Content_Types].xml'] = xml_bytes(ct)
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix='.pptx', dir=target.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as out:
                for item in z.infolist():
                    if item.filename != 'docProps/thumbnail.jpeg':
                        out.writestr(item, replacement.get(item.filename, z.read(item.filename)))
            with zipfile.ZipFile(tmp) as check:
                if check.testzip() is not None:
                    raise ValueError('Invalid PPTX ZIP')
                actual = shape_map(ET.fromstring(check.read('ppt/slides/slide1.xml')))
                for sid, expected in blocks.items():
                    found = '\n'.join(''.join(p.itertext()) for p in actual[sid].findall('p:txBody/a:p',NS))
                    if found != expected:
                        raise ValueError(f'Content verification failed for shape {sid}')
                # A sz outside 100..400000 makes PowerPoint refuse to open the file.
                for element in actual[14].iter(tag('a:rPr')):
                    value = element.get('sz')
                    if value is not None and not 100 <= int(value) <= 400000:
                        raise ValueError(f'Invalid font size sz={value}; sizes are in 1/100 pt')
            os.replace(tmp, target)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
    report = {'output':str(target), 'draft':args.draft, 'missing_fields':missing,
              'template_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),
              'data':data, 'runs':runs, 'fonts':{'chinese':args.cn_font or BODY_CJK,
                                    'latin':args.en_font or BODY_LATIN,
                                    'source':'override' if (args.cn_font or args.en_font) else 'template'},
              'layout':{'body_size':args.body_size,'title_size':args.title_size,
              'meta_size':args.meta_size,'title_lines':args.title_lines,'bio_top':args.bio_top,
              'abstract_leading':args.abstract_leading,'leading':args.leading,'shift_abstract':args.shift_abstract,
              'meta_hanging_indent':args.meta_hanging_indent,'box_height':args.box_height},
              'validation':'Text and ZIP verified; rendering and visual review still required.'}
    manifest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'pptx':str(target),'manifest':str(manifest),'missing':missing},ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--data', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--template', type=Path, default=ROOT/'assets'/'poster-template.pptx')
    parser.add_argument('--draft', action='store_true')
    parser.add_argument('--force', action='store_true')
    parser.add_argument('--body-size', type=float, default=16)
    parser.add_argument('--title-size', type=float, default=20)
    parser.add_argument('--meta-size', type=float, default=16)
    parser.add_argument('--runs', type=Path,
                        help='JSON of per-character runs: {"abstract": [{"range": [60, 61], '
                             '"baseline": -25000, "size_scale": 0.65}]}')
    parser.add_argument('--cn-font', default=None,
                        help='Chinese typeface override for all poster text (default: keep the template fonts)')
    parser.add_argument('--en-font', default=None,
                        help='Latin/number typeface override for all poster text (default: keep the template fonts)')
    parser.add_argument('--title-lines', type=int, choices=(1,2), default=1)
    parser.add_argument('--bio-top', type=float)
    parser.add_argument('--abstract-leading', type=float,
                        help='fixed line spacing (pt) for the abstract, default body-size x 1.25')
    parser.add_argument('--leading',
                        help='line spacing per shape, e.g. "12=26,19=26"; use it to fill the '
                             'abstract/bio band with short content instead of cramming it at the top')
    parser.add_argument('--shift-abstract', type=float,
                        help='move the abstract block (label + body) down by N pt, e.g. 16')
    parser.add_argument('--meta-hanging-indent', action='store_true',
                        help='hanging indent so a wrapped 报告人 line aligns under the name')
    parser.add_argument('--box-height',
                        help='resize text frames, e.g. "25=108,19=137" (shape id = pt)')
    args = parser.parse_args()
    try:
        build(args)
    except (ValueError, OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as e:
        parser.exit(2, f'Error: {e}\n')


if __name__ == '__main__':
    main()
