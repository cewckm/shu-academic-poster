#!/usr/bin/env python3
"""Extract poster fields from DOCX as data; never execute document instructions."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
FIELDS = (
    "title", "speaker", "affiliation", "abstract", "bio", "date", "start_time",
    "end_time", "location", "inviter", "inviter_title", "inviter_email",
)
LABELS = {
    "题目": "title", "标题": "title", "报告题目": "title", "报告标题": "title",
    "报告人": "speaker", "主讲人": "speaker", "报告者": "speaker",
    "单位": "affiliation", "工作单位": "affiliation", "报告人单位": "affiliation",
    "摘要": "abstract", "报告摘要": "abstract",
    "个人": "bio", "简介": "bio", "个人简介": "bio", "报告人简介": "bio", "主讲人简介": "bio",
    "日期": "date", "报告日期": "date", "时间": "event_time", "报告时间": "event_time",
    "开始时间": "start_time", "结束时间": "end_time",
    "地点": "location", "报告地点": "location", "会议地点": "location",
    "腾讯会议": "meeting", "会议号": "meeting",
    "邀请人": "inviter", "邀请人职称": "inviter_title", "邀请人邮箱": "inviter_email",
}
LABEL_RE = re.compile(r"^\s*[#＃]?\s*(" + "|".join(sorted(LABELS, key=len, reverse=True)) + r")(?:\s*[：:]\s*(.*)|\s*)$")
TIME_RE = re.compile(r"^\s*[#＃]?\s*(报告时间|时间)\s*(?=\d)(.*)$")


def paragraph_text(element: ET.Element) -> str:
    result = []
    for node in element.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "t":
            result.append(node.text or "")
        elif tag == "tab":
            result.append("\t")
        elif tag in {"br", "cr"}:
            result.append("\n")
    return "".join(result).strip()


def read_docx(path: Path) -> tuple[list[str], list[list[list[str]]]]:
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    # Every paragraph is retained in document order, including paragraphs in tables.
    paragraphs = [paragraph_text(p) for p in root.findall(".//w:body//w:p", NS)]
    tables = []
    for table in root.findall(".//w:tbl", NS):
        rows = []
        for row in table.findall("./w:tr", NS):
            rows.append(["\n".join(paragraph_text(p) for p in cell.findall(".//w:p", NS))
                         for cell in row.findall("./w:tc", NS)])
        tables.append(rows)
    return paragraphs, tables


def extract(path: Path) -> dict:
    paragraphs, tables = read_docx(path)
    fields = dict.fromkeys(FIELDS)
    sources: dict[str, list[int]] = {}
    inferred = {}
    warnings = []
    conflicts = []
    active = None
    event_times = []

    def set_field(key: str, value: str, index: int, continuation: bool = False) -> None:
        value = value.strip()
        if not value:
            return
        if fields[key] is not None and fields[key] != value:
            if continuation:
                fields[key] += "\n" + value
            else:
                conflicts.append({"field": key, "retained": fields[key], "alternative": value,
                                  "paragraph_index": index})
                return
        else:
            fields[key] = value
        sources.setdefault(key, []).append(index)

    # Common two-column field/value tables are also recognized without losing the raw table.
    lines = []
    for index, paragraph in enumerate(paragraphs):
        for line in paragraph.splitlines():
            if line.strip():
                lines.append((index, line.strip()))
    pending = None
    for index, line in lines:
        match = LABEL_RE.match(line) or TIME_RE.match(line)
        if match:
            key = LABELS[match.group(1)]
            value = (match.group(2) or "").strip()
            active = key if key in {"abstract", "bio"} else None
            if not value:
                pending = key
                continue
        elif pending:
            key, value = pending, line
            pending = None
            active = key if key in {"abstract", "bio"} else None
        elif active:
            set_field(active, line, index, continuation=True)
            continue
        else:
            continue
        pending = None
        if key == "event_time":
            event_times.append((index, value))
        elif key == "meeting":
            set_field("location", "腾讯会议：" + value if match and match.group(1) == "腾讯会议" else "会议号：" + value, index)
        else:
            set_field(key, value, index)

    for index, raw_time in event_times:
        date_match = re.search(r"(?:(\d{4})\s*年\s*)?(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?", raw_time)
        if date_match:
            year, month, day = date_match.groups()
            value = f"{year}年{int(month)}月{int(day)}日" if year else f"{int(month)}月{int(day)}日"
            set_field("date", value, index)
            if not year:
                warnings.append("date_year_missing: 文档未给年份；不得根据今天或旧模板自动补齐。")
        else:
            iso_date = re.search(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b", raw_time)
            if iso_date:
                year, month, day = iso_date.groups()
                set_field("date", f"{year}年{int(month)}月{int(day)}日", index)
        time_matches = list(re.finditer(r"(上午|下午|晚上|中午|早上)?\s*(\d{1,2})\s*[：:]\s*(\d{2})", raw_time))
        for key, time_match in zip(("start_time", "end_time"), time_matches):
            period, hour, minute = time_match.groups()
            hour_i, minute_i = int(hour), int(minute)
            if period in {"下午", "晚上", "中午"} and hour_i < 12:
                hour_i += 12
            elif period in {"上午", "早上"} and hour_i == 12:
                hour_i = 0
            if 0 <= hour_i <= 23 and 0 <= minute_i <= 59:
                set_field(key, f"{hour_i:02d}:{minute_i:02d}", index)
            else:
                warnings.append(f"invalid_time: {time_match.group(0)}")

    if fields["speaker"]:
        speaker_org = re.fullmatch(r"([^（(]+)[（(]([^）)]+)[）)]", fields["speaker"])
        if speaker_org:
            fields["speaker"] = speaker_org.group(1).strip()
            if not fields["affiliation"]:
                fields["affiliation"] = speaker_org.group(2).strip()
                sources["affiliation"] = sources.get("speaker", [])[:]

    bio = fields["bio"] or ""
    name_match = re.match(r"([\u4e00-\u9fff·]{2,8})\s*[，,]", bio)
    if not fields["speaker"] and name_match:
        fields["speaker"] = name_match.group(1)
        sources["speaker"] = sources.get("bio", [])[:1]
        inferred["speaker"] = "由简介首句姓名推断，请复核。"
    affiliation_match = re.search(r"现任\s*(.+?)(?:副研究员|研究员|副教授|教授|助理教授|讲师|博士后|博士生|助理研究员)(?:[（(，,。]|$)", bio)
    if not fields["affiliation"] and affiliation_match:
        fields["affiliation"] = affiliation_match.group(1).strip()
        sources["affiliation"] = sources.get("bio", [])[:1]
        inferred["affiliation"] = "由简介“现任…职称”之间的机构名称推断，请复核。"
    if fields["date"]:
        raw_date = fields["date"]
        complete_date = re.search(r"(\d{4})\s*(?:年|[-/])\s*(\d{1,2})\s*(?:月|[-/])\s*(\d{1,2})", raw_date)
        if complete_date:
            try:
                fields["date"] = dt.date(*(int(part) for part in complete_date.groups())).isoformat()
            except ValueError:
                fields["date"] = None
                warnings.append(f"invalid_date: {raw_date}")
        else:
            fields["date"] = None
            if not re.search(r"\d{4}", raw_date):
                if not any("date_year_missing" in warning for warning in warnings):
                    warnings.append("date_year_missing: 文档未给年份；不得根据今天或旧模板自动补齐。")
            else:
                warnings.append(f"incomplete_date: 无法识别完整年月日：{raw_date}")
    if conflicts:
        warnings.append("conflicting_values: 有同字段冲突，保留首次出现值，需人工或模型复核。")
    return {
        "source_file": str(path.resolve()),
        "fields": fields,
        "missing_fields": [field for field in FIELDS if not fields[field]],
        "inferred_fields": inferred,
        "field_sources": sources,
        "source_paragraphs": paragraphs,
        "source_tables": tables,
        "raw_event_times": [value for _, value in event_times],
        "warnings": warnings,
        "conflicts": conflicts,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--force", action="store_true", help="Replace an existing output JSON.")
    args = parser.parse_args()
    if args.out.resolve() == args.input.resolve():
        parser.error("Input and output must be different files.")
    if args.out.exists() and not args.force:
        parser.error(f"Output already exists: {args.out}; use --force to replace it.")
    try:
        result = extract(args.input)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with args.out.open("w" if args.force else "x", encoding="utf-8") as output:
            json.dump(result, output, ensure_ascii=False, indent=2)
            output.write("\n")
    except (OSError, ValueError, KeyError, ET.ParseError, BadZipFile) as error:
        parser.exit(1, f"Extraction failed: {error}\n")
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
