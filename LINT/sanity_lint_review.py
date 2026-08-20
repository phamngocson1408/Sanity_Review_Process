#!/usr/bin/env python3
"""LINT sanity review bridge.

This tool converts VC Static LINT reports and GUI-generated waiver Tcl into a
Git-friendly CSV review database, an Excel reviewer workbook, and regenerated
waiver Tcl.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import os
import posixpath
import re
import sys
import zipfile
from collections import Counter, OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


BASE_COLUMNS = [
    "issue_id",
    "record_status",
    "review_status",
    "waiver_enabled",
    "waiver_name",
    "review_comment",
    "owner",
    "review_date",
    "tag",
    "severity",
    "goal",
    "module",
    "file",
    "line",
    "hierarchy",
    "object",
    "statement",
    "description",
    "violation",
    "source_report",
    "waiver_user",
    "waiver_timestamp",
    "filter_json",
    "fields_json",
]

KEY_FIELDS = [
    "Tag",
    "Goal",
    "Module",
    "FileName",
    "LineNumber",
    "HIERARCHY",
    "DesignObjHierarchy",
    "Signal",
    "VariableName",
    "ModPortName",
    "FsmName",
    "FsmState",
    "Clock_Register",
    "Statement",
]

OBJECT_FIELDS = [
    "Signal",
    "VariableName",
    "ModPortName",
    "FsmName",
    "FsmState",
    "Clock_Register",
    "RegName",
    "Module_Name",
    "NodeName",
]

FILTER_PRIORITY = [
    "Goal",
    "Module",
    "FileName",
    "LineNumber",
    "Statement",
    "HIERARCHY",
    "DesignObjHierarchy",
    "Signal",
    "VariableName",
    "ModPortName",
    "FsmName",
    "FsmState",
    "Clock_Register",
    "RegName",
    "Module_Name",
    "ExprSize",
]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

REVIEW_SHEET_COLUMNS = ["No.", "Person in Charge", "Date", "Judgment", "Comment"]
CORE_LINT_COLUMNS = [
    "Tag",
    "Description",
    "Violation",
    "Goal",
    "Module",
    "FileName",
    "LineNumber",
    "Statement",
]
MANAGEMENT_COLUMNS = [
    "issue_id",
    "record_status",
    "waiver_enabled",
    "waiver_name",
    "source_report",
    "waiver_user",
    "waiver_timestamp",
    "filter_json",
    "fields_json",
]


def default_paths(base_dir: Path) -> argparse.Namespace:
    return argparse.Namespace(
        full_report=base_dir / "reports" / "report_lint.full.log",
        waived_report=base_dir / "reports" / "report_lint.waived.log",
        waiver_tcl=base_dir / "vc_waiver.tcl",
        old_db=None,
        review_db=base_dir / "data" / "lint_review_db.csv",
        excel=base_dir / "outputs" / "lint_review.xlsx",
        summary=base_dir / "outputs" / "lint_summary.csv",
        generated_waiver=base_dir / "vc_waiver.tcl",
        waiver_audit=base_dir / "outputs" / "waiver_rule_audit.csv",
        user="",
    )


def run_all(base_dir: Path) -> None:
    args = default_paths(base_dir)
    if args.excel.exists():
        import_excel_to_db(args.excel, args.review_db)
        generate_waiver(args.review_db, args.generated_waiver, args.user)
        print(f"imported excel: {args.excel}")
        print(f"updated waiver: {args.generated_waiver}")
    update_review_from_reports(args)


def norm(value: str | None) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def issue_id(fields: dict[str, str]) -> str:
    payload = "|".join(f"{name}={norm(fields.get(name))}" for name in KEY_FIELDS if fields.get(name) is not None)
    digest = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]
    return f"{fields.get('Tag', 'LINT')}_{digest}".replace("/", "_").replace(" ", "_")


def parse_report(path: Path, source_report: str) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    current: OrderedDict[str, str] | None = None
    current_waiver: OrderedDict[str, str] | None = None
    section_tag = ""
    section_severity = ""
    current_key: str | None = None
    in_waiver = False
    section_re = re.compile(r"^\s*([A-Za-z0-9_.-]+)\s+\((\d+)\s+([A-Za-z]+)s?/(\d+)\s+waived\)")
    kv_re = re.compile(r"^\s{2,}([A-Za-z_][A-Za-z0-9_ -]*?)\s+:\s?(.*)$")

    def flush() -> None:
        nonlocal current, current_waiver, current_key, in_waiver
        if not current or "Tag" not in current:
            current = None
            current_waiver = None
            current_key = None
            in_waiver = False
            return
        fields = dict(current)
        waiver = dict(current_waiver or {})
        rows.append(
            {
                "fields": fields,
                "waiver": waiver,
                "source_report": source_report,
                "severity": section_severity,
            }
        )
        current = None
        current_waiver = None
        current_key = None
        in_waiver = False

    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if re.match(r"^\s*-{20,}\s*$", raw):
            flush()
            continue

        section = section_re.match(raw)
        if section:
            section_tag = section.group(1)
            section_severity = section.group(3).lower()
            continue

        if raw.strip() == "Waiver" and current is not None:
            in_waiver = True
            current_key = None
            if current_waiver is None:
                current_waiver = OrderedDict()
            continue

        kv = kv_re.match(raw)
        if kv:
            key = kv.group(1).strip()
            value = kv.group(2).rstrip()
            if key == "Tag":
                flush()
                current = OrderedDict()
                current_waiver = None
                in_waiver = False
            if current is None:
                continue
            target = current_waiver if in_waiver else current
            if target is None:
                continue
            target[key] = value
            current_key = key
            continue

        if current is not None and current_key and raw.startswith(" ") and raw.strip():
            target = current_waiver if in_waiver else current
            if target is not None:
                target[current_key] = f"{target.get(current_key, '')}\n{raw.rstrip()}"

    flush()

    # If a report section was parsed without a per-row severity, keep it useful.
    for row in rows:
        if not row.get("severity"):
            row["severity"] = section_severity
        fields = row["fields"]
        if isinstance(fields, dict) and section_tag and not fields.get("Tag"):
            fields["Tag"] = section_tag
    return rows


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, rows: list[dict[str, str]], columns: list[str] = BASE_COLUMNS) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    extras = sorted({key for row in rows for key in row.keys()} - set(columns))
    all_columns = columns + extras
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def parse_tcl_braced(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        raise ValueError("expected Tcl braced value")
    depth = 0
    out: list[str] = []
    i = start
    while i < len(text):
        ch = text[i]
        if ch == "\\" and i + 1 < len(text):
            if depth > 0:
                out.append(ch)
                out.append(text[i + 1])
            i += 2
            continue
        if ch == "{":
            depth += 1
            if depth > 1:
                out.append(ch)
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return "".join(out), i + 1
            out.append(ch)
        else:
            if depth > 0:
                out.append(ch)
        i += 1
    raise ValueError("unterminated Tcl braced value")


def parse_waiver_tcl(path: Path) -> dict[str, dict[str, object]]:
    rules: dict[str, dict[str, object]] = {}
    if not path.exists():
        return rules
    for lineno, line in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), start=1):
        if not line.strip().startswith("waive_violation"):
            continue
        record: dict[str, object] = {"line": str(lineno), "raw": line}
        pos = 0
        for option in ["-add", "-comment", "-filter", "-app", "-tag", "-user", "-timestamp"]:
            opt_pos = line.find(option, pos)
            if opt_pos == -1:
                continue
            brace_pos = line.find("{", opt_pos)
            if brace_pos == -1:
                continue
            value, end = parse_tcl_braced(line, brace_pos)
            record[option[1:]] = value.strip()
            pos = end
        name = str(record.get("add", "")).strip()
        if not name:
            name = f"waiver_line_{lineno}"
        filter_text = str(record.get("filter", ""))
        record["filter_fields"] = parse_filter(filter_text)
        rules[name] = record
    return rules


def parse_filter(filter_text: str) -> OrderedDict[str, str]:
    fields: OrderedDict[str, str] = OrderedDict()
    # VC waiver filters observed in the sample are simple AND-ed comparisons.
    pattern = re.compile(r'\(?\s*([A-Za-z_][A-Za-z0-9_]*)\s*(==|=~)\s*"((?:\\.|[^"])*)"\s*\)?')
    for key, _op, value in pattern.findall(filter_text):
        fields[key] = value.replace('\\"', '"')
    return fields


def tcl_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")


def filter_expr(filter_fields: dict[str, str]) -> str:
    parts = []
    for key, value in filter_fields.items():
        op = "=~" if "*" in value or "?" in value else "=="
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        parts.append(f'({key} {op} "{escaped}")')
    return " AND ".join(parts)


def fields_for_filter(fields: dict[str, str]) -> OrderedDict[str, str]:
    result: OrderedDict[str, str] = OrderedDict()
    for key in FILTER_PRIORITY:
        value = norm(fields.get(key))
        if value:
            result[key] = value
    return result


def row_from_issue(issue: dict[str, object], waiver_rules: dict[str, dict[str, object]]) -> dict[str, str]:
    fields = issue["fields"]
    waiver = issue.get("waiver") or {}
    assert isinstance(fields, dict)
    assert isinstance(waiver, dict)
    wid = issue_id(fields)
    waiver_name = norm(str(waiver.get("Name", "")))
    rule = waiver_rules.get(waiver_name, {})
    filter_fields = rule.get("filter_fields") if isinstance(rule, dict) else None
    if not isinstance(filter_fields, dict) or not filter_fields:
        filter_fields = fields_for_filter(fields)
    object_value = ""
    for key in OBJECT_FIELDS:
        if norm(fields.get(key)):
            object_value = norm(fields.get(key))
            break
    record_status = "WAIVED" if str(issue.get("source_report", "")) == "waived" else "ACTIVE"
    review_status = "WAIVED" if waiver_name else "UNREVIEWED"
    return {
        "issue_id": wid,
        "record_status": record_status,
        "review_status": review_status,
        "waiver_enabled": "yes" if waiver_name else "no",
        "waiver_name": waiver_name,
        "review_comment": norm(str(waiver.get("Comment") or rule.get("comment") or "")),
        "owner": "",
        "review_date": "",
        "tag": norm(fields.get("Tag")),
        "severity": norm(str(issue.get("severity", ""))),
        "goal": norm(fields.get("Goal")),
        "module": norm(fields.get("Module")),
        "file": norm(fields.get("FileName")),
        "line": norm(fields.get("LineNumber")),
        "hierarchy": norm(fields.get("HIERARCHY") or fields.get("DesignObjHierarchy")),
        "object": object_value,
        "statement": norm(fields.get("Statement")),
        "description": norm(fields.get("Description")),
        "violation": norm(fields.get("Violation")),
        "source_report": str(issue.get("source_report", "")),
        "waiver_user": norm(str(rule.get("user", ""))) if isinstance(rule, dict) else "",
        "waiver_timestamp": norm(str(rule.get("timestamp", ""))) if isinstance(rule, dict) else "",
        "filter_json": json.dumps(filter_fields, ensure_ascii=False),
        "fields_json": json.dumps(fields, ensure_ascii=False),
    }


def collect_current_rows(full_report: Path, waived_report: Path | None, waiver_tcl: Path | None) -> list[dict[str, str]]:
    waiver_rules = parse_waiver_tcl(waiver_tcl) if waiver_tcl else {}
    issues = parse_report(full_report, "full")
    if waived_report and waived_report.exists():
        issues.extend(parse_report(waived_report, "waived"))
    rows_by_id: dict[str, dict[str, str]] = {}
    for issue in issues:
        row = row_from_issue(issue, waiver_rules)
        rows_by_id[row["issue_id"]] = row
    return sorted(rows_by_id.values(), key=lambda r: (r["tag"], r["module"], r["file"], int(r["line"] or 0)))


def similar_issue_key(row: dict[str, str]) -> tuple[str, ...]:
    return (
        norm(row.get("tag")),
        norm(row.get("goal")),
        norm(row.get("module")),
        norm(row.get("file")),
        norm(row.get("hierarchy")),
        norm(row.get("object")),
    )


def preserve_review_fields(row: dict[str, str], old: dict[str, str]) -> None:
    for key in ["review_status", "waiver_enabled", "waiver_name", "review_comment", "owner", "review_date", "filter_json", "waiver_user", "waiver_timestamp"]:
        if old.get(key):
            row[key] = old[key]


def merge_rows(old_rows: list[dict[str, str]], current_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    old_by_id = {row["issue_id"]: row for row in old_rows}
    old_by_similar: dict[tuple[str, ...], list[dict[str, str]]] = {}
    for old in old_rows:
        old_by_similar.setdefault(similar_issue_key(old), []).append(old)
    current_by_id = {row["issue_id"]: row for row in current_rows}
    merged: list[dict[str, str]] = []
    consumed_old_ids: set[str] = set()

    for issue_id_value, current in current_by_id.items():
        old = old_by_id.get(issue_id_value)
        row = dict(current)
        if old:
            preserve_review_fields(row, old)
            row["record_status"] = current.get("record_status") or "ACTIVE"
            consumed_old_ids.add(old["issue_id"])
        elif current.get("record_status") != "WAIVED":
            similar_old = next((candidate for candidate in old_by_similar.get(similar_issue_key(current), []) if candidate.get("issue_id") not in consumed_old_ids), None)
            if similar_old:
                preserve_review_fields(row, similar_old)
                row["record_status"] = "CHANGED"
                consumed_old_ids.add(similar_old["issue_id"])
            else:
                row["record_status"] = "NEW"
        merged.append(row)

    for issue_id_value, old in old_by_id.items():
        if issue_id_value not in current_by_id and issue_id_value not in consumed_old_ids:
            row = dict(old)
            row["record_status"] = "REMOVED"
            merged.append(row)

    status_order = {"NEW": 0, "CHANGED": 1, "ACTIVE": 2, "WAIVED": 3, "REMOVED": 4}
    return sorted(merged, key=lambda r: (status_order.get(r["record_status"], 9), r["tag"], r["module"], r["file"], int(r["line"] or 0)))


def summarize(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    grouped = Counter((r.get("record_status", ""), r.get("review_status", ""), r.get("tag", "")) for r in rows)
    return [
        {"record_status": k[0], "review_status": k[1], "tag": k[2], "count": str(count)}
        for k, count in sorted(grouped.items())
    ]


def audit_waiver_rules(waiver_tcl: Path, rows: list[dict[str, str]], output_path: Path) -> None:
    rules = parse_waiver_tcl(waiver_tcl)
    matched = Counter(row.get("waiver_name", "") for row in rows if row.get("waiver_name"))
    audit_rows: list[dict[str, str]] = []
    for name, rule in sorted(rules.items()):
        count = matched.get(name, 0)
        audit_rows.append(
            {
                "waiver_name": name,
                "status": "ACTIVE" if count else "REDUNDANT",
                "matched_issue_count": str(count),
                "tag": str(rule.get("tag", "")),
                "comment": str(rule.get("comment", "")),
                "user": str(rule.get("user", "")),
                "timestamp": str(rule.get("timestamp", "")),
                "filter_json": json.dumps(rule.get("filter_fields", {}), ensure_ascii=False),
                "raw_tcl": str(rule.get("raw", "")),
            }
        )
    for name in sorted(set(matched) - set(rules)):
        if not name:
            continue
        audit_rows.append(
            {
                "waiver_name": name,
                "status": "MISSING_IN_TCL",
                "matched_issue_count": str(matched[name]),
                "tag": "",
                "comment": "Referenced by waived report but not found in waiver Tcl",
                "user": "",
                "timestamp": "",
                "filter_json": "",
                "raw_tcl": "",
            }
        )
    write_csv(
        output_path,
        audit_rows,
        ["waiver_name", "status", "matched_issue_count", "tag", "comment", "user", "timestamp", "filter_json", "raw_tcl"],
    )


def xlsx_col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def xlsx_col_index(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref)
    if not match:
        return 0
    index = 0
    for ch in match.group(1):
        index = index * 26 + ord(ch) - 64
    return index - 1


def safe_sheet_name(name: str, used: set[str]) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name or "Sheet").strip() or "Sheet"
    base = cleaned[:31]
    candidate = base
    suffix = 1
    while candidate.lower() in used:
        tail = f"_{suffix}"
        candidate = f"{base[:31 - len(tail)]}{tail}"
        suffix += 1
    used.add(candidate.lower())
    return candidate


def sheet_xml(rows: list[list[object]]) -> str:
    if not rows:
        rows = [[""]]
    col_count = max(len(row) for row in rows)
    row_count = len(rows)
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
        "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>",
        "<cols>",
    ]
    widths = {
        1: 7,
        2: 18,
        3: 13,
        4: 14,
        5: 42,
        7: 48,
        11: 64,
        13: 72,
    }
    for col_idx in range(1, col_count + 1):
        width = widths.get(col_idx, 18)
        out.append(f'<col min="{col_idx}" max="{col_idx}" width="{width}" customWidth="1"/>')
    out.extend([
        "</cols>",
        "<sheetData>",
    ])
    for r_idx, row in enumerate(rows, start=1):
        out.append(f'<row r="{r_idx}">')
        for c_idx, value in enumerate(row):
            cell_ref = f"{xlsx_col_name(c_idx)}{r_idx}"
            text = "" if value is None else str(value)
            out.append(f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{html.escape(text)}</t></is></c>')
        out.append("</row>")
    out.extend(["</sheetData>", '<autoFilter ref="A1:{}{}"/>'.format(xlsx_col_name(col_count - 1), row_count), "</worksheet>"])
    return "".join(out)


def write_xlsx(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_items = list(sheets.items())
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for idx, _ in enumerate(sheet_items, start=1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append("</Types>")

    workbook_sheets = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>',
    ]
    rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Relationships xmlns="{NS_PACKAGE_REL}">',
    ]
    used_names: set[str] = set()
    for idx, (name, _rows) in enumerate(sheet_items, start=1):
        safe_name = html.escape(safe_sheet_name(name, used_names))
        workbook_sheets.append(f'<sheet name="{safe_name}" sheetId="{idx}" r:id="rId{idx}"/>')
        rels.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    workbook_sheets.append("</sheets></workbook>")
    rels.append(f'<Relationship Id="rId{len(sheet_items)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    rels.append("</Relationships>")

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '</styleSheet>'
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_PACKAGE_REL}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(workbook_sheets))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        zf.writestr("xl/styles.xml", styles)
        for idx, (_name, rows) in enumerate(sheet_items, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(rows))


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values: list[str] = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        texts = [node.text or "" for node in si.findall(f".//{{{NS_MAIN}}}t")]
        values.append("".join(texts))
    return values


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text_el = cell.find(f".//{{{NS_MAIN}}}t")
        return text_el.text if text_el is not None and text_el.text is not None else ""
    value_el = cell.find(f"{{{NS_MAIN}}}v")
    if value_el is None or value_el.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value_el.text)]
        except (ValueError, IndexError):
            return ""
    return value_el.text


def read_xlsx_sheets(path: Path) -> dict[str, list[dict[str, str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in rels}
        sheets: dict[str, list[dict[str, str]]] = {}
        for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
            sheet_name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{NS_REL}}}id"]
            target = relmap[rel_id]
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            root = ET.fromstring(zf.read(target))
            matrix: list[list[str]] = []
            for row_el in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
                values: list[str] = []
                for cell in row_el.findall(f"{{{NS_MAIN}}}c"):
                    col_idx = xlsx_col_index(cell.attrib.get("r", ""))
                    while len(values) <= col_idx:
                        values.append("")
                    values[col_idx] = cell_text(cell, shared_strings)
                matrix.append(values)
            if not matrix:
                sheets[sheet_name] = []
                continue
            header = matrix[0]
            rows = []
            for row in matrix[1:]:
                if not any(row):
                    continue
                rows.append({header[i]: row[i] if i < len(row) else "" for i in range(len(header)) if header[i]})
            sheets[sheet_name] = rows
        return sheets


def read_xlsx_first_sheet(path: Path) -> list[dict[str, str]]:
    sheets = read_xlsx_sheets(path)
    return next(iter(sheets.values()), [])


def read_review_workbook(path: Path) -> list[dict[str, str]]:
    workbook_rows = read_xlsx_sheets(path)
    imported: list[dict[str, str]] = []
    for sheet_name, sheet_rows in workbook_rows.items():
        if sheet_name in {"Tree Summary", "Instructions", "Summary", "ReviewDB"}:
            if sheet_name == "ReviewDB":
                imported.extend(sheet_rows)
            continue
        for source in sheet_rows:
            if not source.get("issue_id"):
                continue
            try:
                fields = json.loads(source.get("fields_json", "{}") or "{}")
            except json.JSONDecodeError:
                fields = {}
            if not isinstance(fields, dict):
                fields = {}
            fields.update(
                {
                    "Tag": source.get("Tag", ""),
                    "Description": source.get("Description", ""),
                    "Violation": source.get("Violation", ""),
                    "Goal": source.get("Goal", ""),
                    "Module": source.get("Module", ""),
                    "FileName": source.get("FileName", ""),
                    "LineNumber": source.get("LineNumber", ""),
                    "Statement": source.get("Statement", ""),
                }
            )
            row = {col: source.get(col, "") for col in BASE_COLUMNS}
            row.update(
                {
                    "owner": source.get("Person in Charge", source.get("owner", "")),
                    "review_date": source.get("Date", source.get("review_date", "")),
                    "review_status": source.get("Judgment", source.get("review_status", "")),
                    "review_comment": source.get("Comment", source.get("review_comment", "")),
                    "tag": source.get("Tag", source.get("tag", "")),
                    "goal": source.get("Goal", source.get("goal", "")),
                    "module": source.get("Module", source.get("module", "")),
                    "file": source.get("FileName", source.get("file", "")),
                    "line": source.get("LineNumber", source.get("line", "")),
                    "hierarchy": source.get("HIERARCHY") or source.get("DesignObjHierarchy") or source.get("hierarchy", ""),
                    "object": source.get("Signal")
                    or source.get("VariableName")
                    or source.get("ModPortName")
                    or source.get("object", ""),
                    "statement": source.get("Statement", source.get("statement", "")),
                    "description": source.get("Description", source.get("description", "")),
                    "violation": source.get("Violation", source.get("violation", "")),
                    "fields_json": json.dumps(fields, ensure_ascii=False),
                }
            )
            imported.append(row)
    return imported


def tag_sheet_columns(rows: list[dict[str, str]]) -> list[str]:
    dynamic: list[str] = []
    seen = set(REVIEW_SHEET_COLUMNS + MANAGEMENT_COLUMNS)
    for col in CORE_LINT_COLUMNS:
        if col not in seen:
            dynamic.append(col)
            seen.add(col)
    for row in rows:
        try:
            fields = json.loads(row.get("fields_json", "{}") or "{}")
        except json.JSONDecodeError:
            fields = {}
        if not isinstance(fields, dict):
            continue
        for key in fields:
            if key not in seen:
                dynamic.append(key)
                seen.add(key)
    return REVIEW_SHEET_COLUMNS + dynamic + MANAGEMENT_COLUMNS


def row_to_tag_sheet(row: dict[str, str], columns: list[str], index: int) -> list[str]:
    try:
        fields = json.loads(row.get("fields_json", "{}") or "{}")
    except json.JSONDecodeError:
        fields = {}
    if not isinstance(fields, dict):
        fields = {}
    values = {
        "No.": str(index),
        "Person in Charge": row.get("owner", ""),
        "Date": row.get("review_date", ""),
        "Judgment": row.get("review_status", ""),
        "Comment": row.get("review_comment", ""),
        "Tag": row.get("tag", ""),
        "Description": row.get("description", ""),
        "Violation": row.get("violation", ""),
        "Goal": row.get("goal", ""),
        "Module": row.get("module", ""),
        "FileName": row.get("file", ""),
        "LineNumber": row.get("line", ""),
        "Statement": row.get("statement", ""),
    }
    values.update({key: str(value) for key, value in fields.items()})
    for key in MANAGEMENT_COLUMNS:
        values[key] = row.get(key, "")
    return [values.get(col, "") for col in columns]


def tree_summary_matrix(rows: list[dict[str, str]]) -> list[list[str]]:
    header = ["Severity", "Stage", "Tag", "Count", "Waived", "Compressed", "Confirmed", "Remaining"]
    active = [row for row in rows if row.get("record_status") != "REMOVED"]
    grouped: dict[tuple[str, str], dict[str, int | str]] = OrderedDict()
    for row in active:
        key = (row.get("severity", ""), row.get("tag", ""))
        if key not in grouped:
            grouped[key] = {"count": 0, "waived": 0, "stage": ""}
        if row.get("review_status", "").upper() == "WAIVED":
            grouped[key]["waived"] = int(grouped[key]["waived"]) + 1
        else:
            grouped[key]["count"] = int(grouped[key]["count"]) + 1
    matrix = [header]
    for (severity, tag), counts in grouped.items():
        confirmed = int(counts["count"]) + int(counts["waived"])
        matrix.append([severity, str(counts["stage"]), tag, str(counts["count"]), str(counts["waived"]), "0", str(confirmed), str(counts["count"])])
    matrix.append(["TOTAL", "", "", str(sum(1 for r in active if r.get("review_status", "").upper() != "WAIVED")), str(sum(1 for r in active if r.get("review_status", "").upper() == "WAIVED")), "0", str(len(active)), ""])
    return matrix


def export_excel(csv_path: Path, xlsx_path: Path, summary_path: Path | None = None) -> None:
    rows = read_csv(csv_path)
    summary_rows = summarize(rows)
    sheets: OrderedDict[str, list[list[object]]] = OrderedDict()
    sheets["Tree Summary"] = tree_summary_matrix(rows)

    by_tag: OrderedDict[str, list[dict[str, str]]] = OrderedDict()
    for row in rows:
        if row.get("record_status") == "REMOVED":
            continue
        by_tag.setdefault(row.get("tag", "UNKNOWN") or "UNKNOWN", []).append(row)
    for tag, tag_rows in by_tag.items():
        columns = tag_sheet_columns(tag_rows)
        sheets[tag] = [columns] + [row_to_tag_sheet(row, columns, idx) for idx, row in enumerate(tag_rows, start=1)]

    removed_rows = [row for row in rows if row.get("record_status") == "REMOVED"]
    if removed_rows:
        columns = tag_sheet_columns(removed_rows)
        sheets["Removed"] = [columns] + [row_to_tag_sheet(row, columns, idx) for idx, row in enumerate(removed_rows, start=1)]

    sheets["Instructions"] = [
        ["Sanity LINT Review Workbook"],
        ["This workbook follows the report_lint.full.xlsx style: Tree Summary plus one worksheet per tag."],
        ["Reviewer-editable columns are Person in Charge, Date, Judgment, Comment, waiver_enabled, waiver_name, and filter_json."],
        ["Use '*' or '?' in filter_json values to generate Tcl '=~' wildcard filters."],
        ["Git should review LINT/data/lint_review_db.csv, not this binary Excel workbook."],
    ]
    write_xlsx(xlsx_path, sheets)
    if summary_path:
        write_csv(summary_path, summary_rows, ["record_status", "review_status", "tag", "count"])


def generate_waiver(csv_path: Path, output_path: Path, user: str = "") -> None:
    rows = read_csv(csv_path)
    lines = [
        "# Generated by sanity_lint_review.py",
        "# Source: review DB CSV",
        "",
    ]
    seen: set[str] = set()
    for row in rows:
        if row.get("record_status") == "REMOVED":
            continue
        if row.get("waiver_enabled", "").strip().lower() not in {"yes", "y", "true", "1"}:
            continue
        if row.get("review_status", "").strip().upper() not in {"WAIVED", "APPROVED", "APPROVED_WAIVE"}:
            continue
        name = row.get("waiver_name") or row.get("issue_id") or f"{row.get('tag', 'LINT')}_{len(lines)}"
        if name in seen:
            continue
        seen.add(name)
        try:
            filter_fields = json.loads(row.get("filter_json", "{}") or "{}")
        except json.JSONDecodeError:
            filter_fields = {}
        if not isinstance(filter_fields, dict) or not filter_fields:
            try:
                fields = json.loads(row.get("fields_json", "{}") or "{}")
            except json.JSONDecodeError:
                fields = {}
            filter_fields = fields_for_filter(fields if isinstance(fields, dict) else {})
        expr = filter_expr(filter_fields)
        comment = row.get("review_comment", "")
        tag = row.get("tag", "")
        out_user = user or row.get("waiver_user") or os.getenv("USERNAME") or ""
        timestamp = row.get("waiver_timestamp") or "N/A"
        lines.append(
            f"waive_violation -add {{{tcl_escape(name)}}}  "
            f"-comment {{{tcl_escape(comment)}}}  "
            f"-filter {{{expr}}}  -app {{ lint }} -tag {{ {tcl_escape(tag)} }} "
            f"-user {{ {tcl_escape(out_user)} }} -timestamp {{ {tcl_escape(timestamp)} }}"
        )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def import_excel_to_db(excel_path: Path, review_db_path: Path) -> list[dict[str, str]]:
    rows = read_review_workbook(excel_path)
    write_csv(review_db_path, rows)
    return rows


def update_review_from_reports(args: argparse.Namespace) -> list[dict[str, str]]:
    rows = collect_current_rows(args.full_report, args.waived_report, args.waiver_tcl)
    old_db = args.old_db or args.review_db
    if old_db and old_db.exists():
        rows = merge_rows(read_csv(old_db), rows)
    write_csv(args.review_db, rows)
    export_excel(args.review_db, args.excel, args.summary)
    if args.waiver_tcl and args.waiver_audit:
        audit_waiver_rules(args.waiver_tcl, rows, args.waiver_audit)
    print(f"review rows: {len(rows)}")
    print(f"review db   : {args.review_db}")
    print(f"excel       : {args.excel}")
    print(f"summary     : {args.summary}")
    if args.waiver_tcl and args.waiver_audit:
        print(f"waiver audit: {args.waiver_audit}")
    return rows


def cmd_bootstrap(args: argparse.Namespace) -> None:
    rows = collect_current_rows(args.full_report, args.waived_report, args.waiver_tcl)
    if args.old_db and args.old_db.exists():
        rows = merge_rows(read_csv(args.old_db), rows)
    write_csv(args.review_db, rows)
    export_excel(args.review_db, args.excel, args.summary)
    generate_waiver(args.review_db, args.generated_waiver, args.user)
    if args.waiver_tcl and args.waiver_audit:
        audit_waiver_rules(args.waiver_tcl, rows, args.waiver_audit)
    print(f"review rows: {len(rows)}")
    print(f"review db   : {args.review_db}")
    print(f"excel       : {args.excel}")
    print(f"waiver tcl  : {args.generated_waiver}")
    print(f"summary     : {args.summary}")
    if args.waiver_tcl and args.waiver_audit:
        print(f"waiver audit: {args.waiver_audit}")


def cmd_update_review(args: argparse.Namespace) -> None:
    update_review_from_reports(args)


def cmd_parse(args: argparse.Namespace) -> None:
    rows = collect_current_rows(args.full_report, args.waived_report, args.waiver_tcl)
    write_csv(args.output, rows)
    print(f"wrote {len(rows)} rows to {args.output}")


def cmd_merge(args: argparse.Namespace) -> None:
    current = read_csv(args.current_db)
    old = read_csv(args.old_db)
    merged = merge_rows(old, current)
    write_csv(args.output, merged)
    print(f"wrote {len(merged)} rows to {args.output}")


def cmd_export_excel(args: argparse.Namespace) -> None:
    export_excel(args.review_db, args.excel, args.summary)
    print(f"wrote {args.excel}")


def cmd_import_excel(args: argparse.Namespace) -> None:
    rows = import_excel_to_db(args.excel, args.output)
    print(f"wrote {len(rows)} rows to {args.output}")


def cmd_generate(args: argparse.Namespace) -> None:
    generate_waiver(args.review_db, args.output, args.user)
    print(f"wrote {args.output}")


def cmd_generate_from_excel(args: argparse.Namespace) -> None:
    rows = import_excel_to_db(args.excel, args.review_db)
    generate_waiver(args.review_db, args.output, args.user)
    print(f"imported {len(rows)} rows from {args.excel}")
    print(f"wrote {args.output}")


def cmd_audit_waivers(args: argparse.Namespace) -> None:
    rows = read_csv(args.review_db)
    audit_waiver_rules(args.waiver_tcl, rows, args.output)
    print(f"wrote {args.output}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Manage LINT sanity review with CSV/Excel/waiver Tcl.",
        epilog="Run without arguments from the LINT directory to import Excel edits, generate vc_waiver.tcl, then update review DB/Excel from reports.",
    )
    sub = parser.add_subparsers()

    common_reports = argparse.ArgumentParser(add_help=False)
    common_reports.add_argument("--full-report", type=Path, required=True)
    common_reports.add_argument("--waived-report", type=Path)
    common_reports.add_argument("--waiver-tcl", type=Path)

    p = sub.add_parser("run_all", help="Run the default LINT flow without path options.")
    p.set_defaults(func=lambda _args: run_all(Path(__file__).resolve().parent))

    p = sub.add_parser("update-review", parents=[common_reports], help="Merge current reports into the review DB and export reviewer Excel.")
    p.add_argument("--old-db", type=Path)
    p.add_argument("--review-db", type=Path, default=Path("data/lint_review_db.csv"))
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--summary", type=Path, default=Path("outputs/lint_summary.csv"))
    p.add_argument("--waiver-audit", type=Path, default=Path("outputs/waiver_rule_audit.csv"))
    p.set_defaults(func=cmd_update_review)

    p = sub.add_parser("bootstrap", parents=[common_reports], help="Create sample review DB, Excel workbook, summary, and generated waiver Tcl.")
    p.add_argument("--old-db", type=Path)
    p.add_argument("--review-db", type=Path, default=Path("data/lint_review_db.csv"))
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--summary", type=Path, default=Path("outputs/lint_summary.csv"))
    p.add_argument("--generated-waiver", type=Path, default=Path("outputs/generated_vc_waiver.tcl"))
    p.add_argument("--waiver-audit", type=Path, default=Path("outputs/waiver_rule_audit.csv"))
    p.add_argument("--user", default="")
    p.set_defaults(func=cmd_bootstrap)

    p = sub.add_parser("parse", parents=[common_reports], help="Parse current reports into a review DB CSV.")
    p.add_argument("--output", type=Path, required=True)
    p.set_defaults(func=cmd_parse)

    p = sub.add_parser("merge", help="Merge previous review DB with current parsed DB.")
    p.add_argument("--old-db", type=Path, required=True)
    p.add_argument("--current-db", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("export-excel", help="Export review DB CSV to reviewer Excel workbook.")
    p.add_argument("--review-db", type=Path, required=True)
    p.add_argument("--excel", type=Path, required=True)
    p.add_argument("--summary", type=Path)
    p.set_defaults(func=cmd_export_excel)

    p = sub.add_parser("import-excel", help="Import the ReviewDB sheet from Excel back to CSV.")
    p.add_argument("--excel", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.set_defaults(func=cmd_import_excel)

    p = sub.add_parser("generate-waiver", help="Generate VC LINT waiver Tcl from review DB CSV.")
    p.add_argument("--review-db", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--user", default="")
    p.set_defaults(func=cmd_generate)

    p = sub.add_parser("generate-waiver-from-excel", help="Import reviewer Excel to DB, then generate VC LINT waiver Tcl.")
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--review-db", type=Path, default=Path("data/lint_review_db.csv"))
    p.add_argument("--output", type=Path, default=Path("vc_waiver.tcl"))
    p.add_argument("--user", default="")
    p.set_defaults(func=cmd_generate_from_excel)

    p = sub.add_parser("audit-waivers", help="Mark Tcl waiver rules as ACTIVE or REDUNDANT against a review DB.")
    p.add_argument("--review-db", type=Path, required=True)
    p.add_argument("--waiver-tcl", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.set_defaults(func=cmd_audit_waivers)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    argv = list(argv)
    if not argv:
        run_all(Path(__file__).resolve().parent)
        return 0
    parser = build_parser()
    args = parser.parse_args(argv)
    if not hasattr(args, "func"):
        parser.print_help()
        return 2
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
