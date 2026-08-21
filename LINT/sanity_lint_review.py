#!/usr/bin/env python3
"""LINT sanity review bridge.

This tool merges VC Static LINT Excel reports into an Excel reviewer workbook
and regenerates waiver Tcl from reviewer decisions.
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
import subprocess
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
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

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

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def default_paths(base_dir: Path) -> argparse.Namespace:
    return argparse.Namespace(
        full_report=base_dir / "reports" / "report_lint.full.log",
        report_excel=base_dir / "report_lint.full.xlsx",
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
    print("The review flow is split into two explicit commands:")
    print("")
    print("  1) python sanity_lint_review.py prepare-waiver")
    print("  2) Run the sanity tool so reports/report_lint.full.log is updated")
    print("  3) python sanity_lint_review.py merge-report")
    print("")
    print("No files were changed.")


def run_make_excel(base_dir: Path) -> None:
    subprocess.run(["make", "-f", "Makefile", "excel"], cwd=base_dir, check=True)


def generated_files(base_dir: Path, keep_waiver: bool = False) -> list[Path]:
    paths = [
        base_dir / "data" / "current_lint_review_db.csv",
        base_dir / "outputs" / "lint_review.xlsx",
        base_dir / "outputs" / "lint_summary.csv",
        base_dir / "outputs" / "waiver_rule_audit.csv",
        base_dir / "outputs" / "imported_from_excel.csv",
        base_dir / "outputs" / "generated_from_imported_excel.tcl",
        base_dir / "outputs" / "generated_vc_waiver.tcl",
    ]
    if not keep_waiver:
        paths.append(base_dir / "vc_waiver.tcl")
    return paths


def clean_generated_files(base_dir: Path, keep_waiver: bool = False, dry_run: bool = False) -> list[Path]:
    base_dir = base_dir.resolve()
    removed: list[Path] = []
    for path in generated_files(base_dir, keep_waiver=keep_waiver):
        resolved = path.resolve()
        if base_dir != resolved and base_dir not in resolved.parents:
            raise RuntimeError(f"Refusing to clean outside LINT directory: {resolved}")
        if not resolved.exists():
            continue
        if resolved.is_dir():
            continue
        removed.append(resolved)
        if not dry_run:
            resolved.unlink()
    return removed


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


def sanitize_waiver_tag(tag: str) -> str:
    return tag.replace(".", "_").replace("/", "_").replace(" ", "_")


def violation_number(row: dict[str, str]) -> str:
    match = re.search(r"Lint:(\d+)", row.get("violation", ""))
    return match.group(1) if match else ""


def generated_waiver_name(row: dict[str, str], duplicate_waiver_names: set[str], used_names: set[str]) -> str:
    lint_number = violation_number(row)
    current_name = row.get("waiver_name", "").strip()
    if current_name and current_name not in duplicate_waiver_names:
        base = current_name
    elif lint_number:
        base = f"{sanitize_waiver_tag(row.get('tag', 'LINT'))}_{lint_number}"
    else:
        base = current_name or row.get("issue_id") or f"{sanitize_waiver_tag(row.get('tag', 'LINT'))}_{len(used_names) + 1}"

    name = base
    suffix = 1
    while name in used_names:
        suffix += 1
        name = f"{base}_{suffix}"
    used_names.add(name)
    return name


def should_emit_filter(row: dict[str, str]) -> bool:
    # The existing GUI-generated Tcl is mostly one named waiver per violation
    # without filters. W551 examples in vc_waiver.tcl_old use filters to keep
    # waivers specific, so preserve that style for W551.
    return row.get("tag") == "W551"


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
    for key in ["review_status", "review_comment"]:
        if old.get(key):
            row[key] = old[key]
    row["waiver_enabled"] = "yes" if row.get("review_status", "").upper() in {"WAIVED", "APPROVED", "APPROVED_WAIVE"} else "no"


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


def sheet_xml(rows: list[list[object]], editable_headers: set[str] | None = None) -> str:
    if not rows:
        rows = [[""]]
    editable_headers = editable_headers or set()
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
            style = ""
            if r_idx == 1:
                style_id = 1 if text in editable_headers else 2
                style = f' s="{style_id}"'
            out.append(f'<c r="{cell_ref}"{style} t="inlineStr"><is><t xml:space="preserve">{html.escape(text)}</t></is></c>')
        out.append("</row>")
    out.extend(["</sheetData>", '<autoFilter ref="A1:{}{}"/>'.format(xlsx_col_name(col_count - 1), row_count)])
    header = ["" if value is None else str(value) for value in rows[0]]
    if "Judgment" in header and row_count > 1:
        judgment_col = xlsx_col_name(header.index("Judgment"))
        validation_range = f"{judgment_col}2:{judgment_col}{row_count}"
        out.append(
            '<dataValidations count="1">'
            f'<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="{validation_range}">'
            '<formula1>"UNREVIEWED,WAIVED,APPROVED,APPROVED_WAIVE"</formula1>'
            "</dataValidation>"
            "</dataValidations>"
        )
    out.append("</worksheet>")
    return "".join(out)


def write_xlsx(path: Path, sheets: dict[str, list[list[object]]], editable_headers_by_sheet: dict[str, set[str]] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    editable_headers_by_sheet = editable_headers_by_sheet or {}
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
    safe_sheet_names: list[str] = []
    for idx, (name, _rows) in enumerate(sheet_items, start=1):
        safe_name_raw = safe_sheet_name(name, used_names)
        safe_sheet_names.append(safe_name_raw)
        safe_name = html.escape(safe_name_raw)
        workbook_sheets.append(f'<sheet name="{safe_name}" sheetId="{idx}" r:id="rId{idx}"/>')
        rels.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    workbook_sheets.append("</sheets>")
    defined_names: list[str] = []
    for idx, (safe_name, (_name, rows)) in enumerate(zip(safe_sheet_names, sheet_items)):
        if not rows:
            continue
        row_count = len(rows)
        col_count = max(len(row) for row in rows)
        quoted_name = safe_name.replace("'", "''")
        filter_ref = f"'{quoted_name}'!$A$1:${xlsx_col_name(col_count - 1)}${row_count}"
        defined_names.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{idx}" hidden="1">{html.escape(filter_ref)}</definedName>')
    if defined_names:
        workbook_sheets.append("<definedNames>")
        workbook_sheets.extend(defined_names)
        workbook_sheets.append("</definedNames>")
    workbook_sheets.append("</workbook>")
    rels.append(f'<Relationship Id="rId{len(sheet_items)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    rels.append("</Relationships>")

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '</cellXfs>'
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
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(rows, editable_headers_by_sheet.get(_name, set())))


def xlsx_sheet_targets(zf: zipfile.ZipFile) -> OrderedDict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in rels}
    targets: OrderedDict[str, str] = OrderedDict()
    for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
        sheet_name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{NS_REL}}}id"]
        target = relmap[rel_id]
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        targets[sheet_name] = posixpath.normpath(target)
    return targets


def rels_path_for_part(part_name: str) -> str:
    directory = posixpath.dirname(part_name)
    basename = posixpath.basename(part_name)
    return f"{directory}/_rels/{basename}.rels"


def resolve_part_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative_part_target(source_part: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def read_rels(zf: zipfile.ZipFile, rels_path: str) -> ET.Element:
    try:
        return ET.fromstring(zf.read(rels_path))
    except KeyError:
        return ET.Element(f"{{{NS_PACKAGE_REL}}}Relationships")


def next_rel_id(rels_root: ET.Element) -> str:
    used = {
        int(match.group(1))
        for rel in rels_root.findall(f"{{{NS_PACKAGE_REL}}}Relationship")
        if (match := re.fullmatch(r"rId(\d+)", rel.attrib.get("Id", "")))
    }
    value = 1
    while value in used:
        value += 1
    return f"rId{value}"


def merge_content_types(old_zf: zipfile.ZipFile, files: dict[str, bytes], copied_parts: set[str]) -> None:
    if not copied_parts:
        return
    new_root = ET.fromstring(files["[Content_Types].xml"])
    try:
        old_root = ET.fromstring(old_zf.read("[Content_Types].xml"))
    except KeyError:
        old_root = ET.Element(f"{{{NS_CONTENT_TYPES}}}Types")
    existing_defaults = {node.attrib.get("Extension") for node in new_root.findall(f"{{{NS_CONTENT_TYPES}}}Default")}
    existing_overrides = {node.attrib.get("PartName") for node in new_root.findall(f"{{{NS_CONTENT_TYPES}}}Override")}

    needed_exts = {part.rsplit(".", 1)[-1] for part in copied_parts if "." in posixpath.basename(part)}
    for node in old_root.findall(f"{{{NS_CONTENT_TYPES}}}Default"):
        ext = node.attrib.get("Extension")
        if ext in needed_exts and ext not in existing_defaults:
            new_root.append(ET.fromstring(ET.tostring(node)))
            existing_defaults.add(ext)
    for node in old_root.findall(f"{{{NS_CONTENT_TYPES}}}Override"):
        part_name = node.attrib.get("PartName", "").lstrip("/")
        if part_name in copied_parts and node.attrib.get("PartName") not in existing_overrides:
            new_root.append(ET.fromstring(ET.tostring(node)))
            existing_overrides.add(node.attrib.get("PartName"))
    files["[Content_Types].xml"] = ET.tostring(new_root, encoding="utf-8", xml_declaration=True)


def preserve_workbook_drawings(previous_excel: Path | None, xlsx_path: Path) -> None:
    if not previous_excel or not previous_excel.exists() or not xlsx_path.exists():
        return
    with zipfile.ZipFile(xlsx_path, "r") as new_zf:
        files = {name: new_zf.read(name) for name in new_zf.namelist()}
    copied_parts: set[str] = set()
    changed = False

    with zipfile.ZipFile(previous_excel, "r") as old_zf:
        old_targets = xlsx_sheet_targets(old_zf)
        with zipfile.ZipFile(xlsx_path, "r") as new_zf:
            new_targets = xlsx_sheet_targets(new_zf)
        for sheet_name, old_sheet_part in old_targets.items():
            new_sheet_part = new_targets.get(sheet_name)
            if not new_sheet_part:
                continue
            old_sheet_root = ET.fromstring(old_zf.read(old_sheet_part))
            old_drawings = old_sheet_root.findall(f"{{{NS_MAIN}}}drawing")
            if not old_drawings:
                continue
            old_sheet_rels = read_rels(old_zf, rels_path_for_part(old_sheet_part))
            old_rel_by_id = {rel.attrib.get("Id"): rel for rel in old_sheet_rels.findall(f"{{{NS_PACKAGE_REL}}}Relationship")}

            new_sheet_root = ET.fromstring(files[new_sheet_part])
            new_sheet_rels_path = rels_path_for_part(new_sheet_part)
            new_sheet_rels = read_rels_from_bytes(files.get(new_sheet_rels_path))

            for drawing in old_drawings:
                old_rid = drawing.attrib.get(f"{{{NS_REL}}}id")
                old_rel = old_rel_by_id.get(old_rid)
                if old_rel is None:
                    continue
                old_drawing_part = resolve_part_target(old_sheet_part, old_rel.attrib.get("Target", ""))
                if old_drawing_part not in old_zf.namelist():
                    continue

                copy_related_drawing_parts(old_zf, files, old_drawing_part, copied_parts)
                new_rid = next_rel_id(new_sheet_rels)
                ET.SubElement(
                    new_sheet_rels,
                    f"{{{NS_PACKAGE_REL}}}Relationship",
                    {
                        "Id": new_rid,
                        "Type": old_rel.attrib.get("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"),
                        "Target": relative_part_target(new_sheet_part, old_drawing_part),
                    },
                )
                ET.SubElement(new_sheet_root, f"{{{NS_MAIN}}}drawing", {f"{{{NS_REL}}}id": new_rid})
                changed = True

            files[new_sheet_part] = ET.tostring(new_sheet_root, encoding="utf-8", xml_declaration=True)
            files[new_sheet_rels_path] = ET.tostring(new_sheet_rels, encoding="utf-8", xml_declaration=True)

        merge_content_types(old_zf, files, copied_parts)

    if changed:
        with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zf:
            for name, data in files.items():
                out_zf.writestr(name, data)


def read_rels_from_bytes(data: bytes | None) -> ET.Element:
    if data:
        return ET.fromstring(data)
    return ET.Element(f"{{{NS_PACKAGE_REL}}}Relationships")


def copy_related_drawing_parts(old_zf: zipfile.ZipFile, files: dict[str, bytes], drawing_part: str, copied_parts: set[str]) -> None:
    if drawing_part not in files:
        files[drawing_part] = old_zf.read(drawing_part)
    copied_parts.add(drawing_part)
    drawing_rels_path = rels_path_for_part(drawing_part)
    if drawing_rels_path not in old_zf.namelist():
        return
    if drawing_rels_path not in files:
        files[drawing_rels_path] = old_zf.read(drawing_rels_path)
    copied_parts.add(drawing_rels_path)
    drawing_rels = ET.fromstring(old_zf.read(drawing_rels_path))
    for rel in drawing_rels.findall(f"{{{NS_PACKAGE_REL}}}Relationship"):
        target = rel.attrib.get("Target", "")
        if not target or target.startswith("http://") or target.startswith("https://"):
            continue
        target_part = resolve_part_target(drawing_part, target)
        if target_part in old_zf.namelist() and target_part not in files:
            files[target_part] = old_zf.read(target_part)
            copied_parts.add(target_part)


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


def xlsx_row_fields(source: dict[str, str]) -> dict[str, str]:
    ignored = set(REVIEW_SHEET_COLUMNS + MANAGEMENT_COLUMNS + ["record_status", "review_status", "review_comment"])
    return {key: value for key, value in source.items() if key and key not in ignored}


def severity_by_tag_from_workbook(workbook_rows: dict[str, list[dict[str, str]]]) -> dict[str, str]:
    severity_by_tag: dict[str, str] = {}
    for row in workbook_rows.get("Tree Summary", []):
        tag = norm(row.get("Tag"))
        if tag:
            severity_by_tag[tag] = norm(row.get("Severity"))
    return severity_by_tag


def row_from_xlsx_issue(source: dict[str, str], severity_by_tag: dict[str, str] | None = None) -> dict[str, str]:
    fields = xlsx_row_fields(source)
    wid = source.get("issue_id") or issue_id(fields)
    object_value = ""
    for key in OBJECT_FIELDS:
        if norm(fields.get(key)):
            object_value = norm(fields.get(key))
            break
    review_status = norm(source.get("Judgment") or source.get("review_status") or "UNREVIEWED")
    return {
        "issue_id": wid,
        "record_status": norm(source.get("record_status")) or "ACTIVE",
        "review_status": review_status,
        "waiver_enabled": "yes" if review_status.upper() in {"WAIVED", "APPROVED", "APPROVED_WAIVE"} else "no",
        "waiver_name": norm(source.get("waiver_name")),
        "review_comment": norm(source.get("Comment") or source.get("review_comment")),
        "owner": "",
        "review_date": "",
        "tag": norm(fields.get("Tag")),
        "severity": norm((severity_by_tag or {}).get(norm(fields.get("Tag")), "")),
        "goal": norm(fields.get("Goal")),
        "module": norm(fields.get("Module")),
        "file": norm(fields.get("FileName")),
        "line": norm(fields.get("LineNumber")),
        "hierarchy": norm(fields.get("HIERARCHY") or fields.get("DesignObjHierarchy")),
        "object": object_value,
        "statement": norm(fields.get("Statement")),
        "description": norm(fields.get("Description")),
        "violation": norm(fields.get("Violation")),
        "source_report": "full",
        "waiver_user": "",
        "waiver_timestamp": "",
        "filter_json": json.dumps(fields_for_filter(fields), ensure_ascii=False),
        "fields_json": json.dumps(fields, ensure_ascii=False),
    }


def collect_current_rows_from_report_excel(report_excel: Path) -> list[dict[str, str]]:
    workbook_rows = read_xlsx_sheets(report_excel)
    severity_by_tag = severity_by_tag_from_workbook(workbook_rows)
    rows_by_id: dict[str, dict[str, str]] = {}
    for sheet_name, sheet_rows in workbook_rows.items():
        if sheet_name in {"Tree Summary", "Instructions", "Summary", "ReviewDB"}:
            continue
        for source in sheet_rows:
            if not source.get("Tag"):
                continue
            row = row_from_xlsx_issue(source, severity_by_tag)
            rows_by_id[row["issue_id"]] = row
    return sorted(rows_by_id.values(), key=lambda r: (r["tag"], r["module"], r["file"], int(r["line"] or 0)))


def read_review_workbook(path: Path) -> list[dict[str, str]]:
    workbook_rows = read_xlsx_sheets(path)
    severity_by_tag = severity_by_tag_from_workbook(workbook_rows)
    imported: list[dict[str, str]] = []
    for sheet_name, sheet_rows in workbook_rows.items():
        if sheet_name in {"Tree Summary", "Instructions", "Summary", "ReviewDB"}:
            if sheet_name == "ReviewDB":
                imported.extend(sheet_rows)
            continue
        for source in sheet_rows:
            if not source.get("Tag") and not source.get("issue_id"):
                continue
            if not source.get("issue_id"):
                imported.append(row_from_xlsx_issue(source, severity_by_tag))
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


def workbook_headers(path: Path) -> OrderedDict[str, list[str]]:
    headers: OrderedDict[str, list[str]] = OrderedDict()
    if not path.exists():
        return headers
    sheets = read_xlsx_sheets(path)
    for sheet_name, rows in sheets.items():
        if rows:
            headers[sheet_name] = list(rows[0].keys())
    return headers


def issue_id_from_workbook_source(source: dict[str, str]) -> str:
    if source.get("issue_id"):
        return source["issue_id"]
    return issue_id(xlsx_row_fields(source))


def collect_user_extra_columns(previous_excel: Path | None, report_headers: dict[str, list[str]]) -> tuple[dict[str, list[str]], dict[str, dict[str, str]]]:
    extras_by_sheet: dict[str, list[str]] = {}
    values_by_issue: dict[str, dict[str, str]] = {}
    if not previous_excel or not previous_excel.exists():
        return extras_by_sheet, values_by_issue

    old_sheets = read_xlsx_sheets(previous_excel)
    ignored = set(MANAGEMENT_COLUMNS + ["record_status", "review_status", "review_comment"])
    for sheet_name, rows in old_sheets.items():
        if sheet_name in {"Tree Summary", "Instructions", "Summary", "ReviewDB"} or not rows:
            continue
        report_cols = set(report_headers.get(sheet_name, []))
        extras = [col for col in rows[0].keys() if col not in report_cols and col not in ignored]
        if extras:
            extras_by_sheet[sheet_name] = extras
        for source in rows:
            try:
                wid = issue_id_from_workbook_source(source)
            except Exception:
                continue
            values_by_issue[wid] = {col: source.get(col, "") for col in extras}
    return extras_by_sheet, values_by_issue


def report_ordered_sheet_columns(
    tag: str,
    tag_rows: list[dict[str, str]],
    report_headers: dict[str, list[str]],
    extras_by_sheet: dict[str, list[str]],
) -> list[str]:
    columns = list(report_headers.get(tag, []))
    if not columns:
        columns = tag_sheet_columns(tag_rows)
    if "record_status" not in columns:
        insert_at = columns.index("Comment") + 1 if "Comment" in columns else len(columns)
        columns.insert(insert_at, "record_status")
    for extra in extras_by_sheet.get(tag, []):
        if extra not in columns:
            columns.append(extra)
    return columns


def row_to_report_sheet(row: dict[str, str], columns: list[str], index: int, user_extra_values: dict[str, dict[str, str]]) -> list[str]:
    try:
        fields = json.loads(row.get("fields_json", "{}") or "{}")
    except json.JSONDecodeError:
        fields = {}
    if not isinstance(fields, dict):
        fields = {}
    values = {
        "No.": str(index),
        "Person in Charge": "",
        "Date": "",
        "Judgment": row.get("review_status", ""),
        "Comment": row.get("review_comment", ""),
        "record_status": row.get("record_status", ""),
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
    values.update(user_extra_values.get(row.get("issue_id", ""), {}))
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


def export_review_workbook(rows: list[dict[str, str]], xlsx_path: Path, summary_path: Path | None = None, report_excel: Path | None = None, previous_excel: Path | None = None) -> None:
    summary_rows = summarize(rows)
    sheets: OrderedDict[str, list[list[object]]] = OrderedDict()
    editable_headers_by_sheet: dict[str, set[str]] = {}
    sheets["Tree Summary"] = tree_summary_matrix(rows)
    report_headers = workbook_headers(report_excel) if report_excel else OrderedDict()
    extras_by_sheet, user_extra_values = collect_user_extra_columns(previous_excel, report_headers)

    by_tag: OrderedDict[str, list[dict[str, str]]] = OrderedDict()
    for row in rows:
        if row.get("record_status") == "REMOVED":
            continue
        by_tag.setdefault(row.get("tag", "UNKNOWN") or "UNKNOWN", []).append(row)
    for tag, tag_rows in by_tag.items():
        columns = report_ordered_sheet_columns(tag, tag_rows, report_headers, extras_by_sheet)
        sheets[tag] = [columns] + [row_to_report_sheet(row, columns, idx, user_extra_values) for idx, row in enumerate(tag_rows, start=1)]
        editable_headers_by_sheet[tag] = {"Judgment", "Comment", *extras_by_sheet.get(tag, [])}

    removed_rows = [row for row in rows if row.get("record_status") == "REMOVED"]
    if removed_rows:
        columns = tag_sheet_columns(removed_rows)
        sheets["Removed"] = [columns] + [row_to_tag_sheet(row, columns, idx) for idx, row in enumerate(removed_rows, start=1)]
        editable_headers_by_sheet["Removed"] = {"Judgment", "Comment"}

    sheets["Instructions"] = [
        ["Sanity LINT Review Workbook"],
        ["Run make -f Makefile excel to generate report_lint.full.xlsx, then this workbook is merged from it."],
        ["Reviewer-editable report columns are Judgment and Comment."],
        ["User-added columns are preserved for human notes but ignored by vc_waiver.tcl generation."],
        ["This workbook is the review source of truth."],
    ]
    write_xlsx(xlsx_path, sheets, editable_headers_by_sheet)
    preserve_workbook_drawings(previous_excel, xlsx_path)
    if summary_path:
        write_csv(summary_path, summary_rows, ["record_status", "review_status", "tag", "count"])


def export_excel(csv_path: Path, xlsx_path: Path, summary_path: Path | None = None, report_excel: Path | None = None, previous_excel: Path | None = None) -> None:
    export_review_workbook(read_csv(csv_path), xlsx_path, summary_path, report_excel, previous_excel)


def generate_waiver_from_rows(rows: list[dict[str, str]], output_path: Path, user: str = "") -> None:
    lines = [
        "# Generated by sanity_lint_review.py",
        "# Source: lint_review.xlsx",
        "",
    ]
    eligible_rows = [
        row
        for row in rows
        if row.get("record_status") != "REMOVED"
        and row.get("review_status", "").strip().upper() in {"WAIVED", "APPROVED", "APPROVED_WAIVE"}
    ]
    duplicate_waiver_names = {
        name for name, count in Counter(row.get("waiver_name", "").strip() for row in eligible_rows if row.get("waiver_name", "").strip()).items() if count > 1
    }
    used_names: set[str] = set()
    for row in eligible_rows:
        if row.get("record_status") == "REMOVED":
            continue
        name = generated_waiver_name(row, duplicate_waiver_names, used_names)
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
        comment = row.get("review_comment", "")
        tag = row.get("tag", "")
        out_user = user or row.get("waiver_user") or os.getenv("USERNAME") or ""
        timestamp = row.get("waiver_timestamp") or "N/A"
        filter_part = f" -filter {{{filter_expr(filter_fields)}}} " if should_emit_filter(row) and filter_fields else " "
        lines.append(
            f"waive_violation -add {{{tcl_escape(name)}}}  "
            f"-comment {{{tcl_escape(comment)}}} "
            f"{filter_part} -app {{ lint }} -tag {{ {tcl_escape(tag)} }} "
            f"-user {{ {tcl_escape(out_user)} }} -timestamp {{ {tcl_escape(timestamp)} }}"
        )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def generate_waiver(csv_path: Path, output_path: Path, user: str = "") -> None:
    generate_waiver_from_rows(read_csv(csv_path), output_path, user)


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


def update_review_from_report_excel(args: argparse.Namespace, previous_excel: Path | None = None) -> list[dict[str, str]]:
    rows = collect_current_rows_from_report_excel(args.report_excel)
    if previous_excel and previous_excel.exists():
        rows = merge_rows(read_review_workbook(previous_excel), rows)
    export_review_workbook(rows, args.excel, args.summary, report_excel=args.report_excel, previous_excel=previous_excel)
    if args.waiver_tcl and args.waiver_audit:
        audit_waiver_rules(args.waiver_tcl, rows, args.waiver_audit)
    print(f"review rows : {len(rows)}")
    print(f"report excel: {args.report_excel}")
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


def cmd_update_review_excel(args: argparse.Namespace) -> None:
    update_review_from_report_excel(args, previous_excel=args.excel if args.excel.exists() else None)


def cmd_prepare_waiver(args: argparse.Namespace) -> None:
    rows = read_review_workbook(args.excel)
    generate_waiver_from_rows(rows, args.output, args.user)
    print(f"read {len(rows)} rows from {args.excel}")
    print(f"wrote {args.output}")
    print("Next: run the sanity tool to refresh reports/report_lint.full.log.")


def cmd_merge_report(args: argparse.Namespace) -> None:
    run_make_excel(Path(__file__).resolve().parent)
    update_review_from_report_excel(args, previous_excel=args.excel if args.excel.exists() else None)


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


def cmd_clean(args: argparse.Namespace) -> None:
    removed = clean_generated_files(Path(__file__).resolve().parent, keep_waiver=args.keep_waiver, dry_run=args.dry_run)
    action = "would remove" if args.dry_run else "removed"
    if not removed:
        print("no generated files found")
        return
    for path in removed:
        print(f"{action}: {path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Manage LINT sanity review with lint_review.xlsx and waiver Tcl.",
        epilog="Use prepare-waiver first, run the sanity tool, then use merge-report.",
    )
    sub = parser.add_subparsers()

    p = sub.add_parser("run_all", help="Print the two-command review sequence without changing files.")
    p.set_defaults(func=lambda _args: run_all(Path(__file__).resolve().parent))

    p = sub.add_parser("prepare-waiver", help="Import lint_review.xlsx and generate vc_waiver.tcl only.")
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--output", type=Path, default=Path("vc_waiver.tcl"))
    p.add_argument("--user", default="")
    p.set_defaults(func=cmd_prepare_waiver)

    p = sub.add_parser("merge-report", help="Run make -f Makefile excel, then merge report_lint.full.xlsx into lint_review.xlsx.")
    p.add_argument("--report-excel", type=Path, default=Path("report_lint.full.xlsx"))
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--summary", type=Path, default=Path("outputs/lint_summary.csv"))
    p.add_argument("--waiver-tcl", type=Path, default=Path("vc_waiver.tcl"))
    p.add_argument("--waiver-audit", type=Path, default=Path("outputs/waiver_rule_audit.csv"))
    p.set_defaults(func=cmd_merge_report)

    p = sub.add_parser("update-review-excel", help="Merge report_lint.full.xlsx into lint_review.xlsx.")
    p.add_argument("--report-excel", type=Path, default=Path("report_lint.full.xlsx"))
    p.add_argument("--excel", type=Path, default=Path("outputs/lint_review.xlsx"))
    p.add_argument("--summary", type=Path, default=Path("outputs/lint_summary.csv"))
    p.add_argument("--waiver-tcl", type=Path, default=Path("vc_waiver.tcl"))
    p.add_argument("--waiver-audit", type=Path, default=Path("outputs/waiver_rule_audit.csv"))
    p.set_defaults(func=cmd_update_review_excel)

    p = sub.add_parser("clean", help="Delete files generated by sanity_lint_review.py.")
    p.add_argument("--dry-run", action="store_true", help="Print files that would be deleted without deleting them.")
    p.add_argument("--keep-waiver", action="store_true", help="Keep vc_waiver.tcl.")
    p.set_defaults(func=cmd_clean)
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
