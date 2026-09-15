#!/usr/bin/env python3
"""Excel review and waiver management for GCA constraint-analysis reports."""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import os
import re
import zipfile
from collections import Counter, OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


BASE_DIR = Path(__file__).resolve().parent
REVIEW_COLUMNS = [
    "No.", "IP Owner", "Owner Action", "Owner Comment",
    "Filter Mode", "Filter Fields", "Custom Filter",
    "Reviewer", "Reviewer Decision", "Reviewer Comment",
]
MANAGEMENT_COLUMNS = [
    "issue_id", "record_status", "waiver_name", "waiver_user",
    "waiver_timestamp", "source_report", "auto_condition", "condition_json",
]
REPORT_COLUMNS = [
    "Rule", "Severity", "Design", "Scenario", "Issue Index",
    "Waived In Report", "Object", "Message", "Rule Description",
]
RECORD_STATUSES = {"NEW", "CHANGED", "UNCHANGED", "REMOVED"}
DEFAULT_USER = "sanity_gca_review"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def xlsx_col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xlsx_col_index(reference: str) -> int:
    match = re.match(r"([A-Z]+)", reference)
    if not match:
        return 0
    index = 0
    for character in match.group(1):
        index = index * 26 + ord(character) - 64
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


def worksheet_xml(rows: list[list[object]], editable_headers: set[str] | None = None) -> str:
    rows = rows or [[""]]
    editable_headers = editable_headers or set()
    header = ["" if value is None else str(value) for value in rows[0]]
    managed_indexes = {index for index, column in enumerate(header) if column in MANAGEMENT_COLUMNS}
    column_count = max(len(row) for row in rows)
    row_count = len(rows)
    output = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        "<cols>",
    ]
    for column in range(1, column_count + 1):
        width = 64 if column in {4, 7, 18, 19} else 18
        output.append(f'<col min="{column}" max="{column}" width="{width}" customWidth="1"/>')
    output.extend(["</cols>", "<sheetData>"])
    for row_index, row in enumerate(rows, 1):
        output.append(f'<row r="{row_index}">')
        for column_index, value in enumerate(row):
            text = "" if value is None else str(value)
            style = ""
            if row_index == 1:
                style_id = 3 if column_index in managed_indexes else (1 if text in editable_headers else 2)
                style = f' s="{style_id}"'
            reference = f"{xlsx_col_name(column_index)}{row_index}"
            output.append(
                f'<c r="{reference}"{style} t="inlineStr"><is><t xml:space="preserve">'
                f'{html.escape(text)}</t></is></c>'
            )
        output.append("</row>")
    output.extend([
        "</sheetData>",
        f'<autoFilter ref="A1:{xlsx_col_name(column_count - 1)}{row_count}"/>',
    ])
    validations = []
    for column, choices in {
        "Owner Action": "UNREVIEWED,FIXED,WAIVED",
        "Filter Mode": "AUTO,FIELDS,CUSTOM",
        "Reviewer Decision": "PENDING,APPROVED,DISAPPROVED",
    }.items():
        if column in header and row_count > 1:
            letter = xlsx_col_name(header.index(column))
            validations.append(
                f'<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="{letter}2:{letter}{row_count}">'
                f'<formula1>"{choices}"</formula1></dataValidation>'
            )
    if validations:
        output.append(f'<dataValidations count="{len(validations)}">{"".join(validations)}</dataValidations>')
    output.append("</worksheet>")
    return "".join(output)


def write_xlsx(path: Path, sheets: dict[str, list[list[object]]], editable_headers: dict[str, set[str]] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    editable_headers = editable_headers or {}
    items = list(sheets.items())
    used_names: set[str] = set()
    safe_names = [safe_sheet_name(name, used_names) for name, _rows in items]
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    workbook = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>',
    ]
    relationships = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Relationships xmlns="{NS_PACKAGE_REL}">',
    ]
    for index, ((original_name, _rows), safe_name) in enumerate(zip(items, safe_names), 1):
        del original_name
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        workbook.append(f'<sheet name="{html.escape(safe_name)}" sheetId="{index}" r:id="rId{index}"/>')
        relationships.append(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    content_types.append("</Types>")
    workbook.append("</sheets></workbook>")
    relationships.append(
        f'<Relationship Id="rId{len(items) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    relationships.append("</Relationships>")
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>'
        '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>'
    )
    root_relationships = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_PACKAGE_REL}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "".join(content_types))
        archive.writestr("_rels/.rels", root_relationships)
        archive.writestr("xl/workbook.xml", "".join(workbook))
        archive.writestr("xl/_rels/workbook.xml.rels", "".join(relationships))
        archive.writestr("xl/styles.xml", styles)
        for index, (name, rows) in enumerate(items, 1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows, editable_headers.get(name, set())))


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.text or "" for node in item.findall(f".//{{{NS_MAIN}}}t")) for item in root.findall(f"{{{NS_MAIN}}}si")]


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    if cell.attrib.get("t") == "inlineStr":
        node = cell.find(f".//{{{NS_MAIN}}}t")
        return node.text if node is not None and node.text is not None else ""
    node = cell.find(f"{{{NS_MAIN}}}v")
    if node is None or node.text is None:
        return ""
    if cell.attrib.get("t") == "s":
        try:
            return shared_strings[int(node.text)]
        except (ValueError, IndexError):
            return ""
    return node.text


def read_xlsx_sheets(path: Path) -> dict[str, list[dict[str, str]]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_map = {item.attrib["Id"]: item.attrib["Target"].lstrip("/") for item in relationships}
        sheets: dict[str, list[dict[str, str]]] = {}
        for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
            target = relationship_map[sheet.attrib[f"{{{NS_REL}}}id"]]
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            root = ET.fromstring(archive.read(target))
            matrix: list[list[str]] = []
            for row_node in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
                values: list[str] = []
                for cell in row_node.findall(f"{{{NS_MAIN}}}c"):
                    column = xlsx_col_index(cell.attrib.get("r", ""))
                    while len(values) <= column:
                        values.append("")
                    values[column] = cell_text(cell, shared_strings)
                matrix.append(values)
            if not matrix:
                sheets[sheet.attrib["name"]] = []
                continue
            header = matrix[0]
            sheets[sheet.attrib["name"]] = [
                {header[index]: row[index] if index < len(row) else "" for index in range(len(header)) if header[index]}
                for row in matrix[1:] if any(row)
            ]
        return sheets


def norm(value: str | None) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def tcl_quote(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"').replace("$", "\\$") + '"'


def tokenize_tcl_command(line: str) -> list[str]:
    """Split one Tcl command while keeping nested bracket/brace words intact."""
    words: list[str] = []
    start: int | None = None
    braces = brackets = 0
    quoted = escaped = False
    for index, char in enumerate(line):
        if start is None:
            if char.isspace():
                continue
            start = index
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
        elif char == '"' and braces == 0:
            quoted = not quoted
        elif not quoted:
            if char == "{":
                braces += 1
            elif char == "}" and braces:
                braces -= 1
            elif char == "[":
                brackets += 1
            elif char == "]" and brackets:
                brackets -= 1
        if start is not None and braces == 0 and brackets == 0 and not quoted:
            next_char = line[index + 1] if index + 1 < len(line) else " "
            if next_char.isspace():
                words.append(line[start:index + 1])
                start = None
    if start is not None:
        words.append(line[start:])
    return words


def unquote_tcl(word: str) -> str:
    if len(word) >= 2 and ((word[0] == word[-1] == '"') or (word[0] == "{" and word[-1] == "}")):
        word = word[1:-1]
    return word.replace('\\"', '"').replace("\\$", "$").replace("\\\\", "\\")


def parse_waiver_tcl(path: Path) -> list[dict[str, object]]:
    rules: list[dict[str, object]] = []
    if not path.exists():
        return rules
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line.startswith("create_waiver "):
            continue
        words = tokenize_tcl_command(line)
        record: dict[str, object] = {"conditions": []}
        index = 1
        while index < len(words):
            option = words[index]
            if option.startswith("-") and index + 1 < len(words):
                value = words[index + 1]
                if option == "-condition":
                    record["conditions"].append(value)
                else:
                    record[option[1:]] = unquote_tcl(value)
                index += 2
            else:
                index += 1
        if record.get("rule"):
            rules.append(record)
    return rules


def parse_rule_catalog(path: Path) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    if not path.exists():
        return result
    pattern = re.compile(r"^([A-Z][A-Z0-9_]+_\d{4})\s+(error|warning|info)\s+(enabled|disabled)\s+(.*)$", re.I)
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        match = pattern.match(line)
        if match:
            result[match.group(1)] = {
                "severity": match.group(2).lower(),
                "rule_status": match.group(3).lower(),
                "description": norm(match.group(4)),
            }
    return result


def extract_object(message: str) -> str:
    quoted = re.search(r"'([^']+)'", message)
    return quoted.group(1) if quoted else ""


def condition_targets(conditions: list[str]) -> set[str]:
    targets: set[str] = set()
    for condition in conditions:
        for body in re.findall(r"\{([^{}]+)\}", condition):
            targets.update(body.split())
    return targets


def issue_id(issue: dict[str, str]) -> str:
    payload = "|".join(norm(issue.get(key)) for key in ("design", "scenario", "rule", "message"))
    return f"{issue.get('rule', 'GCA')}_{hashlib.sha1(payload.encode('utf-8')).hexdigest()[:12]}"


def parse_gca_report(report_path: Path, rule_path: Path, waiver_path: Path | None = None) -> list[dict[str, str]]:
    catalog = parse_rule_catalog(rule_path)
    design = ""
    scenario = "global"
    severity = rule = rule_description = ""
    issues: list[dict[str, str]] = []
    rule_line = re.compile(r"^\s{4}([A-Z][A-Z0-9_]+_\d{4})\s+(\d+)\s+(\d+)\s+(.*)$")
    detail_line = re.compile(r"^\s{6}(\d+) of (\d+)\s+(\d+)\s+(.*)$")
    severity_line = re.compile(r"^\s{2}(error|warning|info)\s+\d+\s+\d+", re.I)
    scenario_line = re.compile(r"^(\S+)\s+\d+\s+\d+\s+.*scenario violations", re.I)
    in_violations = False
    for raw in report_path.read_text(encoding="utf-8", errors="replace").splitlines():
        if raw.startswith("Scenario Violations"):
            in_violations = True
            continue
        if in_violations and raw.startswith("User Messages"):
            break
        if not in_violations:
            continue
        if raw.startswith("Design:"):
            design = norm(raw.split(":", 1)[1])
            continue
        if raw.startswith("<Global Violations>"):
            scenario = "global"
            continue
        scenario_match = scenario_line.match(raw)
        if scenario_match:
            scenario = scenario_match.group(1)
            continue
        severity_match = severity_line.match(raw)
        if severity_match:
            severity = severity_match.group(1).lower()
            continue
        current_rule = rule_line.match(raw)
        if current_rule:
            rule = current_rule.group(1)
            rule_description = norm(current_rule.group(4))
            continue
        detail = detail_line.match(raw)
        if detail and rule:
            message = norm(detail.group(4))
            info = catalog.get(rule, {})
            issue = {
                "design": design,
                "scenario": scenario,
                "rule": rule,
                "severity": severity or info.get("severity", ""),
                "issue_index": detail.group(1),
                "waived_in_report": "yes" if int(detail.group(3)) else "no",
                "object": extract_object(message),
                "message": message,
                "rule_description": rule_description or info.get("description", ""),
                "source_report": report_path.name,
            }
            issue["issue_id"] = issue_id(issue)
            issues.append(issue)

    waivers = parse_waiver_tcl(waiver_path) if waiver_path else []
    by_rule: dict[str, list[dict[str, object]]] = {}
    for waiver in waivers:
        by_rule.setdefault(str(waiver.get("rule", "")), []).append(waiver)
    matched_ids: set[int] = set()
    unmatched_by_rule: dict[str, list[dict[str, str]]] = {}
    for issue in issues:
        matched = None
        for waiver in by_rule.get(issue["rule"], []):
            targets = condition_targets(waiver.get("conditions", []))
            if issue["object"] and issue["object"] in targets:
                matched = waiver
                break
        if matched is None:
            unmatched_by_rule.setdefault(issue["rule"], []).append(issue)
        else:
            apply_waiver(issue, matched)
            matched_ids.add(id(matched))
    for rule_name, unmatched_issues in unmatched_by_rule.items():
        remaining = [w for w in by_rule.get(rule_name, []) if id(w) not in matched_ids]
        if len(remaining) == len(unmatched_issues):
            for issue, waiver in zip(unmatched_issues, remaining):
                apply_waiver(issue, waiver)
    return issues


def narrow_condition(condition: str, target: str) -> str:
    """Narrow a simple collection condition to one report object."""
    if not target or target not in condition_targets([condition]):
        return condition
    collection = re.search(r"\[(get_ports|get_pins|get_clocks|get_nets|get_cells)\s+\{([^{}]+)\}\]", condition)
    if not collection:
        return condition
    replacement = f"[{collection.group(1)} {{{target}}}]"
    return condition[:collection.start()] + replacement + condition[collection.end():]


def apply_waiver(issue: dict[str, str], waiver: dict[str, object]) -> None:
    conditions = [str(value) for value in waiver.get("conditions", [])]
    if len(conditions) == 1:
        conditions = [narrow_condition(conditions[0], issue.get("object", ""))]
    comment = str(waiver.get("comment", ""))
    metadata = re.match(r"waived by (\S+) on (.+)$", comment, re.I)
    issue.update({
        "waiver_name": str(waiver.get("name", "")),
        "owner_action": "WAIVED",
        "owner_comment": comment,
        "filter_mode": "AUTO",
        "auto_condition": " ".join(f"-condition {value}" for value in conditions),
        "condition_json": json.dumps(conditions, ensure_ascii=False),
        "waiver_user": metadata.group(1) if metadata else "",
        "waiver_timestamp": metadata.group(2) if metadata else "",
    })


def default_row(issue: dict[str, str]) -> dict[str, str]:
    row = dict(issue)
    row.setdefault("record_status", "NEW")
    row.setdefault("owner_action", "UNREVIEWED")
    row.setdefault("owner_comment", "")
    row.setdefault("filter_mode", "AUTO")
    row.setdefault("filter_fields", "")
    row.setdefault("custom_filter", "")
    row.setdefault("ip_owner", "")
    row.setdefault("reviewer", "")
    row.setdefault("reviewer_decision", "PENDING")
    row.setdefault("reviewer_comment", "")
    row.setdefault("waiver_name", "")
    row.setdefault("waiver_user", "")
    row.setdefault("waiver_timestamp", "")
    row.setdefault("auto_condition", infer_condition(row))
    row.setdefault("condition_json", "[]")
    return row


def infer_condition(row: dict[str, str]) -> str:
    message = row.get("message", "")
    obj = row.get("object", "")
    if not obj:
        return ""
    lowered = message.lower()
    if "port" in lowered:
        return f"-condition [list port [get_ports {{{obj}}}]]"
    if "pin" in lowered:
        return f"-condition [list pin [get_pins {{{obj}}}]]"
    if "clock" in lowered:
        return f"-condition [list clock [get_clocks {{{obj}}}]]"
    if "net" in lowered:
        return f"-condition [list net [get_nets {{{obj}}}]]"
    return ""


def preserve_review(row: dict[str, str], old: dict[str, str]) -> None:
    for key in (
        "owner_action", "owner_comment", "filter_mode", "filter_fields", "custom_filter",
        "ip_owner", "reviewer", "reviewer_decision", "reviewer_comment",
        "waiver_name", "waiver_user", "waiver_timestamp", "auto_condition", "condition_json",
    ):
        if old.get(key):
            row[key] = old[key]


def merge_rows(old_rows: list[dict[str, str]], current_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    old_by_id = {row.get("issue_id", ""): row for row in old_rows}
    old_by_similar: dict[tuple[str, str, str, str], list[dict[str, str]]] = {}
    for old in old_rows:
        key = (old.get("design", ""), old.get("scenario", ""), old.get("rule", ""), old.get("object", ""))
        old_by_similar.setdefault(key, []).append(old)
    consumed: set[str] = set()
    merged: list[dict[str, str]] = []
    for current in current_rows:
        row = default_row(current)
        old = old_by_id.get(row["issue_id"])
        if old:
            preserve_review(row, old)
            row["record_status"] = "UNCHANGED"
            consumed.add(old.get("issue_id", ""))
        else:
            key = (row.get("design", ""), row.get("scenario", ""), row.get("rule", ""), row.get("object", ""))
            similar = next((item for item in old_by_similar.get(key, []) if item.get("issue_id", "") not in consumed), None)
            if similar:
                preserve_review(row, similar)
                row["record_status"] = "CHANGED"
                consumed.add(similar.get("issue_id", ""))
        merged.append(row)
    current_ids = {row["issue_id"] for row in current_rows}
    for old in old_rows:
        if old.get("issue_id") not in current_ids and old.get("issue_id") not in consumed:
            removed = dict(old)
            removed["record_status"] = "REMOVED"
            merged.append(removed)
    order = {"NEW": 0, "CHANGED": 1, "UNCHANGED": 2, "REMOVED": 3}
    return sorted(merged, key=lambda row: (order.get(row.get("record_status", ""), 9), row.get("rule", ""), int(row.get("issue_index", "0") or 0)))


def workbook_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    sheets = read_xlsx_sheets(path)
    source = sheets.get("ReviewDB", [])
    rows: list[dict[str, str]] = []
    for item in source:
        row = {
            "issue_id": item.get("issue_id", ""), "record_status": item.get("record_status", ""),
            "waiver_name": item.get("waiver_name", ""), "waiver_user": item.get("waiver_user", ""),
            "waiver_timestamp": item.get("waiver_timestamp", ""), "source_report": item.get("source_report", ""),
            "auto_condition": item.get("auto_condition", ""), "condition_json": item.get("condition_json", "[]"),
            "ip_owner": item.get("IP Owner", ""), "owner_action": item.get("Owner Action", "UNREVIEWED"),
            "owner_comment": item.get("Owner Comment", ""), "filter_mode": item.get("Filter Mode", "AUTO"),
            "filter_fields": item.get("Filter Fields", ""), "custom_filter": item.get("Custom Filter", ""),
            "reviewer": item.get("Reviewer", ""), "reviewer_decision": item.get("Reviewer Decision", "PENDING"),
            "reviewer_comment": item.get("Reviewer Comment", ""), "rule": item.get("Rule", ""),
            "severity": item.get("Severity", ""), "design": item.get("Design", ""), "scenario": item.get("Scenario", ""),
            "issue_index": item.get("Issue Index", ""), "waived_in_report": item.get("Waived In Report", ""),
            "object": item.get("Object", ""), "message": item.get("Message", ""),
            "rule_description": item.get("Rule Description", ""),
        }
        rows.append(row)
    return rows


def row_values(row: dict[str, str], number: int) -> list[str]:
    values = {
        "No.": str(number), "IP Owner": row.get("ip_owner", ""), "Owner Action": row.get("owner_action", ""),
        "Owner Comment": row.get("owner_comment", ""), "Filter Mode": row.get("filter_mode", "AUTO"),
        "Filter Fields": row.get("filter_fields", ""), "Custom Filter": row.get("custom_filter", ""),
        "Reviewer": row.get("reviewer", ""), "Reviewer Decision": row.get("reviewer_decision", ""),
        "Reviewer Comment": row.get("reviewer_comment", ""), "issue_id": row.get("issue_id", ""),
        "record_status": row.get("record_status", ""), "waiver_name": row.get("waiver_name", ""),
        "waiver_user": row.get("waiver_user", ""), "waiver_timestamp": row.get("waiver_timestamp", ""),
        "source_report": row.get("source_report", ""), "auto_condition": row.get("auto_condition", ""),
        "condition_json": row.get("condition_json", "[]"), "Rule": row.get("rule", ""),
        "Severity": row.get("severity", ""), "Design": row.get("design", ""), "Scenario": row.get("scenario", ""),
        "Issue Index": row.get("issue_index", ""), "Waived In Report": row.get("waived_in_report", ""),
        "Object": row.get("object", ""), "Message": row.get("message", ""),
        "Rule Description": row.get("rule_description", ""),
    }
    columns = REVIEW_COLUMNS + MANAGEMENT_COLUMNS + REPORT_COLUMNS
    return [values.get(column, "") for column in columns]


def export_workbook(rows: list[dict[str, str]], path: Path) -> None:
    columns = REVIEW_COLUMNS + MANAGEMENT_COLUMNS + REPORT_COLUMNS
    sheets: OrderedDict[str, list[list[str]]] = OrderedDict()
    counts = Counter((row.get("record_status", ""), row.get("owner_action", ""), row.get("rule", "")) for row in rows)
    sheets["Summary"] = [["record_status", "owner_action", "Rule", "Count"]] + [
        [status, action, rule, str(count)] for (status, action, rule), count in sorted(counts.items())
    ]
    sheets["Instructions"] = [
        ["GCA Sanity Review Workbook"],
        ["Edit green columns only. GCA_confirmed.tcl is generated from Owner Action = WAIVED."],
        ["AUTO preserves or infers native GCA -condition clauses."],
        ["FIELDS accepts comma-separated condition keys, or key=Tcl-expression overrides."],
        ["CUSTOM accepts complete one-or-more '-condition [list ...]' clauses."],
        ["Reviewer Decision records peer review and does not control waiver generation."],
    ]
    active = [row for row in rows if row.get("record_status") != "REMOVED"]
    for rule in sorted({row.get("rule", "UNKNOWN") or "UNKNOWN" for row in active}):
        rule_rows = [row for row in active if (row.get("rule", "UNKNOWN") or "UNKNOWN") == rule]
        sheets[rule[:31]] = [columns] + [row_values(row, index) for index, row in enumerate(rule_rows, 1)]
    removed = [row for row in rows if row.get("record_status") == "REMOVED"]
    if removed:
        sheets["Removed"] = [columns] + [row_values(row, index) for index, row in enumerate(removed, 1)]
    sheets["ReviewDB"] = [columns] + [row_values(row, index) for index, row in enumerate(rows, 1)]
    editable = {name: set(REVIEW_COLUMNS[1:]) for name in sheets if name not in {"Summary", "Instructions", "ReviewDB"}}
    write_xlsx(path, sheets, editable)


def field_conditions(row: dict[str, str]) -> str:
    try:
        originals = json.loads(row.get("condition_json", "[]") or "[]")
    except json.JSONDecodeError:
        originals = []
    by_key: dict[str, str] = {}
    for condition in originals:
        match = re.match(r"\[list\s+(\S+)\s+(.+)\]$", condition)
        if match:
            by_key[match.group(1)] = match.group(2)
    output: list[str] = []
    for item in next(csv.reader([row.get("filter_fields", "")], skipinitialspace=True), []):
        item = item.strip()
        if not item:
            continue
        if "=" in item:
            key, expression = (part.strip() for part in item.split("=", 1))
        else:
            key, expression = item, by_key.get(item, "")
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key) or not expression:
            raise ValueError(f"invalid or unavailable GCA condition field: {item!r}")
        output.append(f"-condition [list {key} {expression}]")
    if not output:
        raise ValueError("Filter Fields must produce at least one GCA condition")
    return " ".join(output)


def waiver_condition(row: dict[str, str]) -> str:
    mode = norm(row.get("filter_mode") or "AUTO").upper()
    if mode == "AUTO":
        condition = row.get("auto_condition", "").strip() or infer_condition(row)
    elif mode in {"FIELD", "FIELDS"}:
        condition = field_conditions(row)
    elif mode == "CUSTOM":
        condition = row.get("custom_filter", "").strip()
    else:
        raise ValueError(f"unsupported Filter Mode: {mode!r}")
    if not condition or not condition.startswith("-condition "):
        raise ValueError("GCA waiver requires one or more -condition clauses")
    return condition


def generate_waiver(rows: list[dict[str, str]], path: Path, user: str = "") -> None:
    eligible = [row for row in rows if row.get("record_status") != "REMOVED" and norm(row.get("owner_action")).upper() == "WAIVED"]
    groups: OrderedDict[tuple[str, str, str, str, str], list[dict[str, str]]] = OrderedDict()
    for row in eligible:
        condition = waiver_condition(row)
        key = (row.get("design", ""), row.get("scenario", ""), row.get("rule", ""), row.get("owner_comment", ""), condition)
        groups.setdefault(key, []).append(row)
    lines = ["# Generated by sanity_gca_review.py", "# Source: outputs/gca_review.xlsx", ""]
    used_names: set[str] = set()
    timestamp = datetime.now().strftime("%a %b %d %H:%M:%S %Y")
    for number, ((design, scenario, rule, comment, condition), shared) in enumerate(groups.items(), 1):
        base = shared[0].get("waiver_name", "").strip() or f"excel_{rule}_{number}"
        name = base
        suffix = 2
        while name in used_names:
            name = f"{base}_{suffix}"
            suffix += 1
        used_names.add(name)
        owner = norm(user) or norm(shared[0].get("waiver_user")) or norm(os.getenv("USERNAME")) or DEFAULT_USER
        saved_time = norm(shared[0].get("waiver_timestamp")) or timestamp
        out_comment = comment or f"waived by {owner} on {saved_time}"
        parts = ["create_waiver", "-name", name, "-design", design]
        if scenario and scenario.lower() != "global":
            parts += ["-scenario", scenario]
        parts += ["-rule", rule, "-comment", tcl_quote(out_comment), condition]
        lines.append(" ".join(parts))
        for row in shared:
            row["waiver_name"] = name
            row["waiver_user"] = owner
            row["waiver_timestamp"] = saved_time
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def command_merge(args: argparse.Namespace) -> None:
    current = [default_row(row) for row in parse_gca_report(args.report, args.rules, args.waiver)]
    old = workbook_rows(args.excel)
    rows = merge_rows(old, current) if old else current
    export_workbook(rows, args.excel)
    print(f"review rows: {len(rows)}")
    print(f"excel      : {args.excel}")


def command_generate(args: argparse.Namespace) -> None:
    rows = workbook_rows(args.excel)
    generate_waiver(rows, args.output, args.user)
    export_workbook(rows, args.excel)
    print(f"read {len(rows)} rows from {args.excel}")
    print(f"wrote {args.output}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage GCA sanity review with Excel and GCA_confirmed.tcl")
    sub = parser.add_subparsers(required=True)
    merge = sub.add_parser("merge_excel", help="Parse the current GCA report and merge it into the review workbook")
    merge.add_argument("--report", type=Path, default=Path("reports/BOS_AESDMA_gca_const_analysis.rpt"))
    merge.add_argument("--rules", type=Path, default=Path("reports/BOS_AESDMA_gca.rpt"))
    merge.add_argument("--waiver", type=Path, default=Path("GCA_confirmed.tcl"))
    merge.add_argument("--excel", type=Path, default=Path("outputs/gca_review.xlsx"))
    merge.set_defaults(func=command_merge)
    generate = sub.add_parser("gen_waiver", help="Generate GCA_confirmed.tcl from the review workbook")
    generate.add_argument("--excel", type=Path, default=Path("outputs/gca_review.xlsx"))
    generate.add_argument("--output", type=Path, default=Path("GCA_confirmed.tcl"))
    generate.add_argument("--user", default="")
    generate.set_defaults(func=command_generate)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
