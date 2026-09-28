import unittest
import argparse
import hashlib
import tempfile
import tkinter
import zipfile
import json
from unittest.mock import patch
from pathlib import Path

from sanity_lint_review import (
    AutoFilterIndex,
    ColumnWidthLayout,
    FormulaCell,
    WORKBOOK_MANAGEMENT_COLUMNS,
    filter_expr,
    format_filter_fields_spec,
    generate_waiver_from_rows,
    collect_user_extra_columns,
    cmd_gen_waiver,
    export_review_workbook,
    collect_current_rows,
    merge_rows,
    normalize_record_status,
    parse_filter_fields_spec,
    parse_filter,
    parse_waiver_tcl,
    read_review_workbook,
    read_workbook_column_layouts,
    review_workbook_has_unexported_changes,
    report_ordered_sheet_columns,
    sheet_xml,
    sheet_tab_color,
    summary_formula_matrix,
    stamp_waiver_workbook_hash,
    tcl_braced_filter,
    tcl_double_quote,
    tree_summary_matrix,
    update_review_from_reports,
    waiver_filter_expression,
    write_xlsx,
)


class WaiverFilterTests(unittest.TestCase):
    def test_qualified_fields_round_trip(self):
        expression = '(Goal == "BOS_LINT_RULE") AND (PropertyList:LintPropertyName == "Property_142")'
        fields = parse_filter(expression)
        self.assertEqual(list(fields), ["Goal", "PropertyList:LintPropertyName"])
        self.assertEqual(fields["PropertyList:LintPropertyName"], "Property_142")
        self.assertEqual(filter_expr(fields), expression)

    def test_legacy_saved_filter_export(self):
        self.assertEqual(
            filter_expr({"LintPropertyName": "Property_142"}),
            '(PropertyList:LintPropertyName == "Property_142")',
        )

    def test_braces_in_filter_values_keep_the_original_statement(self):
        fields = {"Statement": "r_rdata <= {truncated ..."}

        expression = filter_expr(fields)

        self.assertIn("{truncated", expression)
        self.assertEqual(parse_filter(expression), {"Statement": "*r_rdata <= {truncated ...*"})

    def test_tcl_double_quote_preserves_filter_value(self):
        expression = '(Statement == "r_rdata <= {data[3:0], $value ...")'
        interpreter = tkinter.Tcl()
        interpreter.eval('proc capture {value} {set ::captured $value}')

        interpreter.eval(f'capture "{tcl_double_quote(expression)}"')

        self.assertEqual(interpreter.eval('set ::captured'), expression)

    def test_tcl_braced_filter_escapes_only_unmatched_braces(self):
        self.assertEqual(tcl_braced_filter('a {b} c'), 'a {b} c')
        self.assertEqual(tcl_braced_filter('a {b ...'), r'a \{b ...')
        self.assertEqual(tcl_braced_filter('a b} ...'), r'a b\} ...')

    def test_original_w551_filter(self):
        rules = parse_waiver_tcl(Path(__file__).with_name("vc_waiver.tcl_ori"))
        fields = rules["W551_837"]["filter_fields"]
        self.assertEqual(fields["PropertyList:LintPropertyName"], "Property_142")
        self.assertNotIn("LintPropertyName", fields)
        expected = fields.copy()
        expected["Statement"] = "*unique casez (r_st)*"
        self.assertEqual(parse_filter(filter_expr(fields)), expected)


class WorkbookColumnTests(unittest.TestCase):
    def test_rule_sheet_tab_color_comes_from_severity(self):
        for severity, expected in {
            "Fatal": "FFFF0000",
            "Error": "FFFF0000",
            "Warning": "FFFFA500",
            "Info": "FF00B050",
        }.items():
            rows = [["Tag", "Severity"], ["RULE", severity]]
            self.assertEqual(sheet_tab_color("RULE", rows), expected)
            self.assertIn(f'<tabColor rgb="{expected}"/>', sheet_xml(rows, tab_color=expected))

    def test_summary_is_uncolored_and_mixed_severity_uses_highest_level(self):
        self.assertIsNone(sheet_tab_color("Tree Summary", [["Severity"], ["Warning"]]))
        self.assertEqual(
            sheet_tab_color("RULE", [["Severity"], ["Warning"], ["Error"]]),
            "FFFF0000",
        )

    def test_rule_sheet_tab_color_can_come_from_tree_summary(self):
        summary = [["Severity", "Tag"], ["Warning", "W551"]]
        self.assertEqual(sheet_tab_color("W551", [["Tag"], ["W551"]], summary), "FFFFA500")

    def test_workbook_writer_applies_tree_summary_tab_color(self):
        sheets = {
            "Summary": [["Severity", "Tag"], ["Warning", "W551"]],
            "W551": [["Tag"], ["W551"]],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            write_xlsx(path, sheets)
            with zipfile.ZipFile(path) as workbook:
                rule_sheet = workbook.read("xl/worksheets/sheet2.xml").decode("utf-8")
                workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        self.assertIn('<sheetPr><tabColor rgb="FFFFA500"/></sheetPr>', rule_sheet)
        self.assertIn('calcMode="auto"', workbook_xml)
        self.assertIn('fullCalcOnLoad="1"', workbook_xml)

    def test_summary_is_sorted_and_colorized_by_severity(self):
        rows = [
            {"severity": "info", "tag": "INFO_RULE", "record_status": "NEW", "owner_action": "UNREVIEWED", "reviewer_decision": ""},
            {"severity": "warning", "tag": "WARN_RULE", "record_status": "NEW", "owner_action": "UNREVIEWED", "reviewer_decision": "PENDING"},
            {"severity": "error", "tag": "ERROR_RULE", "record_status": "NEW", "owner_action": "FIXED", "reviewer_decision": "DISAPPROVED"},
            {"severity": "fatal", "tag": "FATAL_RULE", "record_status": "NEW", "owner_action": "WAIVED", "reviewer_decision": "APPROVED"},
        ]
        matrix = tree_summary_matrix(rows)
        self.assertEqual(
            matrix[0],
            ["Severity", "Tag", "Count", "Waived", "Unreviewed", "Confirmed", "Pending by reviewer"],
        )
        self.assertEqual([row[0] for row in matrix[1:-1]], ["fatal", "error", "warning", "info"])
        self.assertEqual(matrix[1][2:], ["1", "1", "0", "1", "0"])
        self.assertEqual(matrix[2][2:], ["1", "0", "0", "1", "0"])
        self.assertEqual(matrix[3][2:], ["1", "0", "1", "0", "1"])
        self.assertEqual(matrix[4][2:], ["1", "0", "1", "0", "1"])
        self.assertEqual(matrix[-1], ["TOTAL", "", "4", "1", "2", "2", "2"])
        xml = sheet_xml(matrix, colorize_severity_rows=True)
        self.assertEqual(xml.count(' s="4"'), 2)
        self.assertEqual(xml.count(' s="5"'), 1)
        self.assertEqual(xml.count(' s="6"'), 1)
        self.assertIn('<c r="B2" t="inlineStr">', xml)
        self.assertNotIn(' width="', xml)

    def test_editable_columns_have_no_script_assigned_widths(self):
        rows = [
            ["No.", "IP Owner", "Owner Action", "Owner Comment", "Filter Mode",
             "Filter Fields", "Custom Filter", "Reviewer", "Reviewer Decision",
             "Reviewer Comment", "issue_id", "Tag"],
            ["1", "owner", "WAIVED", "reason", "AUTO", "", "", "reviewer",
             "PENDING", "", "id", "W551"],
        ]
        editable = set(rows[0][1:10])
        xml = sheet_xml(rows, editable_headers=editable)
        self.assertNotIn(' width="', xml)
        self.assertIn('<col min="11" max="11" hidden="1"/>', xml)

    def test_summary_has_one_row_per_tag_and_uses_highest_severity(self):
        rows = [
            {"severity": "info", "tag": "MIXED", "record_status": "NEW", "owner_action": "WAIVED", "reviewer_decision": "APPROVED"},
            {"severity": "warning", "tag": "MIXED", "record_status": "NEW", "owner_action": "UNREVIEWED", "reviewer_decision": "PENDING"},
        ]
        matrix = tree_summary_matrix(rows)
        self.assertEqual(matrix[1], ["warning", "MIXED", "2", "1", "1", "1", "1"])

    def test_summary_formulas_link_to_rule_sheet(self):
        matrix = [
            ["Severity", "Tag", "Count", "Waived", "Unreviewed", "Confirmed", "Pending by reviewer"],
            ["warning", "W551", "2", "1", "1", "1", "1"],
            ["TOTAL", "", "2", "1", "1", "1", "1"],
        ]
        sheets = {
            "Summary": matrix,
            "W551": [
                ["No.", "Owner Action", "Reviewer Decision"],
                ["1", "WAIVED", "APPROVED"],
                ["2", "UNREVIEWED", "PENDING"],
            ],
        }
        linked = summary_formula_matrix(matrix, sheets)
        self.assertEqual(linked[1][2], FormulaCell("COUNTA('W551'!$A$2:$A$3)", "2"))
        self.assertEqual(linked[1][3], FormulaCell('COUNTIF(\'W551\'!$B$2:$B$3,"WAIVED")', "1"))
        self.assertEqual(linked[2][6], FormulaCell("SUM(G2:G2)", "1"))
        xml = sheet_xml(linked, colorize_severity_rows=True)
        self.assertIn("<f>COUNTA('W551'!$A$2:$A$3)</f><v>2</v>", xml)

    def test_script_managed_columns_are_grouped_after_review_fields(self):
        report_headers = {
            "W551": ["No.", "issue_id", "Comment", "record_status", "Tag"]
        }
        extras = {"W551": ["Reviewer Note", "waiver_name"]}

        columns = report_ordered_sheet_columns("W551", [], report_headers, extras)

        self.assertEqual(
            columns,
            [
                "No.", "IP Owner", "Owner Action", "Owner Comment",
                "Filter Mode", "Filter Fields", "Custom Filter", "Reviewer",
                "Reviewer Decision", "Reviewer Comment", *WORKBOOK_MANAGEMENT_COLUMNS,
                "Tag", "Reviewer Note",
            ],
        )
        self.assertNotIn("waiver_enabled", columns)
        self.assertNotIn("filter_json", columns)
        self.assertNotIn("fields_json", columns)

    def test_only_script_managed_headers_have_pale_yellow_style(self):
        rows = [
            ["Comment", *WORKBOOK_MANAGEMENT_COLUMNS, "Tag"],
            ["note", *(["value"] * len(WORKBOOK_MANAGEMENT_COLUMNS)), "W551"],
        ]

        xml = sheet_xml(rows)

        self.assertEqual(xml.count(' s="3"'), len(WORKBOOK_MANAGEMENT_COLUMNS))
        self.assertNotIn(' s="4"', xml)
        self.assertEqual(xml.count(' hidden="1"'), len(WORKBOOK_MANAGEMENT_COLUMNS) - 2)
        self.assertNotIn(' width="', xml)
        self.assertIn("Issue Status", WORKBOOK_MANAGEMENT_COLUMNS)
        self.assertIn("Time Stamp", WORKBOOK_MANAGEMENT_COLUMNS)
        self.assertNotIn("record_status", WORKBOOK_MANAGEMENT_COLUMNS)
        self.assertNotIn("waiver_timestamp", WORKBOOK_MANAGEMENT_COLUMNS)

    def test_issue_headers_are_gray_with_do_not_change_notes(self):
        rows = [["IP Owner", "Tag", "Goal", "Reviewer Note"], ["owner", "W551", "LINT", "check"]]
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            write_xlsx(
                path, {"W551": rows},
                {"W551": {"IP Owner", "Reviewer Note"}},
                report_owned_headers_by_sheet={"W551": {"Tag", "Goal"}},
            )
            with zipfile.ZipFile(path) as workbook:
                sheet = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
                comments = workbook.read("xl/comments1.xml").decode("utf-8")

        self.assertIn('<c r="A1" s="1"', sheet)
        self.assertIn('<c r="B1" s="2"', sheet)
        self.assertIn('<c r="C1" s="2"', sheet)
        self.assertIn('<c r="D1" s="1"', sheet)
        self.assertEqual(comments.count("Do not change"), 2)

    def test_report_fields_are_not_treated_as_user_extra_columns(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            write_xlsx(
                path,
                {"W551": [["Tag", "Goal", "Reviewer Note"], ["W551", "LINT", "check"]]},
            )
            extras, values = collect_user_extra_columns(
                path, {}, {"W551": {"Tag", "Goal"}}
            )

        self.assertEqual(extras["W551"], ["Reviewer Note"])
        self.assertEqual(next(iter(values.values())), {"Reviewer Note": "check"})

    def test_status_headers_have_excel_notes(self):
        rows = [
            ["Owner Action", "Reviewer Decision", "Issue Status", "source_report", "Tag"],
            ["UNREVIEWED", "PENDING", "NEW", "full", "W551"],
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            write_xlsx(path, {"W551": rows})

            with zipfile.ZipFile(path) as workbook:
                comments = workbook.read("xl/comments1.xml").decode("utf-8")
                sheet = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
                note_shapes = workbook.read("xl/drawings/commentsDrawing1.vml").decode("utf-8")

        self.assertIn("NEW: the issue appears for the first time", comments)
        self.assertIn("full: the issue comes from report_lint.full", comments)
        self.assertIn("WAIVED: the IP owner decided to waive", comments)
        self.assertIn("DISAPPROVED: the reviewer rejects", comments)
        self.assertIn('<legacyDrawing r:id="rId2"/>', sheet)
        self.assertIn("width:220pt;height:100pt", note_shapes)

    def test_issue_status_is_visible_and_imports_as_record_status(self):
        rows = [["Issue Status", "Time Stamp", "Tag", "Goal", "Module", "FileName", "LineNumber"],
                ["CHANGED", "01-01-2025 12:00:00", "W551", "LINT", "top", "top.sv", "12"]]
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            write_xlsx(path, {"W551": rows})
            with zipfile.ZipFile(path) as workbook:
                sheet = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
            imported = read_review_workbook(path)

        self.assertNotIn(' width="', sheet)
        self.assertEqual(imported[0]["record_status"], "CHANGED")
        self.assertEqual(imported[0]["waiver_timestamp"], "01-01-2025 12:00:00")
        self.assertNotIn("Issue Status", json.loads(imported[0]["fields_json"]))
        self.assertNotIn("Time Stamp", json.loads(imported[0]["fields_json"]))


class RecordStatusTests(unittest.TestCase):
    @staticmethod
    def row(issue_id, line="1", status="UNCHANGED", object_name=None):
        return {
            "issue_id": issue_id,
            "record_status": status,
            "owner_action": "UNREVIEWED",
            "ip_owner": "owner_a",
            "reviewer_decision": "PENDING",
            "tag": "W551",
            "goal": "LINT",
            "module": "top",
            "file": "top.sv",
            "line": line,
            "hierarchy": "top",
            "object": object_name or issue_id,
        }

    def test_timestamp_and_owner_signature_survive_workbook_round_trip(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "waiver_timestamp": "01-01-2025 12:00:00",
            "waiver_owner_signature": "saved-signature",
            "fields_json": '{"Tag":"W551","Goal":"LINT","Module":"top","FileName":"top.sv","LineNumber":"1"}',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            export_review_workbook([row], path)
            restored = read_review_workbook(path)[0]
            with zipfile.ZipFile(path) as workbook:
                sheet = workbook.read("xl/worksheets/sheet2.xml").decode("utf-8")

        self.assertEqual(restored["waiver_timestamp"], "01-01-2025 12:00:00")
        self.assertEqual(restored["waiver_owner_signature"], "saved-signature")
        self.assertIn("Time Stamp", sheet)
        self.assertIn("waiver_owner_signature", sheet)

    def test_export_preserves_manual_column_widths_by_sheet_and_header(self):
        row = self.row("same")
        row["fields_json"] = '{"Tag":"W551","Goal":"LINT","Module":"top"}'
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            previous = temp / "previous.xlsx"
            output = temp / "updated.xlsx"
            write_xlsx(
                previous,
                {
                    "Summary": [["Severity", "Tag"], ["warning", "W551"]],
                    "W551": [["record_status", "IP Owner", "Tag"], ["UNCHANGED", "owner_a", "W551"]],
                },
                column_layouts_by_sheet={
                    "Summary": ColumnWidthLayout(
                        {"Tag": {"width": "37.5", "customWidth": "1"}}, {},
                        {"defaultRowHeight": "15", "defaultColWidth": "12"},
                    ),
                    "W551": ColumnWidthLayout(
                        {
                            "record_status": {"width": "24.5", "customWidth": "1"},
                            "IP Owner": {"width": "19.25", "customWidth": "1"},
                        }, {}, {},
                    ),
                },
            )
            export_review_workbook([row], output, previous_excel=previous)
            layouts = read_workbook_column_layouts(output)

        self.assertEqual(layouts["Summary"].by_header["Tag"]["width"], "37.5")
        self.assertEqual(layouts["Summary"].sheet_format["defaultColWidth"], "12")
        self.assertEqual(layouts["W551"].by_header["Issue Status"]["width"], "24.5")
        self.assertEqual(layouts["W551"].by_header["IP Owner"]["width"], "19.25")

    def test_missing_workbook_has_no_unexported_changes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            self.assertFalse(
                review_workbook_has_unexported_changes(temp / "missing.xlsx", temp / "waiver.tcl")
            )

    def test_existing_workbook_without_waiver_is_unsafe(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "lint_review.xlsx"
            excel.touch()
            self.assertTrue(review_workbook_has_unexported_changes(excel, temp / "missing.tcl"))

    def test_changed_workbook_is_unsafe(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "lint_review.xlsx"
            waiver = temp / "vc_waiver.tcl"
            excel.write_bytes(b"before")
            waiver.write_text("# Generated by test\n", encoding="utf-8")
            stamp_waiver_workbook_hash(waiver, excel)
            excel.write_bytes(b"after")
            self.assertTrue(review_workbook_has_unexported_changes(excel, waiver))

    def test_matching_workbook_hash_is_safe(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "lint_review.xlsx"
            waiver = temp / "vc_waiver.tcl"
            excel.write_bytes(b"review workbook")
            waiver.write_text("# Generated by test\n", encoding="utf-8")
            stamp_waiver_workbook_hash(waiver, excel)
            self.assertFalse(review_workbook_has_unexported_changes(excel, waiver))

    def test_legacy_waiver_without_hash_is_unsafe(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "lint_review.xlsx"
            waiver = temp / "vc_waiver.tcl"
            excel.touch()
            waiver.write_text("# Generated by older version\n", encoding="utf-8")
            self.assertTrue(review_workbook_has_unexported_changes(excel, waiver))

    def test_gen_waiver_blocks_untracked_tcl_before_reading_workbook(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "review.xlsx"
            excel.write_bytes(b"unchanged workbook")
            waiver = temp / "vc_waiver.tcl"
            waiver.write_text("# manual rule\n", encoding="utf-8")
            args = argparse.Namespace(excel=excel, output=waiver, force=False, no_update_excel=True)
            with patch("sanity_lint_review.read_review_workbook") as read_workbook:
                with self.assertRaisesRegex(SystemExit, "no saved generation hash"):
                    cmd_gen_waiver(args)
            read_workbook.assert_not_called()
            self.assertEqual(waiver.read_text(encoding="utf-8"), "# manual rule\n")
            self.assertEqual(excel.read_bytes(), b"unchanged workbook")
            row = self.row("waive")
            row.update({
                "owner_action": "WAIVED", "owner_comment": "Reason",
                "filter_mode": "CUSTOM", "custom_filter": '(Module == "top")',
            })
            args.force = True
            with patch("sanity_lint_review.read_review_workbook", return_value=[row]):
                cmd_gen_waiver(args)
            backups = list(temp.glob("vc_waiver.tcl.bak.*"))
            self.assertEqual(len(backups), 1)
            self.assertEqual(backups[0].read_text(encoding="utf-8"), "# manual rule\n")
            self.assertTrue((temp / "vc_waiver.tcl.sha256").exists())

    def test_gen_waiver_force_backs_up_manual_edit_and_records_hash(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED", "owner_comment": "Reason",
            "filter_mode": "CUSTOM", "custom_filter": '(Module == "top")',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            excel = temp / "review.xlsx"
            excel.write_bytes(b"review workbook")
            waiver = temp / "vc_waiver.tcl"
            args = argparse.Namespace(excel=excel, output=waiver, force=False, no_update_excel=True)
            with patch("sanity_lint_review.read_review_workbook", return_value=[row]):
                cmd_gen_waiver(args)
                first_generated = waiver.read_bytes()
                self.assertEqual(
                    (temp / "vc_waiver.tcl.sha256").read_text(encoding="utf-8").strip(),
                    hashlib.sha256(first_generated).hexdigest(),
                )
                cmd_gen_waiver(args)
                self.assertEqual(waiver.read_bytes(), first_generated)
                waiver.write_bytes(first_generated + b"# manual edit\n")
                with self.assertRaisesRegex(SystemExit, "has changed since"):
                    cmd_gen_waiver(args)
                self.assertEqual(waiver.read_bytes(), first_generated + b"# manual edit\n")
                args.force = True
                cmd_gen_waiver(args)

            backups = list(temp.glob("vc_waiver.tcl.bak.*"))
            self.assertEqual(len(backups), 1)
            self.assertEqual(backups[0].read_bytes(), first_generated + b"# manual edit\n")
            self.assertNotIn(b"# manual edit", waiver.read_bytes())
            self.assertEqual(
                (temp / "vc_waiver.tcl.sha256").read_text(encoding="utf-8").strip(),
                hashlib.sha256(waiver.read_bytes()).hexdigest(),
            )

    def test_legacy_active_and_waived_statuses_become_unchanged(self):
        self.assertEqual(normalize_record_status("ACTIVE"), "UNCHANGED")
        self.assertEqual(normalize_record_status("WAIVED"), "UNCHANGED")

    def test_merge_uses_only_four_record_statuses(self):
        old_rows = [self.row("same"), self.row("changed", line="2"), self.row("removed", line="3")]
        current_rows = [
            self.row("same", status="WAIVED"),
            self.row("new"),
            self.row("replacement", line="2", object_name="changed"),
        ]

        merged = merge_rows(old_rows, current_rows)
        statuses = {row["issue_id"]: row["record_status"] for row in merged}

        self.assertEqual(statuses["same"], "UNCHANGED")
        self.assertEqual(statuses["replacement"], "CHANGED")
        self.assertEqual(statuses["new"], "NEW")
        self.assertEqual(statuses["removed"], "REMOVED")

    def test_update_from_reports_preserves_existing_auto_filter_mode(self):
        old = self.row("same")
        old.update({"filter_mode": "AUTO", "filter_fields": "", "custom_filter": ""})
        current = self.row("same")
        current.update({
            "filter_mode": "CUSTOM",
            "filter_fields": "",
            "custom_filter": '(Module == "top")',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            previous_excel = temp / "review.xlsx"
            previous_excel.touch()
            args = type("Args", (), {
                "full_report": temp / "full.log",
                "waived_report": temp / "waived.log",
                "waiver_tcl": None,
                "review_db": temp / "review.csv",
                "excel": temp / "output.xlsx",
                "summary": temp / "summary.csv",
                "waiver_audit": None,
            })()
            with (
                patch("sanity_lint_review.collect_current_rows", return_value=[current]),
                patch("sanity_lint_review.read_review_workbook", return_value=[old]),
                patch("sanity_lint_review.write_csv"),
                patch("sanity_lint_review.export_review_workbook") as export,
            ):
                merged = update_review_from_reports(args, previous_excel=previous_excel)

        self.assertEqual(merged[0]["filter_mode"], "AUTO")
        self.assertEqual(merged[0]["custom_filter"], "")
        self.assertEqual(export.call_args.args[0][0]["filter_mode"], "AUTO")
        self.assertEqual(export.call_args.kwargs["previous_excel"], previous_excel)

    def test_waiver_generation_depends_only_on_owner_action(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "reviewer_decision": "DISAPPROVED",
            "waiver_name": "W551_test",
            "owner_comment": "Owner waiver reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module =~ "top_*")',
            "fields_json": '{"Tag": "W551"}',
            "filter_json": "{}",
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")

        self.assertIn("W551_test", generated)
        self.assertIn('-filter {(Module =~ "top_*")}', generated)

    def test_waiver_generation_uses_ip_owner_and_supplies_timestamp(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
            "ip_owner": "owner_a",
            "waiver_user": "legacy_user",
            "waiver_timestamp": "N/A",
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")

        self.assertIn("-user { owner_a }", generated)
        self.assertRegex(
            generated,
            r"-timestamp \{ \d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2} \}",
        )
        self.assertEqual(row["waiver_user"], "owner_a")
        self.assertNotEqual(row["waiver_timestamp"], "N/A")

    def test_timestamp_changes_only_when_owner_changes_issue(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
            "waiver_timestamp": "01-01-2020 00:00:00",
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            first_timestamp = row["waiver_timestamp"]
            first_signature = row["waiver_owner_signature"]
            row["reviewer_comment"] = "Reviewed"
            generate_waiver_from_rows([row], output)
            self.assertEqual(row["waiver_timestamp"], first_timestamp)
            row["owner_comment"] = "Updated reason"
            generate_waiver_from_rows([row], output)
            changed_timestamp = row["waiver_timestamp"]
            self.assertNotEqual(changed_timestamp, first_timestamp)
            self.assertNotEqual(row["waiver_owner_signature"], first_signature)
            row["ip_owner"] = "owner_b"
            generate_waiver_from_rows([row], output)
            self.assertNotEqual(row["waiver_timestamp"], changed_timestamp)
            self.assertIn("-user { owner_b }", output.read_text(encoding="utf-8"))
            row["owner_action"] = "FIXED"
            generate_waiver_from_rows([row], output)
            self.assertEqual(row["waiver_timestamp"], "")
            self.assertNotIn("waive_violation -add", output.read_text(encoding="utf-8"))

    def test_legacy_unchanged_rule_keeps_timestamp_on_first_generation(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
            "waiver_timestamp": "01-01-2020 00:00:00",
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            timestamp = row["waiver_timestamp"]
            row.pop("waiver_owner_signature")
            generate_waiver_from_rows([row], output)

        self.assertEqual(row["waiver_timestamp"], timestamp)

    def test_waiver_generation_requires_ip_owner(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "ip_owner": "",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            with self.assertRaisesRegex(ValueError, "IP Owner is required"):
                generate_waiver_from_rows([row], output)

    def test_same_rule_with_different_ip_owners_is_not_grouped(self):
        first = self.row("first")
        second = self.row("second")
        second["ip_owner"] = "owner_b"
        for row in (first, second):
            row.update({
                "owner_action": "WAIVED",
                "owner_comment": "Shared reason",
                "filter_mode": "CUSTOM",
                "custom_filter": '(Module == "top")',
            })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([first, second], output)
            generated = output.read_text(encoding="utf-8")

        self.assertEqual(generated.count("waive_violation -add"), 2)
        self.assertIn("-user { owner_a }", generated)
        self.assertIn("-user { owner_b }", generated)

    def test_items_owned_by_another_waiver_file_are_never_generated(self):
        row = self.row("external-waiver")
        row.update({
            "owner_action": "WAIVED",
            "tag": "RTL_PRAGMA",
            "waiver_source_file": "rtl_pragma_waiver.tcl",
            "owner_comment": "Do not export this item",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")

        self.assertNotIn("waive_violation -add", generated)
        self.assertNotIn("RTL_PRAGMA", generated)

    def test_collect_current_rows_keeps_only_waivers_from_vc_waiver_tcl(self):
        report_template = """  W551  (0 warnings/2 waived)
  -----------------------------------------------------------------------------
  Tag                 : W551
  Violation           : Lint:{number}
  Module              : top
  LineNumber          : {number}
  Goal                : LINT
  Waiver
    Name              : W551_{number}
    Filename          : {filename}
    State             : Waived
  -----------------------------------------------------------------------------
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            full_report = temp / "full.log"
            waived_report = temp / "waived.log"
            full_report.write_text("", encoding="utf-8")
            waived_report.write_text(
                report_template.format(number=1, filename="vc_waiver.tcl")
                + report_template.format(number=2, filename="rtl_pragma_waiver.tcl"),
                encoding="utf-8",
            )
            rows = collect_current_rows(full_report, waived_report, None)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["violation"], "Lint:1")
        self.assertEqual(rows[0]["waiver_source_file"], "vc_waiver.tcl")

    def test_waiver_filter_uses_gui_braced_format_and_preserves_braces(self):
        row = self.row("waive")
        expression = '(Statement == "r_data <= {a, b};")'
        row.update({
            "owner_action": "WAIVED",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": expression,
            "fields_json": '{"Tag":"W551"}',
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")

        self.assertIn(
            '-filter {(Statement == "r_data <= {a, b};")}',
            generated,
        )
        interpreter = tkinter.Tcl()
        interpreter.eval('proc waive_violation {args} {set ::waiver_args $args}')
        interpreter.eval(generated)
        args = interpreter.splitlist(interpreter.eval('set ::waiver_args'))
        self.assertEqual(args[args.index("-filter") + 1], expression)

    def test_filter_fields_support_row_values_and_wildcards(self):
        fields = {"Goal": "LINT", "Module": "top"}
        selected = parse_filter_fields_spec("Goal, Module=axi_*", fields)

        self.assertEqual(filter_expr(selected), '(Goal == "LINT") AND (Module =~ "axi_*")')

    def test_statement_is_trimmed_and_wrapped_with_wildcards(self):
        self.assertEqual(
            filter_expr({"Statement": "   assign y = select ? a * b : c;   "}),
            '(Statement =~ "*assign y = select ? a * b : c;*")',
        )

    def test_statement_does_not_duplicate_existing_edge_wildcards(self):
        self.assertEqual(
            filter_expr({"Statement": "*assign y = a;*"}),
            '(Statement =~ "*assign y = a;*")',
        )

    def test_filter_fields_override_can_contain_commas(self):
        fields = {"Statement": "original"}
        expected = {"Statement": "r_data <= {a, b};"}
        spec = format_filter_fields_spec(expected, fields)

        self.assertEqual(parse_filter_fields_spec(spec, fields), expected)

    def test_legacy_unquoted_override_with_commas_is_supported(self):
        fields = {"Goal": "LINT"}
        spec = "Goal, Statement=r_data <= {a, b};, Signal=sig"

        selected = parse_filter_fields_spec(spec, fields)

        self.assertEqual(selected["Statement"], "r_data <= {a, b};")
        self.assertEqual(selected["Signal"], "sig")

    def test_identical_waiver_rules_are_emitted_once_and_marked(self):
        first = self.row("first")
        second = self.row("second")
        for row in (first, second):
            row.update({
                "owner_action": "WAIVED",
                "owner_comment": "Shared reason",
                "filter_mode": "CUSTOM",
                "custom_filter": '(Module == "top")',
                "fields_json": '{"Tag":"W551","Module":"top"}',
            })

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            generate_waiver_from_rows([first, second], output)
            generated = output.read_text(encoding="utf-8")

        self.assertEqual(generated.count("waive_violation -add"), 1)
        self.assertEqual(first["waiver_name"], second["waiver_name"])
        self.assertIn("Primary waiver rule", first["_waiver_note"])
        self.assertIn("primary issue first", second["_waiver_note"])

    def test_auto_filter_adds_fields_until_issue_is_unique(self):
        first = self.row("first", object_name="sig_a")
        second = self.row("second", object_name="sig_b")
        first["fields_json"] = '{"Goal":"LINT","Module":"top","Signal":"sig_a"}'
        second["fields_json"] = '{"Goal":"LINT","Module":"top","Signal":"sig_b"}'

        expression = waiver_filter_expression(first, [first, second])

        self.assertEqual(
            expression,
            '(Goal == "LINT") AND (Module == "top") AND (Signal == "sig_a")',
        )

    def test_auto_filter_replaces_statement_indentation_with_wildcards(self):
        first = self.row("first")
        second = self.row("second")
        first["fields_json"] = json.dumps({
            "Goal": "LINT", "Module": "top", "Statement": "        if (enable) begin",
        })
        second["fields_json"] = json.dumps({
            "Goal": "LINT", "Module": "top", "Statement": "    if (other) begin",
        })

        expression = waiver_filter_expression(first, [first, second])

        self.assertIn('(Statement =~ "*if (enable) begin*")', expression)

    def test_auto_filter_excludes_filename_and_line_number(self):
        row = self.row("single")
        row["fields_json"] = (
            '{"Goal":"LINT","Module":"top","FileName":"top.sv",'
            '"LineNumber":"42","Signal":"sig"}'
        )

        expression = waiver_filter_expression(row, [row])

        self.assertNotIn("FileName", expression)
        self.assertNotIn("LineNumber", expression)

    def test_auto_filter_index_reuses_counts_for_many_waivers(self):
        rows = []
        for index in range(100):
            row = self.row(f"issue-{index}", object_name=f"sig_{index}")
            row["fields_json"] = json.dumps({
                "Goal": "LINT", "Module": "top", "Signal": f"sig_{index}",
            })
            rows.append(row)

        auto_filter_index = AutoFilterIndex(rows)
        with patch.object(auto_filter_index, "match_count", wraps=auto_filter_index.match_count) as match_count:
            expressions = [
                waiver_filter_expression(row, rows, auto_filter_index)
                for row in rows
            ]

        self.assertEqual(len(expressions), 100)
        self.assertEqual(len(auto_filter_index._counts), 3)
        self.assertEqual(match_count.call_count, 300)
        self.assertIn('(Signal == "sig_99")', expressions[-1])


if __name__ == "__main__":
    unittest.main()
