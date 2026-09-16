import unittest
import tempfile
import tkinter
import zipfile
import os
import json
from unittest.mock import patch
from pathlib import Path

from sanity_lint_review import (
    AutoFilterIndex,
    WORKBOOK_MANAGEMENT_COLUMNS,
    filter_expr,
    format_filter_fields_spec,
    generate_waiver_from_rows,
    collect_current_rows,
    merge_rows,
    normalize_record_status,
    parse_filter_fields_spec,
    parse_filter,
    parse_waiver_tcl,
    report_ordered_sheet_columns,
    sheet_xml,
    tcl_braced_filter,
    tcl_double_quote,
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

    def test_status_headers_have_excel_notes(self):
        rows = [
            ["Owner Action", "Reviewer Decision", "record_status", "source_report", "Tag"],
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


class RecordStatusTests(unittest.TestCase):
    @staticmethod
    def row(issue_id, line="1", status="UNCHANGED", object_name=None):
        return {
            "issue_id": issue_id,
            "record_status": status,
            "owner_action": "UNREVIEWED",
            "reviewer_decision": "PENDING",
            "tag": "W551",
            "goal": "LINT",
            "module": "top",
            "file": "top.sv",
            "line": line,
            "hierarchy": "top",
            "object": object_name or issue_id,
        }

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

    def test_waiver_generation_supplies_default_user_and_timestamp(self):
        row = self.row("waive")
        row.update({
            "owner_action": "WAIVED",
            "owner_comment": "Reason",
            "filter_mode": "CUSTOM",
            "custom_filter": '(Module == "top")',
            "waiver_user": "",
            "waiver_timestamp": "N/A",
        })
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "waiver.tcl"
            with patch.dict(os.environ, {"USERNAME": ""}):
                generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")

        self.assertIn("-user { sanity_lint_review }", generated)
        self.assertRegex(
            generated,
            r"-timestamp \{ \d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2} \}",
        )
        self.assertEqual(row["waiver_user"], "sanity_lint_review")
        self.assertNotEqual(row["waiver_timestamp"], "N/A")

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
