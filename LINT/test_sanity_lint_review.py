import unittest
import tempfile
import tkinter
import zipfile
from pathlib import Path

from sanity_lint_review import (
    WORKBOOK_MANAGEMENT_COLUMNS,
    filter_expr,
    format_filter_fields_spec,
    generate_waiver_from_rows,
    merge_rows,
    normalize_record_status,
    parse_filter_fields_spec,
    parse_filter,
    parse_waiver_tcl,
    report_ordered_sheet_columns,
    sheet_xml,
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
        self.assertEqual(parse_filter(expression), fields)

    def test_tcl_double_quote_preserves_filter_value(self):
        expression = '(Statement == "r_rdata <= {data[3:0], $value ...")'
        interpreter = tkinter.Tcl()
        interpreter.eval('proc capture {value} {set ::captured $value}')

        interpreter.eval(f'capture "{tcl_double_quote(expression)}"')

        self.assertEqual(interpreter.eval('set ::captured'), expression)

    def test_original_w551_filter(self):
        rules = parse_waiver_tcl(Path(__file__).with_name("vc_waiver.tcl_ori"))
        fields = rules["W551_837"]["filter_fields"]
        self.assertEqual(fields["PropertyList:LintPropertyName"], "Property_142")
        self.assertNotIn("LintPropertyName", fields)
        self.assertEqual(parse_filter(filter_expr(fields)), fields)


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
        self.assertIn(r'-filter "(Module =~ \"top_*\")"', generated)

    def test_filter_fields_support_row_values_and_wildcards(self):
        fields = {"Goal": "LINT", "Module": "top"}
        selected = parse_filter_fields_spec("Goal, Module=axi_*", fields)

        self.assertEqual(filter_expr(selected), '(Goal == "LINT") AND (Module =~ "axi_*")')

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

    def test_auto_filter_excludes_filename_and_line_number(self):
        row = self.row("single")
        row["fields_json"] = (
            '{"Goal":"LINT","Module":"top","FileName":"top.sv",'
            '"LineNumber":"42","Signal":"sig"}'
        )

        expression = waiver_filter_expression(row, [row])

        self.assertNotIn("FileName", expression)
        self.assertNotIn("LineNumber", expression)


if __name__ == "__main__":
    unittest.main()
