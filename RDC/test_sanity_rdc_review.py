import tempfile
import unittest
from collections import Counter
from pathlib import Path
from unittest.mock import patch

from sanity_rdc_review import (
    apply_existing_rdc_waivers,
    auto_filter_fields,
    collect_current_rows_from_report_excel,
    filter_expr,
    generate_waiver_from_rows,
    merge_rows,
    parse_filter,
    parse_waiver_tcl,
    read_review_workbook,
)


BASE_DIR = Path(__file__).resolve().parent


class CurrentDataTests(unittest.TestCase):
    def test_report_workbook_has_all_error_and_warning_issues(self):
        rows = collect_current_rows_from_report_excel(BASE_DIR / "report_rdc.full.xlsx")
        self.assertEqual(len(rows), 162)
        self.assertEqual(len({row["issue_id"] for row in rows}), 162)
        self.assertEqual(
            Counter(row["tag"] for row in rows),
            Counter({
                "RDC_CLOCK_CORRUPT_OBSERVED": 153,
                "SETUP_LIBCELLPIN_UNMODELLED": 6,
                "SETUP_RESET_CONSTANT_ACTIVE": 3,
            }),
        )

    def test_reference_file_migrates_current_rdc_waivers_only(self):
        rows = collect_current_rows_from_report_excel(BASE_DIR / "report_rdc.full.xlsx")
        apply_existing_rdc_waivers(rows, BASE_DIR / "vc_waiver.tcl_ori")
        waived = [row for row in rows if row["owner_action"] == "WAIVED"]
        self.assertEqual(len(waived), 114)
        self.assertTrue(all(row["filter_mode"] == "CUSTOM" for row in waived))
        self.assertTrue(all(row["custom_filter"] for row in waived))

    def test_generated_workbook_round_trip(self):
        rows = read_review_workbook(BASE_DIR / "outputs" / "rdc_review.xlsx")
        self.assertEqual(len(rows), 162)
        self.assertEqual(sum(row["owner_action"] == "WAIVED" for row in rows), 114)


class WaiverFormatTests(unittest.TestCase):
    def test_reference_parser_finds_rdc_rules(self):
        rules = parse_waiver_tcl(BASE_DIR / "vc_waiver.tcl_ori")
        rdc_rules = [rule for rule in rules.values() if str(rule.get("app", "")).strip() == "rdc"]
        self.assertEqual(len(rdc_rules), 203)

    def test_generation_uses_rdc_app_and_preserves_custom_filter(self):
        row = {
            "issue_id": "one", "record_status": "NEW", "owner_action": "WAIVED",
            "owner_comment": "reviewed", "filter_mode": "CUSTOM",
            "custom_filter": '(DestObject == "u_gate/ECK")', "filter_fields": "",
            "tag": "RDC_CLOCK_CORRUPT_OBSERVED", "violation": "RDC:99", "waiver_name": "RDC_CLOCK_CORRUPT_OBSERVED_99",
            "waiver_user": "", "waiver_timestamp": "", "waiver_source_file": "vc_waiver.tcl",
            "fields_json": '{"Tag":"RDC_CLOCK_CORRUPT_OBSERVED","DestObject":"u_gate/ECK"}',
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "vc_waiver.tcl"
            with patch.dict("os.environ", {"USERNAME": "rdc_user"}):
                generate_waiver_from_rows([row], output)
            generated = output.read_text(encoding="utf-8")
            parsed = parse_waiver_tcl(output)
        self.assertIn("-app { rdc }", generated)
        self.assertIn('-filter {(DestObject == "u_gate/ECK")}', generated)
        self.assertEqual(len(parsed), 1)
        self.assertEqual(next(iter(parsed.values()))["app"], "rdc")

    def test_qualified_rdc_filter_fields_round_trip(self):
        fields = {
            "ReasonInfoList:ReasonInfo:ReasonCode": "SYNC_SRC_CONV",
            "RdcSourceResets:ResetName": "RESETN_ACLK",
        }
        self.assertEqual(parse_filter(filter_expr(fields)), fields)


class AutoAndMergeTests(unittest.TestCase):
    @staticmethod
    def row(issue_id, convergence):
        return {
            "issue_id": issue_id, "record_status": "NEW", "tag": "RDC_CLOCK_CORRUPT_OBSERVED",
            "module": "top", "file": "top.sv", "line": "10", "object": convergence,
            "fields_json": (
                '{"Tag":"RDC_CLOCK_CORRUPT_OBSERVED","Module":"top","FileName":"top.sv",'
                f'"LineNumber":"10","SrcObject":"{convergence}"}}'
            ),
            "owner_action": "UNREVIEWED", "reviewer_decision": "PENDING",
        }

    def test_auto_excludes_filename_and_line_and_reaches_rdc_object(self):
        first = self.row("one", "u_top/a/D")
        second = self.row("two", "u_top/b/D")
        selected = auto_filter_fields(first, [first, second])
        self.assertNotIn("FileName", selected)
        self.assertNotIn("LineNumber", selected)
        self.assertEqual(selected["SrcObject"], "u_top/a/D")

    def test_merge_preserves_review_decisions(self):
        old = self.row("one", "u_top/a/D")
        old.update({"owner_action": "WAIVED", "owner_comment": "reason", "reviewer_decision": "APPROVED"})
        current = self.row("one", "u_top/a/D")
        merged = merge_rows([old], [current])
        self.assertEqual(merged[0]["record_status"], "UNCHANGED")
        self.assertEqual(merged[0]["owner_action"], "WAIVED")
        self.assertEqual(merged[0]["reviewer_decision"], "APPROVED")


if __name__ == "__main__":
    unittest.main()
