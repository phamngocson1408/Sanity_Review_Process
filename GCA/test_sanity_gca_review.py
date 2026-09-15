import tempfile
import unittest
from pathlib import Path

from sanity_gca_review import (
    BASE_DIR,
    condition_targets,
    generate_waiver,
    merge_rows,
    parse_gca_report,
    parse_waiver_tcl,
    tokenize_tcl_command,
    waiver_condition,
)


class TclParserTests(unittest.TestCase):
    def test_nested_conditions_remain_one_word(self):
        line = (
            'create_waiver -name w1 -design top -rule NTL_0001 '
            '-comment "reviewed reason" -condition [list port [get_ports {A[1] A[0]}]]'
        )
        words = tokenize_tcl_command(line)
        self.assertIn('[list port [get_ports {A[1] A[0]}]]', words)

    def test_existing_waiver_file_is_parsed(self):
        rules = parse_waiver_tcl(BASE_DIR / "GCA_confirmed.tcl")
        self.assertEqual(len(rules), 284)
        self.assertEqual(rules[0]["rule"], "NTL_0001")
        self.assertEqual(rules[0]["conditions"], ["[list port [get_ports {AESDMA_APB_SI0_PREADY}]]"])


class ReportTests(unittest.TestCase):
    def test_current_report_and_waivers_are_correlated(self):
        rows = parse_gca_report(
            BASE_DIR / "reports" / "BOS_AESDMA_gca_const_analysis.rpt",
            BASE_DIR / "reports" / "BOS_AESDMA_gca.rpt",
            BASE_DIR / "GCA_confirmed.tcl",
        )
        self.assertEqual(len(rows), 234)
        self.assertEqual(sum(row["rule"] == "NTL_0001" for row in rows), 228)
        self.assertEqual(sum(row["rule"] == "EXC_0004" for row in rows), 6)
        self.assertTrue(all(row.get("owner_action") == "WAIVED" for row in rows))
        self.assertTrue(all(row.get("auto_condition") for row in rows))

    def test_generated_file_covers_every_current_ntl_object(self):
        rows = parse_gca_report(
            BASE_DIR / "reports" / "BOS_AESDMA_gca_const_analysis.rpt",
            BASE_DIR / "reports" / "BOS_AESDMA_gca.rpt",
            BASE_DIR / "GCA_confirmed.tcl",
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "GCA_confirmed.tcl"
            generate_waiver(rows, output)
            generated = parse_waiver_tcl(output)
        covered = set()
        for waiver in generated:
            if waiver["rule"] == "NTL_0001":
                covered.update(condition_targets(waiver["conditions"]))
        expected = {row["object"] for row in rows if row["rule"] == "NTL_0001"}
        self.assertEqual(covered, expected)
        self.assertEqual(sum(waiver["rule"] == "EXC_0004" for waiver in generated), 6)


class WorkflowTests(unittest.TestCase):
    def row(self, issue_id: str, message: str = "Output port 'PREADY' is unbuffered."):
        return {
            "issue_id": issue_id, "record_status": "NEW", "owner_action": "WAIVED",
            "owner_comment": "approved", "filter_mode": "AUTO", "filter_fields": "",
            "custom_filter": "", "auto_condition": "-condition [list port [get_ports {PREADY}]]",
            "condition_json": '["[list port [get_ports {PREADY}]]"]', "design": "top",
            "scenario": "global", "rule": "NTL_0001", "object": "PREADY", "message": message,
            "issue_index": "1", "waiver_name": "w1", "waiver_user": "user",
            "waiver_timestamp": "time", "ip_owner": "owner", "reviewer": "reviewer",
            "reviewer_decision": "APPROVED", "reviewer_comment": "ok",
        }

    def test_auto_uses_native_gca_condition(self):
        self.assertEqual(
            waiver_condition(self.row("one")),
            "-condition [list port [get_ports {PREADY}]]",
        )

    def test_custom_requires_a_condition_clause(self):
        row = self.row("one")
        row.update({"filter_mode": "CUSTOM", "custom_filter": "anything"})
        with self.assertRaises(ValueError):
            waiver_condition(row)

    def test_merge_preserves_review_and_marks_changed(self):
        old = self.row("old", "Output port 'PREADY' was unbuffered.")
        current = self.row("new", "Output port 'PREADY' is unbuffered.")
        current.update({"owner_action": "UNREVIEWED", "reviewer_decision": "PENDING"})
        merged = merge_rows([old], [current])
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["record_status"], "CHANGED")
        self.assertEqual(merged[0]["owner_action"], "WAIVED")
        self.assertEqual(merged[0]["reviewer_decision"], "APPROVED")


if __name__ == "__main__":
    unittest.main()
