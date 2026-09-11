import unittest
from pathlib import Path

from sanity_lint_review import filter_expr, parse_filter, parse_waiver_tcl


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

    def test_original_w551_filter(self):
        rules = parse_waiver_tcl(Path(__file__).with_name("vc_waiver.tcl_ori"))
        fields = rules["W551_837"]["filter_fields"]
        self.assertEqual(fields["PropertyList:LintPropertyName"], "Property_142")
        self.assertNotIn("LintPropertyName", fields)
        self.assertEqual(parse_filter(filter_expr(fields)), fields)


if __name__ == "__main__":
    unittest.main()
