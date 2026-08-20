# Sanity Review Process

This repository contains a sample LINT review/waiver flow for Jira SDI-20.

The important idea is:

```text
LINT reports + GUI waiver Tcl
  -> Git-friendly review_db.csv
  -> Excel workbook for reviewer, using the report_lint.full.xlsx style
  -> validated review_db.csv
  -> generated VC LINT waiver Tcl
```

Excel is used as the reviewer UI. The CSV file is the Git-reviewable source of truth.

The generated Excel workbook follows the familiar LINT report layout:

```text
Tree Summary
W216
W240
W528
...
```

Each tag sheet starts with reviewer columns:

```text
No., Person in Charge, Date, Judgment, Comment
```

The management columns needed for import/generate are kept at the right side of each tag sheet.

`vc_waiver.tcl` is generated in GUI-style: one `waive_violation` command per waived issue, with names based on the LINT violation number such as `DeadCode-ML_739`.

## Files

```text
LINT/reports/report_lint.full.log
LINT/reports/report_lint.waived.log
LINT/vc_waiver.tcl
LINT/sanity_lint_review.py
LINT/data/lint_review_db.csv
LINT/outputs/lint_review.xlsx
LINT/outputs/lint_summary.csv
LINT/outputs/waiver_rule_audit.csv
LINT/vc_waiver.tcl
```

## One-shot sample generation without options

Run this from the `LINT` directory:

```powershell
cd LINT
python sanity_lint_review.py
```

This imports existing Excel edits, generates `vc_waiver.tcl`, parses current reports, merges them into `data/lint_review_db.csv`, marks issue state as `NEW` / `CHANGED` / `ACTIVE` / `WAIVED` / `REMOVED`, and exports a refreshed `outputs/lint_review.xlsx`.

For the detailed loop, see:

```text
LINT/REVIEW_FLOW.md
```

This default command is also available as:

```powershell
python sanity_lint_review.py run_all
```

## Normal update flow

1. Parse the newest report:

```bash
python sanity_lint_review.py parse \
  --full-report reports/report_lint.full.log \
  --waived-report reports/report_lint.waived.log \
  --waiver-tcl vc_waiver.tcl \
  --output data/current_lint_review_db.csv
```

2. Merge with the previous Git-managed DB:

```bash
python sanity_lint_review.py merge \
  --old-db data/lint_review_db.csv \
  --current-db data/current_lint_review_db.csv \
  --output data/lint_review_db.csv
```

3. Export reviewer Excel:

```bash
python sanity_lint_review.py export-excel \
  --review-db data/lint_review_db.csv \
  --excel outputs/lint_review.xlsx \
  --summary outputs/lint_summary.csv
```

4. After reviewer edits `Person in Charge`, `Date`, `Judgment`, `Comment`, `waiver_enabled`, `waiver_name`, or `filter_json` in Excel, import it back to CSV:

```bash
python sanity_lint_review.py import-excel \
  --excel outputs/lint_review.xlsx \
  --output data/lint_review_db.csv
```

5. Generate waiver Tcl from the edited Excel workbook:

```bash
python sanity_lint_review.py generate-waiver-from-excel \
  --excel outputs/lint_review.xlsx \
  --review-db data/lint_review_db.csv \
  --output vc_waiver.tcl
```

6. Audit redundant GUI waiver rules:

```bash
python sanity_lint_review.py audit-waivers \
  --review-db data/lint_review_db.csv \
  --waiver-tcl vc_waiver.tcl \
  --output outputs/waiver_rule_audit.csv
```

7. Clean generated files:

```powershell
python sanity_lint_review.py clean
```

Preview first:

```powershell
python sanity_lint_review.py clean --dry-run
```

Keep the generated waiver file:

```powershell
python sanity_lint_review.py clean --keep-waiver
```

## Review DB columns

- `issue_id`: stable hash-based issue identifier.
- `record_status`: `NEW`, `CHANGED`, `ACTIVE`, `WAIVED`, or `REMOVED`.
- `review_status`: `UNREVIEWED`, `WAIVED`, `APPROVED`, or `APPROVED_WAIVE`.
- `waiver_enabled`: `yes` or `no`.
- `waiver_name`: Tcl waiver name.
- `review_comment`: waiver/review reason.
- `filter_json`: fields used to generate the Tcl `-filter`.
- `fields_json`: full parsed LINT fields for audit/debug.

## Redundant waiver removal

`LINT/outputs/waiver_rule_audit.csv` compares the GUI-generated `vc_waiver.tcl` rules with the currently waived issues.

- `ACTIVE`: waiver name is referenced by at least one current waived issue.
- `REDUNDANT`: waiver name is not referenced by any current waived issue and should be reviewed for removal.
- `MISSING_IN_TCL`: waived report references this waiver name, but it was not found in the input `vc_waiver.tcl`.

## Wildcard

Wildcard is controlled in `filter_json`.

Example:

```json
{"Goal":"BOS_LINT_RULE","Module":"AXICRYPT_*","Signal":"I_*"}
```

When a value contains `*` or `?`, the generated Tcl uses `=~`.

```tcl
(Module =~ "AXICRYPT_*")
```

Otherwise it uses exact match:

```tcl
(Module == "AXICRYPT_RCC")
```
