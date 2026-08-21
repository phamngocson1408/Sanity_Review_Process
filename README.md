# Sanity Review Process

This repository contains a sample LINT review/waiver flow for Jira SDI-20.

The important idea is:

```text
report_lint.full.xlsx + reviewer decisions
  -> Git-friendly review_db.csv
  -> Excel workbook for reviewer
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

Each tag sheet follows `report_lint.full.xlsx`. The reviewer-editable report columns are:

```text
Judgment, Comment
```

User-added columns are preserved in `lint_review.xlsx` for human notes, but ignored when `vc_waiver.tcl` is generated.

`vc_waiver.tcl` is generated in GUI-style: one `waive_violation` command per waived issue, with names based on the LINT violation number such as `DeadCode-ML_739`.

## Files

```text
LINT/reports/report_lint.full.log
LINT/report_lint.full.xlsx
LINT/vc_waiver.tcl
LINT/sanity_lint_review.py
LINT/data/lint_review_db.csv
LINT/outputs/lint_review.xlsx
LINT/outputs/lint_summary.csv
LINT/outputs/waiver_rule_audit.csv
LINT/vc_waiver.tcl
```

## Two-command review flow

Run this from the `LINT` directory:

```powershell
cd LINT
python sanity_lint_review.py prepare-waiver
```

This imports existing Excel edits and generates `vc_waiver.tcl`.

Then run the sanity tool so `reports/report_lint.full.log` is refreshed with the new waiver file.

After the sanity tool finishes, run:

```powershell
python sanity_lint_review.py merge-report
```

This runs `make -f Makefile excel`, merges the refreshed `report_lint.full.xlsx` into `data/lint_review_db.csv`, marks issue state as `NEW` / `CHANGED` / `ACTIVE` / `REMOVED`, and exports a refreshed `outputs/lint_review.xlsx`.

For the detailed loop, see:

```text
LINT/REVIEW_FLOW.md
```

Running `python sanity_lint_review.py` without arguments prints the two-command sequence and does not modify files.

## Normal update flow

1. Generate waiver Tcl from reviewer decisions:

```bash
python sanity_lint_review.py prepare-waiver
```

2. Run the sanity tool so the reports are regenerated with the new `vc_waiver.tcl`.

3. Convert and merge the refreshed report workbook:

```bash
python sanity_lint_review.py merge-report
```

4. Reviewer edits `Judgment` or `Comment` in `outputs/lint_review.xlsx`.

5. Repeat from step 1 for the next sanity run.

6. Audit redundant GUI waiver rules when needed:

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
- `record_status`: `NEW`, `CHANGED`, `ACTIVE`, or `REMOVED`.
- `review_status`: `UNREVIEWED`, `WAIVED`, `APPROVED`, or `APPROVED_WAIVE`.
- `review_comment`: waiver/review reason.
- `filter_json`: internally derived fields used to generate the Tcl `-filter`.
- `fields_json`: full parsed LINT fields for audit/debug.

## Redundant waiver removal

`LINT/outputs/waiver_rule_audit.csv` compares the GUI-generated `vc_waiver.tcl` rules with the current review DB.

- `ACTIVE`: waiver name is referenced by at least one current waived issue.
- `REDUNDANT`: waiver name is not referenced by any current waived issue and should be reviewed for removal.
- `MISSING_IN_TCL`: review DB references this waiver name, but it was not found in the input `vc_waiver.tcl`.

## Wildcard

Wildcard is controlled by the internally derived `filter_json`.

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
