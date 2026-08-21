# Sanity Review Process

This repository contains a sample LINT review/waiver flow for Jira SDI-20.

## Weakness Of Previous Procedure

The previous procedure relied mainly on the `vc_waiver.tcl` file and a separate PPT file, instead of a single Excel source of truth. That had several weaknesses:

- `vc_waiver.tcl` is a plain text file, so it cannot carry detailed review context such as screenshots or images of the waived issue.
- The PPT file did carry that detailed context, but it was maintained separately and was not synced regularly with `vc_waiver.tcl`, so the two artifacts drifted apart over time.
- Because review information was split across a text file and a slide deck with no consistent sync, reviewers had no single place to see both the waiver decision and its justification.
- This made review and peer-review difficult: peer reviewers had to cross-check two disconnected files, and outdated or missing PPT updates could hide the real reasoning behind a waiver.

## Proposed Solution

Instead of splitting the waiver decision into `vc_waiver.tcl` and its justification into a rarely-synced PPT file, this flow uses Excel as both the reviewer UI and the review source of truth, so decisions and their supporting context always live together and stay in sync:

```text
report_lint.full.xlsx + reviewer decisions
  -> lint_review.xlsx as the review source of truth
  -> generated VC LINT waiver Tcl
```

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

Green headers are user-editable. Gray headers are report-owned and should not be edited.
The `Judgment` column has a dropdown for the supported review decisions.

User-added columns are preserved in `lint_review.xlsx` for human notes, but ignored when `vc_waiver.tcl` is generated.

`vc_waiver.tcl` is generated in GUI-style: one `waive_violation` command per waived issue, with names based on the LINT violation number such as `DeadCode-ML_739`. It is always regenerated from `lint_review.xlsx`, so it never drifts out of sync the way it could with a manually maintained PPT.

## Supported Features

This environment supports a two-command review loop around the existing sanity flow:

- Generate `vc_waiver.tcl` directly from reviewer decisions in `outputs/lint_review.xlsx`.
- Convert the latest sanity report log into `report_lint.full.xlsx` by running `make -f Makefile excel`.
- Merge `report_lint.full.xlsx` into `outputs/lint_review.xlsx` while preserving reviewer `Judgment` and `Comment`.
- Keep `outputs/lint_review.xlsx` as the single source of truth for review status and waiver decisions.
- Mark issues as `NEW`, `ACTIVE`, `CHANGED`, or `REMOVED` during merge.
- Preserve user-added columns in `lint_review.xlsx` for notes, tracking, links, or images; those columns are ignored by waiver generation.
- Color workbook headers: green for user-editable columns and gray for report-owned columns.
- Enable sheet filters and provide a dropdown list for the `Judgment` column.
- Generate GUI-style VC LINT waiver Tcl from rows whose `Judgment` is `WAIVED`, `APPROVED`, or `APPROVED_WAIVE`.

## Files

```text
LINT/reports/report_lint.full.log
LINT/report_lint.full.xlsx
LINT/vc_waiver.tcl
LINT/sanity_lint_review.py
LINT/outputs/lint_review.xlsx
LINT/outputs/lint_summary.csv
LINT/outputs/waiver_rule_audit.csv
LINT/vc_waiver.tcl
```

## Two-command review flow

Run this from the `LINT` directory:

```powershell
cd LINT
python sanity_lint_review.py gen_waiver
```

This imports existing Excel edits and generates `vc_waiver.tcl`.

Then run the sanity tool so `reports/report_lint.full.log` is refreshed with the new waiver file.

After the sanity tool finishes, run:

```powershell
python sanity_lint_review.py merge_excel
```

This runs `make -f Makefile excel`, merges the refreshed `report_lint.full.xlsx` into `outputs/lint_review.xlsx`, and marks issue state as `NEW` / `CHANGED` / `ACTIVE` / `REMOVED`.

For the detailed loop, see:

```text
LINT/REVIEW_FLOW.md
```

Running `python sanity_lint_review.py` without arguments prints the two-command sequence and does not modify files.

## Normal update flow

1. Generate waiver Tcl from reviewer decisions:

```bash
python sanity_lint_review.py gen_waiver
```

2. Run the sanity tool so the reports are regenerated with the new `vc_waiver.tcl`.

3. Convert and merge the refreshed report workbook:

```bash
python sanity_lint_review.py merge_excel
```

4. Reviewer edits `Judgment` or `Comment` in `outputs/lint_review.xlsx`.

5. Repeat from step 1 for the next sanity run.

6. Clean generated files:

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

## Workbook Metadata

`lint_review.xlsx` keeps the report-owned columns from `report_lint.full.xlsx`, plus `record_status`.

Reviewer-owned columns:

```text
Judgment
Comment
```

The script derives `issue_id` and Tcl filter fields from the report-owned columns when it needs them.

## Redundant waiver removal

`LINT/outputs/waiver_rule_audit.csv` compares the GUI-generated `vc_waiver.tcl` rules with the current review workbook.

- `ACTIVE`: waiver name is referenced by at least one current waived issue.
- `REDUNDANT`: waiver name is not referenced by any current waived issue and should be reviewed for removal.
- `MISSING_IN_TCL`: the review workbook references this waiver name, but it was not found in the input `vc_waiver.tcl`.

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
