# LINT Sanity Review Flow

This document describes the intended review loop for LINT sanity issues and waivers.

## Main Idea

`outputs/lint_review.xlsx` is the reviewer UI and the review source of truth.

`report_lint.full.xlsx` is the tool-generated LINT report.

`vc_waiver.tcl` is generated from approved/waived items in `lint_review.xlsx` and is used by the next sanity check run.

## Directory Layout

```text
LINT/
  reports/
    report_lint.full.log
  report_lint.full.xlsx
  outputs/
    lint_review.xlsx
    lint_summary.csv
    waiver_rule_audit.csv
  sanity_lint_review.py
  vc_waiver.tcl
```

## Normal Loop

1. Run the sanity check script.

The sanity tool creates or refreshes the report log under `reports/`.

2. Generate the waiver file from the reviewer workbook.

From the `LINT` directory:

```powershell
python sanity_lint_review.py prepare-waiver
```

The script does the following:

```text
existing outputs/lint_review.xlsx, if any
  -> generate vc_waiver.tcl
```

3. Run the sanity tool again.

This step must happen after `vc_waiver.tcl` is generated, so the next report reflects the user's latest waiver decisions.

4. Merge the refreshed sanity report into the review workbook.

From the `LINT` directory:

```powershell
python sanity_lint_review.py merge-report
```

The script does the following:

```text
make -f Makefile excel
  -> generate report_lint.full.xlsx from the refreshed reports

report_lint.full.xlsx
  -> parse current issues
  -> merge with outputs/lint_review.xlsx
  -> export outputs/lint_review.xlsx
```

5. Reviewer edits `outputs/lint_review.xlsx`.

The reviewer can update these report columns:

```text
Judgment
Comment
```

The reviewer may also add new columns for human notes. Those columns are preserved in `lint_review.xlsx`, but ignored by `vc_waiver.tcl` generation.

6. Generate the next waiver file.

Run:

```powershell
python sanity_lint_review.py prepare-waiver
```

This reads reviewer edits from `outputs/lint_review.xlsx`, then generates:

```text
vc_waiver.tcl
```

7. Run sanity check again with the new `vc_waiver.tcl`.

The sanity tool creates new reports.

8. Merge the refreshed report again.

```powershell
python sanity_lint_review.py merge-report
```

The script merges the new reports with the previous `lint_review.xlsx`, so old comments/status are preserved.

## Status Meaning

`record_status` is the issue state based on the latest sanity reports.

```text
NEW
  The issue appears in the latest report and was not found in the previous review workbook.

ACTIVE
  The issue appears in the latest full report and already existed in the review workbook.

CHANGED
  The issue no longer has the same exact issue_id, but it still looks like the same logical issue.
  Example: line number or statement changed, while tag/module/file/hierarchy/object stayed the same.
  Previous reviewer fields are preserved and the row is marked for manual confirmation.

WAIVED
  The issue appears in the latest waived report. This means the sanity tool recognized it as waived.

REMOVED
  The issue existed in the previous review workbook but no longer appears in the latest reports.
```

`review_status` is imported from the Excel `Judgment` column.

```text
UNREVIEWED
  No review decision yet.

WAIVED
  Reviewer intends this issue to be waived.

APPROVED / APPROVED_WAIVE
  Also treated as waiver-approved by the generator.
```

## How Review Memory Is Preserved

Each parsed issue gets a stable `issue_id` generated from fields such as:

```text
Tag
Goal
Module
FileName
LineNumber
Hierarchy
Signal / VariableName / ModPortName
Statement
```

When reports are parsed again, the script compares new issues with the existing `outputs/lint_review.xlsx`.

```text
same issue_id
  -> keep previous reviewer fields
  -> update report-derived fields

similar issue key
  -> keep previous reviewer fields
  -> mark as CHANGED

new issue_id
  -> mark as NEW

old issue_id not found in current reports
  -> mark as REMOVED
```

Reviewer fields preserved during merge:

```text
review_status
review_comment
```

## Waiver Generation

The script generates `vc_waiver.tcl` from rows where:

```text
record_status != REMOVED
Judgment in WAIVED / APPROVED / APPROVED_WAIVE
```

The generated Tcl follows the GUI-style waiver file:

```text
one waive_violation command per waived issue
waiver name = <Tag>_<Lint violation number>
```

Example:

```text
Tag: DeadCode-ML
Violation: Lint:739
```

generates:

```tcl
waive_violation -add {DeadCode-ML_739} ...
```

The script does not collapse multiple waived issues that share the same generated waiver name, because the GUI-style `vc_waiver.tcl` keeps issue-level waiver commands.

The script derives Tcl `-filter` fields from the report-owned columns. User-added columns are not used.

Exact match:

```json
{"Module":"AXICRYPT_RCC"}
```

generates:

```tcl
(Module == "AXICRYPT_RCC")
```

Wildcard match:

```json
{"Module":"AXICRYPT_*"}
```

generates:

```tcl
(Module =~ "AXICRYPT_*")
```

## Redundant Waiver Review

`outputs/waiver_rule_audit.csv` compares the current `vc_waiver.tcl` with the latest waived report.

```text
ACTIVE
  Waiver rule is referenced by at least one current waived issue.

REDUNDANT
  Waiver rule exists in vc_waiver.tcl but is not referenced by current waived issues.

MISSING_IN_TCL
  Waived report references a waiver name not found in vc_waiver.tcl.
```

## Cleaning Generated Files

To delete files generated by `sanity_lint_review.py`:

```powershell
python sanity_lint_review.py clean
```

Preview the cleanup without deleting:

```powershell
python sanity_lint_review.py clean --dry-run
```

Keep `vc_waiver.tcl` while deleting DB/Excel/report artifacts:

```powershell
python sanity_lint_review.py clean --keep-waiver
```

The clean command does not delete `reports/` and does not delete `vc_waiver.tcl_old`.
