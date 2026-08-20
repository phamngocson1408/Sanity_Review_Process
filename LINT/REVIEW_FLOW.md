# LINT Sanity Review Flow

This document describes the intended review loop for LINT sanity issues and waivers.

## Main Idea

`lint_review.xlsx` is the reviewer UI.

`data/lint_review_db.csv` is the Git-friendly source of truth that keeps review memory across sanity runs.

`vc_waiver.tcl` is generated from approved/waived items in the review DB and is used by the next sanity check run.

## Directory Layout

```text
LINT/
  reports/
    report_lint.full.log
    report_lint.waived.log
  data/
    lint_review_db.csv
  outputs/
    lint_review.xlsx
    lint_summary.csv
    waiver_rule_audit.csv
  sanity_lint_review.py
  vc_waiver.tcl
```

## Normal Loop

1. Run the sanity check script.

The sanity tool creates report files under `reports/`.

```text
reports/report_lint.full.log
reports/report_lint.waived.log
```

2. Run the review bridge.

From the `LINT` directory:

```powershell
python sanity_lint_review.py
```

The script does the following:

```text
existing outputs/lint_review.xlsx, if any
  -> import reviewer edits into data/lint_review_db.csv
  -> generate vc_waiver.tcl

current reports/*.log
  -> parse current issues
  -> merge with data/lint_review_db.csv
  -> export outputs/lint_review.xlsx
```

3. Reviewer edits `outputs/lint_review.xlsx`.

The reviewer can update these columns:

```text
Person in Charge
Date
Judgment
Comment
waiver_enabled
waiver_name
filter_json
```

The reviewer may also add images/comments for human review. The script preserves structured cell data, but embedded images are not used for waiver generation.

4. Generate the next waiver file.

Run:

```powershell
python sanity_lint_review.py
```

or explicitly:

```powershell
python sanity_lint_review.py generate-waiver-from-excel
```

This imports reviewer edits from `outputs/lint_review.xlsx` into `data/lint_review_db.csv`, then generates:

```text
vc_waiver.tcl
```

5. Run sanity check again with the new `vc_waiver.tcl`.

The sanity tool creates new reports.

6. Run the review bridge again.

```powershell
python sanity_lint_review.py
```

The script merges the new reports with the previous review DB, so old comments/status are preserved.

## Status Meaning

`record_status` is the issue state based on the latest sanity reports.

```text
NEW
  The issue appears in the latest report and was not found in the previous review DB.

ACTIVE
  The issue appears in the latest full report and already existed in the review DB.

CHANGED
  The issue no longer has the same exact issue_id, but it still looks like the same logical issue.
  Example: line number or statement changed, while tag/module/file/hierarchy/object stayed the same.
  Previous reviewer fields are preserved and the row is marked for manual confirmation.

WAIVED
  The issue appears in the latest waived report. This means the sanity tool recognized it as waived.

REMOVED
  The issue existed in the previous review DB but no longer appears in the latest reports.
```

`review_status` is the reviewer decision.

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

When reports are parsed again, the script compares new issues with the existing `data/lint_review_db.csv`.

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
waiver_enabled
waiver_name
review_comment
owner
review_date
filter_json
waiver_user
waiver_timestamp
```

## Waiver Generation

The script generates `vc_waiver.tcl` from rows where:

```text
record_status != REMOVED
waiver_enabled == yes
review_status in WAIVED / APPROVED / APPROVED_WAIVE
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

The script does not collapse multiple waived issues that share the same `Waiver Name` in `report_lint.waived.log`, because the GUI-style `vc_waiver.tcl` keeps issue-level waiver commands.

`filter_json` controls the generated Tcl `-filter`.

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
