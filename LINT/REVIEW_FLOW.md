# LINT Sanity Review Flow

This document describes the intended review loop for LINT sanity issues and waivers.

## Weakness Of Previous Procedure

- Review relied mainly on the `vc_waiver.tcl` file and a separate PPT file, instead of a single Excel source of truth.
- `vc_waiver.tcl` is a plain text file, so it cannot carry detailed review context such as screenshots or images of the waived issue.
- The PPT file did carry that detailed context, but it was not synced regularly with `vc_waiver.tcl`, so the two artifacts drifted apart over time.
- With review information split across a text file and a slide deck, review and peer-review were difficult: reviewers had to cross-check two disconnected files, and stale PPT content could hide the real reasoning behind a waiver.

## Proposed Solution

`outputs/lint_review.xlsx` becomes the single reviewer UI and review source of truth, replacing the split between `vc_waiver.tcl` and the PPT file:

- Being an Excel workbook rather than plain text, it can hold the waiver decision together with rich review context (comments, extra columns, images) in one place, which `vc_waiver.tcl` alone could not carry.
- `report_lint.full.xlsx`, the tool-generated LINT report, is merged into `lint_review.xlsx`, so reviewers and peer reviewers see decisions and their justification side by side instead of cross-checking a separate slide deck.
- `vc_waiver.tcl` is generated from issues that the IP owner marks as `WAIVED` in `lint_review.xlsx` and used by the next sanity check run, so it is always derived from the same up-to-date source instead of drifting out of sync with a manually maintained PPT.

## Supported Features

- `gen_waiver` generates `vc_waiver.tcl` from `outputs/lint_review.xlsx`.
- `merge_excel` runs `make -f Makefile excel` and merges the refreshed `report_lint.full.xlsx` into `outputs/lint_review.xlsx`.
- IP-owner fields are `IP Owner`, `Owner Action`, and `Owner Comment`.
- Reviewer fields are `Reviewer`, `Reviewer Decision`, and `Reviewer Comment`.
- Report-owned columns are refreshed from `report_lint.full.xlsx`.
- Header colors show ownership: green means editable, gray means report-owned, and pale yellow identifies script-managed metadata.
- Filters are enabled on each sheet. `Owner Action` and `Reviewer Decision` provide dropdown lists.
- Excel Notes on status headers describe their supported values.
- Waiver filters support `AUTO`, `FIELDS`, and `CUSTOM` modes.
- Issue change state is tracked with `record_status`.

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
python sanity_lint_review.py gen_waiver
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
python sanity_lint_review.py merge_excel
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

5. The IP owner handles issues in `outputs/lint_review.xlsx`.

The IP owner updates:

```text
IP Owner
Owner Action: UNREVIEWED / FIXED / WAIVED
Owner Comment
Filter Mode: AUTO / FIELDS / CUSTOM
Filter Fields
Custom Filter
```

`Owner Action = WAIVED` tells the script to generate a waiver for that issue. `Owner Action = FIXED` means the IP owner resolved the issue without a waiver.

The filter columns control how that issue is identified in `vc_waiver.tcl`:

- `AUTO` progressively selects report fields until the issue is unique among current issues with the same tag. It deliberately excludes `FileName` and `LineNumber`. If multiple issues have identical supported fields, they intentionally share one waiver rule.
- `FIELDS` uses a comma-separated list such as `Goal, Module, Signal`. A bare field uses the value in that row. `Module=axi_*` overrides the row value, and `*` or `?` selects the Tcl `=~` operator.
- `CUSTOM` uses the complete expression entered in `Custom Filter`.

The script reports an error instead of generating an unfiltered or invalid waiver when the selected mode cannot produce a valid filter.

6. The reviewer evaluates the IP owner's handling.

The reviewer updates:

```text
Reviewer
Reviewer Decision: PENDING / APPROVED / DISAPPROVED
Reviewer Comment
```

The reviewer decision records peer-review results only. It does not control waiver Tcl generation.

Green headers are user-editable. Gray headers are report-owned and should not be edited. Pale-yellow headers identify script-managed metadata and should not be edited manually. Each sheet has filters enabled.

The reviewer may also add new columns for human notes. Those columns are preserved in `lint_review.xlsx`, but ignored by `vc_waiver.tcl` generation.

7. Generate the next waiver file.

Run:

```powershell
python sanity_lint_review.py gen_waiver
```

This reads reviewer edits from `outputs/lint_review.xlsx`, then generates:

```text
vc_waiver.tcl
```

8. Run sanity check again with the new `vc_waiver.tcl`.

The sanity tool creates new reports.

9. Merge the refreshed report again.

```powershell
python sanity_lint_review.py merge_excel
```

The script merges the new reports with the previous `lint_review.xlsx`, so old comments/status are preserved.

## Status Meaning

`record_status` is the issue state based on the latest sanity reports.

```text
NEW
  The issue appears in the latest report and was not found in the previous review workbook.

CHANGED
  The issue no longer has the same exact issue_id, but it still looks like the same logical issue.
  Example: line number or statement changed, while tag/module/file/hierarchy/object stayed the same.
  Previous IP-owner and reviewer fields are preserved and the row is marked for manual confirmation.

UNCHANGED
  The issue still appears in the latest report and its identity-defining attributes have not changed.

REMOVED
  The issue existed in the previous review workbook but no longer appears in the latest reports.
```

`Owner Action` records the IP owner's handling decision.

```text
UNREVIEWED
  The IP owner has not handled the issue yet.

FIXED
  The IP owner fixed the issue without a waiver.

WAIVED
  The IP owner decided to waive the issue. This value enables waiver Tcl generation.
```

`Reviewer Decision` records the independent peer-review result.

```text
PENDING
  The handling has not been reviewed yet.

APPROVED
  The reviewer accepts the IP owner's handling.

DISAPPROVED
  The reviewer rejects the IP owner's handling.
```

`Reviewer Decision` does not affect waiver Tcl generation.

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
  -> keep previous IP-owner and reviewer fields
  -> update report-derived fields

similar issue key
  -> keep previous IP-owner and reviewer fields
  -> mark as CHANGED

new issue_id
  -> mark as NEW

old issue_id not found in current reports
  -> mark as REMOVED
```

Human-entered fields preserved during merge:

```text
IP Owner
Owner Action
Owner Comment
Reviewer
Reviewer Decision
Reviewer Comment
```

## Waiver Generation

The script generates `vc_waiver.tcl` from rows where:

```text
record_status != REMOVED
Owner Action = WAIVED
```

The reviewer decision is deliberately not part of this condition. Waiver Tcl is generated from the IP owner's decision, even when `Reviewer Decision` is `PENDING` or `DISAPPROVED`.

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

Every generated waiver contains a Tcl `-filter`. The IP owner controls it with `Filter Mode`, `Filter Fields`, and `Custom Filter`. User-added columns are not used.

Filter mode priority is explicit: the script uses only the mode selected in `Filter Mode`.

```text
AUTO
  Select report fields until one issue is identified, excluding FileName and LineNumber.
  Issues with identical supported fields share one rule.

FIELDS
  Use the listed report fields and optional user-supplied values or wildcards.

CUSTOM
  Use the user-defined Tcl filter expression unchanged.
```

Rows with the same tag, resulting filter expression, and owner comment are emitted as one `waive_violation` command. The first row is the primary issue. The primary and repeated rows use the same `waiver_name`, and Excel Notes on their `waiver_name` cells identify the shared rule relationship.

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

The `ACTIVE` value below belongs only to the waiver-rule audit. It is not a `record_status` value.

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
