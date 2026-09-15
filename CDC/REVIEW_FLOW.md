# CDC Sanity Review Flow

This directory provides an Excel-based CDC review and waiver workflow modeled after the LINT flow.

## Files

```text
CDC/
  report_cdc.full.xlsx
  reports/
    report_cdc.full.log
    report_cdc.waived.log
  outputs/
    cdc_review.xlsx
    cdc_summary.csv
    waiver_rule_audit.csv
  sanity_cdc_review.py
  vc_waiver.tcl_ori
  vc_waiver.tcl
```

`report_cdc.full.xlsx` is the current issue source. `vc_waiver.tcl_ori` is a read-only migration/reference input. `vc_waiver.tcl` is generated from the reviewed workbook.

The script is standalone and uses only the Python standard library.

## Review Model

Issue lifecycle and human review are independent:

```text
record_status: NEW / CHANGED / UNCHANGED / REMOVED
Owner Action: UNREVIEWED / FIXED / WAIVED
Reviewer Decision: PENDING / APPROVED / DISAPPROVED
```

IP owners edit `IP Owner`, `Owner Action`, `Owner Comment`, `Filter Mode`, `Filter Fields`, and `Custom Filter`. Reviewers edit `Reviewer`, `Reviewer Decision`, and `Reviewer Comment`.

Only non-removed rows with `Owner Action = WAIVED` are written to `vc_waiver.tcl`. Reviewer Decision records peer review and does not control generation.

## Filter Modes

- `AUTO` progressively selects report fields until the issue is unique among current issues with the same tag. It excludes `FileName`, `LineNumber`, `Tag`, `Description`, and `Violation`. CDC object fields such as `ConvergencePoint`, `SrcObject`, `DestObject`, `NetName`, and `SeqElement` are supported.
- `FIELDS` uses a comma-separated list such as `Module, ConvergencePoint`. A bare field uses the report value. `Field=value` overrides it. Values containing `*` or `?` use `=~`; other values use `==`.
- `CUSTOM` emits the complete expression from `Custom Filter` unchanged.

Generation fails rather than writing an unfiltered waiver when the selected mode cannot produce a valid filter.

Existing CDC waivers from `vc_waiver.tcl_ori` are matched by tag and the numeric violation identifier in the waiver name, such as `CDC_COHERENCY_RECONV_SEQ_178` and `CDC:178`. Matching entries are imported as `Owner Action = WAIVED` with `Filter Mode = CUSTOM`, preserving their original comment, filter, user, and timestamp.

The shared reference file also contains non-CDC entries. Migration and audit consider only `waive_violation` commands with `-app { cdc }`. Generated output contains CDC commands only and uses this format:

```tcl
waive_violation -add {<name>}  -comment {<reason>} -filter {<expression>} -app { cdc } -tag { <tag> } -user { <user> } -timestamp { <DD-MM-YYYY HH:MM:SS> }
```

## Normal Loop

Run commands from the `CDC` directory.

1. Refresh `report_cdc.full.xlsx` with the existing report-generation flow.

2. Merge the current report into the review workbook:

```powershell
python sanity_cdc_review.py merge_excel
```

This writes `outputs/cdc_review.xlsx`, preserves previous review decisions, and marks issues as `NEW`, `CHANGED`, `UNCHANGED`, or `REMOVED`.

3. Review `outputs/cdc_review.xlsx`. Green columns are editable, gray columns are report-owned, and pale-yellow columns are script-managed.

4. Generate the CDC waiver file:

```powershell
python sanity_cdc_review.py gen_waiver
```

5. Run CDC again using the new `vc_waiver.tcl`, refresh `report_cdc.full.xlsx`, and repeat `merge_excel`.

To override the migration source on the first merge:

```powershell
python sanity_cdc_review.py merge_excel --waiver-tcl <reference-waiver.tcl>
```

To set the user for newly generated waiver metadata:

```powershell
python sanity_cdc_review.py gen_waiver --user <username>
```

Keep `vc_waiver.tcl_ori` unchanged so it remains available as the original reference.
