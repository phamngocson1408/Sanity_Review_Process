# GCA Sanity Review Flow

This directory provides an Excel-based review and waiver workflow for Galaxy Constraint Analyzer (GCA), following the same ownership and review model used by the LINT flow.

## Files

```text
GCA/
  reports/
    <design>_gca.rpt
    <design>_gca_const_analysis.rpt
  outputs/
    gca_review.xlsx
  GCA_confirmed.tcl
  sanity_gca_review.py
```

The constant-analysis report is the current issue source. The rule report supplies severity and rule descriptions. `GCA_confirmed.tcl` supplies existing native GCA conditions and is regenerated from the reviewed workbook.

## Review Columns

IP owners edit:

```text
IP Owner
Owner Action: UNREVIEWED / FIXED / WAIVED
Owner Comment
Filter Mode: AUTO / FIELDS / CUSTOM
Filter Fields
Custom Filter
```

Reviewers edit:

```text
Reviewer
Reviewer Decision: PENDING / APPROVED / DISAPPROVED
Reviewer Comment
```

Issue lifecycle is tracked independently in `record_status`:

```text
NEW / CHANGED / UNCHANGED / REMOVED
```

Only non-removed rows with `Owner Action = WAIVED` generate `create_waiver` commands. Reviewer Decision records peer review but does not control generation.

## Filter Modes

GCA uses native `-condition [list ...]` clauses rather than the LINT `-filter` expression syntax.

- `AUTO` reuses the native condition from the existing `GCA_confirmed.tcl`. When one imported condition contains multiple objects, a simple port, pin, clock, net, or cell collection is narrowed to the current issue object so each Excel row can be managed independently. For a new issue, the script can infer simple port, pin, clock, or net conditions from the report message. Generation stops with an error if no safe condition is available.
- `FIELDS` uses comma-separated condition keys from the imported condition, such as `port`. A value can be overridden with `key=Tcl-expression`, for example `port=[get_ports {PREADY}]`.
- `CUSTOM` uses complete condition clauses entered in `Custom Filter`, for example `-condition [list port [get_ports {PREADY}]]`. Multiple `-condition` clauses are allowed.

The script never emits an unconditioned GCA waiver.

## Normal Loop

Run commands from the `GCA` directory.

1. After running GCA and refreshing both reports, merge current results into Excel:

```powershell
python sanity_gca_review.py merge_excel
```

This parses the reports, correlates existing conditions from `GCA_confirmed.tcl`, preserves review fields from the previous workbook, and writes `outputs/gca_review.xlsx`.

2. Review the workbook. Green columns are editable; gray columns come from the report; pale-yellow columns are script-managed metadata.

3. Generate the confirmed waiver file:

```powershell
python sanity_gca_review.py gen_waiver
```

This reads `outputs/gca_review.xlsx`, writes `GCA_confirmed.tcl`, and updates generated waiver metadata in the workbook.

4. Run GCA again with the new `GCA_confirmed.tcl`, then repeat `merge_excel`.

For different report names, override the paths explicitly:

```powershell
python sanity_gca_review.py merge_excel `
  --report reports/<design>_gca_const_analysis.rpt `
  --rules reports/<design>_gca.rpt
```

Use `--user` when generating to override the waiver user recorded for newly generated entries:

```powershell
python sanity_gca_review.py gen_waiver --user <username>
```

## Existing Waivers

Existing waiver commands are matched to current issues by rule and by objects referenced in their condition. Simple grouped object collections are split into issue-level AUTO conditions so changing one row cannot unintentionally waive the other objects from the old group. For conditions whose object cannot be recovered from the report text, the script uses positional matching only when the number of unmatched issues and unmatched waiver commands for that rule is identical. This preserves the six complex `EXC_0004` conditions in the supplied data.

Multiple Excel rows that resolve to the same design, scenario, rule, comment, and condition are emitted as one `create_waiver` command. This preserves grouped object conditions without duplicating identical Tcl commands.
