# Workflows

This file is for repeatable operator runbooks. Keep `AGENTS.md` focused on agent behavior and repo rules; keep this file focused on commands, checks, and handoff artifacts.

## ETABS Shear Wall Required Steel

Use this workflow when pulling required boundary steel from a live ETABS model for downstream Excel, SAFE, or schedule work.

1. Run ETABS analysis/design and wait until ETABS is idle.
2. Confirm the intended model is open in ETABS.
3. From this repo, run:

```powershell
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1
```

To guard against pulling from the wrong live model, pass the expected `.EDB` path:

```powershell
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -ExpectedModelPath "C:\path\to\model.EDB"
```

Useful filters:

```powershell
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -Story Story1
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -Pier W4,W6 -Story Story1,Story2
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -OutputDirectory .\out\etabs-shear-wall-required-steel\latest
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -OutputWorkbookPath "C:\path\to\shear-wall-design-required-steel.xlsx"
```

Outputs:

- `shear-wall-design-required-steel.xlsx`: deliverable workbook parallel to the shear-wall design workbook.
- `info.csv`, `raw-station-results.csv`, `story-envelope.csv`, `warnings.csv`, `qa-qc-alignment.csv`: supporting audit files, unless `-NoCsv` is used.

Workbook tabs:

- `SHEAR WALL TABLE OUTPUT`: final schedule-style design table.
- `Master Design Hierarchy`: required steel area hierarchy used for schedule selection.
- `As Master`: required steel by story, pier, and side.
- `Design Master`: required steel, selected design values, selected design IDs, and flipped story order table.
- `Info`: model path, save time, units, source API, row counts, output path, and envelope rule.
- `All Piers`: pier/story required steel envelope.
- Per-pier tabs such as `W4`: same envelope data filtered to one pier.
- `Sanity Checks`: row counts, warning count, hierarchy exceedance count, and missing wall length checks.
- `Raw Station Results`: direct ETABS Top/Bottom station rows from `SapModel.DesignShearWall.GetPierSummaryResults`.
- `Warnings`: ETABS warning/error rows from the live design summary.
- `QA_QC Alignment`: final alignment gate comparing ETABS station maxima, pier/story envelopes, and visible workbook output cells.

Dynamic workbook wiring:

- `Raw Station Results` is the static ETABS source pull.
- `All Piers` uses Excel formulas to envelope required steel, boundary zone lengths, and D/C ratio from `Raw Station Results`.
- `As Master` uses formulas from `All Piers`.
- `Design Master` uses formulas from `As Master` and `Master Design Hierarchy`.
- `SHEAR WALL TABLE OUTPUT` uses formulas from `Design Master`.
- Per-pier sheets use formulas from `All Piers` for required steel and downstream design checks.
- `Sanity Checks` uses formulas for row counts, warning counts, hierarchy exceedance count, and missing wall lengths.

QA/QC alignment:

- The script writes cached ETABS values into formula cells, immediately reopens the workbook, and adds `QA_QC Alignment`.
- The alignment gate compares live ETABS station maxima to the computed envelope, then compares the envelope to `All Piers`, `As Master`, and the per-pier design tabs.
- Default tolerance is `0.005 in^2`. Override only when there is a deliberate rounding requirement:

```powershell
.\scripts\export-etabs-shear-wall-required-steel-pipeline.ps1 -QaToleranceIn2 0.01
```

- If any alignment check fails, the workbook is still written for review, but the script exits with an error and the failures appear at the top of `QA_QC Alignment`.

Checks before using the data:

- Verify the `Info` tab matches the intended ETABS model path.
- Verify the model save time is newer than any Excel workbook you are comparing against.
- Verify `QA_QC Alignment` has no `Error` severity rows.
- Spot-check suspicious values in `Raw Station Results` before using `All Piers`.
- Treat open Excel workbooks as downstream artifacts, not source truth, unless their metadata matches the current pull.

## ETABS Meshing Debug

1. Run ETABS analysis and wait until ETABS is idle.
2. Capture the actual ETABS warning text.
3. Start read-only:

```powershell
.\scripts\diagnose-etabs-meshing.ps1 -WarningTextPath "C:\path\to\warning.txt" -OnlyWarningTargets
```

4. Add temporary point markers only when needed:

```powershell
.\scripts\diagnose-etabs-meshing.ps1 -WarningTextPath "C:\path\to\warning.txt" -OnlyWarningTargets -MarkInModel -UnlockIfLocked
```

5. Add arrow overlays only when a larger visual indicator is needed:

```powershell
.\scripts\diagnose-etabs-meshing.ps1 -WarningTextPath "C:\path\to\warning.txt" -OnlyWarningTargets -ArrowMarkers -UnlockIfLocked
```

6. Clear temporary geometry markers before rerunning analysis:

```powershell
.\scripts\clear-etabs-geometry-markers.ps1
```

## SAFE Instability Debug

Use SAFE-specific scripts instead of adapting ETABS scripts.

```powershell
.\scripts\diagnose-safe-instability.ps1
```

If temporary model edits or cleanup are needed, keep them narrow and use the SAFE-specific utilities in `scripts/`.

## AutoCAD Shear Wall Table Fill

Use the `auto-cad-shear-wall-table-fill` skill for AutoCAD shear wall schedule table work.

Hard requirement:

- The deliverable must be a real editable AutoCAD `TABLE` / `ACAD_TABLE` object.
- A visual replica made from loose `LINE`, `TEXT`, or `MTEXT` geometry is not acceptable.
- Use `C:\Users\gkafka\Documents\hello_filled.dxf` and the skill asset `assets/reference-filled-autocad-table.dxf` as the reference for object type/editability.
- Still verify row count, column count, headers, story labels, and reinforcement values against the ETABS workbook output.
- Prefer modifying/copying the template table so font, table style, borders, row heights, merged cells, and notes are preserved.
- Automate bond-zone filling only. Balance-zone cells can remain blank or as preserved template content unless explicitly requested.

Locked build method:

1. Use `C:\Users\gkafka\Documents\table-template.dxf` as the table source.
2. Open the template with AutoCAD COM and cast the modelspace `AcDbTable` entity to `IAcadTable`.
3. Do not draw a replacement table and do not explode the table.
4. Expand the real table columns to match the workbook bond-zone columns, with one `BALANCE` and one `HORIZONTAL REINF.` column preserved per wall group.
5. Fill only bond-zone body cells from workbook tab `SHEAR WALL TABLE OUTPUT`.
6. Leave balance-zone and horizontal reinforcement body cells blank unless explicitly requested.
7. Preserve `SHEAR WALL SCHEDULE` in the title row. Do not place long project names in the merged title row because AutoCAD can auto-grow the row and overlap the notes.
8. After all merges and fills, explicitly restore template-scale geometry:
   - row heights: title/header/body rows at the template heights
   - title/group text height: about `20.6689`
   - header/body text height: about `14.8816`
   - title/group style: `3-16 ANNOTATIVE`
   - header/body style: `1-8 ANNOTATIVE`
9. Run `GenerateLayout`, restore row and text heights again, then recompute the table block.
10. Move the notes block down only if the final table bounding box overlaps it.
11. Save both DWG and DXF outputs.

Current command pattern:

```powershell
python "$env:USERPROFILE\.codex\skills\auto-cad-shear-wall-table-fill\scripts\fill_template_autocad_table.py" `
  --template "C:\Users\gkafka\Documents\table-template.dxf" `
  --workbook "C:\path\to\shear-wall-design-required-steel.xlsx" `
  --output-dwg "C:\path\to\out\1025-atlantic-shear-wall-table-filled.dwg" `
  --output-dxf "C:\path\to\out\1025-atlantic-shear-wall-table-filled.dxf" `
  --summary "C:\path\to\out\summary.json"
```

Final QA gates:

- DXF contains one real `ACAD_TABLE` and no loose `LINE`/`TEXT` table replica geometry.
- AutoCAD COM reopen reports one editable `AcDbTable`.
- Cell editability smoke test passes by setting and restoring one body cell without saving the test edit.
- Workbook-to-table filled bond-zone values match exactly.
- Balance-zone filled count is zero unless explicitly requested.
- Reopened text heights are still model-space readable, not `0.125`/`0.18` paper-size leftovers.
- Notes block does not overlap the final table bounding box.

1025 Atlantic accepted benchmark:

- Output iteration: `1025-atlantic-template-fill-iteration-06`
- Rows/columns: `26 x 21`
- Bond-zone columns: `14`
- Workbook filled cell match: `256 / 256`
- Balance filled count: `0`
