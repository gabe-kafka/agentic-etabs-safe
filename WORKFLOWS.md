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
- `info.csv`, `raw-station-results.csv`, `story-envelope.csv`, `warnings.csv`: supporting audit files, unless `-NoCsv` is used.

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

Checks before using the data:

- Verify the `Info` tab matches the intended ETABS model path.
- Verify the model save time is newer than any Excel workbook you are comparing against.
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
