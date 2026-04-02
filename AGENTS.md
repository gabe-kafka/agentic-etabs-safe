# Agentic ETABS SAFE

This repo is a script-first CSI tooling workspace for ETABS and SAFE.

The useful components are:
- `scripts/connect-etabs-model.ps1`
- `scripts/diagnose-etabs-instability.ps1`
- `scripts/diagnose-etabs-meshing.ps1`
- `scripts/clear-etabs-geometry-markers.ps1`
- `scripts/copy-load-combinations.ps1`
- `scripts/compute-shell-smax.ps1`
- `scripts/export-wall-fr-workbook.py`
- `scripts/dxf-to-etabs.py` — DXF preprocessor for ETABS import (inspect, validate, classify, clean, floors, split)
- `typical-sizes.md` for modeling context only

Do not treat this workspace as a model-builder framework or workflow-capture repo.
Do not recreate `sessions/`, `templates/`, `PLAN.md`, or `README.md` unless the user explicitly asks.

## Working style

- Prefer direct PowerShell or Python scripts in `scripts/`.
- Prefer direct CSI API usage over wrappers or abstractions.
- Keep scripts small, task-specific, and usable against a live model.
- Prefer ETABS/SAFE database tables and direct result calls over UI scraping.
- Export plain CSV or XLSX outputs when the user needs reports.

## ETABS usage

- Attach to the running ETABS instance when possible.
- Resolve `ETABSv1.dll` from the running ETABS folder or a standard ETABS install path.
- Work directly against the live `SapModel`.
- Favor additive utilities: diagnostics, extraction, reporting, and narrowly scoped model edits.

## ETABS meshing debug workflow

- Use `scripts/diagnose-etabs-meshing.ps1` for repeated shell meshing/debug-geometry work.
- Run the script after ETABS analysis has fully finished and the model is idle again. Do not call marker mode while ETABS is actively solving.
- Use read-only mode first when you only need findings.
- Use `-MarkInModel` only when you want temporary visual markers added to the live model.
- Use `-ArrowMarkers` when you want large temporary debug arrow frames pointing at the suspect locations.
- If the model is locked after analysis, use `-UnlockIfLocked` when running marker mode. This unlocks the ETABS model so temporary debug markers can be added.
- Prefer passing the actual ETABS warning text with `-WarningText` or `-WarningTextPath` so the script can focus on the named problem areas first.
- Prefer `-OnlyWarningTargets` when analysis warnings already identify the suspect slabs/walls.
- The marker workflow adds temporary ETABS special points and debug groups only. It should not add structural elements, but it does modify the model state and should be treated as a temporary debug overlay.
- The arrow workflow adds temporary debug frame objects, fixed debug joints, and dedicated debug material/section properties. Treat it as a temporary visual overlay and clear it before rerunning analysis.
- Clear all temporary markers with `scripts/clear-etabs-geometry-markers.ps1` after review or before handing the model back cleanly.

Recommended sequence:
1. Run ETABS analysis.
2. Wait for the analysis monitor to close and ETABS to become idle.
3. Capture the meshing warning text.
4. Run `scripts/diagnose-etabs-meshing.ps1` with the warning text in read-only mode if you only need the diagnosis.
5. Run `scripts/diagnose-etabs-meshing.ps1 -MarkInModel -OnlyWarningTargets -UnlockIfLocked` if you want temporary in-model point markers after analysis.
6. Run `scripts/diagnose-etabs-meshing.ps1 -ArrowMarkers -OnlyWarningTargets -UnlockIfLocked` if you want large temporary arrow-frame overlays instead.
7. Inspect and fix the flagged slab/wall geometry.
8. Run `scripts/clear-etabs-geometry-markers.ps1`.
9. Rerun ETABS analysis.

## SAFE usage

- Use the same script-first approach for SAFE.
- Create SAFE-specific scripts in `scripts/` instead of trying to force ETABS scripts to do both jobs.
- Resolve the SAFE API DLL from the running SAFE install or a standard SAFE install path.
- Attach to a live SAFE model when possible, then query tables/results directly.
- Keep SAFE and ETABS entry points separate unless the shared logic is truly identical.

## Expectations for future changes

- If a new task is ETABS-only, add or update an ETABS-specific script.
- If a new task is SAFE-only, add a separate SAFE-specific script.
- If a report is needed, prefer generating an output file in a user-visible location.
- Do not introduce a new wrapper library unless the user explicitly asks for one.
