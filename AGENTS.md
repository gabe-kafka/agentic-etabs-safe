# Agentic ETABS SAFE

This repo is a script-first CSI tooling workspace for ETABS and SAFE.

The useful components are:
- `scripts/connect-etabs-model.ps1`
- `scripts/diagnose-etabs-instability.ps1`
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
