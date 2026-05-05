# Agentic ETABS / SAFE

Script-first CSI tooling workspace for ETABS and SAFE.

## ETABS Control Center

A local web UI for live model editing. Start it with:

```powershell
python scripts/serve-etabs-control-center.py --open-browser
```

Runs at `http://127.0.0.1:8765/` by default.

**Current panels:**
- **Story editor** — edit story elevations against a live model. Elevation-only edits use `Story.SetHeight`. Story count changes (add/delete) use the `DatabaseTables` API.
- **Geometry check** — scans the live model for near-coincident joints, off-story joints, and dangling frame ends.

## DatabaseTables API

More ETABS functionality is available through `SapModel.DatabaseTables` than through the standard object API. Several operations blocked by the high-level API (e.g. modifying story definitions on a populated model) can be performed by reading and writing the underlying database tables directly via `GetTableForDisplayArray` / `SetTableForEditingArray` / `ApplyEditedTables`.

## Scripts

| Script | Purpose |
|---|---|
| `connect-etabs-model.ps1` | Attach to a running ETABS instance |
| `get-etabs-stories.ps1` | Read current story data as JSON |
| `set-etabs-stories.ps1` | Write story definitions to a live model |
| `diagnose-etabs-instability.ps1` | Instability diagnostics with learning system |
| `diagnose-etabs-meshing.ps1` | Mesh debug overlays |
| `find-geometric-bugs-etabs.ps1` | Near-coincident joints, off-story joints, dangling frame ends |
| `clear-etabs-geometry-markers.ps1` | Remove temporary debug markers |
| `compute-shell-smax.ps1` | Shell principal stress extraction |
| `export-wall-fr-workbook.py` | Wall force resultant workbook |
| `connect-safe-model.ps1` | Attach to a running SAFE instance |
| `diagnose-safe-instability.ps1` | SAFE instability diagnostics |
| `delete-orphan-joints-safe.ps1` | Remove orphan joints from SAFE model |
