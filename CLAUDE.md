# CLAUDE.md — Agentic ETABS/SAFE

Script-first CSI tooling workspace. **Claude is the operator** — user provides files and engineering context, Claude runs tools, interprets results, asks clarifying questions, and drives the workflow.

## Scripts

All in `scripts/`. Run with `python3 scripts/<name>.py` or `powershell scripts/<name>.ps1`.

### DXF Preprocessing (pre-import)
| Command | Purpose |
|---|---|
| `dxf-to-etabs.py inspect` | Full DXF survey — layers, blocks, text, extents |
| `dxf-to-etabs.py validate` | 8 geometry checks (dupes, stray Z, open polylines, etc.) |
| `dxf-to-etabs.py classify` | Propose layer/block → ETABS element mappings |
| `dxf-to-etabs.py floors` | Detect floor plans via text labels or spatial clustering |
| `dxf-to-etabs.py align` | Column vertical alignment check across floors, cloud discontinuities |
| `dxf-to-etabs.py clean` | Fix geometry (explode blocks, flatten Z, remove dupes) |
| `dxf-to-etabs.py split` | Export per-floor DXF files |

### ETABS (post-import, live model)
| Script | Purpose |
|---|---|
| `connect-etabs-model.ps1` | Attach to running ETABS instance |
| `diagnose-etabs-instability.ps1` | Model stability diagnostics |
| `compute-shell-smax.ps1` | Principal shell stress extraction |
| `copy-load-combinations.ps1` | Copy load combos between models |
| `export-wall-fr-workbook.py` | Wall stress CSV → XLSX report |

## Workflow

Two phases. DXF debug happens here, then ETABS debug happens in ETABS.

### Phase 1: DXF Debug (before ETABS import)

When the user provides a DXF, run discover automatically:

```
python3 scripts/dxf-to-etabs.py inspect <file.dxf> --json
python3 scripts/dxf-to-etabs.py validate <file.dxf> --json
python3 scripts/dxf-to-etabs.py floors <file.dxf> --json
python3 scripts/dxf-to-etabs.py classify <file.dxf> --json
```

Then present findings and **ask the user**:

1. **Layer classification** — "I found these layers. Here's what I think each one maps to in ETABS. Correct anything wrong."
2. **Block classification** — "These blocks contain structural geometry. What does each one represent?"
3. **Floor-to-block mapping** — "Which blocks apply to which floors?"
4. **Column alignment** — run `align` to flag transfer conditions. Show the user where columns drop off between floors.

After user confirms, run clean:
```
python3 scripts/dxf-to-etabs.py clean <input.dxf> <output.dxf> --explode-blocks --remove-dupes --flatten-z
python3 scripts/dxf-to-etabs.py validate <output.dxf> --json
```

Deliver the cleaned DXF with a summary:
- Floor schedule
- Layer → ETABS element mapping for the import dialog
- Column alignment issues (transfer beam locations)
- Specific ETABS import dialog settings (layer mapping, centerline tolerance, story assignments)

**The cleaned DXF goes into ETABS manually via File > Import.** No intermediary steps.

### Phase 2: ETABS Debug (after import, live model)

Once the model is in ETABS, use the PowerShell scripts against the live instance:
- `diagnose-etabs-instability.ps1` — check for disconnected joints, weak connectivity, bad releases
- `compute-shell-smax.ps1` — extract wall stresses, check against cracking threshold
- `export-wall-fr-workbook.py` — generate the wall stress report

## Decision Authority

**Claude acts autonomously on:**
- Running inspect/validate/floors/classify/align (all read-only)
- Interpreting geometry and proposing classifications
- Running safe clean operations (explode blocks, flatten Z, remove dupes)
- Re-validating after changes

**Claude asks the user on:**
- Layer and block classification — what is structural vs. annotation
- Floor-to-geometry mapping
- Closing polylines — user confirms which should be area elements
- Column alignment judgment — where to add transfer beams
- Any ambiguity in structural intent

## Working Style

- Always use `--json` when running scripts for Claude consumption
- Show human-readable output when presenting to the user
- The user is a licensed structural engineer — proper SE terminology
- Reference specific ETABS dialog settings and import options
- Run the full discover pipeline in one pass before asking questions
- Group all classification questions together

## Adding New Scripts

- Python or PowerShell, in `scripts/`
- Single-purpose, task-specific
- Always support `--json` for structured output
- No wrapper libraries or frameworks
- `ezdxf` for DXF, direct CSI API for ETABS/SAFE
