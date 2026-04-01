# CLAUDE.md — Agentic ETABS/SAFE

This repo is a script-first CSI tooling workspace. **Claude is the primary operator** — the user provides DXF files and engineering context, Claude runs the tools, interprets the geometry, asks clarifying questions, and drives the pipeline to a clean ETABS-ready output.

## Scripts

All scripts live in `scripts/`. Run with `python3 scripts/<name>.py`.

| Script | Purpose |
|---|---|
| `validate-dxf.py` | DXF preprocessor: inspect, validate, classify, clean, floors |
| `connect-etabs-model.ps1` | Attach to running ETABS instance |
| `diagnose-etabs-instability.ps1` | Model stability diagnostics |
| `compute-shell-smax.ps1` | Principal shell stress extraction |
| `copy-load-combinations.ps1` | Copy load combos between models |
| `export-wall-fr-workbook.py` | Wall stress CSV → XLSX report |

## DXF → ETABS Agentic Workflow

When the user provides a DXF file, run this pipeline. The key principle: **nothing is assumed — Claude reads the file, proposes interpretations, and confirms with the user before acting.**

### Phase 1: Discover

Run inspect and validate in sequence. Do not ask permission between them.

```
python3 scripts/validate-dxf.py inspect <file.dxf> --json
python3 scripts/validate-dxf.py validate <file.dxf> --json
python3 scripts/validate-dxf.py floors <file.dxf> --json
```

From the inspect output, Claude should understand:
- What layers exist and what entity types are on each
- What blocks are defined and where they're inserted
- What text labels exist and their spatial arrangement
- The coordinate extents and overall drawing layout

### Phase 2: Classify (interactive — ask the user)

This is the critical agentic step. The DXF has layers and blocks, but Claude doesn't know what they mean structurally until it asks. Present the user with a clear proposal and ask them to confirm or correct.

**Layer classification** — propose a mapping:
```
I found these layers in your DXF. Here's my best guess at what each one
maps to in ETABS. Correct anything that's wrong:

  ETABS-WALL    → Shell walls (area elements)
  ETABS-BEAM    → Frame beams (line elements)
  MASTER COL    → Frame columns (point elements)
  ETABS-SLAB    → Floor/slab areas
  0             → Annotation / skip on import
```

**Block classification** — for each block, propose what it represents:
```
These blocks contain structural geometry. What is each one?

  MAIN SW         → Shear wall layout (core walls, all floors?)
  SHEAR WALL 2    → Shear wall variant (which floors?)
  master cols     → Column layout (which floors?)
  COL7            → Column layout variant (which floors?)
  sw2, 3rd, 4th   → Beam/wall layouts by floor range?
```

**Floor-to-block mapping** — the key question for vertical stacking:
```
Your file has 18 floor labels. Some blocks appear multiple times at
different Y positions (floor bands). Which blocks apply to which floors?

For example, 'SHEAR WALL 2' is inserted 3x at Y=8873, Y=23873, Y=28873.
Do those correspond to floors 1ST, 4TH, and 5TH?
```

**Do not proceed past Phase 2 without user confirmation.** The user's answers define the structural intent — getting this wrong means a wrong ETABS model.

### Phase 3: Clean

Run fixes based on what was found in Phase 1. These are safe to run without asking:
- Explode blocks (ETABS can't read them)
- Flatten stray Z coordinates
- Remove exact duplicate entities

Ask before:
- Closing open polylines (user must confirm which ones should be areas vs. lines)
- Deleting entities on Layer 0 (could be intentional reference geometry)

```
python3 scripts/validate-dxf.py clean <input.dxf> <output.dxf> --explode-blocks --remove-dupes --flatten-z
python3 scripts/validate-dxf.py validate <output.dxf> --json
```

Always re-validate after cleaning.

### Phase 4: Summarize & Recommend

After the full pipeline, give the user an actionable summary:

1. **Floor schedule** — all floors detected, their labels, which geometry applies to each
2. **Layer → ETABS mapping** — confirmed classifications for the ETABS import dialog
3. **Block → element mapping** — what each block becomes in ETABS (wall, column, beam)
4. **Issues fixed** — what was cleaned and what still needs attention
5. **ETABS import instructions** — specific dialog settings:
   - Which import mode (Floor Plan vs 3D Model)
   - Layer-to-element mapping for the import dialog
   - Maximum distance between parallel lines (for wall centerline generation)
   - Story assignments
   - Recommended snap tolerance and merge distance

### Phase 5: Per-floor export (optional, on request)

If the user wants individual DXF files per floor for sequential ETABS import:
```
python3 scripts/validate-dxf.py split <file.dxf> <output_dir/> --json
```

## Decision Authority

**Claude acts autonomously on:**
- Running inspect/validate/floors (read-only operations)
- Interpreting geometry and proposing classifications
- Running safe clean operations (explode blocks, flatten Z, remove dupes)
- Re-validating after changes

**Claude asks the user on:**
- Layer classification — what is structural vs. annotation vs. skip
- Block interpretation — what each block represents structurally
- Floor-to-geometry mapping — which elements belong to which floors
- Closing polylines — user must confirm intent
- Any ambiguity in structural intent
- Whether to delete vs. relocate Layer 0 geometry

## Working Style

- Always use `--json` when running scripts for Claude consumption
- Show human-readable output when presenting results to the user
- The user is a licensed structural engineer — use proper SE terminology
- Reference specific ETABS dialog settings and import options
- When multiple interpretations are possible, present the most likely one first with alternatives
- Run the full discover pipeline (inspect + validate + floors) in one pass before asking questions
- Group all classification questions together — don't drip-feed them one at a time

## Adding New Scripts

- Python or PowerShell, in `scripts/`
- Single-purpose, task-specific
- Always support `--json` for structured output
- No wrapper libraries or frameworks
- Direct CSI API calls for ETABS/SAFE scripts
- `ezdxf` is the only DXF dependency
