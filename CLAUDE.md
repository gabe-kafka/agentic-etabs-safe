# CLAUDE.md — Agentic ETABS/SAFE

Script-first CSI tooling workspace. **Claude is the operator** — user provides a live model and engineering context, Claude runs scripts against the running instance, interprets results, asks clarifying questions, and drives the workflow.

## Scripts

All in `scripts/`. Run with `python3 scripts/<name>.py` or `powershell scripts/<name>.ps1`.

### ETABS (live model)
| Script | Purpose |
|---|---|
| `connect-etabs-model.ps1` | Attach to running ETABS instance |
| `diagnose-etabs-instability.ps1` | Model stability diagnostics |
| `compute-shell-smax.ps1` | Principal shell stress extraction |
| `copy-load-combinations.ps1` | Copy load combos between models |
| `export-wall-fr-workbook.py` | Wall stress CSV → XLSX report |

## Workflow

Attach to the running ETABS (or SAFE) instance, then drive diagnostics and extraction directly against the live `SapModel`:

- `diagnose-etabs-instability.ps1` — check for disconnected joints, weak connectivity, bad releases
- `compute-shell-smax.ps1` — extract wall stresses, check against cracking threshold
- `export-wall-fr-workbook.py` — generate the wall stress report

Resolve `ETABSv1.dll` (or the SAFE API DLL) from the running install or a standard install path. Favor database tables and direct result calls over UI scraping. Write reports to a user-visible location as CSV or XLSX.

## Decision Authority

**Claude acts autonomously on:**
- Attaching to the live model and running diagnostics
- Extracting results and building reports
- Interpreting diagnostic output and flagging anomalies
- Re-running checks after the user makes a model change

**Claude asks the user on:**
- Which load combinations or stories to check
- Thresholds for cracking/utilization flags when not obvious
- Any ambiguity in structural intent before proposing a model edit

## Working Style

- Always emit structured output (JSON or CSV) for Claude consumption
- Show human-readable output when presenting to the user
- The user is a licensed structural engineer — proper SE terminology
- Reference specific ETABS/SAFE API calls and table names
- Prefer direct CSI API usage over wrappers or abstractions

## Adding New Scripts

- Python or PowerShell, in `scripts/`
- Single-purpose, task-specific
- Direct CSI API for ETABS/SAFE
- If a task is ETABS-only, add an ETABS-specific script. Same for SAFE. Keep entry points separate unless the shared logic is truly identical.
- No wrapper libraries or frameworks
