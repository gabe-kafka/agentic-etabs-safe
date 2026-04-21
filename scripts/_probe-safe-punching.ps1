param(
    [int]$SafePid,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Attach ---
$process = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
if (-not $process) { throw "No running SAFE process." }

$dll = Join-Path -Path (Split-Path -Parent $process.Path) -ChildPath "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll

$helper = [SAFEv1.cHelper](New-Object SAFEv1.Helper)
$api = $helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", $process.Id)
$sm  = $api.SapModel

# --- Units: set to kip-in ---
# 1=lb-in, 3=kip-in, 4=kip-ft, ...
$null = $sm.SetPresentUnits(3)

$report = [ordered]@{}

# --- Model file / units ---
$report.ModelFile = $sm.GetModelFilename($true)
$report.Units = $sm.GetPresentUnits()
$report.IsLocked = $sm.GetModelIsLocked()

# --- Point count (columns are point objects with supports/reactions) ---
$np = 0
$names = [string[]]::new(0)
$r = $sm.PointObj.GetNameList([ref]$np, [ref]$names)
$report.PointCount = $np

# --- Area objects (slabs) ---
$na = 0
$anames = [string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
$report.AreaObjectCount = $na

# --- Frame objects (columns/beams modeled as frames) ---
$nf = 0
$fnames = [string[]]::new(0)
try {
    $null = $sm.FrameObj.GetNameList([ref]$nf, [ref]$fnames)
    $report.FrameObjectCount = $nf
} catch {
    $report.FrameObjectCount = $null
    $report.FrameObjErr = $_.Exception.Message
}

# --- Column / frame section props (we want c1,c2 for rect columns) ---
# SAFE represents columns as point objects with supports (point supports).
# Column geometric sizes often live on a column section definition attached to point.
# Try to pull list of column supports (SupportSpring/Restraint)

# Count points with restraints or springs
$pointsWithSupport = 0
$pointsWithSpring = 0
for ($i = 0; $i -lt [Math]::Min($np, 5000); $i++) {
    $n = $names[$i]
    $restraint = [bool[]]::new(6)
    $null = $sm.PointObj.GetRestraint($n, [ref]$restraint)
    if ($restraint -contains $true) { $pointsWithSupport++ }
    try {
        $springType = 0
        $null = $sm.PointObj.GetSpringAssignment($n, [ref]$springType)
        if ($springType -gt 0) { $pointsWithSpring++ }
    } catch {}
}
$report.PointsScanned = [Math]::Min($np, 5000)
$report.PointsWithRestraint = $pointsWithSupport
$report.PointsWithSpring = $pointsWithSpring

# --- Load patterns / combos ---
$nlp = 0
$lpNames = [string[]]::new(0)
$null = $sm.LoadPatterns.GetNameList([ref]$nlp, [ref]$lpNames)
$report.LoadPatternCount = $nlp
$report.LoadPatterns = @($lpNames)

$nc = 0
$cNames = [string[]]::new(0)
$null = $sm.RespCombo.GetNameList([ref]$nc, [ref]$cNames)
$report.ComboCount = $nc
$report.Combos = @($cNames)

# --- Concrete design (SAFE slab design; punching shear results) ---
# Common SAFE design results tables:
#   "Concrete Slab Design 1 - Punching Shear Data"
#   "Concrete Slab Design 2 - Punching Shear Perimeter Data"
# Check if design has been run. Try GetSummaryResults from design.
$designSupported = $true
$punchSummary = $null
try {
    # SAFE's design API nests under SapModel.DesignConcreteSlab - confirm availability
    $dcs = $sm.DesignConcreteSlab
    $report.HasDesignConcreteSlab = $null -ne $dcs
    $codes = ""
    try {
        $null = $dcs.GetCode([ref]$codes)
        $report.DesignCode = $codes
    } catch {
        $report.DesignCodeErr = $_.Exception.Message
    }
} catch {
    $report.HasDesignConcreteSlab = $false
    $report.DesignAccessErr = $_.Exception.Message
}

# --- Try pulling the punching shear design table via DatabaseTables ---
$tableName = "Concrete Slab Design - Punching Shear Data"
$version = 0
$fields = [string[]]::new(0)
$nRows = 0
$data = [string[]]::new(0)
try {
    $null = $sm.DatabaseTables.GetTableForDisplayArray($tableName, [ref]$fields, "", [ref]$version, [ref]$fields, [ref]$nRows, [ref]$data)
    $report.PunchingTableRows = $nRows
    $report.PunchingTableFields = @($fields)
    $report.PunchingTableSample = if ($data.Length -gt 0) { @($data[0..([Math]::Min(39,$data.Length-1))]) } else { @() }
} catch {
    $report.PunchingTableErr = $_.Exception.Message
}

# Try alternate name variants
foreach ($alt in @(
    "Concrete Slab Design 1 - Punching Shear Data",
    "Punching Shear Design Data",
    "Concrete Design - Punching Shear",
    "Punching Shear Data 01 - Stations Perimeter Geometry",
    "Design Forces - Concrete Punching"
)) {
    try {
        $v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
        $null = $sm.DatabaseTables.GetTableForDisplayArray($alt, [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
        $report["Alt_$alt"] = "rows=$nr fields=$($f.Length)"
    } catch {}
}

# --- List ALL available database tables (so we can find punching tables) ---
try {
    $allTables = [string[]]::new(0)
    $nAll = 0
    $null = $sm.DatabaseTables.GetAllTables([ref]$nAll, [ref]$allTables)
    $matched = $allTables | Where-Object { $_ -match "(?i)punch|slab.*design|unbalanced|transfer" }
    $report.AllPunchingRelatedTables = @($matched)
    $report.TotalTableCount = $nAll
} catch {
    $report.AllTablesErr = $_.Exception.Message
}

if ($AsJson) {
    $report | ConvertTo-Json -Depth 6
} else {
    $report | Format-List
}
