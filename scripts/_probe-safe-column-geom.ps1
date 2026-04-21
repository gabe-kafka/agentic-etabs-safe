param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
if (-not $proc) { throw "No SAFE running." }
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3) # kip-in

# Pull the punching table → get the 22 column point IDs
$v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
$null = $sm.DatabaseTables.GetTableForDisplayArray("Concrete Slab Design - Punching Shear Data", [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
$nCols = $f.Length
Write-Host "Fields: $nCols; Rows: $nr"
$pointIdx = [array]::IndexOf($f,"Point")

$punchPoints = @()
for ($i=0; $i -lt $nr; $i++) {
    $punchPoints += $d[$i*$nCols + $pointIdx]
}
Write-Host "Punching points: $($punchPoints -join ',')"

# For the first few, probe every API surface that might expose column dimensions
$firstPt = $punchPoints[0]
Write-Host ""
Write-Host "=== Probing point '$firstPt' for column geometry ==="

# 1. PointObj surface
$methods = ($sm.PointObj | Get-Member -MemberType Method).Name | Where-Object { $_ -notmatch "^(get_|set_|GetType|ToString|Equals|GetHashCode)" }
Write-Host ""
Write-Host "PointObj methods containing 'Col' or 'Support':"
$methods | Where-Object { $_ -match "(?i)(column|support|sectioncut|punch)" } | ForEach-Object { Write-Host "  $_" }

# 2. Check if PointObj has column-section assignment
try {
    $name=""
    $null = $sm.PointObj.GetColumnProp($firstPt, [ref]$name)
    Write-Host "GetColumnProp: '$name'"
} catch { Write-Host "GetColumnProp: $($_.Exception.Message)" }

# 3. Walk whole SapModel for 'Column' methods
$smMethods = ($sm | Get-Member -MemberType Property).Name
Write-Host ""
Write-Host "SapModel top-level properties with 'Col','Punch','Frame':"
$smMethods | Where-Object { $_ -match "(?i)(column|punch|frame|propcol)" } | ForEach-Object { Write-Host "  $_" }

# 4. Punching Shear overrides?
Write-Host ""
Write-Host "=== Probing SapModel.PointObj.GetLoadPoint ==="
try {
    $type=0;$ortho=0;$cx=0.0;$cy=0.0;$ang=0.0;$csys=""
    # Many variants of column prop retrieval exist in SAFE API; try them:
    $null = $sm.PointObj.GetColumn($firstPt, [ref]$type, [ref]$cx, [ref]$cy, [ref]$ang, [ref]$ortho, [ref]$csys)
    Write-Host "GetColumn: type=$type cx=$cx cy=$cy ang=$ang ortho=$ortho csys='$csys'"
} catch { Write-Host "GetColumn: $($_.Exception.Message)" }

# 5. PropPointSupport -> column sizes
try {
    $propName=""
    $null = $sm.PointObj.GetSupport($firstPt, [ref]$propName)
    Write-Host "GetSupport prop: '$propName'"
    if ($propName) {
        $a=0.0;$b=0.0;$type=0;$col=0;$cx=0.0;$cy=0.0;$ang=0.0
        try {
            $null = $sm.PropPointSupport.GetPointSupport($propName, [ref]$type, [ref]$col, [ref]$cx, [ref]$cy, [ref]$ang)
            Write-Host "PropPointSupport.GetPointSupport: type=$type col=$col cx=$cx cy=$cy ang=$ang"
        } catch { Write-Host "PropPointSupport.GetPointSupport err: $($_.Exception.Message)" }
    }
} catch { Write-Host "GetSupport err: $($_.Exception.Message)" }

# 6. Try the "Point Supports - Column" database table
foreach ($tbl in @(
    "Point Supports",
    "Point Supports - Column",
    "Point Supports - Column/Point Support Stiffness",
    "Column Property Definitions",
    "Objects and Elements - Joints",
    "Joint Assignments - Supports",
    "Joint Assignments - Column Support",
    "Point Object Loads - Applied",
    "Structure Layout - Point Coordinates",
    "Point Coordinates"
)) {
    try {
        $fv=0;$fl=[string[]]::new(0);$nrl=0;$dl=[string[]]::new(0)
        $null = $sm.DatabaseTables.GetTableForDisplayArray($tbl, [ref]$fl, "", [ref]$fv, [ref]$fl, [ref]$nrl, [ref]$dl)
        Write-Host ("Table '{0}': rows={1} fields={2}" -f $tbl, $nrl, ($fl -join ","))
        if ($nrl -gt 0 -and $nrl -le 3) {
            Write-Host ("  first row: {0}" -f (($dl[0..([Math]::Min($fl.Length-1,$dl.Length-1))]) -join " | "))
        }
    } catch {
        # silently skip tables that don't exist
    }
}
