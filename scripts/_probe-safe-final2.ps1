param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3)  # kip-in

# GetCoordCartesian with 5 args (csys trailing default)
Write-Host "=== Point coords, 5-arg signature ==="
$na=0;$anames=[string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
$nPts=0;$ptNames=[string[]]::new(0)
$null = $sm.AreaObj.GetPoints("1", [ref]$nPts, [ref]$ptNames)
foreach ($p in $ptNames) {
    $x=0.0;$y=0.0;$z=0.0
    try {
        $null = $sm.PointObj.GetCoordCartesian($p, [ref]$x, [ref]$y, [ref]$z, "Global")
        Write-Host ("  pt {0}: ({1}, {2}, {3})" -f $p, $x, $y, $z)
    } catch {
        Write-Host "  err: $($_.Exception.Message)"
        # Inspect signature
        $mi = $sm.PointObj.GetType().GetMethods() | Where-Object { $_.Name -eq "GetCoordCartesian" }
        foreach ($m in $mi) {
            $ps = $m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }
            Write-Host "    sig: $($ps -join ', ')"
        }
        break
    }
}

# PropMaterial.GetNameList with typed arrays
Write-Host ""
Write-Host "=== PropMaterial.GetNameList typed enum ==="
$nmat=0;$mnames=[string[]]::new(0);$mtypes=[SAFEv1.eMatType[]]::new(0)
try {
    $null = $sm.PropMaterial.GetNameList([ref]$nmat, [ref]$mnames, [ref]$mtypes)
    for ($i=0; $i -lt $nmat; $i++) { Write-Host ("  {0}: {1}" -f $mnames[$i], $mtypes[$i]) }
} catch { Write-Host "err: $($_.Exception.Message)" }

# PropMaterial.GetOConcrete — inspect signature
Write-Host ""
Write-Host "=== PropMaterial.GetOConcrete signatures ==="
$mi = $sm.PropMaterial.GetType().GetMethods() | Where-Object { $_.Name -match "GetOConcrete" }
foreach ($m in $mi) {
    $ps = $m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }
    Write-Host "  $($m.Name)($($ps -join ', '))"
}

# Combo method list
Write-Host ""
Write-Host "=== RespCombo methods ==="
($sm.RespCombo | Get-Member -MemberType Method | Where-Object { $_.Name -notmatch "^(get_|set_|GetType$|ToString|Equals|GetHashCode|CreateObjRef|InitializeLifetimeService|GetLifetimeService)" }).Name | Sort-Object

# LoadPatterns methods
Write-Host ""
Write-Host "=== LoadPatterns methods ==="
($sm.LoadPatterns | Get-Member -MemberType Method | Where-Object { $_.Name -match "^Get" }).Name | Sort-Object

# Load patterns: get DL multiplier
Write-Host ""
Write-Host "=== Load patterns ==="
$nlp=0;$lpnames=[string[]]::new(0)
$null = $sm.LoadPatterns.GetNameList([ref]$nlp, [ref]$lpnames)
foreach ($p in $lpnames) {
    $type=0;$selfMult=0.0
    try {
        $null = $sm.LoadPatterns.GetLoadType($p, [ref]$type)
    } catch {}
    try {
        $null = $sm.LoadPatterns.GetSelfWTMultiplier($p, [ref]$selfMult)
    } catch {}
    Write-Host ("  {0}: type={1} selfWtMult={2}" -f $p, $type, $selfMult)
}

# ---- Load hunt: try area 1's load by walking load patterns  ----
Write-Host ""
Write-Host "=== AreaObj.GetLoadUniform on area 1 per pattern ==="
$na=0;$anames=[string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
# try for all areas, by pattern
foreach ($a in $anames) {
    $numLoads=0
    $areaName=[string[]]::new(0); $loadPat=[string[]]::new(0); $csys=[string[]]::new(0)
    $dir=[int[]]::new(0); $val=[double[]]::new(0)
    try {
        $null = $sm.AreaObj.GetLoadUniform($a, [ref]$numLoads, [ref]$areaName, [ref]$loadPat, [ref]$csys, [ref]$dir, [ref]$val, 0)
        if ($numLoads -gt 0) {
            Write-Host ("area {0}: {1} loads" -f $a, $numLoads)
            for ($k=0; $k -lt $numLoads; $k++) {
                Write-Host ("  {0} dir={1} val={2} csys={3}" -f $loadPat[$k], $dir[$k], $val[$k], $csys[$k])
            }
        }
    } catch {}
}

# ItemType = group (any group containing all areas)?
Write-Host ""
Write-Host "=== GetLoadUniform via ItemType=1 (Group) ==="
$ng=0;$gnames=[string[]]::new(0)
try { $null = $sm.GroupDef.GetNameList([ref]$ng, [ref]$gnames) } catch { Write-Host "GroupDef err: $($_.Exception.Message)" }
Write-Host "Groups: $($gnames -join ',')"
foreach ($g in $gnames) {
    $numLoads=0
    $areaName=[string[]]::new(0); $loadPat=[string[]]::new(0); $csys=[string[]]::new(0)
    $dir=[int[]]::new(0); $val=[double[]]::new(0)
    try {
        $null = $sm.AreaObj.GetLoadUniform($g, [ref]$numLoads, [ref]$areaName, [ref]$loadPat, [ref]$csys, [ref]$dir, [ref]$val, 1)
        if ($numLoads -gt 0) {
            Write-Host ("group '{0}': {1} loads" -f $g, $numLoads)
            for ($k=0; $k -lt [Math]::Min($numLoads, 5); $k++) {
                Write-Host ("  area={0} pat={1} dir={2} val={3}" -f $areaName[$k], $loadPat[$k], $dir[$k], $val[$k])
            }
        }
    } catch {}
}

# Directly, is there an "ALL" container? Try with empty string + ItemType=2 (selected)
Write-Host ""
Write-Host "=== Combo DConS2 detail ==="
$count=0;$caseType=[int[]]::new(0);$caseName=[string[]]::new(0);$sf=[double[]]::new(0)
try {
    $null = $sm.RespCombo.GetCaseList("DConS2", [ref]$count, [ref]$caseType, [ref]$caseName, [ref]$sf)
    Write-Host "DConS2 has $count cases"
    for ($i=0; $i -lt $count; $i++) {
        Write-Host ("  caseType={0} case='{1}' sf={2}" -f $caseType[$i], $caseName[$i], $sf[$i])
    }
} catch { Write-Host "err: $($_.Exception.Message)" }

# Slab polygon with Z-coords to confirm planar
Write-Host ""
Write-Host "=== Slab polygon (area 1) XY coords ==="
foreach ($p in $ptNames) {
    $x=0.0;$y=0.0;$z=0.0
    try {
        $null = $sm.PointObj.GetCoordCartesian($p, [ref]$x, [ref]$y, [ref]$z, "Global")
        Write-Host ("  pt {0}: ({1,12:F2}, {2,12:F2}, {3,8:F2})" -f $p, $x, $y, $z)
    } catch {}
}
