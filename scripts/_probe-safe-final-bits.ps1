param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel

# --- 1. GetSlab: use typed enums ---
Write-Host "=== GetSlab with typed enums ==="
$slabType = [SAFEv1.eSlabType]::Slab
$shellType = [SAFEv1.eShellType]::ShellThin
$mat=""; $thickness=0.0; $col=0; $notes=""; $guid=""
try {
    $null = $sm.PropArea.GetSlab("8"" Concrete Slab 5ksi", [ref]$slabType, [ref]$shellType, [ref]$mat, [ref]$thickness, [ref]$col, [ref]$notes, [ref]$guid)
    Write-Host ("  slabType={0} shellType={1} mat='{2}' thickness={3}" -f $slabType, $shellType, $mat, $thickness)
} catch {
    Write-Host "err: $($_.Exception.Message)"
}

# --- 2. Area Loads Uniform via DB table ---
Write-Host ""
Write-Host "=== Area Loads Uniform (DB table) ==="
$null = $sm.SetPresentUnits(2)  # lb-ft: loads in psf
foreach ($tbl in @(
    "Area Loads - Uniform",
    "Area Object Loads - Uniform",
    "Area Load Assignments - Uniform",
    "Area Uniform Loads",
    "Slab Uniform Loads",
    "Uniform Load Sets",
    "Load Assignments - Area - Uniform",
    "Load Case Definitions - Static Load Case"
)) {
    try {
        $v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
        $null = $sm.DatabaseTables.GetTableForDisplayArray($tbl, [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
        if ($nr -gt 0) {
            Write-Host ("'{0}': rows={1}" -f $tbl, $nr)
            Write-Host ("  fields: {0}" -f ($f -join "|"))
            $ncol = $f.Length
            for ($i=0; $i -lt [Math]::Min(3,$nr); $i++) {
                $row = @()
                for ($j=0; $j -lt $ncol; $j++) { $row += $d[$i*$ncol + $j] }
                Write-Host ("  row{0}: {1}" -f $i, ($row -join " | "))
            }
        }
    } catch {}
}

# --- 3. Try GetLoadUniform with present units = lb-ft (psf) ---
Write-Host ""
Write-Host "=== AreaObj.GetLoadUniform in lb-ft units ==="
$na = 0; $anames = [string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
# Area 1 is the slab
foreach ($a in @("1","6","7")) {
    if ($anames -notcontains $a) { continue }
    $numLoads = 0
    $areaName=[string[]]::new(0); $loadPat=[string[]]::new(0); $csys=[string[]]::new(0)
    $dir=[int[]]::new(0); $val=[double[]]::new(0)
    try {
        $null = $sm.AreaObj.GetLoadUniform($a, [ref]$numLoads, [ref]$areaName, [ref]$loadPat, [ref]$csys, [ref]$dir, [ref]$val, 0)
        Write-Host ("area {0}: {1} loads" -f $a, $numLoads)
        for ($k=0; $k -lt $numLoads; $k++) {
            Write-Host ("  {0} dir={1} val={2} psf csys={3}" -f $loadPat[$k], $dir[$k], $val[$k], $csys[$k])
        }
    } catch {
        Write-Host ("area {0} err: {1}" -f $a, $_.Exception.Message)
    }
}

# --- 4. Combo definition ---
Write-Host ""
Write-Host "=== Combo definitions ==="
foreach ($c in @("DConS1","DConS2")) {
    try {
        $type=0
        $null = $sm.RespCombo.GetType_1($c, [ref]$type)
        $count=0;$caseType=[int[]]::new(0);$caseName=[string[]]::new(0);$sf=[double[]]::new(0)
        $null = $sm.RespCombo.GetCaseList($c, [ref]$count, [ref]$caseType, [ref]$caseName, [ref]$sf)
        Write-Host ("Combo '{0}': type={1}, {2} cases" -f $c, $type, $count)
        for ($i=0; $i -lt $count; $i++) {
            Write-Host ("  {0} '{1}' sf={2}" -f $caseType[$i], $caseName[$i], $sf[$i])
        }
    } catch { Write-Host "err: $($_.Exception.Message)" }
}

# --- 5. PropMaterial: try alternate signature for GetNameList ---
Write-Host ""
Write-Host "=== PropMaterial.GetNameList ==="
$nmat = 0; $mnames = [string[]]::new(0); $mtypes = [int[]]::new(0)
try {
    $null = $sm.PropMaterial.GetNameList([ref]$nmat, [ref]$mnames, [ref]$mtypes)
    for ($i=0; $i -lt $nmat; $i++) { Write-Host ("  {0} type={1}" -f $mnames[$i], $mtypes[$i]) }
} catch { Write-Host "err: $($_.Exception.Message)" }

# f'c for 5000Psi
Write-Host ""
Write-Host "=== PropMaterial.GetOConcrete for concretes ==="
$null = $sm.SetPresentUnits(2)  # lb-ft: fc in psf? check
try {
    $fc=0.0;$lw=$false;$fcs=0.0;$fy=0.0;$fys=0.0
    $strain=0.0;$strainUlt=0.0;$ft=0.0;$frictAng=0.0;$dilAng=0.0
    $null = $sm.PropMaterial.GetOConcrete("5000Psi", [ref]$fc, [ref]$lw, [ref]$fcs, [ref]$strain, [ref]$strainUlt, [ref]$ft, [ref]$frictAng, [ref]$dilAng)
    Write-Host ("5000Psi: fc={0} lightweight={1} fcs={2}" -f $fc, $lw, $fcs)
} catch { Write-Host "err: $($_.Exception.Message)" }

# Try with lb-in units for f'c in psi
$null = $sm.SetPresentUnits(1)  # lb-in
try {
    $fc=0.0;$lw=$false;$fcs=0.0
    $strain=0.0;$strainUlt=0.0;$ft=0.0;$frictAng=0.0;$dilAng=0.0
    $null = $sm.PropMaterial.GetOConcrete("5000Psi", [ref]$fc, [ref]$lw, [ref]$fcs, [ref]$strain, [ref]$strainUlt, [ref]$ft, [ref]$frictAng, [ref]$dilAng)
    Write-Host ("5000Psi @ lb-in: fc={0} psi" -f $fc)
} catch { Write-Host "err: $($_.Exception.Message)" }

# --- 6. Get Z coordinates of all area 1 points (to confirm it's the flat slab) ---
Write-Host ""
Write-Host "=== Area 1 point coordinates ==="
$null = $sm.SetPresentUnits(3)  # kip-in
$nPts=0;$ptNames=[string[]]::new(0)
$null = $sm.AreaObj.GetPoints("1", [ref]$nPts, [ref]$ptNames)
for ($i=0; $i -lt $nPts; $i++) {
    $x=0.0;$y=0.0;$z=0.0
    $null = $sm.PointObj.GetCoordCartesian($ptNames[$i], [ref]$x, [ref]$y, [ref]$z)
    Write-Host ("  pt {0}: ({1}, {2}, {3})" -f $ptNames[$i], $x, $y, $z)
}

# --- 7. Is there more than one slab area? Identify slab vs wall areas ---
Write-Host ""
Write-Host "=== Slab-property areas ==="
$slabAreas = @()
for ($i=0; $i -lt $na; $i++) {
    $a = $anames[$i]
    $p = ""
    try { $null = $sm.AreaObj.GetProperty($a, [ref]$p) } catch {}
    if ($p -match "Slab") {
        $slabAreas += $a
        $np=0;$pn=[string[]]::new(0)
        $null = $sm.AreaObj.GetPoints($a, [ref]$np, [ref]$pn)
        Write-Host ("  area {0}: prop='{1}' with {2} points" -f $a, $p, $np)
    }
}

# --- 8. Confirm DConS2 is the governing combo for punching ---
# Sample punching table Combo column
Write-Host ""
Write-Host "=== Punching combo frequency ==="
$v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
$null = $sm.DatabaseTables.GetTableForDisplayArray("Concrete Slab Design - Punching Shear Data", [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
$comboIdx = [array]::IndexOf($f,"Combo")
$combos=@{}
for ($i=0; $i -lt $nr; $i++) {
    $c = $d[$i*$f.Length + $comboIdx]
    $combos[$c] = ($combos[$c] + 1)
}
$combos.GetEnumerator() | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Key, $_.Value) }
