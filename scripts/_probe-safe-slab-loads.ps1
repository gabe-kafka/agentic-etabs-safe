param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3) # kip-in

# --- AreaObj methods ---
Write-Host "=== AreaObj methods (Get*) ==="
($sm.AreaObj | Get-Member -MemberType Method | Where-Object { $_.Name -match "^Get" }).Name | Sort-Object

# --- Area properties ---
Write-Host ""
Write-Host "=== PropArea methods (Get*) ==="
($sm.PropArea | Get-Member -MemberType Method | Where-Object { $_.Name -match "^Get" }).Name | Sort-Object

# --- Enumerate area objects ---
$na = 0; $anames = [string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
Write-Host ""
Write-Host ("Total area objects: {0}" -f $na)

# For each area: get polygon, property (type), loads
Write-Host ""
Write-Host "=== Area polygon + property summary (first 5) ==="
for ($i=0; $i -lt [Math]::Min(5,$na); $i++) {
    $a = $anames[$i]
    $nPts = 0; $ptNames = [string[]]::new(0)
    $null = $sm.AreaObj.GetPoints($a, [ref]$nPts, [ref]$ptNames)
    Write-Host ("Area '{0}': {1} points: {2}" -f $a, $nPts, ($ptNames -join ","))

    $propName = ""
    try {
        $null = $sm.AreaObj.GetProperty($a, [ref]$propName)
        Write-Host "  property: '$propName'"
    } catch { Write-Host "  GetProperty err: $($_.Exception.Message)" }

    $isOpening = $false
    try {
        $null = $sm.AreaObj.GetOpening($a, [ref]$isOpening)
        Write-Host "  isOpening: $isOpening"
    } catch {}
}

# Unique props
$propsUsed = @{}
for ($i=0; $i -lt $na; $i++) {
    $p=""
    try { $null = $sm.AreaObj.GetProperty($anames[$i], [ref]$p) } catch {}
    if ($p) { $propsUsed[$p] = (($propsUsed[$p] | Measure-Object).Count + 1) }
}
Write-Host ""
Write-Host "=== Unique area properties in model ==="
$propsUsed.Keys | ForEach-Object { Write-Host "  $_" }

# Property details: slab thickness
Write-Host ""
Write-Host "=== Slab property details ==="
foreach ($p in $propsUsed.Keys) {
    try {
        $slabType=0;$shellType=0;$mat="";$matAng=0.0;$thickness=0.0;$col=0;$notes="";$guid=""
        $null = $sm.PropArea.GetSlab($p, [ref]$slabType, [ref]$shellType, [ref]$mat, [ref]$thickness, [ref]$col, [ref]$notes, [ref]$guid)
        Write-Host ("  {0}: slabType={1} shellType={2} mat={3} thickness={4}" -f $p, $slabType, $shellType, $mat, $thickness)
    } catch { Write-Host "  $p -- GetSlab err: $($_.Exception.Message)" }
}

# --- Area loads ---
Write-Host ""
Write-Host "=== Area uniform loads ==="
for ($i=0; $i -lt [Math]::Min(5,$na); $i++) {
    $a = $anames[$i]
    try {
        $numLoads = 0
        $areaName=[string[]]::new(0)
        $loadPat=[string[]]::new(0)
        $csys=[string[]]::new(0)
        $dir=[int[]]::new(0)
        $val=[double[]]::new(0)
        $null = $sm.AreaObj.GetLoadUniform($a, [ref]$numLoads, [ref]$areaName, [ref]$loadPat, [ref]$csys, [ref]$dir, [ref]$val, 0)
        Write-Host ("Area '{0}': {1} uniform loads" -f $a, $numLoads)
        for ($k=0; $k -lt $numLoads; $k++) {
            Write-Host ("  {0} dir={1} val={2} csys={3}" -f $loadPat[$k], $dir[$k], $val[$k], $csys[$k])
        }
    } catch { Write-Host "  AreaObj.GetLoadUniform err on $a : $($_.Exception.Message)" }
}

# --- Material f'c ---
Write-Host ""
Write-Host "=== Concrete material f'c ==="
$nmat=0;$mnames=[string[]]::new(0)
$null = $sm.PropMaterial.GetNameList([ref]$nmat, [ref]$mnames)
foreach ($m in $mnames) {
    try {
        $mtype=0;$color=0;$notes="";$guid=""
        $null = $sm.PropMaterial.GetMaterial($m, [ref]$mtype, [ref]$color, [ref]$notes, [ref]$guid)
        # Type 2 = concrete
        if ($mtype -eq 2) {
            $fc=0.0;$fcLight=$false;$fcsFactor=0.0;$fy=0.0;$fys=0.0;$Ec=0.0
            try {
                $null = $sm.PropMaterial.GetOConcrete($m, [ref]$fc, [ref]$fcLight, [ref]$fcsFactor, [ref]$fy, [ref]$fys, [ref]$Ec, [ref]$Ec, [ref]$Ec, [ref]$Ec)
            } catch {
                try {
                    $null = $sm.PropMaterial.GetOConcrete_1($m, [ref]$fc, [ref]$fcLight, [ref]$fcsFactor, [ref]$fy, [ref]$fys, [ref]$Ec, [ref]$Ec, [ref]$Ec, [ref]$Ec)
                } catch {}
            }
            Write-Host ("  {0}: type=concrete fc={1} ksi" -f $m, $fc)
        } else {
            Write-Host ("  {0}: type={1}" -f $m, $mtype)
        }
    } catch { Write-Host "  $m err: $($_.Exception.Message)" }
}

# Frame local axes angle
Write-Host ""
Write-Host "=== Frame local axes (first 3 columns) ==="
$nf=0;$fnames=[string[]]::new(0)
$null = $sm.FrameObj.GetNameList([ref]$nf, [ref]$fnames)
for ($i=0; $i -lt 3; $i++) {
    $fr = $fnames[$i]
    $ang=0.0;$advanced=$false
    try {
        $null = $sm.FrameObj.GetLocalAxes($fr, [ref]$ang, [ref]$advanced)
        Write-Host ("Frame {0}: ang={1} deg (advanced={2})" -f $fr, $ang, $advanced)
    } catch { Write-Host "  err: $($_.Exception.Message)" }
}
