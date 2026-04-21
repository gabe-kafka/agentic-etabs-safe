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

# Select all load patterns for results display
$null = $sm.Results.Setup.DeselectAllCasesAndCombosForOutput()
foreach ($p in @("Dead","Live","SDL","DConS2")) {
    try {
        # LoadPatterns use SetCaseSelectedForOutput
        if ($p -eq "DConS2") {
            $null = $sm.Results.Setup.SetComboSelectedForOutput($p, $true)
        } else {
            $null = $sm.Results.Setup.SetCaseSelectedForOutput($p, $true)
        }
    } catch { Write-Host "sel err $p : $($_.Exception.Message)" }
}

# Get reactions at one punching point per pattern
$probePts = @("26","32","38","44")
Write-Host "=== Results.JointReact for sample punching points ==="
$numberResults=0
$obj=[string[]]::new(0)
$elm=[string[]]::new(0)
$acase=[string[]]::new(0)
$stepType=[string[]]::new(0)
$stepNum=[double[]]::new(0)
$f1=[double[]]::new(0)
$f2=[double[]]::new(0)
$f3=[double[]]::new(0)
$m1=[double[]]::new(0)
$m2=[double[]]::new(0)
$m3=[double[]]::new(0)

foreach ($p in $probePts) {
    try {
        $null = $sm.Results.JointReact(
            $p,
            0,  # ItemType.Object
            [ref]$numberResults,
            [ref]$obj,
            [ref]$elm,
            [ref]$acase,
            [ref]$stepType,
            [ref]$stepNum,
            [ref]$f1, [ref]$f2, [ref]$f3,
            [ref]$m1, [ref]$m2, [ref]$m3
        )
        Write-Host ("point {0}: {1} results" -f $p, $numberResults)
        for ($i=0; $i -lt $numberResults; $i++) {
            Write-Host ("  {0}: F3={1} kip  M1={2} kip-in  M2={3} kip-in" -f $acase[$i], $f3[$i], $m1[$i], $m2[$i])
        }
    } catch { Write-Host "JointReact err $p : $($_.Exception.Message)" }
}

# Try PropMaterial GetOConcrete -- inspect with list of methods
Write-Host ""
Write-Host "=== PropMaterial all methods ==="
($sm.PropMaterial | Get-Member -MemberType Method | Where-Object { $_.Name -match "(?i)(concrete|iso|elastic|material)" -and $_.Name -notmatch "^(get_|set_|add_|remove_)" }).Name | Sort-Object

# Try GetMPIsotropic for the concrete -> gives E
Write-Host ""
Write-Host "=== PropMaterial.GetMPIsotropic / GetOConcrete_1 ==="
try {
    $e=0.0;$u=0.0;$a=0.0
    $null = $sm.PropMaterial.GetMPIsotropic("5000Psi", [ref]$e, [ref]$u, [ref]$a)
    Write-Host ("5000Psi isotropic: E={0} ksi, nu={1}, alpha={2}" -f $e, $u, $a)
} catch { Write-Host "GetMPIsotropic err: $($_.Exception.Message)" }

# GetOConcrete: inspect reflected methods
$mi = $sm.PropMaterial.GetType().GetMethods() | Where-Object { $_.Name -match "^GetOConcrete" -or $_.Name -match "(?i)concrete" }
Write-Host ""
Write-Host "=== All concrete-ish methods (reflected) ==="
foreach ($m in $mi) {
    $ps = $m.GetParameters() | ForEach-Object { "$($_.ParameterType.Name) $($_.Name)" }
    Write-Host "  $($m.Name)($($ps -join ', '))"
}

# ---- Scan ALL database tables with "load" OR "uniform" in name ----
Write-Host ""
Write-Host "=== Database tables with load data ==="
$tryTables = @(
    "Slab Uniform Loads",
    "Slab Loads - Uniform",
    "Area Object Uniform Loads",
    "Floor Loads",
    "Load Assignments - Slab",
    "Load Assignments - Area Uniform",
    "Object Load Assignments - Uniform Loads on Slabs",
    "Uniform Load Assignments - Area",
    "Surface Loads",
    "Surface Load Assignments",
    "Uniform Loads to Frames - Area",
    "Slab Edge Uniform Loads"
)
foreach ($t in $tryTables) {
    try {
        $v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
        $null = $sm.DatabaseTables.GetTableForDisplayArray($t, [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
        if ($nr -gt 0) {
            Write-Host "Table '$t': $nr rows, fields: $($f -join '|')"
            # Show first row
            $cols = $f.Length
            $r = @(); for ($j=0; $j -lt $cols; $j++) { $r += $d[$j] }
            Write-Host ("  row0: {0}" -f ($r -join ' | '))
        }
    } catch {}
}

# Also check AreaObj methods we haven't tried
Write-Host ""
Write-Host "=== AreaObj.GetLoadUniformToFrame probe ==="
try {
    $numLoads=0
    $areaName=[string[]]::new(0); $loadPat=[string[]]::new(0); $csys=[string[]]::new(0)
    $dir=[int[]]::new(0); $val=[double[]]::new(0); $distType=[int[]]::new(0)
    $null = $sm.AreaObj.GetLoadUniformToFrame("1", [ref]$numLoads, [ref]$areaName, [ref]$loadPat, [ref]$csys, [ref]$dir, [ref]$val, [ref]$distType, 0)
    Write-Host "area 1 GetLoadUniformToFrame: $numLoads loads"
    for ($k=0; $k -lt $numLoads; $k++) {
        Write-Host ("  {0} dir={1} val={2} distType={3}" -f $loadPat[$k], $dir[$k], $val[$k], $distType[$k])
    }
} catch { Write-Host "err: $($_.Exception.Message)" }
