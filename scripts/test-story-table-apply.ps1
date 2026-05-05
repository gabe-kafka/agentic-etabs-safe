param(
    [int]$EtabsPid,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$targetStories = @(
    [pscustomobject]@{ Name = "CELLAR";     Elevation = 76.28 }
    [pscustomobject]@{ Name = "1ST FLOOR";  Elevation = 86.28 }
    [pscustomobject]@{ Name = "2ND FLOOR";  Elevation = 96.28 }
    [pscustomobject]@{ Name = "3RD FLOOR";  Elevation = 106.28 }
    [pscustomobject]@{ Name = "4TH FLOOR";  Elevation = 116.28 }
    [pscustomobject]@{ Name = "5TH FLOOR";  Elevation = 126.28 }
    [pscustomobject]@{ Name = "6TH FLOOR";  Elevation = 136.28 }
    [pscustomobject]@{ Name = "7TH FLOOR";  Elevation = 146.28 }
    [pscustomobject]@{ Name = "8TH FLOOR";  Elevation = 156.28 }
    [pscustomobject]@{ Name = "9TH FLOOR";  Elevation = 166.28 }
    [pscustomobject]@{ Name = "10TH FLOOR"; Elevation = 176.28 }
    [pscustomobject]@{ Name = "11TH FLOOR"; Elevation = 186.28 }
    [pscustomobject]@{ Name = "12TH FLOOR"; Elevation = 196.28 }
    [pscustomobject]@{ Name = "13TH FLOOR"; Elevation = 206.28 }
    [pscustomobject]@{ Name = "14TH FLOOR"; Elevation = 216.28 }
    [pscustomobject]@{ Name = "15TH FLOOR"; Elevation = 226.28 }
    [pscustomobject]@{ Name = "16TH FLOOR"; Elevation = 236.28 }
    [pscustomobject]@{ Name = "17TH FLOOR"; Elevation = 246.28 }
    [pscustomobject]@{ Name = "18TH FLOOR"; Elevation = 256.28 }
    [pscustomobject]@{ Name = "ROOF";       Elevation = 266.28 }
)
$baseElevation = 0.0

for ($i = 0; $i -lt $targetStories.Count; $i++) {
    $below = if ($i -eq 0) { $baseElevation } else { $targetStories[$i-1].Elevation }
    $targetStories[$i] | Add-Member -NotePropertyName Height -NotePropertyValue ([Math]::Round($targetStories[$i].Elevation - $below, 6))
}

function Get-EtabsProcess {
    param([int]$Pid)
    $procs = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
    if ($Pid) { return $procs | Where-Object { $_.Id -eq $Pid } | Select-Object -First 1 }
    $w = $procs | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($w) { return $w | Select-Object -First 1 }
    return $procs | Select-Object -First 1
}

function Resolve-EtabsDll {
    param([System.Diagnostics.Process]$Process)
    $c = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $Process) {
        try { $c.Add((Join-Path (Split-Path -Parent $Process.Path) "ETABSv1.dll")) } catch {}
    }
    foreach ($p in @(
        "C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll"
    )) { $c.Add($p) }
    foreach ($path in $c) { if (Test-Path -LiteralPath $path) { return $path } }
    throw "ETABSv1.dll not found."
}

$proc = Get-EtabsProcess -Pid $EtabsPid
if ($null -eq $proc) { throw "No running ETABS process found." }

$dll = Resolve-EtabsDll -Process $proc
Add-Type -Path $dll

$helper = New-Object ETABSv1.Helper
$api    = ([ETABSv1.cHelper]$helper).GetObjectProcess("CSI.ETABS.API.ETABSObject", $proc.Id)
$sap    = $api.SapModel
$db     = $sap.DatabaseTables

Write-Host "Connected ETABS PID $($proc.Id)"

# --- Dump method signatures via all available type paths ---
Write-Host ""
Write-Host "=== DatabaseTables method signatures ==="

$sigSources = [System.Collections.Generic.List[System.Type]]::new()
$sigSources.Add($db.GetType())
foreach ($iface in $db.GetType().GetInterfaces()) { $sigSources.Add($iface) }
foreach ($asm in [System.AppDomain]::CurrentDomain.GetAssemblies()) {
    try {
        foreach ($t in $asm.GetTypes()) {
            if ($t.Name -like "*DatabaseTable*" -or $t.Name -like "*cDatabaseTable*") {
                $sigSources.Add($t)
            }
        }
    } catch {}
}

$seen = [System.Collections.Generic.HashSet[string]]::new()
foreach ($t in $sigSources) {
    try {
        foreach ($m in $t.GetMethods()) {
            if ($m.Name -notmatch "TableForDisplay|TableForEditing|ApplyEdited") { continue }
            $params = $m.GetParameters() | ForEach-Object {
                $byref = if ($_.ParameterType.IsByRef) { "ref " } else { "" }
                "$byref$($_.ParameterType.Name.TrimEnd('&')) $($_.Name)"
            }
            $sig = "$($m.Name)($($params -join ', '))"
            if ($seen.Add($sig)) { Write-Host "  $sig" }
        }
    } catch {}
}

# --- Read Story Definitions table ---
Write-Host ""
Write-Host "=== Reading Story Definitions table ==="

$ff = ""; $ni = 0
[string[]]$flds = @()
$nr = 0
[string[]]$td = @()

$ret = $db.GetTableForDisplayArray("Story Definitions", [ref]$ff, "All", [ref]$ni, [ref]$flds, [ref]$nr, [ref]$td)
Write-Host "  ret=$ret  numItems=$ni  numRecords=$nr  fields=$($flds.Count)"
Write-Host "  Fields: $($flds -join ' | ')"

if ($ret -ne 0 -or $flds.Count -eq 0) { throw "Could not read Story Definitions." }

$nf = $flds.Count
$nameIdx = -1; $heightIdx = -1; $masterIdx = -1; $similarIdx = -1
$spliceAboveIdx = -1; $spliceHeightIdx = -1; $colorIdx = -1

for ($f = 0; $f -lt $nf; $f++) {
    switch -Regex ($flds[$f].ToLower()) {
        "^name$|^story$"             { $nameIdx        = $f }
        "height"                     { $heightIdx       = $f }
        "master"                     { $masterIdx       = $f }
        "similar"                    { $similarIdx      = $f }
        "splice.*above|above splice" { $spliceAboveIdx  = $f }
        "splice.*height|splice ht"   { $spliceHeightIdx = $f }
        "color"                      { $colorIdx        = $f }
    }
}

Write-Host "  nameIdx=$nameIdx  heightIdx=$heightIdx"

if ($nameIdx -lt 0 -or $heightIdx -lt 0) { throw "Cannot map Name/Height. Fields: $($flds -join ', ')" }

$nc = $targetStories.Count
$newData = New-Object string[] ($nc * $nf)
for ($r = 0; $r -lt $nc; $r++) {
    $base = $r * $nf
    $newData[$base + $nameIdx]   = $targetStories[$r].Name
    $newData[$base + $heightIdx] = [string]$targetStories[$r].Height
    if ($masterIdx       -ge 0) { $newData[$base + $masterIdx]       = "Yes" }
    if ($similarIdx      -ge 0) { $newData[$base + $similarIdx]      = "" }
    if ($spliceAboveIdx  -ge 0) { $newData[$base + $spliceAboveIdx]  = "No" }
    if ($spliceHeightIdx -ge 0) { $newData[$base + $spliceHeightIdx] = "0" }
    if ($colorIdx        -ge 0) { $newData[$base + $colorIdx]        = "-1" }
}

Write-Host "  Built $nc records x $nf fields"
Write-Host "  Sample row 0: $(for ($f=0;$f -lt $nf;$f++) { "$($flds[$f])=$($newData[$f])" })"

if ($DryRun) {
    Write-Host ""
    Write-Host "DryRun complete."
    exit 0
}

# --- Try SetTableForEditingArray variants ---
Write-Host ""
Write-Host "=== SetTableForEditingArray attempts ==="

$setOk = $false
$setRet = $null

function Try-Set {
    param([string]$Label, [scriptblock]$Block)
    Write-Host "  [$Label]"
    try {
        $result = & $Block
        Write-Host "  -> SUCCESS ret=$result"
        return $result
    } catch {
        $msg = $_.Exception.Message -replace "`r`n"," " -replace "`n"," "
        Write-Host "  -> FAIL: $msg"
        return $null
    }
}

$r1 = Try-Set "5p tv=ref nc=ref" { $tv=[int]0;$ncr=[int]$nc; $db.SetTableForEditingArray("Story Definitions",[ref]$tv,[ref]$ncr,$flds,$newData) }
if ($null -ne $r1) { $setOk=$true; $setRet=$r1 }

if (-not $setOk) {
    $r2 = Try-Set "5p tv=ref fields nc=ref" { $tv=[int]0;$ncr=[int]$nc; $db.SetTableForEditingArray("Story Definitions",[ref]$tv,$flds,[ref]$ncr,$newData) }
    if ($null -ne $r2) { $setOk=$true; $setRet=$r2 }
}

if (-not $setOk) {
    $r3 = Try-Set "5p tv=val nc=ref" { $ncr=[int]$nc; $db.SetTableForEditingArray("Story Definitions",0,[ref]$ncr,$flds,$newData) }
    if ($null -ne $r3) { $setOk=$true; $setRet=$r3 }
}

if (-not $setOk) {
    $r4 = Try-Set "5p tv=val fields nc=ref" { $ncr=[int]$nc; $db.SetTableForEditingArray("Story Definitions",0,$flds,[ref]$ncr,$newData) }
    if ($null -ne $r4) { $setOk=$true; $setRet=$r4 }
}

if (-not $setOk) {
    $r5 = Try-Set "5p all by-value" { $db.SetTableForEditingArray("Story Definitions",0,$nc,$flds,$newData) }
    if ($null -ne $r5) { $setOk=$true; $setRet=$r5 }
}

if (-not $setOk) {
    $r6 = Try-Set "4p nc=ref" { $ncr=[int]$nc; $db.SetTableForEditingArray("Story Definitions",[ref]$ncr,$flds,$newData) }
    if ($null -ne $r6) { $setOk=$true; $setRet=$r6 }
}

if (-not $setOk) {
    $r7 = Try-Set "4p all by-value" { $db.SetTableForEditingArray("Story Definitions",$nc,$flds,$newData) }
    if ($null -ne $r7) { $setOk=$true; $setRet=$r7 }
}

if (-not $setOk) {
    throw "All SetTableForEditingArray variants failed."
}

Write-Host "  SetTableForEditingArray succeeded."

# --- Try ApplyEditedTables variants ---
Write-Host ""
Write-Host "=== ApplyEditedTables attempts ==="

$applyOk = $false

$a1 = Try-Set "7p ref-outputs" {
    $ak=""; $fe=$false; $ec=0; $wc=0; $ic=0; [string[]]$errs=@()
    $r = $db.ApplyEditedTables($false,[ref]$ak,[ref]$fe,[ref]$ec,[ref]$wc,[ref]$ic,[ref]$errs)
    if ($fe) { throw "Fatal: $(($errs|Select-Object -First 5) -join '; ')" }
    $r
}
if ($null -ne $a1) { $applyOk = $true }

if (-not $applyOk) {
    $a2 = Try-Set "1p DoLocking" { $db.ApplyEditedTables($false) }
    if ($null -ne $a2) { $applyOk = $true }
}

if (-not $applyOk) {
    throw "All ApplyEditedTables variants failed."
}

Write-Host "  ApplyEditedTables succeeded."

# --- Verify ---
Write-Host ""
Write-Host "=== Post-write verification ==="
$ff2=""; $ni2=0; [string[]]$flds2=@(); $nr2=0; [string[]]$td2=@()
$db.GetTableForDisplayArray("Story Definitions",[ref]$ff2,"All",[ref]$ni2,[ref]$flds2,[ref]$nr2,[ref]$td2) | Out-Null
Write-Host "  Post-apply record count: $nr2"
$nf2 = $flds2.Count
for ($r=0; $r -lt $nr2; $r++) {
    $b = $r * $nf2
    $row = for ($f=0;$f -lt $nf2;$f++) { "$($flds2[$f])=$($td2[$b+$f])" }
    Write-Host "  Row $r : $($row -join '  ')"
}

Write-Host ""
Write-Host "Done."
