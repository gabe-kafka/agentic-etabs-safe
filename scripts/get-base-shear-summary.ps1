param(
    [int]$EtabsPid,

    [string[]]$IncludeComboPatterns,

    [switch]$ListOnly,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-EtabsProcess {
    param([int]$RequestedPid)
    $procs = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) {
        return $procs | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }
    $windowed = $procs | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($windowed) { return $windowed | Select-Object -First 1 }
    return $procs | Select-Object -First 1
}

function Resolve-EtabsApiDll {
    param([System.Diagnostics.Process]$Process)
    $candidates = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Process -and -not [string]::IsNullOrWhiteSpace($Process.Path)) {
        $candidates.Add((Join-Path (Split-Path -Parent $Process.Path) "ETABSv1.dll"))
    }
    @(
        "C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll"
    ) | ForEach-Object { $candidates.Add($_) }
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path }
    }
    throw "ETABSv1.dll not found."
}

function Assert-Success {
    param([int]$Ret, [string]$Op)
    if ($Ret -ne 0) { throw "$Op failed: return code $Ret" }
}

function Classify-Combo {
    param([string]$Name)
    $upper = $Name.ToUpperInvariant()
    $isAsd = $upper -match "ASD" -or $upper -match "SERVICE" -or $upper -match "ALLOW"
    $isWind = $upper -match "WIND" -or $upper -match "WL" -or $upper -match "WX" -or $upper -match "WY"
    $isEq = $upper -match "EARTHQUAKE" -or $upper -match "SEISMIC" -or $upper -match "EQ" -or $upper -match "EX" -or $upper -match "EY"
    return [pscustomobject]@{
        IsAsd = $isAsd
        IsWind = $isWind
        IsEarthquake = $isEq
    }
}

$process = Get-EtabsProcess -RequestedPid $EtabsPid
if ($null -eq $process) { throw "No running ETABS process." }

$dll = Resolve-EtabsApiDll -Process $process
Add-Type -Path $dll

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$sap = $api.SapModel

$numCombos = 0
[string[]]$comboNames = @()
Assert-Success ($sap.RespCombo.GetNameList([ref]$numCombos, [ref]$comboNames)) "GetNameList"

$classified = foreach ($n in $comboNames) {
    $c = Classify-Combo -Name $n
    [pscustomobject]@{
        Name = $n
        IsAsd = $c.IsAsd
        IsWind = $c.IsWind
        IsEarthquake = $c.IsEarthquake
    }
}

if ($ListOnly) {
    if ($AsJson) {
        @{
            ModelPath = $sap.GetModelFilename($true)
            Combos = @($classified)
        } | ConvertTo-Json -Depth 5
    }
    else {
        $classified | Format-Table -AutoSize
    }
    return
}

$targets = if ($IncludeComboPatterns) {
    $classified | Where-Object {
        $name = $_.Name
        ($IncludeComboPatterns | Where-Object { $name -like $_ }).Count -gt 0
    }
}
else {
    $classified | Where-Object { $_.IsAsd -and ($_.IsWind -or $_.IsEarthquake) }
}

$targetNames = @($targets | Select-Object -ExpandProperty Name)
if ($targetNames.Count -eq 0) {
    throw "No combos matched. Use -ListOnly to see available combos."
}

Assert-Success ($sap.Results.Setup.DeselectAllCasesAndCombosForOutput()) "Deselect output"
foreach ($name in $targetNames) {
    Assert-Success ($sap.Results.Setup.SetComboSelectedForOutput($name, $true)) "Select combo $name"
}

$numberResults = 0
[string[]]$loadCase = @()
[string[]]$stepType = @()
[double[]]$stepNum = @()
[double[]]$fx = @()
[double[]]$fy = @()
[double[]]$fz = @()
[double[]]$mx = @()
[double[]]$my = @()
[double[]]$mz = @()
[double]$gx = 0
[double]$gy = 0
[double]$gz = 0

$ret = $sap.Results.BaseReact(
    [ref]$numberResults,
    [ref]$loadCase,
    [ref]$stepType,
    [ref]$stepNum,
    [ref]$fx,
    [ref]$fy,
    [ref]$fz,
    [ref]$mx,
    [ref]$my,
    [ref]$mz,
    [ref]$gx,
    [ref]$gy,
    [ref]$gz
)
Assert-Success $ret "BaseReact"

$units = $sap.GetPresentUnits_2
$forceUnit = $null
$lengthUnit = $null
$tempUnit = $null
try {
    $fu = 0; $lu = 0; $tu = 0
    [void]$sap.GetPresentUnits_2([ref]$fu, [ref]$lu, [ref]$tu)
    $forceUnit = $fu
    $lengthUnit = $lu
    $tempUnit = $tu
}
catch { }

$rows = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $numberResults; $i++) {
    $rows.Add([pscustomobject]@{
        Combo = $loadCase[$i]
        StepType = $stepType[$i]
        StepNum = $stepNum[$i]
        Fx = $fx[$i]
        Fy = $fy[$i]
        Fz = $fz[$i]
        Mx = $mx[$i]
        My = $my[$i]
        Mz = $mz[$i]
        AbsFx = [math]::Abs($fx[$i])
        AbsFy = [math]::Abs($fy[$i])
        Resultant = [math]::Sqrt($fx[$i] * $fx[$i] + $fy[$i] * $fy[$i])
    }) | Out-Null
}

function Pick-Max {
    param($items, [string]$Category, [scriptblock]$Filter)
    $filtered = @($items | Where-Object {
        $name = $_.Combo
        $match = $classified | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($null -eq $match) { return $false }
        if ($Category -eq "Wind") { return $match.IsWind -and $match.IsAsd }
        if ($Category -eq "Earthquake") { return $match.IsEarthquake -and $match.IsAsd }
        return $false
    })
    if ($filtered.Count -eq 0) { return $null }
    $byFx = $filtered | Sort-Object -Property AbsFx -Descending | Select-Object -First 1
    $byFy = $filtered | Sort-Object -Property AbsFy -Descending | Select-Object -First 1
    $byResultant = $filtered | Sort-Object -Property Resultant -Descending | Select-Object -First 1
    return [pscustomobject]@{
        Category = $Category
        MaxAbsFx = [pscustomobject]@{ Combo = $byFx.Combo; Fx = $byFx.Fx; Fy = $byFx.Fy; Resultant = $byFx.Resultant }
        MaxAbsFy = [pscustomobject]@{ Combo = $byFy.Combo; Fx = $byFy.Fx; Fy = $byFy.Fy; Resultant = $byFy.Resultant }
        MaxResultant = [pscustomobject]@{ Combo = $byResultant.Combo; Fx = $byResultant.Fx; Fy = $byResultant.Fy; Resultant = $byResultant.Resultant }
        Rows = @($filtered | ForEach-Object { [pscustomobject]@{ Combo = $_.Combo; Fx = $_.Fx; Fy = $_.Fy; Resultant = $_.Resultant } })
    }
}

$summary = [pscustomobject]@{
    ProcessId = $process.Id
    ModelPath = $sap.GetModelFilename($true)
    ForceUnitCode = $forceUnit
    LengthUnitCode = $lengthUnit
    ComboCount = $targetNames.Count
    CombosEvaluated = $targetNames
    Wind = Pick-Max -items $rows -Category "Wind"
    Earthquake = Pick-Max -items $rows -Category "Earthquake"
}

if ($AsJson) {
    $summary | ConvertTo-Json -Depth 8
}
else {
    $summary
}
