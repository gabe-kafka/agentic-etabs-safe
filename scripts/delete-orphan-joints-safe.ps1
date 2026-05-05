param(
    [int]$SafePid,
    [switch]$DryRun,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$dll = Get-ChildItem "C:\Program Files\Computers and Structures" -Recurse -Filter "SAFEv1.dll" -ErrorAction SilentlyContinue |
       Sort-Object FullName -Descending | Select-Object -First 1
if (-not $dll) { throw "SAFEv1.dll not found." }
Add-Type -Path $dll.FullName

$helper = New-Object SAFEv1.Helper
$proc = if ($SafePid) {
    Get-Process -Id $SafePid -ErrorAction Stop
} else {
    Get-Process SAFE -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1
}
if (-not $proc) { throw "No running SAFE process found." }
$api = $helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sap = $api.SapModel

# Build set of joints connected to any element
$connected = New-Object System.Collections.Generic.HashSet[string]

$numF = 0; $fNames = [string[]]@()
$sap.FrameObj.GetNameList([ref]$numF, [ref]$fNames) | Out-Null
foreach ($f in $fNames) {
    $p1=''; $p2=''
    $sap.FrameObj.GetPoints($f,[ref]$p1,[ref]$p2) | Out-Null
    [void]$connected.Add($p1); [void]$connected.Add($p2)
}

$numA = 0; $aNames = [string[]]@()
$sap.AreaObj.GetNameList([ref]$numA, [ref]$aNames) | Out-Null
foreach ($a in $aNames) {
    $nPts=0; $pts=[string[]]@()
    $sap.AreaObj.GetPoints($a,[ref]$nPts,[ref]$pts) | Out-Null
    foreach ($p in $pts) { [void]$connected.Add($p) }
}

$numJ=0; $jNames=[string[]]@()
$sap.PointObj.GetNameList([ref]$numJ,[ref]$jNames) | Out-Null

# Find orphans: not connected, not restrained, no spring
$orphans = @()
foreach ($j in $jNames) {
    if ($connected.Contains($j)) { continue }

    $val=[bool[]]@($false,$false,$false,$false,$false,$false)
    $sap.PointObj.GetRestraint($j,[ref]$val) | Out-Null
    if (@($val | Where-Object {$_}).Count -gt 0) { continue }

    $k=[double[]]@(0,0,0,0,0,0)
    $sap.PointObj.GetSpring($j,[ref]$k) | Out-Null
    if (($k | Measure-Object -Sum).Sum -ne 0) { continue }

    $orphans += $j
}

Write-Host "Orphaned joints found: $($orphans.Count)"

if ($orphans.Count -eq 0) {
    Write-Host "Nothing to delete."
    exit 0
}

if ($DryRun) {
    Write-Host "DryRun mode - no changes made."
    $orphans | ForEach-Object { Write-Host "  Would delete joint $_" }
    exit 0
}

# Delete orphaned joints
$deleted = 0; $failed = @()
foreach ($j in $orphans) {
    $ret = $sap.PointObj.DeleteSpecialPoint($j, [SAFEv1.eItemType]::Object)
    if ($ret -eq 0) { $deleted++ }
    else { $failed += $j }
}

Write-Host "Deleted : $deleted"
if ($failed.Count -gt 0) {
    Write-Host "Failed  : $($failed.Count) - $($failed -join ', ')"
}

$result = [pscustomobject]@{
    OrphansFound  = $orphans.Count
    Deleted       = $deleted
    Failed        = $failed.Count
    FailedJoints  = $failed
}

if ($AsJson) { $result | ConvertTo-Json }
else { $result | Format-List }
