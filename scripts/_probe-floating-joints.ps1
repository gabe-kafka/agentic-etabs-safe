Add-Type -Path 'C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll'
$helper = New-Object SAFEv1.Helper
$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
$api = $helper.GetObjectProcess('CSI.SAFE.API.ETABSObject', $proc.Id)
$sap = $api.SapModel

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
$floating = @($jNames | Where-Object { -not $connected.Contains($_) })

Write-Host "Object-level joints  : $numJ"
Write-Host "Connected to elements: $($connected.Count)"
Write-Host "Floating joints      : $($floating.Count)"
Write-Host ""

# Sample 15 floating joints
Write-Host "Sample floating joints (coords, restraint, spring):"
$rows = @()
foreach ($j in ($floating | Select-Object -First 15)) {
    $x=0.0;$y=0.0;$z=0.0
    $sap.PointObj.GetCoordCartesian($j,[ref]$x,[ref]$y,[ref]$z,'Global') | Out-Null

    $val=[bool[]]@($false,$false,$false,$false,$false,$false)
    $sap.PointObj.GetRestraint($j,[ref]$val) | Out-Null
    $restrained = (@($val | Where-Object {$_}).Count -gt 0)

    $k=[double[]]@(0,0,0,0,0,0)
    $sap.PointObj.GetSpring($j,[ref]$k) | Out-Null
    $hasSpring = (($k | Measure-Object -Sum).Sum -ne 0)

    $rows += [pscustomobject]@{
        Joint      = $j
        X          = [Math]::Round($x,1)
        Y          = [Math]::Round($y,1)
        Z          = [Math]::Round($z,1)
        Restrained = $restrained
        Spring     = $hasSpring
    }
}
$rows | Format-Table -AutoSize

# Count how many floating joints have restraints or springs
$nRestrained = 0; $nSpring = 0
foreach ($j in $floating) {
    $val=[bool[]]@($false,$false,$false,$false,$false,$false)
    $sap.PointObj.GetRestraint($j,[ref]$val) | Out-Null
    if (@($val | Where-Object {$_}).Count -gt 0) { $nRestrained++ }

    $k=[double[]]@(0,0,0,0,0,0)
    $sap.PointObj.GetSpring($j,[ref]$k) | Out-Null
    if (($k | Measure-Object -Sum).Sum -ne 0) { $nSpring++ }
}
Write-Host "Of $($floating.Count) floating joints:"
Write-Host "  With restraints : $nRestrained"
Write-Host "  With springs    : $nSpring"
Write-Host "  Truly isolated  : $($floating.Count - $nRestrained - $nSpring)"
