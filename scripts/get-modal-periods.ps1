param([switch]$AsJson)
$ErrorActionPreference = "Stop"

$process = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if ($null -eq $process) { throw "No ETABS process found." }
$exePath = $process.Path
if ([string]::IsNullOrWhiteSpace($exePath)) {
    $exePath = "C:\Program Files\Computers and Structures\ETABS 20\ETABS.exe"
}
$dll = Join-Path (Split-Path -Parent $exePath) "ETABSv1.dll"
Add-Type -Path $dll

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$sap = $api.SapModel

[void]$sap.Results.Setup.DeselectAllCasesAndCombosForOutput()

$numCases = 0
[string[]]$caseNames = @()
[void]$sap.LoadCases.GetNameList([ref]$numCases, [ref]$caseNames, [ETABSv1.eLoadCaseType]::Modal)

$rows = New-Object System.Collections.Generic.List[object]
foreach ($case in $caseNames) {
    [void]$sap.Results.Setup.SetCaseSelectedForOutput($case, $true)

    $n = 0
    [string[]]$lc = @(); [string[]]$st = @(); [double[]]$sn = @()
    [double[]]$period = @(); [double[]]$freq = @(); [double[]]$circFreq = @(); [double[]]$eigenVal = @()
    $ret = $sap.Results.ModalPeriod([ref]$n, [ref]$lc, [ref]$st, [ref]$sn, [ref]$period, [ref]$freq, [ref]$circFreq, [ref]$eigenVal)
    if ($ret -ne 0) {
        Write-Host "ModalPeriod ret=$ret for $case (analysis not run?)"
        continue
    }
    for ($i = 0; $i -lt $n; $i++) {
        $rows.Add([pscustomobject]@{
            Case = $lc[$i]
            Mode = $sn[$i]
            Period = $period[$i]
            Freq = $freq[$i]
        }) | Out-Null
    }

    $nm = 0
    [string[]]$lc2 = @(); [string[]]$st2 = @(); [double[]]$sn2 = @()
    [double[]]$ux = @(); [double[]]$uy = @(); [double[]]$uz = @()
    [double[]]$sumUx = @(); [double[]]$sumUy = @(); [double[]]$sumUz = @()
    [double[]]$rx = @(); [double[]]$ry = @(); [double[]]$rz = @()
    [double[]]$sumRx = @(); [double[]]$sumRy = @(); [double[]]$sumRz = @()
    $ret2 = $sap.Results.ModalParticipatingMassRatios([ref]$nm, [ref]$lc2, [ref]$st2, [ref]$sn2,
        [ref]$ux, [ref]$uy, [ref]$uz, [ref]$sumUx, [ref]$sumUy, [ref]$sumUz,
        [ref]$rx, [ref]$ry, [ref]$rz, [ref]$sumRx, [ref]$sumRy, [ref]$sumRz)
    if ($ret2 -eq 0) {
        for ($i = 0; $i -lt $nm; $i++) {
            $idx = $i
            $row = $rows[$idx]
            if ($null -ne $row -and $row.Mode -eq $sn2[$i]) {
                Add-Member -InputObject $row -NotePropertyName Ux -NotePropertyValue $ux[$i] -Force
                Add-Member -InputObject $row -NotePropertyName Uy -NotePropertyValue $uy[$i] -Force
                Add-Member -InputObject $row -NotePropertyName SumUx -NotePropertyValue $sumUx[$i] -Force
                Add-Member -InputObject $row -NotePropertyName SumUy -NotePropertyValue $sumUy[$i] -Force
            }
        }
    }
}

$xDominant = $rows | Sort-Object -Property Ux -Descending | Select-Object -First 1
$yDominant = $rows | Sort-Object -Property Uy -Descending | Select-Object -First 1

$summary = [pscustomobject]@{
    ModelPath = $sap.GetModelFilename($true)
    Modes = @($rows)
    XDominantMode = $xDominant
    YDominantMode = $yDominant
    T1 = ($rows | Where-Object { $_.Mode -eq 1 } | Select-Object -First 1).Period
}

if ($AsJson) { $summary | ConvertTo-Json -Depth 6 }
else { $summary }
