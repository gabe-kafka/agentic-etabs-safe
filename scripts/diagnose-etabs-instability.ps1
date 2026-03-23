param(
    [int]$EtabsPid,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-EtabsApiDll {
    param(
        [System.Diagnostics.Process]$Process
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($null -ne $Process) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($Process.Path)) {
                $candidates.Add((Join-Path -Path (Split-Path -Parent $Process.Path) -ChildPath "ETABSv1.dll"))
            }
        }
        catch {
        }
    }

    @(
        "C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 20\ETABSv1.dll"
    ) | ForEach-Object {
        $candidates.Add($_)
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    throw "ETABSv1.dll was not found."
}

function Get-EtabsProcess {
    param(
        [int]$RequestedPid
    )

    $processes = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) {
        return $processes | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }

    $withWindow = $processes | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($withWindow) {
        return $withWindow | Select-Object -First 1
    }

    return $processes | Select-Object -First 1
}

function Get-LogSummary {
    param(
        [string]$ModelPath
    )

    if ([string]::IsNullOrWhiteSpace($ModelPath)) {
        return $null
    }

    $logPath = [System.IO.Path]::ChangeExtension($ModelPath, ".LOG")
    if (-not (Test-Path -LiteralPath $logPath)) {
        return $null
    }

    $lines = Get-Content -LiteralPath $logPath
    $warningLines = $lines | Where-Object {
        $_ -match "WARNING" -or $_ -match "UNSTABLE" -or $_ -match "ILL-CONDITIONED" -or $_ -match "NEGATIVE EIGENVALUES"
    }

    [pscustomobject]@{
        LogPath = $logPath
        Warnings = @($warningLines)
        HasInstabilityWarning = ($lines -match "THE STRUCTURE IS UNSTABLE OR ILL-CONDITIONED").Count -gt 0
        Completed = ($lines -match "A N A L Y S I S   C O M P L E T E").Count -gt 0
    }
}

function Get-CaseStatuses {
    param(
        $SapModel
    )

    $numberItems = 0
    [string[]]$caseNames = @()
    [int[]]$statuses = @()
    $SapModel.Analyze.GetCaseStatus([ref]$numberItems, [ref]$caseNames, [ref]$statuses) | Out-Null

    $rows = @()
    for ($i = 0; $i -lt $numberItems; $i++) {
        $rows += [pscustomobject]@{
            Case = $caseNames[$i]
            Status = $statuses[$i]
        }
    }

    return $rows
}

function Get-SolverOptions {
    param(
        $SapModel
    )

    $solverType = 0
    $solverProcessType = 0
    $parallelRuns = 0
    $responseFileSizeMaxMb = 0
    $analysisThreads = 0
    $stiffCase = ""

    $SapModel.Analyze.GetSolverOption_3(
        [ref]$solverType,
        [ref]$solverProcessType,
        [ref]$parallelRuns,
        [ref]$responseFileSizeMaxMb,
        [ref]$analysisThreads,
        [ref]$stiffCase) | Out-Null

    [pscustomobject]@{
        SolverType = $solverType
        SolverProcessType = $solverProcessType
        ParallelRuns = $parallelRuns
        ResponseFileSizeMaxMb = $responseFileSizeMaxMb
        AnalysisThreads = $analysisThreads
        StiffCase = $stiffCase
    }
}

function Get-SupportSummary {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$pointNames = @()
    $SapModel.PointObj.GetNameList([ref]$numberNames, [ref]$pointNames) | Out-Null

    $supports = @()
    $weakConnectivityPoints = @()

    foreach ($pointName in $pointNames) {
        [bool[]]$restraint = @($false, $false, $false, $false, $false, $false)
        $SapModel.PointObj.GetRestraint($pointName, [ref]$restraint) | Out-Null

        $x = 0.0
        $y = 0.0
        $z = 0.0
        $SapModel.PointObj.GetCoordCartesian($pointName, [ref]$x, [ref]$y, [ref]$z, "Global") | Out-Null

        $connectedCount = 0
        [int[]]$objectTypes = @()
        [string[]]$objectNames = @()
        [int[]]$pointNumbers = @()
        $SapModel.PointObj.GetConnectivity($pointName, [ref]$connectedCount, [ref]$objectTypes, [ref]$objectNames, [ref]$pointNumbers) | Out-Null

        if ($restraint -contains $true) {
            $supports += [pscustomobject]@{
                Point = $pointName
                X = [math]::Round($x, 3)
                Y = [math]::Round($y, 3)
                Z = [math]::Round($z, 3)
                Restraint = ($restraint | ForEach-Object { if ($_){1}else{0} }) -join ""
                ConnectivityCount = $connectedCount
            }
        }

        if (-not ($restraint -contains $true) -and $connectedCount -le 1) {
            $weakConnectivityPoints += [pscustomobject]@{
                Point = $pointName
                X = [math]::Round($x, 3)
                Y = [math]::Round($y, 3)
                Z = [math]::Round($z, 3)
                ConnectivityCount = $connectedCount
                ObjectNames = $objectNames -join ","
            }
        }
    }

    [pscustomobject]@{
        SupportCount = $supports.Count
        Supports = @($supports)
        WeakConnectivityPoints = @($weakConnectivityPoints)
    }
}

function Get-FrameReleaseSummary {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$frameNames = @()
    $SapModel.FrameObj.GetNameList([ref]$numberNames, [ref]$frameNames) | Out-Null

    $releasedFrames = @()

    foreach ($frameName in $frameNames) {
        [bool[]]$iReleases = @($false, $false, $false, $false, $false, $false)
        [bool[]]$jReleases = @($false, $false, $false, $false, $false, $false)
        [double[]]$startValues = @(0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
        [double[]]$endValues = @(0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
        $SapModel.FrameObj.GetReleases($frameName, [ref]$iReleases, [ref]$jReleases, [ref]$startValues, [ref]$endValues) | Out-Null

        if (($iReleases -contains $true) -or ($jReleases -contains $true)) {
            $section = ""
            $autoSelect = ""
            $SapModel.FrameObj.GetSection($frameName, [ref]$section, [ref]$autoSelect) | Out-Null

            $pointI = ""
            $pointJ = ""
            $SapModel.FrameObj.GetPoints($frameName, [ref]$pointI, [ref]$pointJ) | Out-Null

            $label = ""
            $story = ""
            $SapModel.FrameObj.GetLabelFromName($frameName, [ref]$label, [ref]$story) | Out-Null

            $designOrientation = [ETABSv1.eFrameDesignOrientation]::Null
            $SapModel.FrameObj.GetDesignOrientation($frameName, [ref]$designOrientation) | Out-Null

            $releasedFrames += [pscustomobject]@{
                Name = $frameName
                Label = $label
                Story = $story
                Section = $section
                DesignOrientation = $designOrientation.ToString()
                PointI = $pointI
                PointJ = $pointJ
                IReleases = ($iReleases | ForEach-Object { if ($_){1}else{0} }) -join ""
                JReleases = ($jReleases | ForEach-Object { if ($_){1}else{0} }) -join ""
            }
        }
    }

    $dualEndMomentReleases = @($releasedFrames | Where-Object {
        $_.IReleases.Substring(4, 1) -eq "1" -and $_.IReleases.Substring(5, 1) -eq "1" -and
        $_.JReleases.Substring(4, 1) -eq "1" -and $_.JReleases.Substring(5, 1) -eq "1"
    })

    [pscustomobject]@{
        ReleasedFrameCount = $releasedFrames.Count
        ReleasedFrames = @($releasedFrames)
        DualEndMomentReleaseCount = $dualEndMomentReleases.Count
        DualEndMomentReleases = $dualEndMomentReleases
    }
}

function Get-FramePropertySummary {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$frameNames = @()
    $SapModel.FrameObj.GetNameList([ref]$numberNames, [ref]$frameNames) | Out-Null

    $propertySet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($frameName in $frameNames) {
        $propertyName = ""
        $autoSelect = ""
        $SapModel.FrameObj.GetSection($frameName, [ref]$propertyName, [ref]$autoSelect) | Out-Null
        if (-not [string]::IsNullOrWhiteSpace($propertyName)) {
            $propertySet.Add($propertyName) | Out-Null
        }
    }

    $suspectProperties = @()

    foreach ($propertyName in $propertySet) {
        [double[]]$modifiers = @(1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0)
        $SapModel.PropFrame.GetModifiers($propertyName, [ref]$modifiers) | Out-Null

        $area = 0.0
        $as2 = 0.0
        $as3 = 0.0
        $torsion = 0.0
        $i22 = 0.0
        $i33 = 0.0
        $s22 = 0.0
        $s33 = 0.0
        $z22 = 0.0
        $z33 = 0.0
        $r22 = 0.0
        $r33 = 0.0
        $SapModel.PropFrame.GetSectProps(
            $propertyName,
            [ref]$area,
            [ref]$as2,
            [ref]$as3,
            [ref]$torsion,
            [ref]$i22,
            [ref]$i33,
            [ref]$s22,
            [ref]$s33,
            [ref]$z22,
            [ref]$z33,
            [ref]$r22,
            [ref]$r33) | Out-Null

        $badModifiers = @()
        for ($i = 0; $i -lt $modifiers.Count; $i++) {
            if ($modifiers[$i] -le 0.0) {
                $badModifiers += "Modifier[$i]=$($modifiers[$i])"
            }
        }

        $badProps = @()
        if ($area -le 0.0) { $badProps += "Area=$area" }
        if ($torsion -le 0.0) { $badProps += "Torsion=$torsion" }
        if ($i22 -le 0.0) { $badProps += "I22=$i22" }
        if ($i33 -le 0.0) { $badProps += "I33=$i33" }

        if ($badModifiers.Count -gt 0 -or $badProps.Count -gt 0) {
            $suspectProperties += [pscustomobject]@{
                Property = $propertyName
                BadModifiers = $badModifiers -join "; "
                BadSectionProperties = $badProps -join "; "
            }
        }
    }

    [pscustomobject]@{
        SuspectFramePropertyCount = $suspectProperties.Count
        SuspectFrameProperties = @($suspectProperties)
    }
}

function Get-AreaPropertySummary {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$areaNames = @()
    $SapModel.AreaObj.GetNameList([ref]$numberNames, [ref]$areaNames) | Out-Null

    $propertySet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($areaName in $areaNames) {
        $propertyName = ""
        $SapModel.AreaObj.GetProperty($areaName, [ref]$propertyName) | Out-Null
        if (-not [string]::IsNullOrWhiteSpace($propertyName)) {
            $propertySet.Add($propertyName) | Out-Null
        }
    }

    $suspectProperties = @()

    foreach ($propertyName in $propertySet) {
        [double[]]$modifiers = @(1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0)
        $SapModel.PropArea.GetModifiers($propertyName, [ref]$modifiers) | Out-Null

        $badModifiers = @()
        for ($i = 0; $i -lt $modifiers.Count; $i++) {
            if ($modifiers[$i] -le 0.0) {
                $badModifiers += "Modifier[$i]=$($modifiers[$i])"
            }
        }

        if ($badModifiers.Count -gt 0) {
            $suspectProperties += [pscustomobject]@{
                Property = $propertyName
                BadModifiers = $badModifiers -join "; "
            }
        }
    }

    [pscustomobject]@{
        SuspectAreaPropertyCount = $suspectProperties.Count
        SuspectAreaProperties = @($suspectProperties)
    }
}

$process = Get-EtabsProcess -RequestedPid $EtabsPid
if ($null -eq $process) {
    throw "No running ETABS process was found."
}

$apiDllPath = Resolve-EtabsApiDll -Process $process
Add-Type -Path $apiDllPath

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$sapModel = $api.SapModel

$modelPath = [System.IO.Path]::Combine($sapModel.GetModelFilepath(), ([System.IO.Path]::GetFileNameWithoutExtension($sapModel.GetModelFilename($true)) + ".EDB"))

$diagnostic = [pscustomobject]@{
    ProcessId = $process.Id
    ProcessVersion = $process.MainModule.FileVersionInfo.FileVersion
    ApiDllPath = $apiDllPath
    ModelPath = $modelPath
    ModelLocked = $sapModel.GetModelIsLocked()
    SolverOptions = Get-SolverOptions -SapModel $sapModel
    CaseStatuses = Get-CaseStatuses -SapModel $sapModel
    LogSummary = Get-LogSummary -ModelPath $modelPath
    Supports = Get-SupportSummary -SapModel $sapModel
    FrameReleases = Get-FrameReleaseSummary -SapModel $sapModel
    FrameProperties = Get-FramePropertySummary -SapModel $sapModel
    AreaProperties = Get-AreaPropertySummary -SapModel $sapModel
}

if ($AsJson) {
    $diagnostic | ConvertTo-Json -Depth 8
}
else {
    $diagnostic
}
