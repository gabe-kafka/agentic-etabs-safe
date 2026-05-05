param(
    [int]$SafePid,
    [string]$KnowledgePath,
    [string]$RecordResolution,
    [string[]]$ResolutionTags,
    [switch]$RunStandardSolverCheck,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path -Path $PSScriptRoot -ChildPath "_instability-learning.ps1")

function Resolve-SafeApiDll {
    param(
        [System.Diagnostics.Process]$Process
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($null -ne $Process) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($Process.Path)) {
                $candidates.Add((Join-Path -Path (Split-Path -Parent $Process.Path) -ChildPath "SAFEv1.dll"))
            }
        }
        catch {
        }
    }

    @(
        "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll",
        "C:\Program Files\Computers and Structures\SAFE 20\SAFEv1.dll",
        "C:\Program Files\Computers and Structures\SAFE 16\SAFEv1.dll",
        "C:\Program Files (x86)\Computers and Structures\SAFE 21\SAFEv1.dll",
        "C:\Program Files (x86)\Computers and Structures\SAFE 20\SAFEv1.dll",
        "C:\Program Files (x86)\Computers and Structures\SAFE 16\SAFEv1.dll"
    ) | ForEach-Object {
        $candidates.Add($_)
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    throw "SAFEv1.dll was not found."
}

function Get-SafeProcess {
    param(
        [int]$RequestedPid
    )

    $processes = Get-Process SAFE -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) {
        return $processes | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }

    $withWindow = $processes | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($withWindow) {
        return $withWindow | Select-Object -First 1
    }

    return $processes | Select-Object -First 1
}

function Get-ProcessVersion {
    param(
        [System.Diagnostics.Process]$Process
    )

    if ($null -eq $Process) {
        return $null
    }

    try {
        $fileVersion = $Process.MainModule.FileVersionInfo
        return "$($fileVersion.FileMajorPart).$($fileVersion.FileMinorPart).$($fileVersion.FileBuildPart)"
    }
    catch {
        return $null
    }
}

function Resolve-CanonicalModelPath {
    param(
        [string]$RawModelPath,
        [string]$ModelDirectory
    )

    if ([string]::IsNullOrWhiteSpace($RawModelPath)) {
        return $null
    }

    if ([System.IO.Path]::GetExtension($RawModelPath) -ieq ".FDB") {
        return $RawModelPath
    }

    if (-not [string]::IsNullOrWhiteSpace($ModelDirectory)) {
        $candidate = Join-Path -Path $ModelDirectory -ChildPath ("{0}.FDB" -f [System.IO.Path]::GetFileNameWithoutExtension($RawModelPath))
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $RawModelPath
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
    $warningLines = @(
        $lines | Where-Object {
            $_ -match "WARNING" -or $_ -match "UNSTABLE" -or $_ -match "ILL-CONDITIONED" -or $_ -match "NEGATIVE EIGENVALUES"
        }
    )

    [pscustomobject]@{
        LogPath                = $logPath
        Warnings               = $warningLines
        HasInstabilityWarning  = ($lines -match "THE STRUCTURE IS UNSTABLE OR ILL-CONDITIONED").Count -gt 0
        AdvancedSolverInUse    = ($lines -match "USING THE ADVANCED SOLVER").Count -gt 0
        Completed              = ($lines -match "A N A L Y S I S   C O M P L E T E").Count -gt 0
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
            Case   = $caseNames[$i]
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
        SolverType            = $solverType
        SolverProcessType     = $solverProcessType
        ParallelRuns          = $parallelRuns
        ResponseFileSizeMaxMb = $responseFileSizeMaxMb
        AnalysisThreads       = $analysisThreads
        StiffCase             = $stiffCase
    }
}

function Get-PointConnectivitySummary {
    param(
        $SapModel,
        $StripEndpointIndex = $null
    )

    $numberNames = 0
    [string[]]$pointNames = @()
    $SapModel.PointObj.GetNameList([ref]$numberNames, [ref]$pointNames) | Out-Null

    $supports = @()
    $floatingPoints = @()
    $weakConnectivityPoints = @()
    $pointsWithSpring = 0
    $translationalSupportCount = 0
    $fixedLikeSupportCount = 0

    foreach ($pointName in $pointNames) {
        [bool[]]$restraint = @($false, $false, $false, $false, $false, $false)
        $SapModel.PointObj.GetRestraint($pointName, [ref]$restraint) | Out-Null

        [double[]]$spring = @(0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
        $SapModel.PointObj.GetSpring($pointName, [ref]$spring) | Out-Null

        $isSpecialPoint = $false
        $SapModel.PointObj.GetSpecialPoint($pointName, [ref]$isSpecialPoint) | Out-Null

        $isStripEndpoint = $false
        [string[]]$stripNames = @()
        if ($null -ne $StripEndpointIndex -and $StripEndpointIndex.ContainsKey($pointName)) {
            $isStripEndpoint = $true
            $stripNames = @($StripEndpointIndex[$pointName])
        }

        $x = 0.0
        $y = 0.0
        $z = 0.0
        $SapModel.PointObj.GetCoordCartesian($pointName, [ref]$x, [ref]$y, [ref]$z, "Global") | Out-Null

        $connectedCount = 0
        [int[]]$objectTypes = @()
        [string[]]$objectNames = @()
        [int[]]$pointNumbers = @()
        $SapModel.PointObj.GetConnectivity($pointName, [ref]$connectedCount, [ref]$objectTypes, [ref]$objectNames, [ref]$pointNumbers) | Out-Null

        $hasRestraint = $restraint -contains $true
        $hasSpring = [math]::Abs((($spring | Measure-Object -Sum).Sum)) -gt 0.0

        if ($hasSpring) {
            $pointsWithSpring++
        }

        if ($hasRestraint) {
            $hasTranslationalRestraint = $restraint[0] -or $restraint[1] -or $restraint[2]
            $hasRotationalRestraint = $restraint[3] -or $restraint[4] -or $restraint[5]

            if ($hasTranslationalRestraint -and -not $hasRotationalRestraint) {
                $translationalSupportCount++
            }
            elseif ($hasTranslationalRestraint -and $hasRotationalRestraint) {
                $fixedLikeSupportCount++
            }

            $supports += [pscustomobject]@{
                Point             = $pointName
                X                 = [math]::Round($x, 3)
                Y                 = [math]::Round($y, 3)
                Z                 = [math]::Round($z, 3)
                Restraint         = ($restraint | ForEach-Object { if ($_){1}else{0} }) -join ""
                ConnectivityCount = $connectedCount
                IsSpecialPoint    = $isSpecialPoint
                IsStripEndpoint   = $isStripEndpoint
                StripNames        = @($stripNames)
            }
        }

        if (-not $hasRestraint -and -not $hasSpring -and $connectedCount -eq 0) {
            $floatingPoints += [pscustomobject]@{
                Point          = $pointName
                X              = [math]::Round($x, 3)
                Y              = [math]::Round($y, 3)
                Z              = [math]::Round($z, 3)
                IsSpecialPoint = $isSpecialPoint
                IsStripEndpoint = $isStripEndpoint
                StripNames      = @($stripNames)
            }
        }

        if (-not $hasRestraint -and -not $hasSpring -and $connectedCount -le 1) {
            $weakConnectivityPoints += [pscustomobject]@{
                Point             = $pointName
                X                 = [math]::Round($x, 3)
                Y                 = [math]::Round($y, 3)
                Z                 = [math]::Round($z, 3)
                ConnectivityCount = $connectedCount
                ConnectedObjects  = $objectNames -join ","
                IsSpecialPoint    = $isSpecialPoint
                IsStripEndpoint   = $isStripEndpoint
                StripNames        = @($stripNames)
            }
        }
    }

    $floatingElevationBreakdown = @(
        $floatingPoints |
            Group-Object Z |
            Sort-Object -Property @{ Expression = "Count"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
            ForEach-Object {
                [pscustomobject]@{
                    Elevation = $_.Name
                    Count     = $_.Count
                }
            }
    )

    [pscustomobject]@{
        PointCount                 = $numberNames
        SupportCount               = $supports.Count
        TranslationalSupportCount  = $translationalSupportCount
        FixedLikeSupportCount      = $fixedLikeSupportCount
        PointsWithSpringCount      = $pointsWithSpring
        Supports                   = @($supports)
        FloatingPointCount         = $floatingPoints.Count
        FloatingPoints             = @($floatingPoints)
        SpecialFloatingPointCount  = @($floatingPoints | Where-Object { $_.IsSpecialPoint }).Count
        StripEndpointFloatingPointCount = @($floatingPoints | Where-Object { $_.IsStripEndpoint }).Count
        NonStripFloatingPointCount = @($floatingPoints | Where-Object { -not $_.IsStripEndpoint }).Count
        WeakConnectivityPointCount = $weakConnectivityPoints.Count
        WeakConnectivityPoints     = @($weakConnectivityPoints)
        FloatingElevationBreakdown = $floatingElevationBreakdown
    }
}

function Get-StripEndpointSummary {
    param(
        $SapModel
    )

    [string[]]$fieldKeyList = @()
    $tableVersion = 0
    [string[]]$fields = @()
    $numberRecords = 0
    [string[]]$tableData = @()

    $tableKey = "Strip Object Connectivity"
    $ret = $SapModel.DatabaseTables.GetTableForDisplayArray(
        $tableKey,
        [ref]$fieldKeyList,
        "",
        [ref]$tableVersion,
        [ref]$fields,
        [ref]$numberRecords,
        [ref]$tableData)

    if ($ret -ne 0 -or $numberRecords -le 0 -or $fields.Count -eq 0) {
        return [pscustomobject]@{
            StripCount              = 0
            EndpointAssignmentCount = 0
            EndpointUniquePointCount = 0
            StripRows               = @()
            EndpointIndex           = @{}
        }
    }

    $fieldCount = $fields.Count
    $nameIndex = [Array]::IndexOf($fields, "Name")
    $startIndex = [Array]::IndexOf($fields, "StartPoint")
    $endIndex = [Array]::IndexOf($fields, "EndPoint")
    $layerIndex = [Array]::IndexOf($fields, "Layer")

    $rows = @()
    $endpointIndex = @{}

    for ($rowIndex = 0; $rowIndex -lt $numberRecords; $rowIndex++) {
        $offset = $rowIndex * $fieldCount
        $name = $tableData[$offset + $nameIndex]
        $startPoint = $tableData[$offset + $startIndex]
        $endPoint = $tableData[$offset + $endIndex]
        $layer = if ($layerIndex -ge 0) { $tableData[$offset + $layerIndex] } else { $null }

        $rows += [pscustomobject]@{
            Name       = $name
            StartPoint = $startPoint
            EndPoint   = $endPoint
            Layer      = $layer
        }

        foreach ($pointName in @($startPoint, $endPoint)) {
            if ([string]::IsNullOrWhiteSpace($pointName)) {
                continue
            }

            if (-not $endpointIndex.ContainsKey($pointName)) {
                $endpointIndex[$pointName] = @()
            }

            $endpointIndex[$pointName] = @($endpointIndex[$pointName] + $name | Sort-Object -Unique)
        }
    }

    [pscustomobject]@{
        StripCount               = $rows.Count
        EndpointAssignmentCount  = $rows.Count * 2
        EndpointUniquePointCount = $endpointIndex.Keys.Count
        StripRows                = @($rows)
        EndpointIndex            = $endpointIndex
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

            $releasedFrames += [pscustomobject]@{
                Name      = $frameName
                Label     = $label
                Story     = $story
                Section   = $section
                PointI    = $pointI
                PointJ    = $pointJ
                IReleases = ($iReleases | ForEach-Object { if ($_){1}else{0} }) -join ""
                JReleases = ($jReleases | ForEach-Object { if ($_){1}else{0} }) -join ""
            }
        }
    }

    $dualEndMomentReleases = @(
        $releasedFrames | Where-Object {
            $_.IReleases.Substring(4, 1) -eq "1" -and $_.IReleases.Substring(5, 1) -eq "1" -and
            $_.JReleases.Substring(4, 1) -eq "1" -and $_.JReleases.Substring(5, 1) -eq "1"
        }
    )

    [pscustomobject]@{
        ReleasedFrameCount      = $releasedFrames.Count
        ReleasedFrames          = @($releasedFrames)
        DualEndMomentReleaseCount = $dualEndMomentReleases.Count
        DualEndMomentReleases   = $dualEndMomentReleases
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
        for ($i = 0; $i -lt [Math]::Min($modifiers.Count, 8); $i++) {
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
                Property             = $propertyName
                BadModifiers         = $badModifiers -join "; "
                BadSectionProperties = $badProps -join "; "
            }
        }
    }

    [pscustomobject]@{
        SuspectFramePropertyCount = $suspectProperties.Count
        SuspectFrameProperties    = @($suspectProperties)
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
        for ($i = 0; $i -lt [Math]::Min($modifiers.Count, 8); $i++) {
            if ($modifiers[$i] -le 0.0) {
                $badModifiers += "Modifier[$i]=$($modifiers[$i])"
            }
        }

        if ($badModifiers.Count -gt 0) {
            $suspectProperties += [pscustomobject]@{
                Property     = $propertyName
                BadModifiers = $badModifiers -join "; "
            }
        }
    }

    [pscustomobject]@{
        SuspectAreaPropertyCount = $suspectProperties.Count
        SuspectAreaProperties    = @($suspectProperties)
    }
}

function Get-ConstraintSummary {
    param(
        $SapModel
    )

    $numberPoints = 0
    [string[]]$pointNames = @()
    $SapModel.PointObj.GetNameList([ref]$numberPoints, [ref]$pointNames) | Out-Null

    $numberConstraints = 0
    [string[]]$constraintNames = @()
    $SapModel.ConstraintDef.GetNameList([ref]$numberConstraints, [ref]$constraintNames) | Out-Null

    $constraints = @()

    foreach ($constraintName in $constraintNames) {
        $axis = [SAFEv1.eConstraintAxis]::AutoAxis
        $coordinateSystem = ""
        $definitionRet = $SapModel.ConstraintDef.GetDiaphragm($constraintName, [ref]$axis, [ref]$coordinateSystem)

        $assignedPoints = @()
        foreach ($pointName in $pointNames) {
            $diaphragmOption = [SAFEv1.eDiaphragmOption]::Disconnect
            $diaphragmName = ""
            $SapModel.PointObj.GetDiaphragm($pointName, [ref]$diaphragmOption, [ref]$diaphragmName) | Out-Null

            if ($diaphragmOption -eq [SAFEv1.eDiaphragmOption]::DefinedDiaphragm -and
                [string]::Equals($diaphragmName, $constraintName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $assignedPoints += $pointName
            }
        }

        $constraints += [pscustomobject]@{
            Name               = $constraintName
            Type               = if ($definitionRet -eq 0) { "Diaphragm" } else { "Unknown" }
            Axis               = if ($definitionRet -eq 0) { $axis.ToString() } else { $null }
            CoordinateSystem   = $coordinateSystem
            AssignedPointCount = $assignedPoints.Count
            AssignedPoints     = $assignedPoints
        }
    }

    [pscustomobject]@{
        ConstraintCount = $constraints.Count
        Constraints     = @($constraints)
    }
}

function Get-AreaAspectRatioSummary {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$areaNames = @()
    $SapModel.AreaObj.GetNameList([ref]$numberNames, [ref]$areaNames) | Out-Null

    $badAspect = @()

    foreach ($areaName in $areaNames) {
        $numberPoints = 0
        [string[]]$pointNames = @()
        $SapModel.AreaObj.GetPoints($areaName, [ref]$numberPoints, [ref]$pointNames) | Out-Null

        if ($numberPoints -lt 3) {
            continue
        }

        $xs = @()
        $ys = @()

        foreach ($pointName in $pointNames) {
            $x = 0.0
            $y = 0.0
            $z = 0.0
            $SapModel.PointObj.GetCoordCartesian($pointName, [ref]$x, [ref]$y, [ref]$z, "Global") | Out-Null
            $xs += $x
            $ys += $y
        }

        $dx = ($xs | Measure-Object -Maximum).Maximum - ($xs | Measure-Object -Minimum).Minimum
        $dy = ($ys | Measure-Object -Maximum).Maximum - ($ys | Measure-Object -Minimum).Minimum
        $minDimension = [Math]::Min($dx, $dy)
        $maxDimension = [Math]::Max($dx, $dy)

        if ($minDimension -gt 0.0 -and ($maxDimension / $minDimension) -gt 20.0) {
            $badAspect += [pscustomobject]@{
                Area  = $areaName
                Ratio = [math]::Round(($maxDimension / $minDimension), 2)
            }
        }
    }

    [pscustomobject]@{
        BadAspectCount = $badAspect.Count
        BadAspectAreas = @($badAspect | Sort-Object -Property @{ Expression = "Ratio"; Descending = $true })
    }
}

function Get-StandardSolverProbe {
    param(
        $SapModel,
        [string]$ModelPath,
        [switch]$Enabled
    )

    if (-not $Enabled) {
        return [pscustomobject]@{
            Executed                   = $false
            OriginalSolverType         = $null
            RunReturn                  = $null
            LogPath                    = [System.IO.Path]::ChangeExtension($ModelPath, ".LOG")
            ZeroDiagonalRowCount       = 0
            ZeroDiagonalRows           = @()
            ZeroDiagonalUniqueJoints   = @()
            ZeroDiagonalUniqueJointCount = 0
            LossOfAccuracyRowCount     = 0
            LossOfAccuracyRows         = @()
        }
    }

    $logPath = [System.IO.Path]::ChangeExtension($ModelPath, ".LOG")
    $originalLocked = $SapModel.GetModelIsLocked()
    $originalSolverType = 0
    $originalForce32BitSolver = $false
    $originalStiffCase = ""
    $SapModel.Analyze.GetSolverOption([ref]$originalSolverType, [ref]$originalForce32BitSolver, [ref]$originalStiffCase) | Out-Null

    $unlockRet = $SapModel.SetModelIsLocked($false)
    if ($unlockRet -ne 0) {
        throw "SetModelIsLocked(false) returned $unlockRet."
    }

    try {
        $setRet = $SapModel.Analyze.SetSolverOption(0, $false, $originalStiffCase)
        if ($setRet -ne 0) {
            throw "SetSolverOption(0, ...) returned $setRet."
        }

        $runRet = $SapModel.Analyze.RunAnalysis()
        Start-Sleep -Seconds 1

        $zeroDiagonalRows = @()
        $lossOfAccuracyRows = @()

        if (Test-Path -LiteralPath $logPath) {
            foreach ($line in (Get-Content -LiteralPath $logPath)) {
                if ($line -match '^\s*Joint\s+(\S+)\s+(\S+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+(.*)$') {
                    $row = [pscustomobject]@{
                        Label = $matches[1]
                        DOF   = $matches[2]
                        X     = [double]$matches[3]
                        Y     = [double]$matches[4]
                        Z     = [double]$matches[5]
                        Detail = $matches[6].Trim()
                        Raw   = $line.Trim()
                    }

                    if ($line -match 'Diag = 0') {
                        $zeroDiagonalRows += $row
                    }
                    elseif ($line -match 'Lost accuracy') {
                        $lossOfAccuracyRows += $row
                    }
                }
            }
        }

        $zeroDiagonalUniqueJoints = @(
            $zeroDiagonalRows |
                Where-Object { $_.Label -notlike "~*" } |
                ForEach-Object { $_.Label } |
                Sort-Object -Unique
        )

        return [pscustomobject]@{
            Executed                     = $true
            OriginalSolverType           = $originalSolverType
            RunReturn                    = $runRet
            LogPath                      = $logPath
            ZeroDiagonalRowCount         = $zeroDiagonalRows.Count
            ZeroDiagonalRows             = @($zeroDiagonalRows)
            ZeroDiagonalUniqueJoints     = $zeroDiagonalUniqueJoints
            ZeroDiagonalUniqueJointCount = $zeroDiagonalUniqueJoints.Count
            LossOfAccuracyRowCount       = $lossOfAccuracyRows.Count
            LossOfAccuracyRows           = @($lossOfAccuracyRows)
        }
    }
    finally {
        $null = $SapModel.SetModelIsLocked($false)
        $null = $SapModel.Analyze.SetSolverOption($originalSolverType, $originalForce32BitSolver, $originalStiffCase)
        if ($originalLocked) {
            $null = $SapModel.SetModelIsLocked($true)
        }
    }
}

function Get-DiagnosticFindings {
    param(
        $Diagnostic
    )

    $findings = New-Object System.Collections.Generic.List[pscustomobject]

    if ($Diagnostic.LogSummary -and $Diagnostic.LogSummary.HasInstabilityWarning) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Solver" -Signature "ill-conditioned-log" -Detail "Analysis log reports that the structure is unstable or ill-conditioned." -Evidence $Diagnostic.LogSummary.LogPath)) | Out-Null
    }

    if ($Diagnostic.Supports.SupportCount -eq 0) {
        $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Supports" -Signature "no-supports" -Detail "No restrained points were found in the SAFE object model.")) | Out-Null
    }
    elseif ($Diagnostic.Supports.FixedLikeSupportCount -eq 0 -and $Diagnostic.Supports.TranslationalSupportCount -lt 3) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Supports" -Signature "light-support-pattern" -Detail ("Only {0} translational support point(s) were found and none are fixed-like." -f $Diagnostic.Supports.TranslationalSupportCount))) | Out-Null
    }

    if ($Diagnostic.StandardSolverProbe.Executed -and $Diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJointCount -gt 0) {
        $stripEndpointZeroDiagonal = @(
            $Diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJoints |
                Where-Object { $Diagnostic.StripEndpoints.EndpointIndex.ContainsKey($_) }
        )

        $nonStripZeroDiagonal = @(
            $Diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJoints |
                Where-Object { -not $Diagnostic.StripEndpoints.EndpointIndex.ContainsKey($_) }
        )

        if ($stripEndpointZeroDiagonal.Count -gt 0) {
            $detail = if ($nonStripZeroDiagonal.Count -eq 0) {
                ("Standard solver found {0} unique zero-diagonal joints, and all of them are strip endpoint point objects." -f $Diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJointCount)
            }
            else {
                ("Standard solver found {0} zero-diagonal strip endpoint joint(s) plus {1} additional non-strip joint(s)." -f $stripEndpointZeroDiagonal.Count, $nonStripZeroDiagonal.Count)
            }

            $evidenceRows = @(
                $Diagnostic.StandardSolverProbe.ZeroDiagonalRows |
                    Where-Object { $Diagnostic.StripEndpoints.EndpointIndex.ContainsKey($_.Label) } |
                    Select-Object -First 12
            )

            $stripSamples = @(
                $stripEndpointZeroDiagonal |
                    Select-Object -First 8 |
                    ForEach-Object {
                        [pscustomobject]@{
                            Point  = $_
                            Strips = @($Diagnostic.StripEndpoints.EndpointIndex[$_])
                        }
                    }
            )

            $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Strip Endpoint Joints" -Signature "strip-endpoint-zero-diagonal" -Detail $detail -Evidence ([pscustomobject]@{
                ZeroDiagonalRows = $evidenceRows
                StripSamples     = $stripSamples
            }))) | Out-Null
        }

        if ($nonStripZeroDiagonal.Count -gt 0) {
            $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Connectivity" -Signature "nonstrip-zero-diagonal-joints" -Detail ("Standard solver found {0} zero-diagonal joint(s) that are not strip endpoints." -f $nonStripZeroDiagonal.Count) -Evidence (@($nonStripZeroDiagonal | Select-Object -First 12)))) | Out-Null
        }
    }
    elseif ($Diagnostic.Supports.NonStripFloatingPointCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Connectivity" -Signature "floating-points" -Detail ("{0} non-strip point object(s) have zero connectivity, no restraint, and no spring." -f $Diagnostic.Supports.NonStripFloatingPointCount) -Evidence (@($Diagnostic.Supports.FloatingPoints | Where-Object { -not $_.IsStripEndpoint } | Select-Object -First 12)))) | Out-Null
    }

    if ($Diagnostic.Supports.StripEndpointFloatingPointCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Strip Endpoint Joints" -Signature "strip-endpoint-point-objects" -Detail ("{0} strip endpoint point object(s) have zero connectivity, no restraint, and no spring." -f $Diagnostic.Supports.StripEndpointFloatingPointCount) -Evidence (@($Diagnostic.Supports.FloatingPoints | Where-Object { $_.IsStripEndpoint } | Select-Object -First 12)))) | Out-Null
    }

    if ($Diagnostic.StandardSolverProbe.Executed -and $Diagnostic.StandardSolverProbe.LossOfAccuracyRowCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Numerics" -Signature "loss-of-accuracy-joint" -Detail ("Standard solver reported loss of accuracy at {0} joint/DOF row(s)." -f $Diagnostic.StandardSolverProbe.LossOfAccuracyRowCount) -Evidence (@($Diagnostic.StandardSolverProbe.LossOfAccuracyRows | Select-Object -First 10)))) | Out-Null
    }

    if ($Diagnostic.FrameReleases.DualEndMomentReleaseCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Frame Release" -Signature "dual-end-moment-release" -Detail ("{0} frame object(s) release M2 and M3 at both ends, which creates a local mechanism." -f $Diagnostic.FrameReleases.DualEndMomentReleaseCount) -Evidence (@($Diagnostic.FrameReleases.DualEndMomentReleases | Select-Object -First 10)))) | Out-Null
    }

    if ($Diagnostic.FrameProperties.SuspectFramePropertyCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Frame Property" -Signature "nonpositive-frame-property" -Detail ("{0} frame property definition(s) contain nonpositive stiffness terms or modifiers." -f $Diagnostic.FrameProperties.SuspectFramePropertyCount) -Evidence (@($Diagnostic.FrameProperties.SuspectFrameProperties | Select-Object -First 10)))) | Out-Null
    }

    if ($Diagnostic.AreaProperties.SuspectAreaPropertyCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Area Property" -Signature "nonpositive-area-modifier" -Detail ("{0} area property definition(s) contain nonpositive modifiers." -f $Diagnostic.AreaProperties.SuspectAreaPropertyCount) -Evidence (@($Diagnostic.AreaProperties.SuspectAreaProperties | Select-Object -First 10)))) | Out-Null
    }

    $singlePointConstraints = @($Diagnostic.Constraints.Constraints | Where-Object { $_.AssignedPointCount -eq 1 })
    if ($singlePointConstraints.Count -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "ERROR" -Category "Constraint" -Signature "single-point-constraint" -Detail ("{0} constraint definition(s) are assigned to exactly one point." -f $singlePointConstraints.Count) -Evidence (@($singlePointConstraints | Select-Object -First 10)))) | Out-Null
    }

    $unusedConstraints = @($Diagnostic.Constraints.Constraints | Where-Object { $_.AssignedPointCount -eq 0 })
    if ($unusedConstraints.Count -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "INFO" -Category "Constraint" -Signature "unused-constraint-definition" -Detail ("{0} constraint definition(s) are defined but not assigned to any points." -f $unusedConstraints.Count) -Evidence (@($unusedConstraints | Select-Object -First 10)))) | Out-Null
    }

    if ($Diagnostic.AreaAspectRatios.BadAspectCount -gt 0) {
        $findings.Add((New-InstabilityFinding -Severity "WARN" -Category "Mesh" -Signature "high-area-aspect-ratio" -Detail ("{0} area object(s) have a bounding-box aspect ratio greater than 20." -f $Diagnostic.AreaAspectRatios.BadAspectCount) -Evidence (@($Diagnostic.AreaAspectRatios.BadAspectAreas | Select-Object -First 10)))) | Out-Null
    }

    return @($findings)
}

function Get-LimitedRows {
    param(
        [object[]]$Rows,
        [int]$Limit = 25
    )

    return @(@($Rows) | Select-Object -First $Limit)
}

$process = Get-SafeProcess -RequestedPid $SafePid
if ($null -eq $process) {
    throw "No running SAFE process was found."
}

$apiDllPath = Resolve-SafeApiDll -Process $process
Add-Type -Path $apiDllPath

$helper = [SAFEv1.cHelper](New-Object SAFEv1.Helper)
$api = $helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", $process.Id)
$sapModel = $api.SapModel

$rawModelPath = $sapModel.GetModelFilename($true)
$modelDirectory = $sapModel.GetModelFilepath()
$modelPath = Resolve-CanonicalModelPath -RawModelPath $rawModelPath -ModelDirectory $modelDirectory
$stripEndpointSummary = Get-StripEndpointSummary -SapModel $sapModel

$diagnostic = [pscustomobject]@{
    ProcessId        = $process.Id
    ProcessVersion   = Get-ProcessVersion -Process $process
    ApiVersion       = $helper.GetOAPIVersionNumber()
    ApiDllPath       = $apiDllPath
    ModelPath        = $modelPath
    ModelLocked      = $sapModel.GetModelIsLocked()
    SolverOptions    = Get-SolverOptions -SapModel $sapModel
    CaseStatuses     = Get-CaseStatuses -SapModel $sapModel
    LogSummary       = Get-LogSummary -ModelPath $modelPath
    StripEndpoints   = $stripEndpointSummary
    Supports         = Get-PointConnectivitySummary -SapModel $sapModel -StripEndpointIndex $stripEndpointSummary.EndpointIndex
    FrameReleases    = Get-FrameReleaseSummary -SapModel $sapModel
    FrameProperties  = Get-FramePropertySummary -SapModel $sapModel
    AreaProperties   = Get-AreaPropertySummary -SapModel $sapModel
    Constraints      = Get-ConstraintSummary -SapModel $sapModel
    AreaAspectRatios = Get-AreaAspectRatioSummary -SapModel $sapModel
    StandardSolverProbe = Get-StandardSolverProbe -SapModel $sapModel -ModelPath $modelPath -Enabled:$RunStandardSolverCheck
}

$findings = @(Get-DiagnosticFindings -Diagnostic $diagnostic)
$knowledgePath = Resolve-InstabilityKnowledgePath -KnowledgePath $KnowledgePath
$knowledgeStore = Import-InstabilityKnowledgeStore -Path $knowledgePath
Add-InstabilityObservation -Store $knowledgeStore -Platform "SAFE" -ModelPath $modelPath -Findings $findings
Add-InstabilityResolution -Store $knowledgeStore -Platform "SAFE" -ModelPath $modelPath -Findings $findings -Resolution $RecordResolution -Tags $ResolutionTags
Export-InstabilityKnowledgeStore -Store $knowledgeStore -Path $knowledgePath

$mostLikelyBug = $findings | Sort-Object { Get-InstabilitySeverityRank -Severity $_.Severity } -Descending | Select-Object -First 1
$learningSummary = New-InstabilityLearningSummary -Store $knowledgeStore -KnowledgePath $knowledgePath -Platform "SAFE" -Findings $findings

$result = [pscustomobject]@{
    ProcessId        = $diagnostic.ProcessId
    ProcessVersion   = $diagnostic.ProcessVersion
    ApiVersion       = $diagnostic.ApiVersion
    ApiDllPath       = $diagnostic.ApiDllPath
    ModelPath        = $diagnostic.ModelPath
    ModelLocked      = $diagnostic.ModelLocked
    SolverOptions    = $diagnostic.SolverOptions
    CaseStatuses     = $diagnostic.CaseStatuses
    LogSummary       = $diagnostic.LogSummary
    Supports         = [pscustomobject]@{
        PointCount                 = $diagnostic.Supports.PointCount
        SupportCount               = $diagnostic.Supports.SupportCount
        TranslationalSupportCount  = $diagnostic.Supports.TranslationalSupportCount
        FixedLikeSupportCount      = $diagnostic.Supports.FixedLikeSupportCount
        PointsWithSpringCount      = $diagnostic.Supports.PointsWithSpringCount
        FloatingPointCount         = $diagnostic.Supports.FloatingPointCount
        SpecialFloatingPointCount  = $diagnostic.Supports.SpecialFloatingPointCount
        StripEndpointFloatingPointCount = $diagnostic.Supports.StripEndpointFloatingPointCount
        NonStripFloatingPointCount = $diagnostic.Supports.NonStripFloatingPointCount
        WeakConnectivityPointCount = $diagnostic.Supports.WeakConnectivityPointCount
        FloatingElevationBreakdown = $diagnostic.Supports.FloatingElevationBreakdown
        SupportSamples             = Get-LimitedRows -Rows $diagnostic.Supports.Supports -Limit 15
        FloatingPoints             = Get-LimitedRows -Rows $diagnostic.Supports.FloatingPoints -Limit 25
        WeakConnectivityPoints     = Get-LimitedRows -Rows $diagnostic.Supports.WeakConnectivityPoints -Limit 25
    }
    StripEndpoints   = [pscustomobject]@{
        StripCount               = $diagnostic.StripEndpoints.StripCount
        EndpointAssignmentCount  = $diagnostic.StripEndpoints.EndpointAssignmentCount
        EndpointUniquePointCount = $diagnostic.StripEndpoints.EndpointUniquePointCount
        StripRows                = Get-LimitedRows -Rows $diagnostic.StripEndpoints.StripRows -Limit 25
    }
    FrameReleases    = [pscustomobject]@{
        ReleasedFrameCount        = $diagnostic.FrameReleases.ReleasedFrameCount
        DualEndMomentReleaseCount = $diagnostic.FrameReleases.DualEndMomentReleaseCount
        ReleasedFrames            = Get-LimitedRows -Rows $diagnostic.FrameReleases.ReleasedFrames -Limit 25
        DualEndMomentReleases     = Get-LimitedRows -Rows $diagnostic.FrameReleases.DualEndMomentReleases -Limit 25
    }
    FrameProperties  = [pscustomobject]@{
        SuspectFramePropertyCount = $diagnostic.FrameProperties.SuspectFramePropertyCount
        SuspectFrameProperties    = Get-LimitedRows -Rows $diagnostic.FrameProperties.SuspectFrameProperties -Limit 25
    }
    AreaProperties   = [pscustomobject]@{
        SuspectAreaPropertyCount = $diagnostic.AreaProperties.SuspectAreaPropertyCount
        SuspectAreaProperties    = Get-LimitedRows -Rows $diagnostic.AreaProperties.SuspectAreaProperties -Limit 25
    }
    Constraints      = [pscustomobject]@{
        ConstraintCount = $diagnostic.Constraints.ConstraintCount
        Constraints     = @(
            @($diagnostic.Constraints.Constraints) |
                ForEach-Object {
                    [pscustomobject]@{
                        Name               = $_.Name
                        Type               = $_.Type
                        Axis               = $_.Axis
                        CoordinateSystem   = $_.CoordinateSystem
                        AssignedPointCount = $_.AssignedPointCount
                        AssignedPoints     = Get-LimitedRows -Rows $_.AssignedPoints -Limit 25
                    }
                }
        )
    }
    AreaAspectRatios = [pscustomobject]@{
        BadAspectCount = $diagnostic.AreaAspectRatios.BadAspectCount
        BadAspectAreas = Get-LimitedRows -Rows $diagnostic.AreaAspectRatios.BadAspectAreas -Limit 25
    }
    StandardSolverProbe = [pscustomobject]@{
        Executed                     = $diagnostic.StandardSolverProbe.Executed
        OriginalSolverType           = $diagnostic.StandardSolverProbe.OriginalSolverType
        RunReturn                    = $diagnostic.StandardSolverProbe.RunReturn
        LogPath                      = $diagnostic.StandardSolverProbe.LogPath
        ZeroDiagonalRowCount         = $diagnostic.StandardSolverProbe.ZeroDiagonalRowCount
        ZeroDiagonalUniqueJointCount = $diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJointCount
        ZeroDiagonalUniqueJoints     = Get-LimitedRows -Rows $diagnostic.StandardSolverProbe.ZeroDiagonalUniqueJoints -Limit 50
        ZeroDiagonalRows             = Get-LimitedRows -Rows $diagnostic.StandardSolverProbe.ZeroDiagonalRows -Limit 25
        LossOfAccuracyRowCount       = $diagnostic.StandardSolverProbe.LossOfAccuracyRowCount
        LossOfAccuracyRows           = Get-LimitedRows -Rows $diagnostic.StandardSolverProbe.LossOfAccuracyRows -Limit 25
    }
    Findings         = $findings
    MostLikelyBug    = $mostLikelyBug
    Learning         = $learningSummary
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
}
else {
    $result
}
