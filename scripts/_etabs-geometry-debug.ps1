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

function Connect-EtabsSession {
    param(
        [int]$EtabsPid
    )

    $process = Get-EtabsProcess -RequestedPid $EtabsPid
    if ($null -eq $process) {
        throw "No running ETABS process was found."
    }

    $apiDllPath = Resolve-EtabsApiDll -Process $process
    Add-Type -Path $apiDllPath

    $helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
    $api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
    $sapModel = $api.SapModel

    $rawModelPath = $sapModel.GetModelFilename($true)
    $modelDirectory = $sapModel.GetModelFilepath()
    $modelPath = if (-not [string]::IsNullOrWhiteSpace($modelDirectory) -and -not [string]::IsNullOrWhiteSpace($rawModelPath)) {
        $candidate = Join-Path -Path $modelDirectory -ChildPath ("{0}.EDB" -f [System.IO.Path]::GetFileNameWithoutExtension($rawModelPath))
        if (Test-Path -LiteralPath $candidate) {
            (Resolve-Path -LiteralPath $candidate).Path
        }
        else {
            $rawModelPath
        }
    }
    else {
        $rawModelPath
    }

    [pscustomobject]@{
        Process = $process
        ApiDllPath = $apiDllPath
        Helper = $helper
        Api = $api
        SapModel = $sapModel
        ModelPath = $modelPath
    }
}

function Get-DefaultMeshingThresholds {
    param(
        [int]$PresentUnitsEnum
    )

    switch ($PresentUnitsEnum) {
        1 { return [pscustomobject]@{ ShortEdge = 6.0; DuplicateVertex = 1.0; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 12.0; MarkerOffset = 3.0 } }
        2 { return [pscustomobject]@{ ShortEdge = 0.5; DuplicateVertex = 0.08; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 1.0; MarkerOffset = 0.25 } }
        3 { return [pscustomobject]@{ ShortEdge = 6.0; DuplicateVertex = 1.0; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 12.0; MarkerOffset = 3.0 } }
        4 { return [pscustomobject]@{ ShortEdge = 0.5; DuplicateVertex = 0.08; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 1.0; MarkerOffset = 0.25 } }
        5 { return [pscustomobject]@{ ShortEdge = 150.0; DuplicateVertex = 25.0; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 300.0; MarkerOffset = 75.0 } }
        6 { return [pscustomobject]@{ ShortEdge = 15.0; DuplicateVertex = 2.5; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 30.0; MarkerOffset = 7.5 } }
        7 { return [pscustomobject]@{ ShortEdge = 0.15; DuplicateVertex = 0.025; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 0.30; MarkerOffset = 0.075 } }
        8 { return [pscustomobject]@{ ShortEdge = 150.0; DuplicateVertex = 25.0; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 300.0; MarkerOffset = 75.0 } }
        9 { return [pscustomobject]@{ ShortEdge = 15.0; DuplicateVertex = 2.5; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 30.0; MarkerOffset = 7.5 } }
        10 { return [pscustomobject]@{ ShortEdge = 0.15; DuplicateVertex = 0.025; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 0.30; MarkerOffset = 0.075 } }
        default { return [pscustomobject]@{ ShortEdge = 6.0; DuplicateVertex = 1.0; SliverAngleDeg = 10.0; CollinearAngleToleranceDeg = 1.0; RedundantEdge = 12.0; MarkerOffset = 3.0 } }
    }
}

function Get-InteriorAngleDegrees {
    param(
        [pscustomobject]$PointA,
        [pscustomobject]$PointB,
        [pscustomobject]$PointC
    )

    $v1x = $PointA.X - $PointB.X
    $v1y = $PointA.Y - $PointB.Y
    $v1z = $PointA.Z - $PointB.Z
    $v2x = $PointC.X - $PointB.X
    $v2y = $PointC.Y - $PointB.Y
    $v2z = $PointC.Z - $PointB.Z
    $len1 = [Math]::Sqrt(($v1x * $v1x) + ($v1y * $v1y) + ($v1z * $v1z))
    $len2 = [Math]::Sqrt(($v2x * $v2x) + ($v2y * $v2y) + ($v2z * $v2z))

    if ($len1 -le 0.0 -or $len2 -le 0.0) {
        return $null
    }

    $dot = ($v1x * $v2x) + ($v1y * $v2y) + ($v1z * $v2z)
    $cosine = $dot / ($len1 * $len2)
    $cosine = [Math]::Max(-1.0, [Math]::Min(1.0, $cosine))
    return [Math]::Acos($cosine) * (180.0 / [Math]::PI)
}

function Get-PointDistance {
    param(
        [pscustomobject]$PointA,
        [pscustomobject]$PointB
    )

    $dx = $PointB.X - $PointA.X
    $dy = $PointB.Y - $PointA.Y
    $dz = $PointB.Z - $PointA.Z
    return [Math]::Sqrt(($dx * $dx) + ($dy * $dy) + ($dz * $dz))
}

function Get-EdgeMidpoint {
    param(
        [pscustomobject]$PointA,
        [pscustomobject]$PointB
    )

    [pscustomobject]@{
        X = ($PointA.X + $PointB.X) / 2.0
        Y = ($PointA.Y + $PointB.Y) / 2.0
        Z = ($PointA.Z + $PointB.Z) / 2.0
    }
}

function New-MarkerCoordinate {
    param(
        [double]$X,
        [double]$Y,
        [double]$Z,
        [string]$Kind
    )

    [pscustomobject]@{
        X = $X
        Y = $Y
        Z = $Z
        Kind = $Kind
    }
}

function Parse-EtabsMeshingWarningTargets {
    param(
        [string]$WarningText,
        [string]$WarningTextPath
    )

    $raw = $null
    if (-not [string]::IsNullOrWhiteSpace($WarningText)) {
        $raw = $WarningText
    }
    elseif (-not [string]::IsNullOrWhiteSpace($WarningTextPath)) {
        if (-not (Test-Path -LiteralPath $WarningTextPath -PathType Leaf)) {
            throw "Warning text path '$WarningTextPath' does not exist."
        }

        $raw = Get-Content -LiteralPath $WarningTextPath -Raw
    }

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return @()
    }

    $matches = [regex]::Matches($raw, 'At\s+(?<Story>[^\r\n]+?)\s+area\s+(?<Label>[^\r\n(]+?)\s+\(')
    $targets = New-Object System.Collections.ArrayList
    $seen = New-Object System.Collections.Generic.HashSet[string]

    foreach ($match in $matches) {
        $story = $match.Groups["Story"].Value.Trim()
        $label = $match.Groups["Label"].Value.Trim()
        if ([string]::IsNullOrWhiteSpace($story) -or [string]::IsNullOrWhiteSpace($label)) {
            continue
        }

        $key = ("{0}|{1}" -f $story.ToUpperInvariant(), $label.ToUpperInvariant())
        if (-not $seen.Add($key)) {
            continue
        }

        $targets.Add([pscustomobject]@{
            Story = $story
            Label = $label
        }) | Out-Null
    }

    return @($targets)
}

function Get-EtabsAreaObjects {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$areaNames = @()
    $SapModel.AreaObj.GetNameList([ref]$numberNames, [ref]$areaNames) | Out-Null

    $areas = New-Object System.Collections.ArrayList

    foreach ($areaName in $areaNames) {
        $label = ""
        $story = ""
        $SapModel.AreaObj.GetLabelFromName($areaName, [ref]$label, [ref]$story) | Out-Null

        $property = ""
        $SapModel.AreaObj.GetProperty($areaName, [ref]$property) | Out-Null

        $designOrientation = [ETABSv1.eAreaDesignOrientation]::Null
        $SapModel.AreaObj.GetDesignOrientation($areaName, [ref]$designOrientation) | Out-Null
        if ($designOrientation -notin @([ETABSv1.eAreaDesignOrientation]::Floor, [ETABSv1.eAreaDesignOrientation]::Wall, [ETABSv1.eAreaDesignOrientation]::Ramp_DO_NOT_USE)) {
            continue
        }

        $numberPoints = 0
        [string[]]$pointNames = @()
        $SapModel.AreaObj.GetPoints($areaName, [ref]$numberPoints, [ref]$pointNames) | Out-Null
        if ($numberPoints -lt 3) {
            continue
        }

        $points = New-Object System.Collections.ArrayList
        foreach ($pointName in $pointNames) {
            $x = 0.0
            $y = 0.0
            $z = 0.0
            $SapModel.PointObj.GetCoordCartesian($pointName, [ref]$x, [ref]$y, [ref]$z, "Global") | Out-Null
            $points.Add([pscustomobject]@{
                Point = $pointName
                X = $x
                Y = $y
                Z = $z
            }) | Out-Null
        }

        $areas.Add([pscustomobject]@{
            Name = $areaName
            Label = $label
            Story = $story
            Property = $property
            DesignOrientation = $designOrientation.ToString()
            Points = @($points)
        }) | Out-Null
    }

    return @($areas)
}

function New-AreaGeometryFinding {
    param(
        [string]$Severity,
        [string]$Category,
        [string]$AreaName,
        [string]$Label,
        [string]$Story,
        [string]$Property,
        [string]$DesignOrientation,
        [string]$Detail,
        [double]$MetricValue,
        [double]$Threshold,
        [int]$PointIndex,
        [int]$EdgeIndex,
        [pscustomobject[]]$Points,
        [pscustomobject[]]$MarkerCoordinates,
        [bool]$MatchesWarningTarget
    )

    [pscustomobject]@{
        Severity = $Severity
        Category = $Category
        AreaName = $AreaName
        Label = $Label
        Story = $Story
        Property = $Property
        DesignOrientation = $DesignOrientation
        Detail = $Detail
        MetricValue = [Math]::Round($MetricValue, 6)
        Threshold = [Math]::Round($Threshold, 6)
        PointIndex = $PointIndex
        EdgeIndex = $EdgeIndex
        Points = @($Points | ForEach-Object {
            [pscustomobject]@{
                Point = $_.Point
                X = [Math]::Round($_.X, 4)
                Y = [Math]::Round($_.Y, 4)
                Z = [Math]::Round($_.Z, 4)
            }
        })
        MarkerCoordinates = @($MarkerCoordinates | ForEach-Object {
            [pscustomobject]@{
                X = [Math]::Round($_.X, 4)
                Y = [Math]::Round($_.Y, 4)
                Z = [Math]::Round($_.Z, 4)
                Kind = $_.Kind
            }
        })
        MatchesWarningTarget = $MatchesWarningTarget
    }
}

function Get-EtabsMeshingFindings {
    param(
        $SapModel,
        [double]$ShortEdgeThreshold,
        [double]$DuplicateVertexThreshold,
        [double]$SliverAngleDeg,
        [double]$CollinearAngleToleranceDeg,
        [double]$RedundantEdgeThreshold,
        [pscustomobject[]]$WarningTargets
    )

    $warningKeys = New-Object System.Collections.Generic.HashSet[string]
    foreach ($target in @($WarningTargets)) {
        $warningKeys.Add(("{0}|{1}" -f $target.Story.ToUpperInvariant(), $target.Label.ToUpperInvariant())) | Out-Null
    }

    $areas = Get-EtabsAreaObjects -SapModel $SapModel
    $findings = New-Object System.Collections.ArrayList
    $areaSummaries = New-Object System.Collections.ArrayList

    foreach ($area in $areas) {
        $points = @($area.Points)
        $pointCount = $points.Count
        $areaKey = ("{0}|{1}" -f $area.Story.ToUpperInvariant(), $area.Label.ToUpperInvariant())
        $matchesWarning = $warningKeys.Contains($areaKey)
        $areaFindings = New-Object System.Collections.ArrayList

        for ($i = 0; $i -lt $pointCount; $i++) {
            $current = $points[$i]
            $next = $points[($i + 1) % $pointCount]
            $edgeLength = Get-PointDistance -PointA $current -PointB $next
            $midpoint = Get-EdgeMidpoint -PointA $current -PointB $next

            if ($edgeLength -le $DuplicateVertexThreshold) {
                $severity = if ($edgeLength -le ($DuplicateVertexThreshold / 4.0)) { "ERROR" } else { "WARN" }
                $areaFindings.Add((New-AreaGeometryFinding -Severity $severity -Category "NearDuplicateVertex" -AreaName $area.Name -Label $area.Label -Story $area.Story -Property $area.Property -DesignOrientation $area.DesignOrientation -Detail ("Edge {0} is effectively collapsed at length {1}." -f ($i + 1), [Math]::Round($edgeLength, 4)) -MetricValue $edgeLength -Threshold $DuplicateVertexThreshold -PointIndex ($i + 1) -EdgeIndex ($i + 1) -Points @($current, $next) -MarkerCoordinates @((New-MarkerCoordinate -X $current.X -Y $current.Y -Z $current.Z -Kind "Vertex"), (New-MarkerCoordinate -X $next.X -Y $next.Y -Z $next.Z -Kind "Vertex"), (New-MarkerCoordinate -X $midpoint.X -Y $midpoint.Y -Z $midpoint.Z -Kind "Midpoint")) -MatchesWarningTarget $matchesWarning)) | Out-Null
            }
            elseif ($edgeLength -le $ShortEdgeThreshold) {
                $severity = if ($edgeLength -le ($ShortEdgeThreshold / 3.0)) { "ERROR" } else { "WARN" }
                $areaFindings.Add((New-AreaGeometryFinding -Severity $severity -Category "ShortEdge" -AreaName $area.Name -Label $area.Label -Story $area.Story -Property $area.Property -DesignOrientation $area.DesignOrientation -Detail ("Edge {0} length {1} is below the short-edge threshold." -f ($i + 1), [Math]::Round($edgeLength, 4)) -MetricValue $edgeLength -Threshold $ShortEdgeThreshold -PointIndex ($i + 1) -EdgeIndex ($i + 1) -Points @($current, $next) -MarkerCoordinates @((New-MarkerCoordinate -X $current.X -Y $current.Y -Z $current.Z -Kind "Vertex"), (New-MarkerCoordinate -X $next.X -Y $next.Y -Z $next.Z -Kind "Vertex"), (New-MarkerCoordinate -X $midpoint.X -Y $midpoint.Y -Z $midpoint.Z -Kind "Midpoint")) -MatchesWarningTarget $matchesWarning)) | Out-Null
            }
        }

        for ($i = 0; $i -lt $pointCount; $i++) {
            $previous = $points[($i + $pointCount - 1) % $pointCount]
            $current = $points[$i]
            $next = $points[($i + 1) % $pointCount]
            $incomingLength = Get-PointDistance -PointA $previous -PointB $current
            $outgoingLength = Get-PointDistance -PointA $current -PointB $next
            $angle = Get-InteriorAngleDegrees -PointA $previous -PointB $current -PointC $next
            if ($null -eq $angle) {
                continue
            }

            if ($angle -le $SliverAngleDeg -and ($incomingLength -le ($ShortEdgeThreshold * 8.0) -or $outgoingLength -le ($ShortEdgeThreshold * 8.0))) {
                $severity = if ($angle -le ($SliverAngleDeg / 2.0)) { "ERROR" } else { "WARN" }
                $areaFindings.Add((New-AreaGeometryFinding -Severity $severity -Category "SliverCorner" -AreaName $area.Name -Label $area.Label -Story $area.Story -Property $area.Property -DesignOrientation $area.DesignOrientation -Detail ("Vertex {0} interior angle {1} deg is a likely sliver corner." -f ($i + 1), [Math]::Round($angle, 3)) -MetricValue $angle -Threshold $SliverAngleDeg -PointIndex ($i + 1) -EdgeIndex ($i + 1) -Points @($previous, $current, $next) -MarkerCoordinates @((New-MarkerCoordinate -X $current.X -Y $current.Y -Z $current.Z -Kind "Vertex")) -MatchesWarningTarget $matchesWarning)) | Out-Null
            }

            if ([Math]::Abs(180.0 - $angle) -le $CollinearAngleToleranceDeg -and ($incomingLength -le $RedundantEdgeThreshold -or $outgoingLength -le $RedundantEdgeThreshold)) {
                $areaFindings.Add((New-AreaGeometryFinding -Severity "WARN" -Category "RedundantVertex" -AreaName $area.Name -Label $area.Label -Story $area.Story -Property $area.Property -DesignOrientation $area.DesignOrientation -Detail ("Vertex {0} is nearly collinear ({1} deg) and likely redundant for meshing." -f ($i + 1), [Math]::Round($angle, 3)) -MetricValue $angle -Threshold (180.0 - $CollinearAngleToleranceDeg) -PointIndex ($i + 1) -EdgeIndex ($i + 1) -Points @($previous, $current, $next) -MarkerCoordinates @((New-MarkerCoordinate -X $current.X -Y $current.Y -Z $current.Z -Kind "Vertex")) -MatchesWarningTarget $matchesWarning)) | Out-Null
            }
        }

        if ($matchesWarning -and $areaFindings.Count -eq 0) {
            $centroidX = ($points | Measure-Object -Property X -Average).Average
            $centroidY = ($points | Measure-Object -Property Y -Average).Average
            $centroidZ = ($points | Measure-Object -Property Z -Average).Average
            $areaFindings.Add((New-AreaGeometryFinding -Severity "INFO" -Category "WarningAreaReview" -AreaName $area.Name -Label $area.Label -Story $area.Story -Property $area.Property -DesignOrientation $area.DesignOrientation -Detail "The area was named in an ETABS meshing warning but did not trip the geometric heuristics. Review the full shell mesh on this object." -MetricValue 0.0 -Threshold 0.0 -PointIndex 0 -EdgeIndex 0 -Points @($points) -MarkerCoordinates @((New-MarkerCoordinate -X $centroidX -Y $centroidY -Z $centroidZ -Kind "Centroid")) -MatchesWarningTarget $true)) | Out-Null
        }

        foreach ($finding in $areaFindings) {
            $findings.Add($finding) | Out-Null
        }

        if ($areaFindings.Count -gt 0) {
            $score = 0
            foreach ($finding in $areaFindings) {
                switch ($finding.Category) {
                    "NearDuplicateVertex" { $score += 5 }
                    "ShortEdge" { $score += 4 }
                    "SliverCorner" { $score += 3 }
                    "RedundantVertex" { $score += 2 }
                    "WarningAreaReview" { $score += 1 }
                }

                if ($finding.Severity -eq "ERROR") {
                    $score += 2
                }
            }

            $areaSummaries.Add([pscustomobject]@{
                AreaName = $area.Name
                Label = $area.Label
                Story = $area.Story
                Property = $area.Property
                DesignOrientation = $area.DesignOrientation
                FindingCount = $areaFindings.Count
                WarningMatch = $matchesWarning
                Score = $score
                Categories = (@($areaFindings | Select-Object -ExpandProperty Category -Unique) -join ", ")
            }) | Out-Null
        }
    }

    [pscustomobject]@{
        Areas = @($areas)
        Findings = @($findings | Sort-Object @{ Expression = { if ($_.MatchesWarningTarget) { 0 } else { 1 } } }, @{ Expression = { switch ($_.Severity) { "ERROR" { 0 } "WARN" { 1 } default { 2 } } } }, Story, Label, Category, EdgeIndex)
        AreaSummaries = @($areaSummaries | Sort-Object @{ Expression = { -$_.WarningMatch } }, @{ Expression = { -$_.Score } }, Story, Label)
    }
}

function Ensure-EtabsGroup {
    param(
        $SapModel,
        [string]$GroupName,
        [int]$Color
    )

    $SapModel.GroupDef.SetGroup_1(
        $GroupName,
        $Color,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true,
        $true) | Out-Null
}

function Remove-EtabsGeometryMarkers {
    param(
        $SapModel,
        [string]$GroupPrefix = "DBG_GEOM"
    )

    $groupNamesToDelete = @(
        "$GroupPrefix`_MARKERS",
        "$GroupPrefix`_AREAS",
        "$GroupPrefix`_WARNING_AREAS"
    )
    $arrowFrameGroup = "$GroupPrefix`_ARROW_FRAMES"
    $arrowPointGroup = "$GroupPrefix`_ARROW_POINTS"
    $arrowMaterial = "$GroupPrefix`_ARROW_MAT"
    $arrowSection = "$GroupPrefix`_ARROW_SEC"

    try {
        $SapModel.FrameObj.Delete($arrowFrameGroup, [ETABSv1.eItemType]::Group) | Out-Null
    }
    catch {
    }

    try {
        $SapModel.PointObj.DeleteSpecialPoint($arrowPointGroup, [ETABSv1.eItemType]::Group) | Out-Null
    }
    catch {
    }

    foreach ($groupName in $groupNamesToDelete) {
        try {
            $SapModel.PointObj.DeleteSpecialPoint($groupName, [ETABSv1.eItemType]::Group) | Out-Null
        }
        catch {
        }
    }

    foreach ($groupName in $groupNamesToDelete) {
        try {
            $SapModel.GroupDef.Delete($groupName) | Out-Null
        }
        catch {
        }
    }

    foreach ($groupName in @($arrowFrameGroup, $arrowPointGroup)) {
        try {
            $SapModel.GroupDef.Delete($groupName) | Out-Null
        }
        catch {
        }
    }

    try {
        $SapModel.PropFrame.Delete($arrowSection) | Out-Null
    }
    catch {
    }

    try {
        $SapModel.PropMaterial.Delete($arrowMaterial) | Out-Null
    }
    catch {
    }
}

function Add-EtabsGeometryMarkers {
    param(
        $SapModel,
        [pscustomobject[]]$Findings,
        [string]$GroupPrefix = "DBG_GEOM",
        [switch]$SelectInModel,
        [int]$MaxMarkers = 200
    )

    Remove-EtabsGeometryMarkers -SapModel $SapModel -GroupPrefix $GroupPrefix

    $markerGroup = "$GroupPrefix`_MARKERS"
    $areaGroup = "$GroupPrefix`_AREAS"
    $warningAreaGroup = "$GroupPrefix`_WARNING_AREAS"

    Ensure-EtabsGroup -SapModel $SapModel -GroupName $markerGroup -Color 255
    Ensure-EtabsGroup -SapModel $SapModel -GroupName $areaGroup -Color 65535
    Ensure-EtabsGroup -SapModel $SapModel -GroupName $warningAreaGroup -Color 16711680

    $SapModel.SelectObj.ClearSelection() | Out-Null

    $markerRecords = New-Object System.Collections.ArrayList
    $markerKeys = New-Object System.Collections.Generic.HashSet[string]
    $usedNames = New-Object System.Collections.Generic.HashSet[string]
    foreach ($finding in @($Findings)) {
        $SapModel.AreaObj.SetGroupAssign($finding.AreaName, $areaGroup, $true, [ETABSv1.eItemType]::Objects) | Out-Null
        if ($finding.MatchesWarningTarget) {
            $SapModel.AreaObj.SetGroupAssign($finding.AreaName, $warningAreaGroup, $true, [ETABSv1.eItemType]::Objects) | Out-Null
        }

        if ($SelectInModel) {
            $SapModel.AreaObj.SetSelected($finding.AreaName, $true, [ETABSv1.eItemType]::Objects) | Out-Null
        }

        foreach ($marker in @($finding.MarkerCoordinates)) {
            if ($markerRecords.Count -ge $MaxMarkers) {
                break
            }

            $key = "{0:F4}|{1:F4}|{2:F4}|{3}" -f $marker.X, $marker.Y, $marker.Z, $marker.Kind
            if (-not $markerKeys.Add($key)) {
                continue
            }

            $pointName = ""
            $userNameBase = "{0}_{1}_{2}" -f $GroupPrefix, $finding.Category.ToUpperInvariant(), ($marker.Kind.ToUpperInvariant())
            $suffix = $markerRecords.Count + 1
            $candidateName = "{0}_{1:D3}" -f $userNameBase, $suffix
            while (-not $usedNames.Add($candidateName)) {
                $suffix++
                $candidateName = "{0}_{1:D3}" -f $userNameBase, $suffix
            }

            $ret = $SapModel.PointObj.AddCartesian($marker.X, $marker.Y, $marker.Z, [ref]$pointName, $candidateName, "Global", $true, 0)
            if ($ret -ne 0) {
                continue
            }

            $SapModel.PointObj.SetSpecialPoint($pointName, $true, [ETABSv1.eItemType]::Objects) | Out-Null
            $SapModel.PointObj.SetGroupAssign($pointName, $markerGroup, $true, [ETABSv1.eItemType]::Objects) | Out-Null
            if ($SelectInModel) {
                $SapModel.PointObj.SetSelected($pointName, $true, [ETABSv1.eItemType]::Objects) | Out-Null
            }

            $markerRecords.Add([pscustomobject]@{
                PointName = $pointName
                UserName = $candidateName
                X = [Math]::Round($marker.X, 4)
                Y = [Math]::Round($marker.Y, 4)
                Z = [Math]::Round($marker.Z, 4)
                Kind = $marker.Kind
            }) | Out-Null
        }
    }

    [pscustomobject]@{
        GroupPrefix = $GroupPrefix
        MarkerGroup = $markerGroup
        AreaGroup = $areaGroup
        WarningAreaGroup = $warningAreaGroup
        MarkerCount = $markerRecords.Count
        Markers = @($markerRecords)
    }
}

function Get-DebugArrowSettings {
    param(
        [int]$PresentUnitsEnum
    )

    switch ($PresentUnitsEnum) {
        1 { return [pscustomobject]@{ ShaftLength = 120.0; HeadLength = 36.0; HeadWidth = 48.0; TipGap = 12.0; SectionDepth = 6.0; SectionWidth = 6.0 } }
        2 { return [pscustomobject]@{ ShaftLength = 10.0; HeadLength = 3.0; HeadWidth = 4.0; TipGap = 1.0; SectionDepth = 0.5; SectionWidth = 0.5 } }
        3 { return [pscustomobject]@{ ShaftLength = 120.0; HeadLength = 36.0; HeadWidth = 48.0; TipGap = 12.0; SectionDepth = 6.0; SectionWidth = 6.0 } }
        4 { return [pscustomobject]@{ ShaftLength = 10.0; HeadLength = 3.0; HeadWidth = 4.0; TipGap = 1.0; SectionDepth = 0.5; SectionWidth = 0.5 } }
        5 { return [pscustomobject]@{ ShaftLength = 3000.0; HeadLength = 900.0; HeadWidth = 1200.0; TipGap = 300.0; SectionDepth = 150.0; SectionWidth = 150.0 } }
        6 { return [pscustomobject]@{ ShaftLength = 300.0; HeadLength = 90.0; HeadWidth = 120.0; TipGap = 30.0; SectionDepth = 15.0; SectionWidth = 15.0 } }
        7 { return [pscustomobject]@{ ShaftLength = 3.0; HeadLength = 0.9; HeadWidth = 1.2; TipGap = 0.3; SectionDepth = 0.15; SectionWidth = 0.15 } }
        default { return [pscustomobject]@{ ShaftLength = 120.0; HeadLength = 36.0; HeadWidth = 48.0; TipGap = 12.0; SectionDepth = 6.0; SectionWidth = 6.0 } }
    }
}

function Ensure-DebugArrowResources {
    param(
        $SapModel,
        [string]$GroupPrefix,
        [int]$PresentUnitsEnum
    )

    $settings = Get-DebugArrowSettings -PresentUnitsEnum $PresentUnitsEnum
    $materialName = "$GroupPrefix`_ARROW_MAT"
    $sectionName = "$GroupPrefix`_ARROW_SEC"

    $SapModel.PropMaterial.SetMaterial($materialName, [ETABSv1.eMatType]::Steel, 0, "", "") | Out-Null
    $SapModel.PropMaterial.SetMPIsotropic($materialName, 29000000.0, 0.3, 0.0000065, 0.0) | Out-Null
    $SapModel.PropMaterial.SetWeightAndMass($materialName, 1, 0.0, 0.0) | Out-Null
    $SapModel.PropFrame.SetRectangle($sectionName, $materialName, $settings.SectionDepth, $settings.SectionWidth, 0, "", "") | Out-Null

    [double[]]$modifiers = @(1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0)
    $SapModel.PropFrame.SetModifiers($sectionName, [ref]$modifiers) | Out-Null

    [pscustomobject]@{
        MaterialName = $materialName
        SectionName = $sectionName
        Settings = $settings
    }
}

function Get-DebugArrowVector2d {
    param(
        [pscustomobject]$Finding,
        [pscustomobject]$Marker
    )

    $centroidX = 0.0
    $centroidY = 0.0
    $hasCentroid = $false

    if (@($Finding.Points).Count -gt 0) {
        $centroidX = ($Finding.Points | Measure-Object -Property X -Average).Average
        $centroidY = ($Finding.Points | Measure-Object -Property Y -Average).Average
        $hasCentroid = $true
    }

    $dx = 0.0
    $dy = 0.0
    if ($hasCentroid) {
        $dx = $Marker.X - $centroidX
        $dy = $Marker.Y - $centroidY
    }

    $len = [Math]::Sqrt(($dx * $dx) + ($dy * $dy))
    if ($len -le 0.0001) {
        $dx = 1.0
        $dy = 1.0
        $len = [Math]::Sqrt(2.0)
    }

    [pscustomobject]@{
        X = $dx / $len
        Y = $dy / $len
    }
}

function Add-DebugArrowPoint {
    param(
        $SapModel,
        [double]$X,
        [double]$Y,
        [double]$Z,
        [string]$UserName,
        [string]$PointGroup
    )

    $pointName = ""
    $ret = $SapModel.PointObj.AddCartesian($X, $Y, $Z, [ref]$pointName, $UserName, "Global", $true, 0)
    if ($ret -ne 0) {
        throw "PointObj.AddCartesian returned $ret for '$UserName'."
    }

    $SapModel.PointObj.SetSpecialPoint($pointName, $true, [ETABSv1.eItemType]::Objects) | Out-Null
    $SapModel.PointObj.SetGroupAssign($pointName, $PointGroup, $true, [ETABSv1.eItemType]::Objects) | Out-Null
    [bool[]]$restraint = @($true, $true, $true, $true, $true, $true)
    $SapModel.PointObj.SetRestraint($pointName, [ref]$restraint, [ETABSv1.eItemType]::Objects) | Out-Null

    return $pointName
}

function Add-EtabsGeometryArrowMarkers {
    param(
        $SapModel,
        [pscustomobject[]]$Findings,
        [string]$GroupPrefix = "DBG_GEOM",
        [int]$PresentUnitsEnum
    )

    $arrowFrameGroup = "$GroupPrefix`_ARROW_FRAMES"
    $arrowPointGroup = "$GroupPrefix`_ARROW_POINTS"
    Ensure-EtabsGroup -SapModel $SapModel -GroupName $arrowFrameGroup -Color 16753920
    Ensure-EtabsGroup -SapModel $SapModel -GroupName $arrowPointGroup -Color 16753920

    $resource = Ensure-DebugArrowResources -SapModel $SapModel -GroupPrefix $GroupPrefix -PresentUnitsEnum $PresentUnitsEnum
    $settings = $resource.Settings

    $createdArrows = New-Object System.Collections.ArrayList
    $markerKeys = New-Object System.Collections.Generic.HashSet[string]
    $sequence = 1

    foreach ($finding in @($Findings)) {
        foreach ($marker in @($finding.MarkerCoordinates)) {
            $key = "{0:F4}|{1:F4}|{2:F4}|{3}" -f $marker.X, $marker.Y, $marker.Z, $marker.Kind
            if (-not $markerKeys.Add($key)) {
                continue
            }

            $direction = Get-DebugArrowVector2d -Finding $finding -Marker $marker
            $perpX = -$direction.Y
            $perpY = $direction.X

            $tipX = $marker.X + ($direction.X * $settings.TipGap)
            $tipY = $marker.Y + ($direction.Y * $settings.TipGap)
            $tailX = $marker.X + ($direction.X * ($settings.TipGap + $settings.ShaftLength))
            $tailY = $marker.Y + ($direction.Y * ($settings.TipGap + $settings.ShaftLength))
            $headBaseX = $marker.X + ($direction.X * ($settings.TipGap + $settings.HeadLength))
            $headBaseY = $marker.Y + ($direction.Y * ($settings.TipGap + $settings.HeadLength))
            $leftX = $headBaseX + ($perpX * ($settings.HeadWidth / 2.0))
            $leftY = $headBaseY + ($perpY * ($settings.HeadWidth / 2.0))
            $rightX = $headBaseX - ($perpX * ($settings.HeadWidth / 2.0))
            $rightY = $headBaseY - ($perpY * ($settings.HeadWidth / 2.0))
            $z = $marker.Z

            $tailPoint = Add-DebugArrowPoint -SapModel $SapModel -X $tailX -Y $tailY -Z $z -UserName ("{0}_AT_{1:D3}" -f $GroupPrefix, $sequence) -PointGroup $arrowPointGroup
            $tipPoint = Add-DebugArrowPoint -SapModel $SapModel -X $tipX -Y $tipY -Z $z -UserName ("{0}_AP_{1:D3}" -f $GroupPrefix, $sequence) -PointGroup $arrowPointGroup
            $leftPoint = Add-DebugArrowPoint -SapModel $SapModel -X $leftX -Y $leftY -Z $z -UserName ("{0}_AL_{1:D3}" -f $GroupPrefix, $sequence) -PointGroup $arrowPointGroup
            $rightPoint = Add-DebugArrowPoint -SapModel $SapModel -X $rightX -Y $rightY -Z $z -UserName ("{0}_AR_{1:D3}" -f $GroupPrefix, $sequence) -PointGroup $arrowPointGroup

            $shaftName = ""
            $leftName = ""
            $rightName = ""
            $SapModel.FrameObj.AddByPoint($tailPoint, $tipPoint, [ref]$shaftName, $resource.SectionName, ("{0}_SHAFT_{1:D3}" -f $GroupPrefix, $sequence)) | Out-Null
            $SapModel.FrameObj.AddByPoint($leftPoint, $tipPoint, [ref]$leftName, $resource.SectionName, ("{0}_HEADL_{1:D3}" -f $GroupPrefix, $sequence)) | Out-Null
            $SapModel.FrameObj.AddByPoint($rightPoint, $tipPoint, [ref]$rightName, $resource.SectionName, ("{0}_HEADR_{1:D3}" -f $GroupPrefix, $sequence)) | Out-Null

            foreach ($frameName in @($shaftName, $leftName, $rightName)) {
                if (-not [string]::IsNullOrWhiteSpace($frameName)) {
                    $SapModel.FrameObj.SetGroupAssign($frameName, $arrowFrameGroup, $true, [ETABSv1.eItemType]::Objects) | Out-Null
                }
            }

            $createdArrows.Add([pscustomobject]@{
                Sequence = $sequence
                AreaName = $finding.AreaName
                Label = $finding.Label
                Story = $finding.Story
                Kind = $marker.Kind
                Tip = [pscustomobject]@{ X = [Math]::Round($tipX, 4); Y = [Math]::Round($tipY, 4); Z = [Math]::Round($z, 4) }
                Tail = [pscustomobject]@{ X = [Math]::Round($tailX, 4); Y = [Math]::Round($tailY, 4); Z = [Math]::Round($z, 4) }
                Left = [pscustomobject]@{ X = [Math]::Round($leftX, 4); Y = [Math]::Round($leftY, 4); Z = [Math]::Round($z, 4) }
                Right = [pscustomobject]@{ X = [Math]::Round($rightX, 4); Y = [Math]::Round($rightY, 4); Z = [Math]::Round($z, 4) }
                Frames = @($shaftName, $leftName, $rightName)
            }) | Out-Null

            $sequence++
        }
    }

    [pscustomobject]@{
        ArrowFrameGroup = $arrowFrameGroup
        ArrowPointGroup = $arrowPointGroup
        MaterialName = $resource.MaterialName
        SectionName = $resource.SectionName
        ArrowCount = @($createdArrows).Count
        Arrows = @($createdArrows)
    }
}
