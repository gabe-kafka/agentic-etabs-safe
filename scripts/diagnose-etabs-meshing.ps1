param(
    [int]$EtabsPid,
    [string]$WarningText,
    [string]$WarningTextPath,
    [double]$ShortEdgeThreshold = -1,
    [double]$DuplicateVertexThreshold = -1,
    [double]$SliverAngleDeg = -1,
    [double]$CollinearAngleToleranceDeg = -1,
    [double]$RedundantEdgeThreshold = -1,
    [switch]$MarkInModel,
    [switch]$ArrowMarkers,
    [switch]$SelectInModel,
    [switch]$OnlyWarningTargets,
    [switch]$UnlockIfLocked,
    [string]$GroupPrefix = "DBG_GEOM",
    [int]$MaxMarkers = 200,
    [string]$CsvOut,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path -Path $PSScriptRoot -ChildPath "_etabs-geometry-debug.ps1")

$connection = Connect-EtabsSession -EtabsPid $EtabsPid
$sapModel = $connection.SapModel
$presentUnits = [int]($sapModel.GetPresentUnits())
$defaults = Get-DefaultMeshingThresholds -PresentUnitsEnum $presentUnits

if ($ShortEdgeThreshold -lt 0) { $ShortEdgeThreshold = $defaults.ShortEdge }
if ($DuplicateVertexThreshold -lt 0) { $DuplicateVertexThreshold = $defaults.DuplicateVertex }
if ($SliverAngleDeg -lt 0) { $SliverAngleDeg = $defaults.SliverAngleDeg }
if ($CollinearAngleToleranceDeg -lt 0) { $CollinearAngleToleranceDeg = $defaults.CollinearAngleToleranceDeg }
if ($RedundantEdgeThreshold -lt 0) { $RedundantEdgeThreshold = $defaults.RedundantEdge }

$warningTargets = Parse-EtabsMeshingWarningTargets -WarningText $WarningText -WarningTextPath $WarningTextPath
$diagnostic = Get-EtabsMeshingFindings `
    -SapModel $sapModel `
    -ShortEdgeThreshold $ShortEdgeThreshold `
    -DuplicateVertexThreshold $DuplicateVertexThreshold `
    -SliverAngleDeg $SliverAngleDeg `
    -CollinearAngleToleranceDeg $CollinearAngleToleranceDeg `
    -RedundantEdgeThreshold $RedundantEdgeThreshold `
    -WarningTargets $warningTargets

$findingsToMark = @($diagnostic.Findings)
if ($OnlyWarningTargets -and @($warningTargets).Count -gt 0) {
    $findingsToMark = @($diagnostic.Findings | Where-Object { $_.MatchesWarningTarget })
}
elseif (@($warningTargets).Count -gt 0) {
    $warningFindings = @($diagnostic.Findings | Where-Object { $_.MatchesWarningTarget })
    if ($warningFindings.Count -gt 0) {
        $findingsToMark = $warningFindings
    }
}

$markerResult = $null
$arrowResult = $null
$modelWasLocked = $sapModel.GetModelIsLocked()
$unlockedForMarkers = $false
if ($MarkInModel -or $ArrowMarkers) {
    if ($modelWasLocked) {
        if (-not $UnlockIfLocked) {
            throw "Marker or arrow overlay mode requires an unlocked ETABS model. Re-run with -UnlockIfLocked if you want the script to unlock the model and place temporary debug overlays."
        }

        $unlockRet = $sapModel.SetModelIsLocked($false)
        if ($unlockRet -ne 0) {
            throw "SetModelIsLocked(false) returned $unlockRet."
        }

        $unlockedForMarkers = $true
    }

    if ($MarkInModel) {
        $markerResult = Add-EtabsGeometryMarkers `
            -SapModel $sapModel `
            -Findings $findingsToMark `
            -GroupPrefix $GroupPrefix `
            -SelectInModel:$SelectInModel `
            -MaxMarkers $MaxMarkers
    }

    if ($ArrowMarkers) {
        $arrowResult = Add-EtabsGeometryArrowMarkers `
            -SapModel $sapModel `
            -Findings $findingsToMark `
            -GroupPrefix $GroupPrefix `
            -PresentUnitsEnum $presentUnits
    }
}

$result = [pscustomobject]@{
    ProcessId = $connection.Process.Id
    ProcessVersion = $connection.Process.MainModule.FileVersionInfo.FileVersion
    ApiDllPath = $connection.ApiDllPath
    ModelPath = $connection.ModelPath
    PresentUnitsEnum = $presentUnits
    ModelWasLocked = $modelWasLocked
    UnlockedForMarkers = $unlockedForMarkers
    Thresholds = [pscustomobject]@{
        ShortEdge = $ShortEdgeThreshold
        DuplicateVertex = $DuplicateVertexThreshold
        SliverAngleDeg = $SliverAngleDeg
        CollinearAngleToleranceDeg = $CollinearAngleToleranceDeg
        RedundantEdge = $RedundantEdgeThreshold
    }
    WarningTargets = @($warningTargets)
    AreaCount = @($diagnostic.Areas).Count
    FindingCount = @($diagnostic.Findings).Count
    MarkedFindingCount = @($findingsToMark).Count
    Findings = @($diagnostic.Findings)
    AreaSummaries = @($diagnostic.AreaSummaries)
    MarkerResult = $markerResult
    ArrowResult = $arrowResult
}

if (-not [string]::IsNullOrWhiteSpace($CsvOut)) {
    $flatRows = foreach ($finding in @($diagnostic.Findings)) {
        $pointText = @($finding.Points | ForEach-Object {
            "{0}({1},{2},{3})" -f $_.Point, $_.X, $_.Y, $_.Z
        }) -join "; "

        [pscustomobject]@{
            Severity = $finding.Severity
            Category = $finding.Category
            Story = $finding.Story
            Label = $finding.Label
            AreaName = $finding.AreaName
            Property = $finding.Property
            DesignOrientation = $finding.DesignOrientation
            MetricValue = $finding.MetricValue
            Threshold = $finding.Threshold
            EdgeIndex = $finding.EdgeIndex
            PointIndex = $finding.PointIndex
            MatchesWarningTarget = $finding.MatchesWarningTarget
            Detail = $finding.Detail
            Points = $pointText
        }
    }

    $flatRows | Export-Csv -LiteralPath $CsvOut -NoTypeInformation
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
}
else {
    $result
}
