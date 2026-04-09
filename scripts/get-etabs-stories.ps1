param(
    [int]$EtabsPid,
    [ValidateSet("Model", "Feet")]
    [string]$LengthUnit = "Model",
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path -Path $PSScriptRoot -ChildPath "_etabs-geometry-debug.ps1")

function Get-ObjectNames {
    param(
        $ObjectApi,
        [string]$Label
    )

    $numberNames = 0
    [string[]]$names = @()
    $ret = $ObjectApi.GetNameList([ref]$numberNames, [ref]$names)
    if ($ret -ne 0) {
        throw "GetNameList failed for $Label with return code $ret."
    }

    return @($names)
}

function Get-ItemCount {
    param(
        $Items
    )

    return (@($Items) | Measure-Object).Count
}

function Get-ModelLengthUnitName {
    param(
        [int]$PresentUnitsEnum
    )

    switch ($PresentUnitsEnum) {
        1 { return "in" }
        2 { return "ft" }
        3 { return "in" }
        4 { return "ft" }
        5 { return "mm" }
        6 { return "cm" }
        7 { return "m" }
        8 { return "mm" }
        9 { return "cm" }
        10 { return "m" }
        11 { return "m" }
        12 { return "cm" }
        13 { return "mm" }
        14 { return "m" }
        15 { return "cm" }
        16 { return "mm" }
        default { throw "Unsupported ETABS present units enum '$PresentUnitsEnum'." }
    }
}

function Get-FeetPerModelUnit {
    param(
        [int]$PresentUnitsEnum
    )

    switch (Get-ModelLengthUnitName -PresentUnitsEnum $PresentUnitsEnum) {
        "in" { return (1.0 / 12.0) }
        "ft" { return 1.0 }
        "mm" { return 0.00328083989501312 }
        "cm" { return 0.0328083989501312 }
        "m" { return 3.28083989501312 }
        default { throw "Unable to resolve feet conversion for present units enum '$PresentUnitsEnum'." }
    }
}

function Convert-LengthFromModel {
    param(
        [double]$Value,
        [int]$PresentUnitsEnum,
        [string]$TargetLengthUnit
    )

    switch ($TargetLengthUnit) {
        "Feet" { return ($Value * (Get-FeetPerModelUnit -PresentUnitsEnum $PresentUnitsEnum)) }
        default { return $Value }
    }
}

$connection = Connect-EtabsSession -EtabsPid $EtabsPid
$sapModel = $connection.SapModel
$presentUnits = [int]($sapModel.GetPresentUnits())
$modelLengthUnit = Get-ModelLengthUnitName -PresentUnitsEnum $presentUnits

$baseElevation = 0.0
$storyCount = 0
[string[]]$storyNames = @()
[double[]]$storyElevations = @()
[double[]]$storyHeights = @()
[bool[]]$isMasterStory = @()
[string[]]$similarToStory = @()
[bool[]]$spliceAbove = @()
[double[]]$spliceHeight = @()
[int[]]$storyColors = @()

$ret = $sapModel.Story.GetStories_2(
    [ref]$baseElevation,
    [ref]$storyCount,
    [ref]$storyNames,
    [ref]$storyElevations,
    [ref]$storyHeights,
    [ref]$isMasterStory,
    [ref]$similarToStory,
    [ref]$spliceAbove,
    [ref]$spliceHeight,
    [ref]$storyColors)

if ($ret -ne 0) {
    throw "Story.GetStories_2 returned $ret."
}

$stories = for ($i = 0; $i -lt $storyCount; $i++) {
    [pscustomobject]@{
        Index = $i
        Name = $storyNames[$i]
        Elevation = Convert-LengthFromModel -Value $storyElevations[$i] -PresentUnitsEnum $presentUnits -TargetLengthUnit $LengthUnit
        Height = Convert-LengthFromModel -Value $storyHeights[$i] -PresentUnitsEnum $presentUnits -TargetLengthUnit $LengthUnit
        IsMasterStory = $isMasterStory[$i]
        SimilarToStory = $similarToStory[$i]
        SpliceAbove = $spliceAbove[$i]
        SpliceHeight = Convert-LengthFromModel -Value $spliceHeight[$i] -PresentUnitsEnum $presentUnits -TargetLengthUnit $LengthUnit
        Color = $storyColors[$i]
    }
}

$result = [pscustomobject]@{
    ProcessId = $connection.Process.Id
    ProcessVersion = $connection.Process.MainModule.FileVersionInfo.FileVersion
    ApiDllPath = $connection.ApiDllPath
    ModelPath = $connection.ModelPath
    PresentUnitsEnum = $presentUnits
    ModelLengthUnit = $modelLengthUnit
    LengthUnit = if ($LengthUnit -eq "Feet") { "ft" } else { $modelLengthUnit }
    ModelLocked = $sapModel.GetModelIsLocked()
    BaseElevation = Convert-LengthFromModel -Value $baseElevation -PresentUnitsEnum $presentUnits -TargetLengthUnit $LengthUnit
    StoryCount = $storyCount
    Stories = @($stories)
    PhysicalCounts = [pscustomobject]@{
        Frames = Get-ItemCount (Get-ObjectNames -ObjectApi $sapModel.FrameObj -Label "frame objects")
        Areas = Get-ItemCount (Get-ObjectNames -ObjectApi $sapModel.AreaObj -Label "area objects")
        Links = Get-ItemCount (Get-ObjectNames -ObjectApi $sapModel.LinkObj -Label "link objects")
        Tendons = Get-ItemCount (Get-ObjectNames -ObjectApi $sapModel.TendonObj -Label "tendon objects")
        Points = Get-ItemCount (Get-ObjectNames -ObjectApi $sapModel.PointObj -Label "point objects")
    }
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
}
else {
    $result
}
