param(
    [int]$EtabsPid,
    [string]$StoryJsonPath,
    [double]$BaseElevation = [double]::NaN,
    [ValidateSet("Model", "Feet")]
    [string]$LengthUnit = "Model",
    [switch]$UnlockIfLocked,
    [switch]$SkipBackup,
    [switch]$Save,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path -Path $PSScriptRoot -ChildPath "_etabs-geometry-debug.ps1")

function Assert-Success {
    param(
        [int]$ReturnCode,
        [string]$Action
    )

    if ($ReturnCode -ne 0) {
        throw "$Action returned $ReturnCode."
    }
}

function Get-ItemCount {
    param(
        $Items
    )

    return (@($Items) | Measure-Object).Count
}

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

function Convert-LengthToModel {
    param(
        [double]$Value,
        [int]$PresentUnitsEnum,
        [string]$SourceLengthUnit
    )

    switch ($SourceLengthUnit) {
        "Feet" { return ($Value / (Get-FeetPerModelUnit -PresentUnitsEnum $PresentUnitsEnum)) }
        default { return $Value }
    }
}

function Get-PhysicalCounts {
    param(
        $SapModel
    )

    return [pscustomobject]@{
        Frames = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.FrameObj -Label "frame objects")
        Areas = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.AreaObj -Label "area objects")
        Links = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.LinkObj -Label "link objects")
        Tendons = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.TendonObj -Label "tendon objects")
        Points = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.PointObj -Label "point objects")
    }
}

function Get-CurrentStoryData {
    param(
        $SapModel,
        [int]$PresentUnitsEnum,
        [string]$LengthUnit
    )

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

    $ret = $SapModel.Story.GetStories_2(
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
            Elevation = Convert-LengthFromModel -Value $storyElevations[$i] -PresentUnitsEnum $PresentUnitsEnum -TargetLengthUnit $LengthUnit
            Height = Convert-LengthFromModel -Value $storyHeights[$i] -PresentUnitsEnum $PresentUnitsEnum -TargetLengthUnit $LengthUnit
            IsMasterStory = $isMasterStory[$i]
            SimilarToStory = $similarToStory[$i]
            SpliceAbove = $spliceAbove[$i]
            SpliceHeight = Convert-LengthFromModel -Value $spliceHeight[$i] -PresentUnitsEnum $PresentUnitsEnum -TargetLengthUnit $LengthUnit
            Color = $storyColors[$i]
        }
    }

    return [pscustomobject]@{
        BaseElevation = Convert-LengthFromModel -Value $baseElevation -PresentUnitsEnum $PresentUnitsEnum -TargetLengthUnit $LengthUnit
        StoryCount = $storyCount
        Stories = @($stories)
    }
}

function Get-BackupPath {
    param(
        [string]$ModelPath
    )

    $directory = Split-Path -Parent $ModelPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ModelPath)
    $extension = [System.IO.Path]::GetExtension($ModelPath)
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path -Path $directory -ChildPath ("{0}.pre-set-stories.{1}{2}" -f $baseName, $timestamp, $extension)
}

function ConvertTo-StoryInputRows {
    param(
        [object]$Payload
    )

    $payloadProperties = @($Payload.PSObject.Properties.Name)
    if ($payloadProperties -contains "Stories") {
        return @($Payload.Stories)
    }

    return @($Payload)
}

function Get-PayloadBaseElevation {
    param(
        [object]$Payload
    )

    $payloadProperties = @($Payload.PSObject.Properties.Name)
    if ($payloadProperties -contains "BaseElevation" -and $null -ne $Payload.BaseElevation) {
        return [double]$Payload.BaseElevation
    }

    return $null
}

function Get-DefaultStoryColor {
    param(
        [int]$Index
    )

    $palette = @(255, 16776960, 65280, 16711680, 8421504, 65535, 16711935)
    return $palette[$Index % $palette.Count]
}

function New-StorySet {
    param(
        [object[]]$InputRows,
        [double]$ResolvedBaseElevation,
        [object[]]$ExistingStories
    )

    if ((Get-ItemCount $InputRows) -eq 0) {
        throw "No story rows were provided."
    }

    $existingStoryIndex = @{}
    foreach ($existingStory in @($ExistingStories)) {
        $existingStoryIndex[$existingStory.Name.ToUpperInvariant()] = $existingStory
    }

    $normalizedRows = @()
    $seenNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $originalIndex = 0

    foreach ($row in @($InputRows)) {
        $rowProperties = @($row.PSObject.Properties.Name)
        if (-not ($rowProperties -contains "Name")) {
            throw "Each story row must contain a Name property."
        }
        if (-not ($rowProperties -contains "Elevation")) {
            throw "Each story row must contain an Elevation property."
        }

        $name = [string]$row.Name
        $name = $name.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw "Story names cannot be blank."
        }

        if (-not $seenNames.Add($name)) {
            throw "Duplicate story name '$name' was provided."
        }

        $elevation = [double]$row.Elevation
        $normalizedRows += [pscustomobject]@{
            OriginalIndex = $originalIndex
            Name = $name
            Elevation = $elevation
        }
        $originalIndex++
    }

    $sortedRows = @($normalizedRows | Sort-Object Elevation, OriginalIndex)

    for ($i = 1; $i -lt $sortedRows.Count; $i++) {
        if ($sortedRows[$i].Elevation -le $sortedRows[$i - 1].Elevation) {
            throw "Story elevations must be strictly increasing after sorting. '$($sortedRows[$i - 1].Name)' and '$($sortedRows[$i].Name)' are invalid."
        }
    }

    $firstHeight = $sortedRows[0].Elevation - $ResolvedBaseElevation
    if ($firstHeight -le 0.0) {
        throw "The first story elevation ($($sortedRows[0].Elevation)) must be above the base elevation ($ResolvedBaseElevation)."
    }

    $storyNames = New-Object 'string[]' $sortedRows.Count
    $storyHeights = New-Object 'double[]' $sortedRows.Count
    $isMasterStory = New-Object 'bool[]' $sortedRows.Count
    $similarToStory = New-Object 'string[]' $sortedRows.Count
    $spliceAbove = New-Object 'bool[]' $sortedRows.Count
    $spliceHeight = New-Object 'double[]' $sortedRows.Count
    $storyColors = New-Object 'int[]' $sortedRows.Count
    $previewStories = @()

    for ($i = 0; $i -lt $sortedRows.Count; $i++) {
        $row = $sortedRows[$i]
        $height = if ($i -eq 0) {
            $row.Elevation - $ResolvedBaseElevation
        }
        else {
            $row.Elevation - $sortedRows[$i - 1].Elevation
        }

        if ($height -le 0.0) {
            throw "Computed story height for '$($row.Name)' is $height. Heights must be positive."
        }

        $existing = $existingStoryIndex[$row.Name.ToUpperInvariant()]
        $storyNames[$i] = $row.Name
        $storyHeights[$i] = $height
        $isMasterStory[$i] = if ($null -ne $existing) { [bool]$existing.IsMasterStory } else { $true }
        $similarToStory[$i] = if ($null -ne $existing -and -not $isMasterStory[$i]) { [string]$existing.SimilarToStory } else { "" }
        $spliceAbove[$i] = if ($null -ne $existing) { [bool]$existing.SpliceAbove } else { $false }
        $spliceHeight[$i] = if ($null -ne $existing) { [double]$existing.SpliceHeight } else { 0.0 }
        $storyColors[$i] = if ($null -ne $existing) { [int]$existing.Color } else { Get-DefaultStoryColor -Index $i }

        $previewStories += [pscustomobject]@{
            Index = $i
            Name = $row.Name
            Elevation = $row.Elevation
            Height = $height
            IsMasterStory = $isMasterStory[$i]
            SimilarToStory = $similarToStory[$i]
            SpliceAbove = $spliceAbove[$i]
            SpliceHeight = $spliceHeight[$i]
            Color = $storyColors[$i]
        }
    }

    return [pscustomobject]@{
        StoryNames = $storyNames
        StoryHeights = $storyHeights
        IsMasterStory = $isMasterStory
        SimilarToStory = $similarToStory
        SpliceAbove = $spliceAbove
        SpliceHeight = $spliceHeight
        StoryColors = $storyColors
        Stories = @($previewStories)
        OrderChanged = (@($normalizedRows | Sort-Object OriginalIndex | ForEach-Object { $_.Name }) -join "|") -ne (@($sortedRows | ForEach-Object { $_.Name }) -join "|")
    }
}

if ([string]::IsNullOrWhiteSpace($StoryJsonPath)) {
    throw "StoryJsonPath is required."
}

if (-not (Test-Path -LiteralPath $StoryJsonPath -PathType Leaf)) {
    throw "Story JSON file '$StoryJsonPath' was not found."
}

$payload = Get-Content -LiteralPath $StoryJsonPath -Raw | ConvertFrom-Json
$inputRows = ConvertTo-StoryInputRows -Payload $payload

$connection = Connect-EtabsSession -EtabsPid $EtabsPid
$sapModel = $connection.SapModel
$presentUnits = [int]($sapModel.GetPresentUnits())
$modelLengthUnit = Get-ModelLengthUnitName -PresentUnitsEnum $presentUnits
$modelWasLocked = $sapModel.GetModelIsLocked()
$unlockedForEdit = $false

if ($modelWasLocked) {
    if (-not $UnlockIfLocked) {
        throw "Setting ETABS stories requires an unlocked model. Re-run with -UnlockIfLocked if you want the script to unlock the model first."
    }

    Assert-Success ($sapModel.SetModelIsLocked($false)) "SetModelIsLocked(false)"
    $unlockedForEdit = $true
}

$prePhysicalCounts = Get-PhysicalCounts -SapModel $sapModel
if ($prePhysicalCounts.Frames -gt 0 -or $prePhysicalCounts.Areas -gt 0 -or $prePhysicalCounts.Links -gt 0 -or $prePhysicalCounts.Tendons -gt 0 -or $prePhysicalCounts.Points -gt 0) {
    throw "Story.SetStories_2 can only be used when no structural objects exist. Current counts: Frames=$($prePhysicalCounts.Frames), Areas=$($prePhysicalCounts.Areas), Links=$($prePhysicalCounts.Links), Tendons=$($prePhysicalCounts.Tendons), Points=$($prePhysicalCounts.Points)."
}

$preStories = Get-CurrentStoryData -SapModel $sapModel -PresentUnitsEnum $presentUnits -LengthUnit $LengthUnit
$payloadBaseElevation = Get-PayloadBaseElevation -Payload $payload
$resolvedBaseElevation = if (-not [double]::IsNaN($BaseElevation)) {
    $BaseElevation
}
elseif ($null -ne $payloadBaseElevation) {
    [double]$payloadBaseElevation
}
else {
    [double]$preStories.BaseElevation
}

$storySet = New-StorySet -InputRows $inputRows -ResolvedBaseElevation $resolvedBaseElevation -ExistingStories $preStories.Stories
$resolvedBaseElevationModel = Convert-LengthToModel -Value $resolvedBaseElevation -PresentUnitsEnum $presentUnits -SourceLengthUnit $LengthUnit
[double[]]$storyHeightsModel = @($storySet.StoryHeights | ForEach-Object {
    Convert-LengthToModel -Value ([double]$_) -PresentUnitsEnum $presentUnits -SourceLengthUnit $LengthUnit
})
[double[]]$spliceHeightModel = @($storySet.SpliceHeight | ForEach-Object {
    Convert-LengthToModel -Value ([double]$_) -PresentUnitsEnum $presentUnits -SourceLengthUnit $LengthUnit
})

$backupPath = $null
$backupCreated = $false
$backupError = $null
if (-not $SkipBackup -and -not [string]::IsNullOrWhiteSpace($connection.ModelPath) -and (Test-Path -LiteralPath $connection.ModelPath)) {
    if ($Save) {
        Assert-Success ($sapModel.File.Save($connection.ModelPath)) "Save model before backup"
    }

    try {
        $backupPath = Get-BackupPath -ModelPath $connection.ModelPath
        Copy-Item -LiteralPath $connection.ModelPath -Destination $backupPath -Force
        $backupCreated = $true
    }
    catch {
        $backupError = $_.Exception.Message
    }
}

$setRet = $sapModel.Story.SetStories_2(
    $resolvedBaseElevationModel,
    $storySet.StoryNames.Length,
    [ref]$storySet.StoryNames,
    [ref]$storyHeightsModel,
    [ref]$storySet.IsMasterStory,
    [ref]$storySet.SimilarToStory,
    [ref]$storySet.SpliceAbove,
    [ref]$spliceHeightModel,
    [ref]$storySet.StoryColors)

Assert-Success $setRet "Story.SetStories_2"

$saveReturnCode = $null
if ($Save -and -not [string]::IsNullOrWhiteSpace($connection.ModelPath)) {
    $saveReturnCode = $sapModel.File.Save($connection.ModelPath)
    Assert-Success $saveReturnCode "Save model after setting stories"
}

$postStories = Get-CurrentStoryData -SapModel $sapModel -PresentUnitsEnum $presentUnits -LengthUnit $LengthUnit
$postPhysicalCounts = Get-PhysicalCounts -SapModel $sapModel

$result = [pscustomobject]@{
    ProcessId = $connection.Process.Id
    ProcessVersion = $connection.Process.MainModule.FileVersionInfo.FileVersion
    ApiDllPath = $connection.ApiDllPath
    ModelPath = $connection.ModelPath
    PresentUnitsEnum = $presentUnits
    ModelLengthUnit = $modelLengthUnit
    LengthUnit = if ($LengthUnit -eq "Feet") { "ft" } else { $modelLengthUnit }
    ModelWasLocked = $modelWasLocked
    UnlockedForEdit = $unlockedForEdit
    SaveRequested = [bool]$Save
    SaveReturnCode = $saveReturnCode
    BackupRequested = [bool](-not $SkipBackup)
    BackupCreated = $backupCreated
    BackupPath = $backupPath
    BackupError = $backupError
    ResolvedBaseElevation = $resolvedBaseElevation
    ResolvedBaseElevationModel = $resolvedBaseElevationModel
    OrderChanged = $storySet.OrderChanged
    PreStoryCount = $preStories.StoryCount
    PreStories = @($preStories.Stories)
    RequestedStories = @($storySet.Stories)
    PostStoryCount = $postStories.StoryCount
    PostStories = @($postStories.Stories)
    PrePhysicalCounts = $prePhysicalCounts
    PostPhysicalCounts = $postPhysicalCounts
    Completed = ($postStories.StoryCount -eq $storySet.Stories.Count)
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
}
else {
    $result
}
