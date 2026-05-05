param(
    [int]$EtabsPid,

    [string]$OutputDirectory,

    [string]$ExpectedModelPath,

    [string[]]$Pier = @(),

    [string[]]$Story = @(),

    [switch]$OpenOutputDirectory,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-Success {
    param(
        [int]$ReturnCode,
        [string]$Operation
    )

    if ($ReturnCode -ne 0) {
        throw "$Operation failed with return code $ReturnCode."
    }
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

    throw "ETABSv1.dll was not found in the running ETABS folder or a standard ETABS install path."
}

function Resolve-CanonicalModelPath {
    param(
        [string]$RawModelPath,
        [string]$ModelDirectory
    )

    if ([string]::IsNullOrWhiteSpace($RawModelPath)) {
        return $null
    }

    if ([System.IO.Path]::GetExtension($RawModelPath) -ieq ".EDB") {
        return $RawModelPath
    }

    if (-not [string]::IsNullOrWhiteSpace($ModelDirectory)) {
        $candidate = Join-Path -Path $ModelDirectory -ChildPath ("{0}.EDB" -f [System.IO.Path]::GetFileNameWithoutExtension($RawModelPath))
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $RawModelPath
}

function Get-CurrentModelPath {
    param(
        $Api
    )

    $rawModelPath = $Api.SapModel.GetModelFilename($true)
    $modelDirectory = $Api.SapModel.GetModelFilepath()
    return Resolve-CanonicalModelPath -RawModelPath $rawModelPath -ModelDirectory $modelDirectory
}

function Resolve-ExistingPathOrNull {
    param(
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if (Test-Path -LiteralPath $PathValue) {
        return (Resolve-Path -LiteralPath $PathValue).Path
    }

    return $PathValue
}

function Get-PresentUnits {
    param(
        $SapModel
    )

    $force = [ETABSv1.eForce]::lb
    $length = [ETABSv1.eLength]::inch
    $temperature = [ETABSv1.eTemperature]::F
    Assert-Success ($SapModel.GetPresentUnits_2([ref]$force, [ref]$length, [ref]$temperature)) "Get present units"

    return [pscustomobject]@{
        Force = $force.ToString()
        Length = $length.ToString()
        Temperature = $temperature.ToString()
    }
}

function Get-LengthToFootFactor {
    param(
        [string]$LengthUnit
    )

    switch -Regex ($LengthUnit) {
        "^inch$" { return (1.0 / 12.0) }
        "^ft$" { return 1.0 }
        "^mm$" { return 0.0032808398950131233 }
        "^cm$" { return 0.032808398950131233 }
        "^m$" { return 3.2808398950131235 }
        default { throw "Unsupported length unit: $LengthUnit" }
    }
}

function Convert-AreaToSquareInches {
    param(
        [AllowNull()]
        [object]$Value,
        [string]$LengthUnit
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    $lengthFactor = Get-LengthToFootFactor -LengthUnit $LengthUnit
    return ([double]$Value) * [Math]::Pow($lengthFactor * 12.0, 2.0)
}

function Convert-LengthToFeet {
    param(
        [AllowNull()]
        [object]$Value,
        [string]$LengthUnit
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    return ([double]$Value) * (Get-LengthToFootFactor -LengthUnit $LengthUnit)
}

function Round-OrNull {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Digits = 3
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    return [Math]::Round([double]$Value, $Digits)
}

function Normalize-DesignMessage {
    param(
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message) -or $Message -eq "No Message") {
        return ""
    }

    return $Message.Trim()
}

function Add-Message {
    param(
        [System.Collections.Generic.List[string]]$Messages,
        [string]$Message
    )

    $normalized = Normalize-DesignMessage -Message $Message
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    if (-not $Messages.Contains($normalized)) {
        $Messages.Add($normalized) | Out-Null
    }
}

function Test-NameFilter {
    param(
        [string]$Value,
        [string[]]$Filters
    )

    if ($null -eq $Filters -or @($Filters).Count -eq 0) {
        return $true
    }

    foreach ($filter in $Filters) {
        if ($Value -like $filter) {
            return $true
        }
    }

    return $false
}

function Get-PierSortKey {
    param(
        [string]$PierLabel
    )

    if ($PierLabel -match "^([A-Za-z]+)(\d+)$") {
        return "{0}|{1:D8}" -f $Matches[1].ToUpperInvariant(), [int]$Matches[2]
    }

    return $PierLabel.ToUpperInvariant()
}

function Get-StorySortKey {
    param(
        [string]$StoryName
    )

    if ([string]::IsNullOrWhiteSpace($StoryName)) {
        return 500000
    }

    $normalized = $StoryName.Trim().ToLowerInvariant()
    if ($normalized -eq "roof") {
        return 200000
    }

    if ($normalized -eq "cellar") {
        return -1000
    }

    if ($normalized -match "^story\s*(\d+)$") {
        return [int]$Matches[1]
    }

    if ($normalized -match "^b(?:asement)?\s*(\d+)$") {
        return -1 * [int]$Matches[1]
    }

    return 100000
}

function New-DefaultOutputDirectory {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path -Path (Join-Path -Path (Get-Location) -ChildPath "out\etabs-shear-wall-required-steel") -ChildPath $timestamp
}

function Export-Rows {
    param(
        [string]$Path,
        [object[]]$Rows
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        [System.IO.File]::WriteAllText($Path, "", ([System.Text.UTF8Encoding]::new($false)))
        return
    }

    @($Rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Get-ShearWallStationRows {
    param(
        $SapModel,
        [string]$LengthUnit,
        [string[]]$PierFilter,
        [string[]]$StoryFilter
    )

    [string[]]$storyNames = @()
    [string[]]$pierLabels = @()
    [string[]]$stations = @()
    [string[]]$designTypes = @()
    [string[]]$pierSecTypes = @()
    [string[]]$edgeBars = @()
    [string[]]$endBars = @()
    [double[]]$barSpacing = @()
    [double[]]$reinfPercent = @()
    [double[]]$currPercent = @()
    [double[]]$dcRatio = @()
    [string[]]$pierLegs = @()
    [double[]]$legX1 = @()
    [double[]]$legY1 = @()
    [double[]]$legX2 = @()
    [double[]]$legY2 = @()
    [double[]]$edgeLeft = @()
    [double[]]$edgeRight = @()
    [double[]]$asLeft = @()
    [double[]]$asRight = @()
    [double[]]$shearAv = @()
    [double[]]$stressCompLeft = @()
    [double[]]$stressCompRight = @()
    [double[]]$stressLimitLeft = @()
    [double[]]$stressLimitRight = @()
    [double[]]$cDepthLeft = @()
    [double[]]$cLimitLeft = @()
    [double[]]$cDepthRight = @()
    [double[]]$cLimitRight = @()
    [double[]]$inelasticRotDemand = @()
    [double[]]$inelasticRotCapacity = @()
    [double[]]$normCompStress = @()
    [double[]]$normCompStressLimit = @()
    [double[]]$cDepth = @()
    [double[]]$bZoneL = @()
    [double[]]$bZoneR = @()
    [double[]]$bZoneLength = @()
    [string[]]$warnMsg = @()
    [string[]]$errMsg = @()

    Assert-Success (
        $SapModel.DesignShearWall.GetPierSummaryResults(
            [ref]$storyNames,
            [ref]$pierLabels,
            [ref]$stations,
            [ref]$designTypes,
            [ref]$pierSecTypes,
            [ref]$edgeBars,
            [ref]$endBars,
            [ref]$barSpacing,
            [ref]$reinfPercent,
            [ref]$currPercent,
            [ref]$dcRatio,
            [ref]$pierLegs,
            [ref]$legX1,
            [ref]$legY1,
            [ref]$legX2,
            [ref]$legY2,
            [ref]$edgeLeft,
            [ref]$edgeRight,
            [ref]$asLeft,
            [ref]$asRight,
            [ref]$shearAv,
            [ref]$stressCompLeft,
            [ref]$stressCompRight,
            [ref]$stressLimitLeft,
            [ref]$stressLimitRight,
            [ref]$cDepthLeft,
            [ref]$cLimitLeft,
            [ref]$cDepthRight,
            [ref]$cLimitRight,
            [ref]$inelasticRotDemand,
            [ref]$inelasticRotCapacity,
            [ref]$normCompStress,
            [ref]$normCompStressLimit,
            [ref]$cDepth,
            [ref]$bZoneL,
            [ref]$bZoneR,
            [ref]$bZoneLength,
            [ref]$warnMsg,
            [ref]$errMsg
        )
    ) "Get shear wall pier summary results"

    $rows = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $storyNames.Length; $index++) {
        if (-not (Test-NameFilter -Value $pierLabels[$index] -Filters $PierFilter)) {
            continue
        }

        if (-not (Test-NameFilter -Value $storyNames[$index] -Filters $StoryFilter)) {
            continue
        }

        $warn = Normalize-DesignMessage -Message $warnMsg[$index]
        $err = Normalize-DesignMessage -Message $errMsg[$index]
        $rows.Add([pscustomobject]@{
            Pier = $pierLabels[$index]
            Story = $storyNames[$index]
            Station = $stations[$index]
            PierLeg = $pierLegs[$index]
            DesignType = $designTypes[$index]
            PierSecType = $pierSecTypes[$index]
            EdgeBar = $edgeBars[$index]
            EndBar = $endBars[$index]
            BarSpacing_SourceLengthUnits = (Round-OrNull -Value $barSpacing[$index])
            RequiredReinfPercent = (Round-OrNull -Value $reinfPercent[$index] -Digits 6)
            CurrentReinfPercent = (Round-OrNull -Value $currPercent[$index] -Digits 6)
            DCRatio = (Round-OrNull -Value $dcRatio[$index] -Digits 6)
            AsLeft_SourceLengthSquared = (Round-OrNull -Value $asLeft[$index] -Digits 6)
            AsRight_SourceLengthSquared = (Round-OrNull -Value $asRight[$index] -Digits 6)
            AsLeft_in2 = (Round-OrNull -Value (Convert-AreaToSquareInches -Value $asLeft[$index] -LengthUnit $LengthUnit))
            AsRight_in2 = (Round-OrNull -Value (Convert-AreaToSquareInches -Value $asRight[$index] -LengthUnit $LengthUnit))
            ShearAv_SourceLengthSquaredPerLength = (Round-OrNull -Value $shearAv[$index] -Digits 6)
            EdgeLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $edgeLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            EdgeRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $edgeRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneL_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneL[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneR_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneR[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneLength_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneLength[$index] -LengthUnit $LengthUnit) -Digits 4)
            StressCompLeft = (Round-OrNull -Value $stressCompLeft[$index])
            StressCompRight = (Round-OrNull -Value $stressCompRight[$index])
            StressLimitLeft = (Round-OrNull -Value $stressLimitLeft[$index])
            StressLimitRight = (Round-OrNull -Value $stressLimitRight[$index])
            CDepthLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cDepthLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            CLimitLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cLimitLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            CDepthRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cDepthRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            CLimitRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cLimitRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            WarnMsg = $warn
            ErrMsg = $err
        }) | Out-Null
    }

    return $rows.ToArray()
}

function Get-EnvelopeRows {
    param(
        [object[]]$StationRows
    )

    $entries = @{}
    foreach ($row in $StationRows) {
        $key = "{0}|{1}" -f $row.Pier, $row.Story
        if (-not $entries.ContainsKey($key)) {
            $entries[$key] = [ordered]@{
                Pier = $row.Pier
                Story = $row.Story
                RequiredSteelLeft_in2 = $null
                LeftControlStation = ""
                LeftControlPierLeg = ""
                RequiredSteelRight_in2 = $null
                RightControlStation = ""
                RightControlPierLeg = ""
                MaxBZoneL_ft = $null
                MaxBZoneR_ft = $null
                MaxDCRatio = $null
                DesignTypes = New-Object System.Collections.Generic.List[string]
                PierSecTypes = New-Object System.Collections.Generic.List[string]
                Messages = New-Object System.Collections.Generic.List[string]
            }
        }

        $entry = $entries[$key]
        if (-not [string]::IsNullOrWhiteSpace($row.DesignType) -and -not $entry["DesignTypes"].Contains($row.DesignType)) {
            $entry["DesignTypes"].Add($row.DesignType) | Out-Null
        }

        if (-not [string]::IsNullOrWhiteSpace($row.PierSecType) -and -not $entry["PierSecTypes"].Contains($row.PierSecType)) {
            $entry["PierSecTypes"].Add($row.PierSecType) | Out-Null
        }

        Add-Message -Messages $entry["Messages"] -Message $row.WarnMsg
        Add-Message -Messages $entry["Messages"] -Message $row.ErrMsg

        if ($null -ne $row.AsLeft_in2 -and ($null -eq $entry["RequiredSteelLeft_in2"] -or $row.AsLeft_in2 -gt $entry["RequiredSteelLeft_in2"])) {
            $entry["RequiredSteelLeft_in2"] = $row.AsLeft_in2
            $entry["LeftControlStation"] = $row.Station
            $entry["LeftControlPierLeg"] = $row.PierLeg
        }

        if ($null -ne $row.AsRight_in2 -and ($null -eq $entry["RequiredSteelRight_in2"] -or $row.AsRight_in2 -gt $entry["RequiredSteelRight_in2"])) {
            $entry["RequiredSteelRight_in2"] = $row.AsRight_in2
            $entry["RightControlStation"] = $row.Station
            $entry["RightControlPierLeg"] = $row.PierLeg
        }

        if ($null -ne $row.BZoneL_ft -and ($null -eq $entry["MaxBZoneL_ft"] -or $row.BZoneL_ft -gt $entry["MaxBZoneL_ft"])) {
            $entry["MaxBZoneL_ft"] = $row.BZoneL_ft
        }

        if ($null -ne $row.BZoneR_ft -and ($null -eq $entry["MaxBZoneR_ft"] -or $row.BZoneR_ft -gt $entry["MaxBZoneR_ft"])) {
            $entry["MaxBZoneR_ft"] = $row.BZoneR_ft
        }

        if ($null -ne $row.DCRatio -and ($null -eq $entry["MaxDCRatio"] -or $row.DCRatio -gt $entry["MaxDCRatio"])) {
            $entry["MaxDCRatio"] = $row.DCRatio
        }
    }

    $rows = foreach ($entry in $entries.Values) {
        [pscustomobject]@{
            Pier = $entry["Pier"]
            Story = $entry["Story"]
            RequiredSteelLeft_in2 = $entry["RequiredSteelLeft_in2"]
            LeftControlStation = $entry["LeftControlStation"]
            LeftControlPierLeg = $entry["LeftControlPierLeg"]
            RequiredSteelRight_in2 = $entry["RequiredSteelRight_in2"]
            RightControlStation = $entry["RightControlStation"]
            RightControlPierLeg = $entry["RightControlPierLeg"]
            MaxBZoneL_ft = $entry["MaxBZoneL_ft"]
            MaxBZoneR_ft = $entry["MaxBZoneR_ft"]
            MaxDCRatio = $entry["MaxDCRatio"]
            DesignTypes = ($entry["DesignTypes"] -join " | ")
            PierSecTypes = ($entry["PierSecTypes"] -join " | ")
            Messages = ($entry["Messages"] -join " | ")
        }
    }

    return @($rows | Sort-Object @{ Expression = { Get-PierSortKey -PierLabel $_.Pier } }, @{ Expression = { Get-StorySortKey -StoryName $_.Story } })
}

function Get-WarningRows {
    param(
        [object[]]$StationRows
    )

    return @($StationRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.WarnMsg) -or -not [string]::IsNullOrWhiteSpace($_.ErrMsg)
        } | Sort-Object @{ Expression = { Get-PierSortKey -PierLabel $_.Pier } }, @{ Expression = { Get-StorySortKey -StoryName $_.Story } }, Station)
}

$process = Get-EtabsProcess -RequestedPid $EtabsPid
if ($null -eq $process) {
    throw "No running ETABS process was found."
}

$apiDllPath = Resolve-EtabsApiDll -Process $process
Add-Type -Path $apiDllPath

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$modelPath = Get-CurrentModelPath -Api $api
$resolvedExpectedModelPath = Resolve-ExistingPathOrNull -PathValue $ExpectedModelPath
if (-not [string]::IsNullOrWhiteSpace($resolvedExpectedModelPath) -and -not [string]::Equals($modelPath, $resolvedExpectedModelPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The live ETABS model does not match ExpectedModelPath. Live: '$modelPath'. Expected: '$resolvedExpectedModelPath'."
}

$units = Get-PresentUnits -SapModel $api.SapModel
$resolvedOutputDirectory = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { New-DefaultOutputDirectory } else { $OutputDirectory }
New-Item -ItemType Directory -Force -Path $resolvedOutputDirectory | Out-Null

$modelFile = if (-not [string]::IsNullOrWhiteSpace($modelPath) -and (Test-Path -LiteralPath $modelPath -PathType Leaf)) { Get-Item -LiteralPath $modelPath } else { $null }
$modelLastWriteTime = if ($null -ne $modelFile) { $modelFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
$modelFileLength = if ($null -ne $modelFile) { $modelFile.Length } else { "" }
$stationRows = Get-ShearWallStationRows -SapModel $api.SapModel -LengthUnit $units.Length -PierFilter $Pier -StoryFilter $Story
$stationRowCount = @($stationRows).Count
if (-not $stationRows -or $stationRowCount -eq 0) {
    throw "No shear wall pier summary results were returned from the active model for the selected filters."
}

$envelopeRows = Get-EnvelopeRows -StationRows $stationRows
$warningRows = Get-WarningRows -StationRows $stationRows
$envelopeRowCount = @($envelopeRows).Count
$warningRowCount = @($warningRows).Count
$infoRows = @(
    [pscustomobject]@{ Field = "ExportedAtLocal"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") },
    [pscustomobject]@{ Field = "ProcessId"; Value = $process.Id },
    [pscustomobject]@{ Field = "MainWindowTitle"; Value = $process.MainWindowTitle },
    [pscustomobject]@{ Field = "ModelPath"; Value = $modelPath },
    [pscustomobject]@{ Field = "ModelLastWriteTime"; Value = $modelLastWriteTime },
    [pscustomobject]@{ Field = "ModelFileLength"; Value = $modelFileLength },
    [pscustomobject]@{ Field = "ExpectedModelPath"; Value = $resolvedExpectedModelPath },
    [pscustomobject]@{ Field = "ApiDllPath"; Value = $apiDllPath },
    [pscustomobject]@{ Field = "ForceUnits"; Value = $units.Force },
    [pscustomobject]@{ Field = "LengthUnits"; Value = $units.Length },
    [pscustomobject]@{ Field = "OutputSteelAreaUnits"; Value = "in^2" },
    [pscustomobject]@{ Field = "OutputLengthUnits"; Value = "ft" },
    [pscustomobject]@{ Field = "PierFilter"; Value = ($Pier -join ", ") },
    [pscustomobject]@{ Field = "StoryFilter"; Value = ($Story -join ", ") },
    [pscustomobject]@{ Field = "RawStationRowCount"; Value = $stationRowCount },
    [pscustomobject]@{ Field = "StoryEnvelopeRowCount"; Value = $envelopeRowCount },
    [pscustomobject]@{ Field = "WarningRowCount"; Value = $warningRowCount },
    [pscustomobject]@{ Field = "Source"; Value = "SapModel.DesignShearWall.GetPierSummaryResults" },
    [pscustomobject]@{ Field = "EnvelopeRule"; Value = "Per pier/story, use max AsLeft and max AsRight across ETABS station rows; keep controlling station columns." }
)

$infoPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "info.csv"
$stationPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "raw-station-results.csv"
$envelopePath = Join-Path -Path $resolvedOutputDirectory -ChildPath "story-envelope.csv"
$warningPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "warnings.csv"

Export-Rows -Path $infoPath -Rows $infoRows
Export-Rows -Path $stationPath -Rows $stationRows
Export-Rows -Path $envelopePath -Rows $envelopeRows
Export-Rows -Path $warningPath -Rows $warningRows

if ($OpenOutputDirectory) {
    Start-Process -FilePath $resolvedOutputDirectory | Out-Null
}

$result = [pscustomobject]@{
    ProcessId = $process.Id
    ModelPath = $modelPath
    OutputDirectory = (Resolve-Path -LiteralPath $resolvedOutputDirectory).Path
    InfoCsv = $infoPath
    RawStationResultsCsv = $stationPath
    StoryEnvelopeCsv = $envelopePath
    WarningsCsv = $warningPath
    ForceUnits = $units.Force
    LengthUnits = $units.Length
    RawStationRowCount = $stationRowCount
    StoryEnvelopeRowCount = $envelopeRowCount
    WarningRowCount = $warningRowCount
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
