param(
    [string]$ResultName = "ASD ENV",

    [ValidateSet("Combo", "Case")]
    [string]$ResultType = "Combo",

    [string]$AreaObject = "all",

    [switch]$SelectedOnly,

    [switch]$WallOnly,

    [int]$EtabsPid,

    [int]$Top = 20,

    [string]$DetailCsvPath,

    [string]$EnvelopeCsvPath,

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

function Get-LoadCaseNames {
    param(
        $SapModel
    )

    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($caseType in [System.Enum]::GetValues([ETABSv1.eLoadCaseType])) {
        $numberNames = 0
        [string[]]$caseNames = @()

        try {
            $ret = $SapModel.LoadCases.GetNameList([ref]$numberNames, [ref]$caseNames, $caseType)
            if ($ret -eq 0) {
                foreach ($caseName in $caseNames) {
                    if (-not [string]::IsNullOrWhiteSpace($caseName)) {
                        $names.Add($caseName) | Out-Null
                    }
                }
            }
        }
        catch {
        }
    }

    return [string[]]($names | ForEach-Object { $_ })
}

function Get-ComboNames {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$comboNames = @()
    Assert-Success ($SapModel.RespCombo.GetNameList([ref]$numberNames, [ref]$comboNames)) "Get load combination names"
    return @($comboNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-OutputSelectionState {
    param(
        $SapModel
    )

    $selectedCases = New-Object System.Collections.Generic.List[string]
    foreach ($caseName in (Get-LoadCaseNames -SapModel $SapModel)) {
        $selected = $false
        Assert-Success ($SapModel.Results.Setup.GetCaseSelectedForOutput($caseName, [ref]$selected)) "Get output selection for case $caseName"
        if ($selected) {
            $selectedCases.Add($caseName) | Out-Null
        }
    }

    $selectedCombos = New-Object System.Collections.Generic.List[string]
    foreach ($comboName in (Get-ComboNames -SapModel $SapModel)) {
        $selected = $false
        Assert-Success ($SapModel.Results.Setup.GetComboSelectedForOutput($comboName, [ref]$selected)) "Get output selection for combo $comboName"
        if ($selected) {
            $selectedCombos.Add($comboName) | Out-Null
        }
    }

    return [pscustomobject]@{
        Cases = @($selectedCases)
        Combos = @($selectedCombos)
    }
}

function Restore-OutputSelectionState {
    param(
        $SapModel,
        [pscustomobject]$State
    )

    Assert-Success ($SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()) "Deselect all cases and combos for output"

    foreach ($caseName in $State.Cases) {
        Assert-Success ($SapModel.Results.Setup.SetCaseSelectedForOutput($caseName, $true)) "Restore output selection for case $caseName"
    }

    foreach ($comboName in $State.Combos) {
        Assert-Success ($SapModel.Results.Setup.SetComboSelectedForOutput($comboName, $true)) "Restore output selection for combo $comboName"
    }
}

function Get-ComboTypeName {
    param(
        [int]$ComboType
    )

    switch ($ComboType) {
        0 { return "LinearAdd" }
        1 { return "Envelope" }
        2 { return "AbsoluteAdd" }
        3 { return "SRSS" }
        4 { return "RangeAdd" }
        default { return "Unknown($ComboType)" }
    }
}

function Get-LeafKey {
    param(
        [string]$SourceType,
        [string]$SourceName,
        [double]$ScaleFactor
    )

    return "{0}|{1}|{2}" -f $SourceType.ToUpperInvariant(), $SourceName.ToUpperInvariant(), ("{0:R}" -f $ScaleFactor)
}

function Expand-EnvelopeLeafSources {
    param(
        $SapModel,
        [string]$ComboName,
        [double]$ParentScaleFactor,
        [System.Collections.Generic.List[object]]$LeafSources,
        [System.Collections.Generic.HashSet[string]]$LeafKeys,
        [System.Collections.Generic.HashSet[string]]$VisitedEnvelopeCombos,
        [System.Collections.Generic.List[object]]$ExpansionTrace
    )

    if (-not $VisitedEnvelopeCombos.Add($ComboName)) {
        return
    }

    $comboType = 0
    Assert-Success ($SapModel.RespCombo.GetTypeCombo($ComboName, [ref]$comboType)) "Get combo type for $ComboName"

    if ($comboType -ne 1) {
        $leafKey = Get-LeafKey -SourceType "Combo" -SourceName $ComboName -ScaleFactor $ParentScaleFactor
        if ($LeafKeys.Add($leafKey)) {
            $LeafSources.Add([pscustomobject]@{
                SourceType = "Combo"
                SourceName = $ComboName
                ScaleFactor = [double]$ParentScaleFactor
            }) | Out-Null
        }

        return
    }

    $ExpansionTrace.Add([pscustomobject]@{
        Combo = $ComboName
        ComboType = "Envelope"
        ScaleFactor = [double]$ParentScaleFactor
    }) | Out-Null

    $itemCount = 0
    [ETABSv1.eCNameType[]]$itemTypes = @()
    [string[]]$itemNames = @()
    [double[]]$scaleFactors = @()
    Assert-Success ($SapModel.RespCombo.GetCaseList($ComboName, [ref]$itemCount, [ref]$itemTypes, [ref]$itemNames, [ref]$scaleFactors)) "Get combo items for $ComboName"

    for ($i = 0; $i -lt $itemCount; $i++) {
        $itemTypeName = $itemTypes[$i].ToString()
        $itemName = $itemNames[$i]
        $itemScaleFactor = [double]$scaleFactors[$i]
        $netScaleFactor = [double]($ParentScaleFactor * $itemScaleFactor)

        if ($itemTypeName -eq "LoadCase") {
            $leafKey = Get-LeafKey -SourceType "Case" -SourceName $itemName -ScaleFactor $netScaleFactor
            if ($LeafKeys.Add($leafKey)) {
                $LeafSources.Add([pscustomobject]@{
                    SourceType = "Case"
                    SourceName = $itemName
                    ScaleFactor = $netScaleFactor
                }) | Out-Null
            }

            continue
        }

        if ($itemTypeName -eq "LoadCombo") {
            $childComboType = 0
            Assert-Success ($SapModel.RespCombo.GetTypeCombo($itemName, [ref]$childComboType)) "Get combo type for $itemName"

            if ($childComboType -eq 1) {
                Expand-EnvelopeLeafSources -SapModel $SapModel -ComboName $itemName -ParentScaleFactor $netScaleFactor -LeafSources $LeafSources -LeafKeys $LeafKeys -VisitedEnvelopeCombos $VisitedEnvelopeCombos -ExpansionTrace $ExpansionTrace
            }
            else {
                $leafKey = Get-LeafKey -SourceType "Combo" -SourceName $itemName -ScaleFactor $netScaleFactor
                if ($LeafKeys.Add($leafKey)) {
                    $LeafSources.Add([pscustomobject]@{
                        SourceType = "Combo"
                        SourceName = $itemName
                        ScaleFactor = $netScaleFactor
                    }) | Out-Null
                }
            }

            continue
        }

        throw "Unsupported combo item type '$itemTypeName' in combo '$ComboName'."
    }
}

function Resolve-RequestedSources {
    param(
        $SapModel,
        [string]$Name,
        [string]$Type
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw "ResultName cannot be blank."
    }

    if ($Type -eq "Case") {
        return [pscustomobject]@{
            RequestedResultName = $Name
            RequestedResultType = "Case"
            RequestedComboType = $null
            SourceMode = "Direct"
            Sources = @(
                [pscustomobject]@{
                    SourceType = "Case"
                    SourceName = $Name
                    ScaleFactor = 1.0
                }
            )
            CasesToSelect = @($Name)
            CombosToSelect = @()
            ExpansionTrace = @()
        }
    }

    $comboType = 0
    Assert-Success ($SapModel.RespCombo.GetTypeCombo($Name, [ref]$comboType)) "Get combo type for $Name"

    if ($comboType -ne 1) {
        return [pscustomobject]@{
            RequestedResultName = $Name
            RequestedResultType = "Combo"
            RequestedComboType = Get-ComboTypeName -ComboType $comboType
            SourceMode = "Direct"
            Sources = @(
                [pscustomobject]@{
                    SourceType = "Combo"
                    SourceName = $Name
                    ScaleFactor = 1.0
                }
            )
            CasesToSelect = @()
            CombosToSelect = @($Name)
            ExpansionTrace = @()
        }
    }

    $leafSources = New-Object System.Collections.Generic.List[object]
    $leafKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $visitedEnvelopeCombos = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $expansionTrace = New-Object System.Collections.Generic.List[object]

    Expand-EnvelopeLeafSources -SapModel $SapModel -ComboName $Name -ParentScaleFactor 1.0 -LeafSources $leafSources -LeafKeys $leafKeys -VisitedEnvelopeCombos $visitedEnvelopeCombos -ExpansionTrace $expansionTrace

    $casesToSelect = @($leafSources | Where-Object { $_.SourceType -eq "Case" } | Select-Object -ExpandProperty SourceName -Unique)
    $combosToSelect = @($leafSources | Where-Object { $_.SourceType -eq "Combo" } | Select-Object -ExpandProperty SourceName -Unique)

    return [pscustomobject]@{
        RequestedResultName = $Name
        RequestedResultType = "Combo"
        RequestedComboType = "Envelope"
        SourceMode = "ExpandedEnvelope"
        Sources = $leafSources.ToArray()
        CasesToSelect = $casesToSelect
        CombosToSelect = $combosToSelect
        ExpansionTrace = $expansionTrace.ToArray()
    }
}

function Select-RequestedOutputs {
    param(
        $SapModel,
        [pscustomobject]$RequestedSources
    )

    Assert-Success ($SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()) "Deselect all cases and combos for output"

    foreach ($caseName in $RequestedSources.CasesToSelect) {
        Assert-Success ($SapModel.Results.Setup.SetCaseSelectedForOutput($caseName, $true)) "Select output case $caseName"
    }

    foreach ($comboName in $RequestedSources.CombosToSelect) {
        Assert-Success ($SapModel.Results.Setup.SetComboSelectedForOutput($comboName, $true)) "Select output combo $comboName"
    }
}

function Get-PrincipalMax {
    param(
        [double]$S11,
        [double]$S22,
        [double]$S12
    )

    $avg = ($S11 + $S22) / 2.0
    $radius = [math]::Sqrt([math]::Pow(($S11 - $S22) / 2.0, 2.0) + [math]::Pow($S12, 2.0))
    return $avg + $radius
}

function Get-PrincipalMin {
    param(
        [double]$S11,
        [double]$S22,
        [double]$S12
    )

    $avg = ($S11 + $S22) / 2.0
    $radius = [math]::Sqrt([math]::Pow(($S11 - $S22) / 2.0, 2.0) + [math]::Pow($S12, 2.0))
    return $avg - $radius
}

function Get-AreaStressShellResults {
    param(
        $SapModel,
        [string]$AreaObjectName,
        [switch]$UseSelectedObjectsOnly
    )

    $name = "ALL"
    $itemType = [ETABSv1.eItemTypeElm]::GroupElm

    if ($UseSelectedObjectsOnly) {
        $name = ""
        $itemType = [ETABSv1.eItemTypeElm]::SelectionElm
    }
    elseif (-not [string]::IsNullOrWhiteSpace($AreaObjectName) -and $AreaObjectName -notin @("all", "*")) {
        if ($AreaObjectName.Contains("*") -or $AreaObjectName.Contains("?")) {
            $name = "ALL"
            $itemType = [ETABSv1.eItemTypeElm]::GroupElm
        }
        else {
            $name = $AreaObjectName
            $itemType = [ETABSv1.eItemTypeElm]::ObjectElm
        }
    }

    $numberResults = 0
    [string[]]$obj = @()
    [string[]]$elm = @()
    [string[]]$pointElm = @()
    [string[]]$loadCase = @()
    [string[]]$stepType = @()
    [double[]]$stepNum = @()
    [double[]]$s11Top = @()
    [double[]]$s22Top = @()
    [double[]]$s12Top = @()
    [double[]]$sMaxTop = @()
    [double[]]$sMinTop = @()
    [double[]]$sAngleTop = @()
    [double[]]$svmTop = @()
    [double[]]$s11Bot = @()
    [double[]]$s22Bot = @()
    [double[]]$s12Bot = @()
    [double[]]$sMaxBot = @()
    [double[]]$sMinBot = @()
    [double[]]$sAngleBot = @()
    [double[]]$svmBot = @()
    [double[]]$s13Avg = @()
    [double[]]$s23Avg = @()
    [double[]]$sMaxAvg = @()
    [double[]]$sAngleAvg = @()

    Assert-Success ($SapModel.Results.AreaStressShell(
            $name,
            $itemType,
            [ref]$numberResults,
            [ref]$obj,
            [ref]$elm,
            [ref]$pointElm,
            [ref]$loadCase,
            [ref]$stepType,
            [ref]$stepNum,
            [ref]$s11Top,
            [ref]$s22Top,
            [ref]$s12Top,
            [ref]$sMaxTop,
            [ref]$sMinTop,
            [ref]$sAngleTop,
            [ref]$svmTop,
            [ref]$s11Bot,
            [ref]$s22Bot,
            [ref]$s12Bot,
            [ref]$sMaxBot,
            [ref]$sMinBot,
            [ref]$sAngleBot,
            [ref]$svmBot,
            [ref]$s13Avg,
            [ref]$s23Avg,
            [ref]$sMaxAvg,
            [ref]$sAngleAvg)) "Get shell stress results"

    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $numberResults; $i++) {
        $rows.Add([pscustomobject]@{
            ObjectName = $obj[$i]
            ElementName = $elm[$i]
            PointElement = $pointElm[$i]
            OutputName = $loadCase[$i]
            StepType = $stepType[$i]
            StepNumber = [double]$stepNum[$i]
            S11Top = [double]$s11Top[$i]
            S22Top = [double]$s22Top[$i]
            S12Top = [double]$s12Top[$i]
            S11Bottom = [double]$s11Bot[$i]
            S22Bottom = [double]$s22Bot[$i]
            S12Bottom = [double]$s12Bot[$i]
        }) | Out-Null
    }

    return [object[]]($rows | ForEach-Object { $_ })
}

function Get-WallAreaObjectNames {
    param(
        $SapModel,
        [string]$AreaObjectFilter
    )

    $numberNames = 0
    [string[]]$areaNames = @()
    Assert-Success ($SapModel.AreaObj.GetNameList([ref]$numberNames, [ref]$areaNames)) "Get area object names"

    $matchedNames = New-Object System.Collections.Generic.List[string]
    $useFilter = -not [string]::IsNullOrWhiteSpace($AreaObjectFilter) -and $AreaObjectFilter -notin @("all", "*")
    $useWildcard = $useFilter -and ($AreaObjectFilter.Contains("*") -or $AreaObjectFilter.Contains("?"))

    foreach ($areaName in $areaNames) {
        if ($useFilter) {
            if ($useWildcard) {
                if ($areaName -notlike $AreaObjectFilter) {
                    continue
                }
            }
            elseif (-not [string]::Equals($areaName, $AreaObjectFilter, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
        }

        $designOrientation = [ETABSv1.eAreaDesignOrientation]::Null
        $ret = $SapModel.AreaObj.GetDesignOrientation($areaName, [ref]$designOrientation)
        if ($ret -ne 0) {
            continue
        }

        if ($designOrientation.ToString() -eq "Wall") {
            $matchedNames.Add($areaName) | Out-Null
        }
    }

    return [string[]]($matchedNames | ForEach-Object { $_ })
}

function Get-AreaStressShellResultsForObjects {
    param(
        $SapModel,
        [string[]]$ObjectNames
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($objectName in $ObjectNames) {
        foreach ($row in (Get-AreaStressShellResults -SapModel $SapModel -AreaObjectName $objectName)) {
            $rows.Add($row) | Out-Null
        }
    }

    return [object[]]($rows | ForEach-Object { $_ })
}

function Get-StoryLookup {
    param(
        $SapModel,
        [string[]]$ObjectNames
    )

    $storyLookup = @{}
    foreach ($objectName in ($ObjectNames | Sort-Object -Unique)) {
        $label = ""
        $story = ""
        try {
            $ret = $SapModel.AreaObj.GetLabelFromName($objectName, [ref]$label, [ref]$story)
            if ($ret -eq 0) {
                $storyLookup[$objectName] = [pscustomobject]@{
                    Label = $label
                    Story = $story
                }
            }
            else {
                $storyLookup[$objectName] = [pscustomobject]@{
                    Label = ""
                    Story = ""
                }
            }
        }
        catch {
            $storyLookup[$objectName] = [pscustomobject]@{
                Label = ""
                Story = ""
            }
        }
    }

    return $storyLookup
}

function Get-RequestedSourceLookup {
    param(
        [pscustomobject]$RequestedSources
    )

    $lookup = @{}
    foreach ($source in $RequestedSources.Sources) {
        $key = "{0}|{1}" -f $source.SourceType.ToUpperInvariant(), $source.SourceName.ToUpperInvariant()
        if (-not $lookup.ContainsKey($key)) {
            $lookup[$key] = New-Object System.Collections.Generic.List[object]
        }

        $lookup[$key].Add($source) | Out-Null
    }

    return $lookup
}

function New-ComputedDetailRows {
    param(
        [object[]]$RawRows,
        [pscustomobject]$RequestedSources,
        [hashtable]$StoryLookup
    )

    $selectedCaseNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($caseName in $RequestedSources.CasesToSelect) {
        $selectedCaseNames.Add($caseName) | Out-Null
    }

    $selectedComboNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($comboName in $RequestedSources.CombosToSelect) {
        $selectedComboNames.Add($comboName) | Out-Null
    }

    $requestedSourceLookup = Get-RequestedSourceLookup -RequestedSources $RequestedSources
    $detailRows = New-Object System.Collections.Generic.List[object]

    foreach ($rawRow in $RawRows) {
        if ($selectedComboNames.Contains($rawRow.OutputName)) {
            $outputType = "Combo"
        }
        elseif ($selectedCaseNames.Contains($rawRow.OutputName)) {
            $outputType = "Case"
        }
        else {
            continue
        }

        $lookupKey = "{0}|{1}" -f $outputType.ToUpperInvariant(), $rawRow.OutputName.ToUpperInvariant()
        if (-not $requestedSourceLookup.ContainsKey($lookupKey)) {
            continue
        }

        $storyInfo = if ($StoryLookup.ContainsKey($rawRow.ObjectName)) { $StoryLookup[$rawRow.ObjectName] } else { [pscustomobject]@{ Label = ""; Story = "" } }

        foreach ($requestedSource in $requestedSourceLookup[$lookupKey]) {
            $scaledS11Top = [double]($requestedSource.ScaleFactor * $rawRow.S11Top)
            $scaledS22Top = [double]($requestedSource.ScaleFactor * $rawRow.S22Top)
            $scaledS12Top = [double]($requestedSource.ScaleFactor * $rawRow.S12Top)
            $scaledS11Bottom = [double]($requestedSource.ScaleFactor * $rawRow.S11Bottom)
            $scaledS22Bottom = [double]($requestedSource.ScaleFactor * $rawRow.S22Bottom)
            $scaledS12Bottom = [double]($requestedSource.ScaleFactor * $rawRow.S12Bottom)

            $computedSMaxTop = Get-PrincipalMax -S11 $scaledS11Top -S22 $scaledS22Top -S12 $scaledS12Top
            $computedSMinTop = Get-PrincipalMin -S11 $scaledS11Top -S22 $scaledS22Top -S12 $scaledS12Top
            $computedSMaxBottom = Get-PrincipalMax -S11 $scaledS11Bottom -S22 $scaledS22Bottom -S12 $scaledS12Bottom
            $computedSMinBottom = Get-PrincipalMin -S11 $scaledS11Bottom -S22 $scaledS22Bottom -S12 $scaledS12Bottom

            $governingSurface = if ($computedSMaxTop -ge $computedSMaxBottom) { "Top" } else { "Bottom" }
            $governingSMax = [double]([math]::Max($computedSMaxTop, $computedSMaxBottom))
            $governingSMin = if ($governingSurface -eq "Top") { $computedSMinTop } else { $computedSMinBottom }

            $detailRows.Add([pscustomobject]@{
                RequestedResultName = $RequestedSources.RequestedResultName
                RequestedResultType = $RequestedSources.RequestedResultType
                RequestedComboType = $RequestedSources.RequestedComboType
                SourceMode = $RequestedSources.SourceMode
                Story = $storyInfo.Story
                Label = $storyInfo.Label
                ObjectName = $rawRow.ObjectName
                ElementName = $rawRow.ElementName
                PointElement = $rawRow.PointElement
                OutputType = $outputType
                OutputName = $rawRow.OutputName
                ScaleFactor = [double]$requestedSource.ScaleFactor
                StepType = $rawRow.StepType
                StepNumber = [double]$rawRow.StepNumber
                S11Top = $scaledS11Top
                S22Top = $scaledS22Top
                S12Top = $scaledS12Top
                ComputedSMaxTop = [double]$computedSMaxTop
                ComputedSMinTop = [double]$computedSMinTop
                S11Bottom = $scaledS11Bottom
                S22Bottom = $scaledS22Bottom
                S12Bottom = $scaledS12Bottom
                ComputedSMaxBottom = [double]$computedSMaxBottom
                ComputedSMinBottom = [double]$computedSMinBottom
                GoverningSurface = $governingSurface
                GoverningSMax = $governingSMax
                GoverningSMin = [double]$governingSMin
            }) | Out-Null
        }
    }

    return [object[]]($detailRows | ForEach-Object { $_ })
}

function New-EnvelopeRows {
    param(
        [object[]]$DetailRows
    )

    $groups = $DetailRows | Group-Object Story, Label, ObjectName, ElementName, PointElement
    $envelopeRows = New-Object System.Collections.Generic.List[object]

    foreach ($group in $groups) {
        $rows = @($group.Group)
        $maxTopRow = $rows | Sort-Object ComputedSMaxTop -Descending | Select-Object -First 1
        $maxBottomRow = $rows | Sort-Object ComputedSMaxBottom -Descending | Select-Object -First 1
        $governingRow = $rows | Sort-Object GoverningSMax -Descending | Select-Object -First 1

        $envelopeRows.Add([pscustomobject]@{
            RequestedResultName = $governingRow.RequestedResultName
            RequestedResultType = $governingRow.RequestedResultType
            RequestedComboType = $governingRow.RequestedComboType
            SourceMode = $governingRow.SourceMode
            Story = $governingRow.Story
            Label = $governingRow.Label
            ObjectName = $governingRow.ObjectName
            ElementName = $governingRow.ElementName
            PointElement = $governingRow.PointElement
            EnvelopeSMaxTop = [double]$maxTopRow.ComputedSMaxTop
            EnvelopeTopOutputType = $maxTopRow.OutputType
            EnvelopeTopOutputName = $maxTopRow.OutputName
            EnvelopeTopScaleFactor = [double]$maxTopRow.ScaleFactor
            EnvelopeTopStepType = $maxTopRow.StepType
            EnvelopeTopStepNumber = [double]$maxTopRow.StepNumber
            EnvelopeSMaxBottom = [double]$maxBottomRow.ComputedSMaxBottom
            EnvelopeBottomOutputType = $maxBottomRow.OutputType
            EnvelopeBottomOutputName = $maxBottomRow.OutputName
            EnvelopeBottomScaleFactor = [double]$maxBottomRow.ScaleFactor
            EnvelopeBottomStepType = $maxBottomRow.StepType
            EnvelopeBottomStepNumber = [double]$maxBottomRow.StepNumber
            GoverningSurface = $governingRow.GoverningSurface
            GoverningSMax = [double]$governingRow.GoverningSMax
            GoverningOutputType = $governingRow.OutputType
            GoverningOutputName = $governingRow.OutputName
            GoverningScaleFactor = [double]$governingRow.ScaleFactor
            GoverningStepType = $governingRow.StepType
            GoverningStepNumber = [double]$governingRow.StepNumber
        }) | Out-Null
    }

    return [object[]]($envelopeRows | ForEach-Object { $_ })
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
$originalOutputSelection = Get-OutputSelectionState -SapModel $sapModel

try {
    $requestedSources = Resolve-RequestedSources -SapModel $sapModel -Name $ResultName -Type $ResultType
    Select-RequestedOutputs -SapModel $sapModel -RequestedSources $requestedSources

    $queriedObjectNames = @()

    if ($WallOnly) {
        $queriedObjectNames = Get-WallAreaObjectNames -SapModel $sapModel -AreaObjectFilter $AreaObject
        if ($queriedObjectNames.Count -eq 0) {
            throw "No wall area objects matched the requested filter."
        }

        $rawRows = Get-AreaStressShellResultsForObjects -SapModel $sapModel -ObjectNames $queriedObjectNames
    }
    else {
        $rawRows = Get-AreaStressShellResults -SapModel $sapModel -AreaObjectName $AreaObject -UseSelectedObjectsOnly:$SelectedOnly

        if ($SelectedOnly -and $AreaObject -notin @("all", "*")) {
            $rawRows = @($rawRows | Where-Object { $_.ObjectName -like $AreaObject })
        }
        elseif (-not $SelectedOnly -and -not [string]::IsNullOrWhiteSpace($AreaObject) -and $AreaObject -notin @("all", "*") -and ($AreaObject.Contains("*") -or $AreaObject.Contains("?"))) {
            $rawRows = @($rawRows | Where-Object { $_.ObjectName -like $AreaObject })
        }
    }

    $storyLookup = Get-StoryLookup -SapModel $sapModel -ObjectNames @($rawRows | Select-Object -ExpandProperty ObjectName -Unique)
    $detailRows = New-ComputedDetailRows -RawRows $rawRows -RequestedSources $requestedSources -StoryLookup $storyLookup

    if ($detailRows.Count -eq 0) {
        throw "No shell stress rows matched the requested output selection."
    }

    $envelopeRows = New-EnvelopeRows -DetailRows $detailRows

    if (-not [string]::IsNullOrWhiteSpace($DetailCsvPath)) {
        $resolvedDetailCsvPath = [System.IO.Path]::GetFullPath($DetailCsvPath)
        $detailRows | Export-Csv -LiteralPath $resolvedDetailCsvPath -NoTypeInformation
        $DetailCsvPath = $resolvedDetailCsvPath
    }

    if (-not [string]::IsNullOrWhiteSpace($EnvelopeCsvPath)) {
        $resolvedEnvelopeCsvPath = [System.IO.Path]::GetFullPath($EnvelopeCsvPath)
        $envelopeRows | Export-Csv -LiteralPath $resolvedEnvelopeCsvPath -NoTypeInformation
        $EnvelopeCsvPath = $resolvedEnvelopeCsvPath
    }

    $maxRow = $envelopeRows | Sort-Object GoverningSMax -Descending | Select-Object -First 1
    $units = Get-PresentUnits -SapModel $sapModel
    $currentModelPath = Get-CurrentModelPath -Api $api

    $result = [pscustomobject]@{
        ModelPath = $currentModelPath
        EtabsProcessId = $process.Id
        RequestedResultName = $requestedSources.RequestedResultName
        RequestedResultType = $requestedSources.RequestedResultType
        RequestedComboType = $requestedSources.RequestedComboType
        SourceMode = $requestedSources.SourceMode
        Units = $units
        SelectedOnly = [bool]$SelectedOnly
        WallOnly = [bool]$WallOnly
        AreaObjectFilter = $AreaObject
        QueriedObjectCount = @($rawRows | Select-Object -ExpandProperty ObjectName -Unique).Count
        SelectedCasesForOutput = @($requestedSources.CasesToSelect)
        SelectedCombosForOutput = @($requestedSources.CombosToSelect)
        SourceExpansion = @($requestedSources.Sources)
        ExpansionTrace = @($requestedSources.ExpansionTrace)
        DetailRowCount = $detailRows.Count
        EnvelopeRowCount = $envelopeRows.Count
        MaxComputedSMax = [double]$maxRow.GoverningSMax
        MaxComputedSMaxLocation = [pscustomobject]@{
            Story = $maxRow.Story
            Label = $maxRow.Label
            ObjectName = $maxRow.ObjectName
            ElementName = $maxRow.ElementName
            PointElement = $maxRow.PointElement
            GoverningSurface = $maxRow.GoverningSurface
            GoverningOutputType = $maxRow.GoverningOutputType
            GoverningOutputName = $maxRow.GoverningOutputName
            GoverningScaleFactor = [double]$maxRow.GoverningScaleFactor
            GoverningStepType = $maxRow.GoverningStepType
            GoverningStepNumber = [double]$maxRow.GoverningStepNumber
        }
        DetailCsvPath = $DetailCsvPath
        EnvelopeCsvPath = $EnvelopeCsvPath
        TopEnvelopeRows = @($envelopeRows | Sort-Object GoverningSMax -Descending | Select-Object -First $Top)
    }

    if ($AsJson) {
        $result | ConvertTo-Json -Depth 8
    }
    else {
        $summary = [pscustomobject]@{
            ModelPath = $result.ModelPath
            RequestedResult = $result.RequestedResultName
            RequestedResultType = $result.RequestedResultType
            RequestedComboType = $result.RequestedComboType
            SourceMode = $result.SourceMode
            DetailRowCount = $result.DetailRowCount
            EnvelopeRowCount = $result.EnvelopeRowCount
            MaxComputedSMax = $result.MaxComputedSMax
            Units = "{0}/{1}^2" -f $result.Units.Force, $result.Units.Length
            MaxLocation = "{0} | {1} | {2} | {3}" -f $result.MaxComputedSMaxLocation.Story, $result.MaxComputedSMaxLocation.ObjectName, $result.MaxComputedSMaxLocation.ElementName, $result.MaxComputedSMaxLocation.PointElement
            GoverningSource = "{0} {1} x {2}" -f $result.MaxComputedSMaxLocation.GoverningOutputType, $result.MaxComputedSMaxLocation.GoverningOutputName, $result.MaxComputedSMaxLocation.GoverningScaleFactor
        }

        $summary
        ""
        $result.TopEnvelopeRows |
            Select-Object Story, ObjectName, ElementName, PointElement, GoverningSurface, GoverningSMax, GoverningOutputName, GoverningScaleFactor |
            Format-Table -AutoSize
    }
}
finally {
    Restore-OutputSelectionState -SapModel $sapModel -State $originalOutputSelection
}
