param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SourceModelPath,

    [int]$TargetEtabsPid,

    [switch]$ReplaceExisting,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

    throw "ETABSv1.dll was not found."
}

function Assert-Success {
    param(
        [int]$ReturnCode,
        [string]$Operation
    )

    if ($ReturnCode -ne 0) {
        throw "$Operation failed with return code $ReturnCode."
    }
}

function Get-ComboDefinitions {
    param(
        $SapModel
    )

    $numberNames = 0
    [string[]]$comboNames = @()
    Assert-Success ($SapModel.RespCombo.GetNameList([ref]$numberNames, [ref]$comboNames)) "Get combo names"

    $definitions = New-Object System.Collections.Generic.List[object]

    foreach ($comboName in $comboNames) {
        $comboType = 0
        Assert-Success ($SapModel.RespCombo.GetTypeCombo($comboName, [ref]$comboType)) "Get combo type for $comboName"

        $itemCount = 0
        [ETABSv1.eCNameType[]]$itemTypes = @()
        [string[]]$itemNames = @()
        [double[]]$scaleFactors = @()
        Assert-Success ($SapModel.RespCombo.GetCaseList($comboName, [ref]$itemCount, [ref]$itemTypes, [ref]$itemNames, [ref]$scaleFactors)) "Get combo items for $comboName"

        $items = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $itemCount; $i++) {
            $items.Add([pscustomobject]@{
                Type = $itemTypes[$i].ToString()
                Name = $itemNames[$i]
                ScaleFactor = [double]$scaleFactors[$i]
            }) | Out-Null
        }

        $definitions.Add([pscustomobject]@{
            Name = $comboName
            ComboType = $comboType
            Items = @($items)
        }) | Out-Null
    }

    return @($definitions)
}

function Get-LoadCaseNames {
    param(
        $SapModel
    )

    $names = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

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

    return $names
}

function Get-ComboSignature {
    param(
        [object]$Definition
    )

    return @(
        $Definition.Items |
        Sort-Object Type, Name, ScaleFactor |
        ForEach-Object { "{0}|{1}|{2}" -f $_.Type, $_.Name, ("{0:R}" -f $_.ScaleFactor) }
    ) -join ";"
}

$resolvedSourceModelPath = (Resolve-Path -LiteralPath $SourceModelPath).Path
$targetProcess = Get-EtabsProcess -RequestedPid $TargetEtabsPid
if ($null -eq $targetProcess) {
    throw "No running ETABS target process was found."
}

$apiDllPath = Resolve-EtabsApiDll -Process $targetProcess
Add-Type -Path $apiDllPath

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$targetApi = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $targetProcess.Id)
$targetSapModel = $targetApi.SapModel

$targetComboDefinitions = Get-ComboDefinitions -SapModel $targetSapModel
$targetComboMap = @{}
foreach ($definition in $targetComboDefinitions) {
    $targetComboMap[$definition.Name] = $definition
}

$targetLoadCases = Get-LoadCaseNames -SapModel $targetSapModel

$sourceApi = $null
$sourceExited = $false

try {
    $sourceApi = $helper.CreateObject($targetProcess.Path)
    Assert-Success ($sourceApi.ApplicationStart()) "Start source ETABS instance"
    try { $sourceApi.Hide() | Out-Null } catch {}

    $sourceSapModel = $sourceApi.SapModel
    Assert-Success ($sourceSapModel.File.OpenFile($resolvedSourceModelPath)) "Open source model"

    $sourceComboDefinitions = Get-ComboDefinitions -SapModel $sourceSapModel

    $added = New-Object System.Collections.Generic.List[string]
    $skippedExisting = New-Object System.Collections.Generic.List[string]
    $skippedSameDefinition = New-Object System.Collections.Generic.List[string]
    $conflicts = New-Object System.Collections.Generic.List[object]
    $missingDependencies = New-Object System.Collections.Generic.List[object]

    $sourceComboNameSet = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($definition in $sourceComboDefinitions) {
        $sourceComboNameSet.Add($definition.Name) | Out-Null
    }

    foreach ($definition in $sourceComboDefinitions) {
        $missingLoadCasesForCombo = @(
            $definition.Items |
            Where-Object {
                $_.Type -eq "LoadCase" -and -not $targetLoadCases.Contains($_.Name)
            } |
            Select-Object -ExpandProperty Name -Unique
        )

        if ($missingLoadCasesForCombo.Count -gt 0) {
            $missingDependencies.Add([pscustomobject]@{
                Combo = $definition.Name
                MissingLoadCases = $missingLoadCasesForCombo
            }) | Out-Null
        }
    }

    $blockedCombos = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $missingDependencies) {
        $blockedCombos.Add($entry.Combo) | Out-Null
    }

    foreach ($definition in $sourceComboDefinitions) {
        if ($blockedCombos.Contains($definition.Name)) {
            continue
        }

        if ($targetComboMap.ContainsKey($definition.Name)) {
            $targetSignature = Get-ComboSignature -Definition $targetComboMap[$definition.Name]
            $sourceSignature = Get-ComboSignature -Definition $definition

            if ($targetSignature -eq $sourceSignature) {
                $skippedSameDefinition.Add($definition.Name) | Out-Null
                continue
            }

            if (-not $ReplaceExisting) {
                $conflicts.Add([pscustomobject]@{
                    Combo = $definition.Name
                    SourceDefinition = $sourceSignature
                    TargetDefinition = $targetSignature
                }) | Out-Null
                $skippedExisting.Add($definition.Name) | Out-Null
                continue
            }

            Assert-Success ($targetSapModel.RespCombo.Delete($definition.Name)) "Delete existing combo $($definition.Name)"
            $targetComboMap.Remove($definition.Name) | Out-Null
        }
    }

    foreach ($definition in $sourceComboDefinitions) {
        if ($blockedCombos.Contains($definition.Name)) {
            continue
        }

        if (-not $targetComboMap.ContainsKey($definition.Name)) {
            Assert-Success ($targetSapModel.RespCombo.Add($definition.Name, $definition.ComboType)) "Add combo $($definition.Name)"
            $targetComboMap[$definition.Name] = [pscustomobject]@{
                Name = $definition.Name
                ComboType = $definition.ComboType
                Items = @()
            }
        }
    }

    foreach ($definition in $sourceComboDefinitions) {
        if ($blockedCombos.Contains($definition.Name)) {
            continue
        }

        if ($skippedExisting.Contains($definition.Name) -and -not $ReplaceExisting) {
            continue
        }

        foreach ($item in $definition.Items) {
            $itemType = [ETABSv1.eCNameType]::$($item.Type)
            Assert-Success ($targetSapModel.RespCombo.SetCaseList($definition.Name, [ref]$itemType, $item.Name, $item.ScaleFactor)) "Add combo item $($item.Name) to $($definition.Name)"
        }

        if (-not $added.Contains($definition.Name) -and -not $skippedSameDefinition.Contains($definition.Name)) {
            $added.Add($definition.Name) | Out-Null
        }
    }

    $result = [pscustomobject]@{
        TargetProcessId = $targetProcess.Id
        TargetModelPath = [System.IO.Path]::Combine($targetSapModel.GetModelFilepath(), ([System.IO.Path]::GetFileNameWithoutExtension($targetSapModel.GetModelFilename($true)) + ".EDB"))
        SourceModelPath = $resolvedSourceModelPath
        AddedCount = $added.Count
        Added = @($added)
        SkippedSameDefinitionCount = $skippedSameDefinition.Count
        SkippedSameDefinition = @($skippedSameDefinition)
        SkippedExistingConflictCount = $skippedExisting.Count
        SkippedExistingConflict = @($skippedExisting)
        MissingDependencyCount = $missingDependencies.Count
        MissingDependencies = @($missingDependencies)
        ReplaceExisting = [bool]$ReplaceExisting
    }

    if ($AsJson) {
        $result | ConvertTo-Json -Depth 8
    }
    else {
        $result
    }
}
finally {
    if ($null -ne $sourceApi -and -not $sourceExited) {
        try {
            $sourceApi.ApplicationExit($false) | Out-Null
            $sourceExited = $true
        }
        catch {
        }
    }
}
