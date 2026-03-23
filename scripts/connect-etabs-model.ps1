param(
    [Parameter(Position = 0)]
    [string]$ModelPath,

    [int]$EtabsPid,

    [switch]$OpenIfDifferent,

    [switch]$LaunchIfNeeded,

    [string]$EtabsExePath,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-EtabsApiDll {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$PreferredExePath
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

    if (-not [string]::IsNullOrWhiteSpace($PreferredExePath)) {
        $resolvedExe = (Resolve-Path -LiteralPath $PreferredExePath).Path
        $candidates.Add((Join-Path -Path (Split-Path -Parent $resolvedExe) -ChildPath "ETABSv1.dll"))
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

function Resolve-ModelPath {
    param(
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if (Test-Path -LiteralPath $PathValue -PathType Container) {
        $edb = Get-ChildItem -LiteralPath $PathValue -Filter *.EDB -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -eq $edb) {
            throw "No .EDB file was found under '$PathValue'."
        }

        return $edb.FullName
    }

    if (Test-Path -LiteralPath $PathValue -PathType Leaf) {
        return (Resolve-Path -LiteralPath $PathValue).Path
    }

    throw "Model path '$PathValue' does not exist."
}

function Get-EtabsProcesses {
    Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
}

function Get-PreferredProcess {
    param(
        [int]$RequestedPid
    )

    $processes = Get-EtabsProcesses
    if ($RequestedPid) {
        return $processes | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }

    $windowed = $processes | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($windowed) {
        return $windowed | Select-Object -First 1
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
        $fvi = $Process.MainModule.FileVersionInfo
        return "$($fvi.FileMajorPart).$($fvi.FileMinorPart).$($fvi.FileBuildPart)"
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

    $extension = [System.IO.Path]::GetExtension($RawModelPath)
    if ($extension -ieq ".EDB") {
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

function Connect-Or-LaunchEtabs {
    param(
        [ETABSv1.cHelper]$Helper,
        [System.Diagnostics.Process]$Process,
        [switch]$Launch,
        [string]$ExePath
    )

    if ($null -ne $Process) {
        return @{
            Api = $Helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $Process.Id)
            Action = "attached"
        }
    }

    if (-not $Launch) {
        throw "No running ETABS process was found."
    }

    $api = if ([string]::IsNullOrWhiteSpace($ExePath)) {
        $Helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
    }
    else {
        $resolvedExe = (Resolve-Path -LiteralPath $ExePath).Path
        $Helper.CreateObject($resolvedExe)
    }

    $startResult = $api.ApplicationStart()
    if ($startResult -ne 0) {
        throw "ApplicationStart returned $startResult."
    }

    Start-Sleep -Seconds 2

    return @{
        Api = $api
        Action = "launched"
    }
}

$resolvedModelPath = Resolve-ModelPath -PathValue $ModelPath
$process = Get-PreferredProcess -RequestedPid $EtabsPid
$apiDllPath = Resolve-EtabsApiDll -Process $process -PreferredExePath $EtabsExePath
Add-Type -Path $apiDllPath

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$connection = Connect-Or-LaunchEtabs -Helper $helper -Process $process -Launch:$LaunchIfNeeded -ExePath $EtabsExePath
$api = $connection.Api

$currentProcess = if ($null -ne $process) { $process } else { Get-PreferredProcess -RequestedPid 0 }
$apiVersion = $helper.GetOAPIVersionNumber()
$rawModelPath = $api.SapModel.GetModelFilename($true)
$modelDirectory = $api.SapModel.GetModelFilepath()
$currentModelPath = Resolve-CanonicalModelPath -RawModelPath $rawModelPath -ModelDirectory $modelDirectory
$pathMatch = $false
$switchResult = $null

if ($resolvedModelPath) {
    $pathMatch = [string]::Equals($currentModelPath, $resolvedModelPath, [System.StringComparison]::OrdinalIgnoreCase)

    if (-not $pathMatch -and $OpenIfDifferent) {
        $switchResult = $api.SapModel.File.OpenFile($resolvedModelPath)
        if ($switchResult -ne 0) {
            throw "OpenFile returned $switchResult for '$resolvedModelPath'."
        }

        Start-Sleep -Seconds 1
        $rawModelPath = $api.SapModel.GetModelFilename($true)
        $modelDirectory = $api.SapModel.GetModelFilepath()
        $currentModelPath = Resolve-CanonicalModelPath -RawModelPath $rawModelPath -ModelDirectory $modelDirectory
        $pathMatch = [string]::Equals($currentModelPath, $resolvedModelPath, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

$result = [pscustomobject]@{
    Action = if ($switchResult -ne $null) { "opened-model" } else { $connection.Action }
    ProcessId = if ($currentProcess) { $currentProcess.Id } else { $null }
    ProcessVersion = Get-ProcessVersion -Process $currentProcess
    ApiVersion = $apiVersion
    MainWindowTitle = if ($currentProcess) { $currentProcess.MainWindowTitle } else { $null }
    ActiveModelPath = $currentModelPath
    ActiveModelPathRaw = $rawModelPath
    ExpectedModelPath = $resolvedModelPath
    PathMatch = if ($resolvedModelPath) { $pathMatch } else { $null }
    ApiDllPath = $apiDllPath
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
