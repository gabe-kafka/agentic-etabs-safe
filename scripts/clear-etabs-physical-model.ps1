param(
    [int]$EtabsPid,
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

function Get-ModelCounts {
    param(
        $SapModel
    )

    [pscustomobject]@{
        Frames = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.FrameObj -Label "frame objects")
        Areas = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.AreaObj -Label "area objects")
        Links = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.LinkObj -Label "link objects")
        Tendons = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.TendonObj -Label "tendon objects")
        Points = Get-ItemCount (Get-ObjectNames -ObjectApi $SapModel.PointObj -Label "point objects")
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
    return Join-Path -Path $directory -ChildPath ("{0}.pre-clear-physical.{1}{2}" -f $baseName, $timestamp, $extension)
}

function Remove-NamedObjects {
    param(
        $ObjectApi,
        [string[]]$Names,
        [string]$Label
    )

    $deleted = 0
    $failed = @()

    foreach ($name in @($Names)) {
        $ret = $ObjectApi.Delete($name, [ETABSv1.eItemType]::Objects)
        if ($ret -eq 0) {
            $deleted++
        }
        else {
            $failed += [pscustomobject]@{
                Name = $name
                ReturnCode = $ret
            }
        }
    }

    return [pscustomobject]@{
        Label = $Label
        Requested = Get-ItemCount $Names
        Deleted = $deleted
        Failed = Get-ItemCount $failed
        Failures = @($failed)
    }
}

function Get-PointFailureDetail {
    param(
        $SapModel,
        [string]$PointName,
        [int]$SetSpecialPointReturnCode,
        [int]$DeleteReturnCode
    )

    $isSpecial = $false
    try {
        $SapModel.PointObj.GetSpecialPoint($PointName, [ref]$isSpecial) | Out-Null
    }
    catch {
    }

    [bool[]]$restraint = @($false, $false, $false, $false, $false, $false)
    try {
        $SapModel.PointObj.GetRestraint($PointName, [ref]$restraint) | Out-Null
    }
    catch {
    }

    $connectivityCount = 0
    [int[]]$objectTypes = @()
    [string[]]$objectNames = @()
    [int[]]$pointNumbers = @()
    try {
        $SapModel.PointObj.GetConnectivity($PointName, [ref]$connectivityCount, [ref]$objectTypes, [ref]$objectNames, [ref]$pointNumbers) | Out-Null
    }
    catch {
    }

    [pscustomobject]@{
        Name = $PointName
        SetSpecialPointReturnCode = $SetSpecialPointReturnCode
        DeleteReturnCode = $DeleteReturnCode
        IsSpecialPoint = $isSpecial
        ConnectivityCount = $connectivityCount
        HasRestraint = (Get-ItemCount ($restraint | Where-Object { $_ })) -gt 0
        ConnectedObjectNames = @($objectNames)
    }
}

function Remove-AllPoints {
    param(
        $SapModel
    )

    $deleted = 0
    $failed = @()
    $passes = 0

    while ($passes -lt 5) {
        $passes++
        $pointNames = @(Get-ObjectNames -ObjectApi $SapModel.PointObj -Label "point objects")
        if ((Get-ItemCount $pointNames) -eq 0) {
            break
        }

        $deletedThisPass = 0
        foreach ($pointName in $pointNames) {
            $setRet = $SapModel.PointObj.SetSpecialPoint($pointName, $true, [ETABSv1.eItemType]::Objects)
            $deleteRet = $SapModel.PointObj.DeleteSpecialPoint($pointName, [ETABSv1.eItemType]::Objects)

            if ($deleteRet -eq 0) {
                $deleted++
                $deletedThisPass++
            }
            else {
                $failed += Get-PointFailureDetail -SapModel $SapModel -PointName $pointName -SetSpecialPointReturnCode $setRet -DeleteReturnCode $deleteRet
            }
        }

        if ($deletedThisPass -eq 0) {
            break
        }

        $failed = @()
    }

    $remainingPoints = @(Get-ObjectNames -ObjectApi $SapModel.PointObj -Label "point objects")
    if ((Get-ItemCount $remainingPoints) -gt 0) {
        $failed = @()
        foreach ($pointName in $remainingPoints) {
            $setRet = $SapModel.PointObj.SetSpecialPoint($pointName, $true, [ETABSv1.eItemType]::Objects)
            $deleteRet = $SapModel.PointObj.DeleteSpecialPoint($pointName, [ETABSv1.eItemType]::Objects)
            if ($deleteRet -eq 0) {
                $deleted++
            }
            else {
                $failed += Get-PointFailureDetail -SapModel $SapModel -PointName $pointName -SetSpecialPointReturnCode $setRet -DeleteReturnCode $deleteRet
            }
        }
    }

    return [pscustomobject]@{
        Requested = $deleted + (Get-ItemCount $failed)
        Deleted = $deleted
        Failed = Get-ItemCount $failed
        Passes = $passes
        Failures = @($failed)
    }
}

$connection = Connect-EtabsSession -EtabsPid $EtabsPid
$sapModel = $connection.SapModel
$modelWasLocked = $sapModel.GetModelIsLocked()
$unlockedForDelete = $false

if ($modelWasLocked) {
    if (-not $UnlockIfLocked) {
        throw "Deleting ETABS objects requires an unlocked model. Re-run with -UnlockIfLocked if you want the script to unlock the model first."
    }

    Assert-Success ($sapModel.SetModelIsLocked($false)) "SetModelIsLocked(false)"
    $unlockedForDelete = $true
}

$preCounts = Get-ModelCounts -SapModel $sapModel

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

$linkDelete = Remove-NamedObjects -ObjectApi $sapModel.LinkObj -Names (Get-ObjectNames -ObjectApi $sapModel.LinkObj -Label "link objects") -Label "link objects"
$frameDelete = Remove-NamedObjects -ObjectApi $sapModel.FrameObj -Names (Get-ObjectNames -ObjectApi $sapModel.FrameObj -Label "frame objects") -Label "frame objects"
$areaDelete = Remove-NamedObjects -ObjectApi $sapModel.AreaObj -Names (Get-ObjectNames -ObjectApi $sapModel.AreaObj -Label "area objects") -Label "area objects"
$pointDelete = Remove-AllPoints -SapModel $sapModel

$saveReturnCode = $null
if ($Save -and -not [string]::IsNullOrWhiteSpace($connection.ModelPath)) {
    $saveReturnCode = $sapModel.File.Save($connection.ModelPath)
    Assert-Success $saveReturnCode "Save model after clearing physical objects"
}

$postCounts = Get-ModelCounts -SapModel $sapModel
$warnings = @()
if ($preCounts.Tendons -gt 0 -or $postCounts.Tendons -gt 0) {
    $warnings += "Tendon objects are present, but this script does not delete them because the ETABS 20 tendon API surface in this workspace does not expose a direct delete method."
}

$result = [pscustomobject]@{
    ProcessId = $connection.Process.Id
    ProcessVersion = $connection.Process.MainModule.FileVersionInfo.FileVersion
    ApiDllPath = $connection.ApiDllPath
    ModelPath = $connection.ModelPath
    ModelWasLocked = $modelWasLocked
    UnlockedForDelete = $unlockedForDelete
    SaveRequested = [bool]$Save
    SaveReturnCode = $saveReturnCode
    BackupRequested = [bool](-not $SkipBackup)
    BackupCreated = $backupCreated
    BackupPath = $backupPath
    BackupError = $backupError
    PreCounts = $preCounts
    Deletes = [pscustomobject]@{
        Links = $linkDelete
        Frames = $frameDelete
        Areas = $areaDelete
        Points = $pointDelete
    }
    PostCounts = $postCounts
    Warnings = @($warnings)
    Completed = ($postCounts.Frames -eq 0 -and $postCounts.Areas -eq 0 -and $postCounts.Links -eq 0 -and $postCounts.Points -eq 0 -and $postCounts.Tendons -eq 0)
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
}
else {
    $result
}
