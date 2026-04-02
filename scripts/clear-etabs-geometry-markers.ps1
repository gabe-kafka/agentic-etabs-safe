param(
    [int]$EtabsPid,
    [string]$GroupPrefix = "DBG_GEOM",
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path -Path $PSScriptRoot -ChildPath "_etabs-geometry-debug.ps1")

$connection = Connect-EtabsSession -EtabsPid $EtabsPid
$sapModel = $connection.SapModel

Remove-EtabsGeometryMarkers -SapModel $sapModel -GroupPrefix $GroupPrefix

$result = [pscustomobject]@{
    ProcessId = $connection.Process.Id
    ModelPath = $connection.ModelPath
    GroupPrefix = $GroupPrefix
    Cleared = $true
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
