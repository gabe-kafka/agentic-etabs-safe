param(
    [int]$Port = 8765,
    [string]$Host = "127.0.0.1",
    [switch]$NoBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = (Get-Command python -ErrorAction Stop).Source
$serverScript = Join-Path -Path $PSScriptRoot -ChildPath "serve-etabs-control-center.py"

$arguments = @(
    $serverScript,
    "--host", $Host,
    "--port", $Port
)

if (-not $NoBrowser) {
    $arguments += "--open-browser"
}

Start-Process -FilePath $python -ArgumentList $arguments -WorkingDirectory $repoRoot | Out-Null

[pscustomobject]@{
    Started = $true
    Url = "http://$Host`:$Port/"
    Script = $serverScript
} | Format-List
