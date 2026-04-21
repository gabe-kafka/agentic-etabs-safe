param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
if (-not $proc) { throw "No SAFE running." }
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3)

# Enumerate available tables via overload
Write-Host "=== DatabaseTables methods ==="
$sm.DatabaseTables | Get-Member -MemberType Method | Where-Object { $_.Name -match "^(Get|List)" } | Select-Object -ExpandProperty Name

Write-Host ""
Write-Host "=== Try GetAvailableTables ==="
try {
    $n=0;$tnames=[string[]]::new(0);$tselected=[bool[]]::new(0);$tempFlag=[int[]]::new(0)
    # Attempt common signatures
    $null = $sm.DatabaseTables.GetAvailableTables([ref]$n, [ref]$tnames, [ref]$tempFlag)
    Write-Host "Count: $n"
    $tnames | Where-Object { $_ -match "(?i)(punch|column|support|point|material|slab|property|area|object.*and.*element|load|concrete)" } | Sort-Object | ForEach-Object { Write-Host "  $_" }
} catch { Write-Host "Err: $($_.Exception.Message)" }

Write-Host ""
Write-Host "=== PointObj all methods ==="
($sm.PointObj | Get-Member -MemberType Method | Where-Object { $_.Name -notmatch "^(get_|set_|add_|remove_|GetType|ToString|Equals|GetHashCode)" }).Name | Sort-Object

Write-Host ""
Write-Host "=== SapModel top-level properties ==="
($sm | Get-Member -MemberType Property).Name | Sort-Object
