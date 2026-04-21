param(
    [int]$SafePid,
    [string]$OutJson = "C:\Users\gkafka\Desktop\agentic-etabs-safe\out\safe_walls.json"
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3)

# Determine slab Z (use area 1 point 1 as reference)
$x=0.0;$y=0.0;$zSlab=0.0
$null = $sm.PointObj.GetCoordCartesian("1", [ref]$x, [ref]$y, [ref]$zSlab, "Global")

$na=0;$anames=[string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)

$walls = @()

for ($i=0; $i -lt $na; $i++) {
    $a = $anames[$i]
    $prop=""
    try { $null = $sm.AreaObj.GetProperty($a, [ref]$prop) } catch { continue }
    if ($prop -notmatch "^Wall") { continue }

    $nPts=0; $ptNames=[string[]]::new(0)
    $null = $sm.AreaObj.GetPoints($a, [ref]$nPts, [ref]$ptNames)

    # Pull all 4 points
    $coords = @()
    foreach ($p in $ptNames) {
        $x=0.0;$y=0.0;$z=0.0
        $null = $sm.PointObj.GetCoordCartesian($p, [ref]$x, [ref]$y, [ref]$z, "Global")
        $coords += ,@{ pt=$p; x=[double]$x; y=[double]$y; z=[double]$z }
    }

    # Wall panel: 4 corners. Take the 2 points at slab Z (these define the 2D trace).
    $atSlab = $coords | Where-Object { [Math]::Abs($_.z - $zSlab) -lt 0.5 }
    if ($atSlab.Count -lt 2) {
        # Fallback: take the 2 points with highest Z (top of wall)
        $atSlab = $coords | Sort-Object -Property { -$_.z } | Select-Object -First 2
    }
    # Usually the wall has 2 top + 2 bottom. Take the 2 top (at slab) → 2D line segment.
    $seg = @()
    foreach ($p in $atSlab | Select-Object -First 2) {
        $seg += ,@([double]$p.x, [double]$p.y)
    }

    $walls += [ordered]@{
        id     = "W$a"
        safe_area = $a
        prop   = $prop
        seg    = $seg
    }
}

$out = [ordered]@{
    extracted_at = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK")
    slab_z_in    = $zSlab
    walls        = $walls
}

$outDir = Split-Path -Parent $OutJson
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$out | ConvertTo-Json -Depth 8 | Set-Content -Path $OutJson -Encoding UTF8

Write-Host ("Extracted {0} walls -> {1}" -f $walls.Count, $OutJson)
$walls | Select-Object -First 6 | ForEach-Object {
    Write-Host ("  {0} ({1}): seg=({2:F0},{3:F0}) -> ({4:F0},{5:F0})" -f $_.id, $_.prop, $_.seg[0][0], $_.seg[0][1], $_.seg[1][0], $_.seg[1][1])
}
