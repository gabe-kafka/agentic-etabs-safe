param(
    [string]$InputJson = "C:\Users\gkafka\Desktop\agentic-etabs-safe\out\safe_ground_truth.json",
    [string]$WallsJson = "C:\Users\gkafka\Desktop\agentic-etabs-safe\out\safe_walls.json",
    [string]$OutDxf    = "C:\Users\gkafka\Desktop\agentic-etabs-safe\out\1025-atlantic.dxf"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$gt = Get-Content -Raw -Path $InputJson | ConvertFrom-Json
$wallsData = if (Test-Path $WallsJson) { Get-Content -Raw -Path $WallsJson | ConvertFrom-Json } else { $null }

# Helper: emit a DXF group (code + value) as two lines.
function Code([int]$c, $v) { "$c`n$v" }

$sb = New-Object System.Text.StringBuilder

# ------ HEADER ------
[void]$sb.AppendLine(@"
0
SECTION
2
HEADER
9
`$ACADVER
1
AC1015
9
`$INSUNITS
70
1
0
ENDSEC
"@)

# ------ TABLES (layers) ------
# Define the three layers the webapp expects
[void]$sb.AppendLine(@"
0
SECTION
2
TABLES
0
TABLE
2
LAYER
70
5
0
LAYER
2
0
70
0
62
7
6
CONTINUOUS
0
LAYER
2
SLAB
70
0
62
3
6
CONTINUOUS
0
LAYER
2
COLUMN-REAL
70
0
62
1
6
CONTINUOUS
0
LAYER
2
COLUMN-LABEL
70
0
62
2
6
CONTINUOUS
0
LAYER
2
SHEAR-WALL
70
0
62
5
6
CONTINUOUS
0
ENDTAB
0
ENDSEC
"@)

# ------ ENTITIES ------
[void]$sb.AppendLine("0`nSECTION`n2`nENTITIES")

# 1) Slab outer -> LWPOLYLINE on SLAB layer
function Emit-LWPolyline([string]$layer, $ring) {
    $n = $ring.Count
    [void]$sb.AppendLine("0`nLWPOLYLINE")
    [void]$sb.AppendLine("8`n$layer")
    [void]$sb.AppendLine("100`nAcDbEntity")
    [void]$sb.AppendLine("100`nAcDbPolyline")
    [void]$sb.AppendLine("90`n$n")    # vertex count
    [void]$sb.AppendLine("70`n1")     # closed
    foreach ($v in $ring) {
        [void]$sb.AppendLine("10`n$($v[0])")
        [void]$sb.AppendLine("20`n$($v[1])")
    }
}

Emit-LWPolyline "SLAB" $gt.slab.outer_in
foreach ($h in $gt.slab.holes_in) { Emit-LWPolyline "SLAB" $h }

# 2) Columns -> POINT + TEXT label
foreach ($c in $gt.columns) {
    $x = $c.position_in[0]
    $y = $c.position_in[1]
    # POINT
    [void]$sb.AppendLine("0`nPOINT")
    [void]$sb.AppendLine("8`nCOLUMN-REAL")
    [void]$sb.AppendLine("10`n$x")
    [void]$sb.AppendLine("20`n$y")
    [void]$sb.AppendLine("30`n0.0")

    # TEXT (label near the point - within 36" per ingest logic)
    [void]$sb.AppendLine("0`nTEXT")
    [void]$sb.AppendLine("8`nCOLUMN-LABEL")
    [void]$sb.AppendLine("10`n$($x + 6)")
    [void]$sb.AppendLine("20`n$($y + 6)")
    [void]$sb.AppendLine("30`n0.0")
    [void]$sb.AppendLine("40`n4.0")      # text height 4"
    [void]$sb.AppendLine("1`n$($c.id)")
}

# 3) Walls as LINE entities on SHEAR-WALL layer
$wallCount = 0
if ($wallsData -and $wallsData.walls) {
    foreach ($w in $wallsData.walls) {
        $a = $w.seg[0]; $b = $w.seg[1]
        [void]$sb.AppendLine("0`nLINE")
        [void]$sb.AppendLine("8`nSHEAR-WALL")
        [void]$sb.AppendLine("10`n$($a[0])")
        [void]$sb.AppendLine("20`n$($a[1])")
        [void]$sb.AppendLine("30`n0.0")
        [void]$sb.AppendLine("11`n$($b[0])")
        [void]$sb.AppendLine("21`n$($b[1])")
        [void]$sb.AppendLine("31`n0.0")
        $wallCount++
    }
}

[void]$sb.AppendLine("0`nENDSEC")
[void]$sb.AppendLine("0`nEOF")

$outDir = Split-Path -Parent $OutDxf
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# Write ASCII (DXF libraries expect single-byte encoding)
[System.IO.File]::WriteAllText($OutDxf, $sb.ToString(), [System.Text.Encoding]::ASCII)

Write-Host "DXF written: $OutDxf"
Write-Host ("  Slab   : {0}-pt outer, {1} holes" -f $gt.slab.outer_in.Count, $gt.slab.holes_in.Count)
Write-Host ("  Columns: {0} POINT + {0} TEXT" -f $gt.columns.Count)
Write-Host ("  Walls  : {0} LINE segments on SHEAR-WALL layer" -f $wallCount)
Write-Host ""
Write-Host "Column IDs emitted:"
$gt.columns.id -join ", "
