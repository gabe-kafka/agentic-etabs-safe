param(
    [int]$SafePid,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-SafeApiDll {
    param([System.Diagnostics.Process]$Process)

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Process) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($Process.Path)) {
                $candidates.Add((Join-Path -Path (Split-Path -Parent $Process.Path) -ChildPath "SAFEv1.dll"))
            }
        } catch { }
    }
    @(
        "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll",
        "C:\Program Files\Computers and Structures\SAFE 20\SAFEv1.dll",
        "C:\Program Files (x86)\Computers and Structures\SAFE 21\SAFEv1.dll",
        "C:\Program Files (x86)\Computers and Structures\SAFE 20\SAFEv1.dll"
    ) | ForEach-Object { $candidates.Add($_) }

    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path }
    }
    throw "SAFEv1.dll not found."
}

function Get-PreferredProcess {
    param([int]$RequestedPid)
    $procs = Get-Process SAFE -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) {
        return $procs | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }
    $windowed = $procs | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($windowed) { return $windowed | Select-Object -First 1 }
    return $procs | Select-Object -First 1
}

$process = Get-PreferredProcess -RequestedPid $SafePid
if ($null -eq $process) { throw "No running SAFE process found." }

$apiDllPath = Resolve-SafeApiDll -Process $process
Add-Type -Path $apiDllPath

$helper = [SAFEv1.cHelper](New-Object SAFEv1.Helper)
$api = $helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", $process.Id)
$sap = $api.SapModel

$typeMap = @{
    1 = "Point"; 2 = "Frame"; 3 = "Cable"; 4 = "Tendon";
    5 = "Area";  6 = "Solid"; 7 = "Link"
}

$num = 0
$objType = [int[]]@()
$objName = [string[]]@()
$null = $sap.SelectObj.GetSelected([ref]$num, [ref]$objType, [ref]$objName)

$frames = New-Object System.Collections.Generic.List[object]
$points = New-Object System.Collections.Generic.List[object]
$areas  = New-Object System.Collections.Generic.List[object]
$others = New-Object System.Collections.Generic.List[object]

for ($i = 0; $i -lt $num; $i++) {
    $t = [int]$objType[$i]
    $n = $objName[$i]

    switch ($t) {
        2 {
            $p1 = ""; $p2 = ""
            $null = $sap.FrameObj.GetPoints($n, [ref]$p1, [ref]$p2)

            $x1 = 0.0; $y1 = 0.0; $z1 = 0.0
            $x2 = 0.0; $y2 = 0.0; $z2 = 0.0
            $null = $sap.PointObj.GetCoordCartesian($p1, [ref]$x1, [ref]$y1, [ref]$z1, "Global")
            $null = $sap.PointObj.GetCoordCartesian($p2, [ref]$x2, [ref]$y2, [ref]$z2, "Global")
            $len = [Math]::Sqrt([Math]::Pow($x2-$x1,2) + [Math]::Pow($y2-$y1,2) + [Math]::Pow($z2-$z1,2))

            $label = ""; $story = ""
            try { $null = $sap.FrameObj.GetLabelFromName($n, [ref]$label, [ref]$story) } catch {}

            $sect = ""; $sAuto = ""
            try { $null = $sap.FrameObj.GetSection($n, [ref]$sect, [ref]$sAuto) } catch {}

            $frames.Add([pscustomobject]@{
                Name    = $n
                Label   = $label
                Story   = $story
                Section = $sect
                P1Name  = $p1
                P2Name  = $p2
                P1      = @($x1, $y1, $z1)
                P2      = @($x2, $y2, $z2)
                Length  = $len
            })
        }
        1 {
            $x = 0.0; $y = 0.0; $z = 0.0
            $null = $sap.PointObj.GetCoordCartesian($n, [ref]$x, [ref]$y, [ref]$z, "Global")

            $label = ""; $story = ""
            try { $null = $sap.PointObj.GetLabelFromName($n, [ref]$label, [ref]$story) } catch {}

            $points.Add([pscustomobject]@{
                Name  = $n
                Label = $label
                Story = $story
                Coord = @($x, $y, $z)
            })
        }
        5 {
            $numPts = 0
            $ptNames = [string[]]@()
            try { $null = $sap.AreaObj.GetPoints($n, [ref]$numPts, [ref]$ptNames) } catch {}

            $label = ""; $story = ""
            try { $null = $sap.AreaObj.GetLabelFromName($n, [ref]$label, [ref]$story) } catch {}

            $areas.Add([pscustomobject]@{
                Name      = $n
                Label     = $label
                Story     = $story
                NumPoints = $numPts
                Points    = $ptNames
            })
        }
        default {
            $others.Add([pscustomobject]@{
                Name = $n
                Type = if ($typeMap.ContainsKey($t)) { $typeMap[$t] } else { "Type$t" }
            })
        }
    }
}

$units = $sap.GetPresentUnits()

$result = [pscustomobject]@{
    ProcessId       = $process.Id
    MainWindowTitle = $process.MainWindowTitle
    Units           = $units
    Counts          = @{
        Total  = $num
        Frames = $frames.Count
        Points = $points.Count
        Areas  = $areas.Count
        Other  = $others.Count
    }
    Frames = $frames
    Points = $points
    Areas  = $areas
    Other  = $others
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 8
} else {
    Write-Output ("SAFE PID {0}  |  Units code {1}" -f $process.Id, $units)
    Write-Output ("Total selected: {0}  (Frames={1} Points={2} Areas={3} Other={4})" -f `
        $num, $frames.Count, $points.Count, $areas.Count, $others.Count)
    Write-Output ""

    if ($frames.Count -gt 0) {
        Write-Output "Frames:"
        foreach ($f in $frames) {
            Write-Output ("  [{0}] label={1} story={2} section='{3}' len={4:F3}" -f `
                $f.Name, $f.Label, $f.Story, $f.Section, $f.Length)
            Write-Output ("       {0}({1:F3},{2:F3},{3:F3}) -> {4}({5:F3},{6:F3},{7:F3})" -f `
                $f.P1Name, $f.P1[0], $f.P1[1], $f.P1[2], $f.P2Name, $f.P2[0], $f.P2[1], $f.P2[2])
        }
    }
    if ($points.Count -gt 0) {
        Write-Output "Points:"
        foreach ($p in $points) {
            Write-Output ("  [{0}] label={1} story={2}  ({3:F3},{4:F3},{5:F3})" -f `
                $p.Name, $p.Label, $p.Story, $p.Coord[0], $p.Coord[1], $p.Coord[2])
        }
    }
    if ($areas.Count -gt 0) {
        Write-Output "Areas:"
        foreach ($a in $areas) {
            Write-Output ("  [{0}] label={1} story={2} numPts={3}" -f $a.Name, $a.Label, $a.Story, $a.NumPoints)
        }
    }
    if ($others.Count -gt 0) {
        Write-Output "Other:"
        foreach ($o in $others) {
            Write-Output ("  [{0}] {1}" -f $o.Name, $o.Type)
        }
    }
}
