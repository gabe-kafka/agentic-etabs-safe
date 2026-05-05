param(
    [int]$EtabsPid,

    # Near-coincident tolerance in model length units. If omitted, auto-selected
    # based on model units (0.5 in, 0.04 ft, 10 mm, or 0.01 m).
    [double]$Tolerance = -1,

    # Z-direction tolerance for off-story check. Defaults to same as $Tolerance.
    [double]$ZTolerance = -1,

    [switch]$SkipOffStory,
    [switch]$SkipFrameGaps,

    # Write near-coincident pairs to this CSV path.
    [string]$CsvOut,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Connect
# ---------------------------------------------------------------------------

function Resolve-EtabsApiDll {
    param([System.Diagnostics.Process]$Process)

    $candidates = [System.Collections.Generic.List[string]]::new()

    if ($null -ne $Process) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($Process.Path)) {
                $candidates.Add((Join-Path (Split-Path -Parent $Process.Path) "ETABSv1.dll"))
            }
        } catch {}
    }

    foreach ($p in @(
        "C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 20\ETABSv1.dll"
    )) { $candidates.Add($p) }

    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path }
    }
    throw "ETABSv1.dll not found."
}

function Get-EtabsProcess {
    param([int]$RequestedPid)
    $procs = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) { return $procs | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1 }
    $windowed = $procs | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($windowed) { return $windowed | Select-Object -First 1 }
    return $procs | Select-Object -First 1
}

# ---------------------------------------------------------------------------
# Unit helpers
# ---------------------------------------------------------------------------

# eUnits enum values from the ETABS API:
#   1=lb-in  2=lb-ft  3=kip-in  4=kip-ft  5=kN-mm  6=kN-cm  7=kN-m  8=N-mm  9=N-cm  10=N-m  11=tf-m  12=tf-cm  13=tf-mm  14=kgf-m  15=kgf-cm  16=kgf-mm
function Get-DefaultTolerance {
    param([int]$Units)

    switch ($Units) {
        1  { return 0.5 }   # lb-in   -> 0.5 in
        2  { return 0.04 }  # lb-ft   -> ~0.5 in in feet
        3  { return 0.5 }   # kip-in  -> 0.5 in
        4  { return 0.04 }  # kip-ft
        5  { return 10.0 }  # kN-mm   -> 10 mm
        6  { return 1.0 }   # kN-cm   -> 1 cm
        7  { return 0.01 }  # kN-m    -> 0.01 m
        8  { return 10.0 }  # N-mm
        9  { return 1.0 }   # N-cm
        10 { return 0.01 }  # N-m
        11 { return 0.01 }  # tf-m
        12 { return 1.0 }   # tf-cm
        13 { return 10.0 }  # tf-mm
        14 { return 0.01 }  # kgf-m
        15 { return 1.0 }   # kgf-cm
        16 { return 10.0 }  # kgf-mm
        default { return 0.5 }
    }
}

# ---------------------------------------------------------------------------
# Near-coincident joint detection (spatial bucket approach, O(n log n))
# ---------------------------------------------------------------------------

function Find-NearCoincidentJoints {
    param(
        [hashtable]$JointCoords,   # name -> @(X,Y,Z)
        [double]$Tol
    )

    # Build a spatial hash: bucket key -> list of joint names.
    # Each joint is placed in the bucket corresponding to floor(coord/Tol).
    # Then for each joint, check its own bucket and the 26 neighboring buckets.

    $buckets = @{}

    foreach ($name in $JointCoords.Keys) {
        $c = $JointCoords[$name]
        $bx = [long][Math]::Floor($c[0] / $Tol)
        $by = [long][Math]::Floor($c[1] / $Tol)
        $bz = [long][Math]::Floor($c[2] / $Tol)
        $key = "$bx|$by|$bz"
        if (-not $buckets.ContainsKey($key)) { $buckets[$key] = [System.Collections.Generic.List[string]]::new() }
        $buckets[$key].Add($name)
    }

    $pairs = [System.Collections.Generic.List[pscustomobject]]::new()
    $seen  = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($name in $JointCoords.Keys) {
        $c  = $JointCoords[$name]
        $bx = [long][Math]::Floor($c[0] / $Tol)
        $by = [long][Math]::Floor($c[1] / $Tol)
        $bz = [long][Math]::Floor($c[2] / $Tol)

        for ($dx = -1; $dx -le 1; $dx++) {
            for ($dy = -1; $dy -le 1; $dy++) {
                for ($dz = -1; $dz -le 1; $dz++) {
                    $key = "$($bx+$dx)|$($by+$dy)|$($bz+$dz)"
                    if (-not $buckets.ContainsKey($key)) { continue }

                    foreach ($other in $buckets[$key]) {
                        if ($other -eq $name) { continue }

                        # Canonical pair key to avoid duplicates
                        $pairKey = if ([string]::CompareOrdinal($name, $other) -lt 0) { "$name|$other" } else { "$other|$name" }
                        if (-not $seen.Add($pairKey)) { continue }

                        $oc = $JointCoords[$other]
                        $dx3 = $c[0] - $oc[0]
                        $dy3 = $c[1] - $oc[1]
                        $dz3 = $c[2] - $oc[2]
                        $dist = [Math]::Sqrt($dx3*$dx3 + $dy3*$dy3 + $dz3*$dz3)

                        if ($dist -le $Tol) {
                            $pairs.Add([pscustomobject]@{
                                JointA   = $name
                                JointB   = $other
                                Distance = [Math]::Round($dist, 6)
                                AX = [Math]::Round($c[0],  4)
                                AY = [Math]::Round($c[1],  4)
                                AZ = [Math]::Round($c[2],  4)
                                BX = [Math]::Round($oc[0], 4)
                                BY = [Math]::Round($oc[1], 4)
                                BZ = [Math]::Round($oc[2], 4)
                                DeltaX = [Math]::Round([Math]::Abs($dx3), 6)
                                DeltaY = [Math]::Round([Math]::Abs($dy3), 6)
                                DeltaZ = [Math]::Round([Math]::Abs($dz3), 6)
                            })
                        }
                    }
                }
            }
        }
    }

    return @($pairs | Sort-Object Distance)
}

# ---------------------------------------------------------------------------
# Off-story joint detection
# ---------------------------------------------------------------------------

function Get-StoryElevations {
    param($SapModel)

    $baseElevation = 0.0
    $numStories   = 0
    [string[]]$storyNames = @()
    [double[]]$storyElevs = @()
    [double[]]$storyHeights = @()
    [bool[]]$isMasterStory = @()
    [string[]]$similarTo = @()
    [bool[]]$spliceAbove = @()
    [double[]]$spliceHeight = @()
    [int[]]$color = @()

    $ret = $SapModel.Story.GetStories_2(
        [ref]$baseElevation,
        [ref]$numStories,
        [ref]$storyNames,
        [ref]$storyElevs,
        [ref]$storyHeights,
        [ref]$isMasterStory,
        [ref]$similarTo,
        [ref]$spliceAbove,
        [ref]$spliceHeight,
        [ref]$color)

    if ($ret -ne 0) { return @() }

    $stories = @()
    $stories += [pscustomobject]@{
        Name      = "Base"
        Elevation = $baseElevation
    }

    for ($i = 0; $i -lt $numStories; $i++) {
        $stories += [pscustomobject]@{
            Name      = $storyNames[$i]
            Elevation = $storyElevs[$i]
        }
    }
    return $stories
}

function Find-OffStoryJoints {
    param(
        [hashtable]$JointCoords,
        [pscustomobject[]]$Stories,
        [double]$ZTol
    )

    if (@($Stories).Count -eq 0) { return @() }

    $storyElevs = @($Stories | ForEach-Object { $_.Elevation })

    $offStory = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($name in $JointCoords.Keys) {
        $z = $JointCoords[$name][2]
        $minDelta = ($storyElevs | ForEach-Object { [Math]::Abs($_ - $z) } | Measure-Object -Minimum).Minimum
        $nearestStory = $Stories | Sort-Object { [Math]::Abs($_.Elevation - $z) } | Select-Object -First 1

        if ($minDelta -gt $ZTol) {
            $offStory.Add([pscustomobject]@{
                Joint        = $name
                Z            = [Math]::Round($z, 4)
                NearestStory = $nearestStory.Name
                NearestElev  = [Math]::Round($nearestStory.Elevation, 4)
                Delta        = [Math]::Round($minDelta, 6)
                X            = [Math]::Round($JointCoords[$name][0], 4)
                Y            = [Math]::Round($JointCoords[$name][1], 4)
            })
        }
    }

    return @($offStory | Sort-Object Delta -Descending)
}

# ---------------------------------------------------------------------------
# Frame endpoint gap detection
# Frame endpoints are defined by joint names, so a "gap" here means the
# frame's assigned joint exists but its coordinates differ from the frame
# object's connectivity point reported by GetCoordCartesian on the joint.
# More practically: check if a frame's I or J joint is suspiciously far from
# every area corner (slab/wall) within the same bounding region.
# The simpler and more useful check: find frames whose I/J joint has only
# that one frame connected to it AND no restraint - a dangling end.
# ---------------------------------------------------------------------------

function Find-DanglingFrameEnds {
    param($SapModel, [hashtable]$JointCoords)

    $numNames = 0
    [string[]]$frameNames = @()
    $SapModel.FrameObj.GetNameList([ref]$numNames, [ref]$frameNames) | Out-Null

    # Build connectivity map: joint -> count of objects
    $connectivity = @{}
    foreach ($name in $JointCoords.Keys) {
        $connCount = 0
        [int[]]$objTypes  = @()
        [string[]]$objNames = @()
        [int[]]$ptNums    = @()
        $SapModel.PointObj.GetConnectivity($name, [ref]$connCount, [ref]$objTypes, [ref]$objNames, [ref]$ptNums) | Out-Null
        $connectivity[$name] = $connCount
    }

    $danglers = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($frameName in $frameNames) {
        $ptI = ""; $ptJ = ""
        $SapModel.FrameObj.GetPoints($frameName, [ref]$ptI, [ref]$ptJ) | Out-Null

        foreach ($endLabel in @("I", "J")) {
            $pt = if ($endLabel -eq "I") { $ptI } else { $ptJ }

            if (-not $connectivity.ContainsKey($pt)) { continue }
            if ($connectivity[$pt] -gt 1) { continue }  # connected to something else - fine

            # Only one connection (this frame). Check for restraint.
            [bool[]]$restraint = @($false,$false,$false,$false,$false,$false)
            $SapModel.PointObj.GetRestraint($pt, [ref]$restraint) | Out-Null
            if ($restraint -contains $true) { continue }  # it's a support - fine

            $c = $JointCoords[$pt]
            $label = ""; $story = ""
            $SapModel.FrameObj.GetLabelFromName($frameName, [ref]$label, [ref]$story) | Out-Null

            $section = ""; $auto = ""
            $SapModel.FrameObj.GetSection($frameName, [ref]$section, [ref]$auto) | Out-Null

            $danglers.Add([pscustomobject]@{
                Frame   = $frameName
                Label   = $label
                Story   = $story
                Section = $section
                End     = $endLabel
                Joint   = $pt
                X       = [Math]::Round($c[0], 4)
                Y       = [Math]::Round($c[1], 4)
                Z       = [Math]::Round($c[2], 4)
            })
        }
    }

    return @($danglers | Sort-Object Story, Label)
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$process = Get-EtabsProcess -RequestedPid $EtabsPid
if ($null -eq $process) { throw "No running ETABS process found." }

$apiDllPath = Resolve-EtabsApiDll -Process $process
Add-Type -Path $apiDllPath

$helper   = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api      = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$sapModel = $api.SapModel

# Detect units and set tolerances
$presentUnits = [int]($sapModel.GetPresentUnits())
$autoTol = Get-DefaultTolerance -Units $presentUnits

if ($Tolerance  -lt 0) { $Tolerance  = $autoTol }
if ($ZTolerance -lt 0) { $ZTolerance = $Tolerance }

if (-not $AsJson) {
    Write-Host "Model units enum : $presentUnits"
    Write-Host "XY/Z tolerance   : $Tolerance"
    Write-Host ""
}

# Load all joint coordinates into a hashtable
$numJoints = 0
[string[]]$jointNames = @()
$sapModel.PointObj.GetNameList([ref]$numJoints, [ref]$jointNames) | Out-Null

if (-not $AsJson) { Write-Host "Loading $numJoints joint coordinates..." }
$jointCoords = @{}
foreach ($j in $jointNames) {
    $x = 0.0; $y = 0.0; $z = 0.0
    $sapModel.PointObj.GetCoordCartesian($j, [ref]$x, [ref]$y, [ref]$z, "Global") | Out-Null
    $jointCoords[$j] = @($x, $y, $z)
}

# --- Check 1: Near-coincident joints ---
if (-not $AsJson) { Write-Host "Checking near-coincident joints..." }
$coincidentPairs = Find-NearCoincidentJoints -JointCoords $jointCoords -Tol $Tolerance

# --- Check 2: Off-story joints ---
$offStoryJoints = @()
$stories = @()
if (-not $SkipOffStory) {
    if (-not $AsJson) { Write-Host "Loading story elevations..." }
    $stories = Get-StoryElevations -SapModel $sapModel
    if (@($stories).Count -gt 0) {
        if (-not $AsJson) { Write-Host "Checking off-story joints ($(@($stories).Count) stories)..." }
        $offStoryJoints = Find-OffStoryJoints -JointCoords $jointCoords -Stories $stories -ZTol $ZTolerance
    } else {
        if (-not $AsJson) { Write-Host "No stories found - skipping off-story check." }
    }
}

# --- Check 3: Dangling frame ends ---
$danglingEnds = @()
if (-not $SkipFrameGaps) {
    if (-not $AsJson) { Write-Host "Checking for dangling frame ends..." }
    $danglingEnds = Find-DanglingFrameEnds -SapModel $sapModel -JointCoords $jointCoords
}

if (-not $AsJson) { Write-Host "" }

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

$result = [pscustomobject]@{
    ProcessId            = $process.Id
    ModelPath            = [System.IO.Path]::Combine($sapModel.GetModelFilepath(), ([System.IO.Path]::GetFileNameWithoutExtension($sapModel.GetModelFilename($true)) + ".EDB"))
    PresentUnitsEnum     = $presentUnits
    Tolerance            = $Tolerance
    ZTolerance           = $ZTolerance
    JointCount           = $numJoints
    StoryCount           = @($stories).Count

    NearCoincidentCount  = @($coincidentPairs).Count
    NearCoincidentPairs  = $coincidentPairs

    OffStoryCount        = @($offStoryJoints).Count
    OffStoryJoints       = $offStoryJoints

    DanglingEndCount     = @($danglingEnds).Count
    DanglingEnds         = $danglingEnds
}

# Console summary
if (-not $AsJson) {
    Write-Host "========================================="
    Write-Host " GEOMETRIC BUG SUMMARY"
    Write-Host "========================================="
    Write-Host ("  Near-coincident joint pairs : {0}" -f $result.NearCoincidentCount)
    Write-Host ("  Off-story joints            : {0}" -f $result.OffStoryCount)
    Write-Host ("  Dangling frame ends         : {0}" -f $result.DanglingEndCount)
    Write-Host ""

    if (@($coincidentPairs).Count -gt 0) {
        Write-Host "--- Near-Coincident Pairs (top 20) ---"
        $coincidentPairs | Select-Object -First 20 | Format-Table JointA, JointB, Distance, DeltaX, DeltaY, DeltaZ, AX, AY, AZ -AutoSize
    }

    if (@($offStoryJoints).Count -gt 0) {
        Write-Host "--- Off-Story Joints (top 20) ---"
        $offStoryJoints | Select-Object -First 20 | Format-Table Joint, X, Y, Z, NearestStory, NearestElev, Delta -AutoSize
    }

    if (@($danglingEnds).Count -gt 0) {
        Write-Host "--- Dangling Frame Ends (top 20) ---"
        $danglingEnds | Select-Object -First 20 | Format-Table Frame, Label, Story, Section, End, Joint, X, Y, Z -AutoSize
    }
}

# Optional CSV export of near-coincident pairs
if (-not [string]::IsNullOrWhiteSpace($CsvOut) -and @($coincidentPairs).Count -gt 0) {
    $coincidentPairs | Export-Csv -LiteralPath $CsvOut -NoTypeInformation
    if (-not $AsJson) { Write-Host "Near-coincident pairs written to: $CsvOut" }
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 10
} else {
    $result
}
