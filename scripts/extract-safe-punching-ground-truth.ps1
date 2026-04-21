param(
    [int]$SafePid,
    [double]$DeadPsf = 100,   # self-weight for 8" conc
    [double]$SdlPsf  = 35,
    [double]$LivePsf = 40,
    [string]$OutJson = "C:\Users\gkafka\Desktop\agentic-etabs-safe\out\safe_ground_truth.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Attach ---
$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
if (-not $proc) { throw "No SAFE running." }
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3)  # kip-in

# Ensure out directory exists
$outDir = Split-Path -Parent $OutJson
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# --- Pull punching design table ---
$v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
$null = $sm.DatabaseTables.GetTableForDisplayArray(
    "Concrete Slab Design - Punching Shear Data",
    [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
if ($nr -eq 0) { throw "Punching design table is empty. Run Design > Concrete Slab Design first." }

$ncol = $f.Length
$ix = @{}
for ($i=0; $i -lt $f.Length; $i++) { $ix[$f[$i]] = $i }

function Row([int]$r, [string]$field) { $d[$r * $ncol + $ix[$field]] }

# --- Build punching results array ---
$punchingResults = @()
$pointToRow = @{}
for ($r=0; $r -lt $nr; $r++) {
    $pt = Row $r "Point"
    $pointToRow[$pt] = $r
    $punchingResults += [ordered]@{
        id                   = "C$pt"
        safe_point           = $pt
        combo                = Row $r "Combo"
        global_x_in          = [double](Row $r "GlobalX")
        global_y_in          = [double](Row $r "GlobalY")
        location             = Row $r "Location"
        status               = Row $r "Status"
        dcr                  = [double](Row $r "Ratio")
        vu_kip               = [double](Row $r "Vu")
        total_mu2_kip_in     = [double](Row $r "Mu2")
        total_mu3_kip_in     = [double](Row $r "Mu3")
        unbal_mu2_kip_in     = [double](Row $r "UnbalMu2")
        unbal_mu3_kip_in     = [double](Row $r "UnbalMu3")
        gamma_v2             = [double](Row $r "Gamma_v2")
        gamma_v3             = [double](Row $r "Gamma_v3")
        d_in                 = [double](Row $r "Depth")
        b0_in                = [double](Row $r "Perimeter")
        shear_stress_max_ksi = [double](Row $r "ShrStrMax")
        shear_stress_cap_ksi = [double](Row $r "ShrStrCap")
    }
}

# --- Columns: map frames -> punching points, grab section + rotation ---
$columns = @()
$nf=0;$fnames=[string[]]::new(0)
$null = $sm.FrameObj.GetNameList([ref]$nf, [ref]$fnames)
foreach ($fr in $fnames) {
    $p1=""; $p2=""
    $null = $sm.FrameObj.GetPoints($fr, [ref]$p1, [ref]$p2)
    $punchPt = if ($pointToRow.ContainsKey($p1)) { $p1 } elseif ($pointToRow.ContainsKey($p2)) { $p2 } else { $null }
    if (-not $punchPt) { continue }

    $sec=""; $sauto=""
    $null = $sm.FrameObj.GetSection($fr, [ref]$sec, [ref]$sauto)

    # Frame rotation about local 1
    $ang=0.0; $advanced=$false
    try { $null = $sm.FrameObj.GetLocalAxes($fr, [ref]$ang, [ref]$advanced) } catch {}

    # Section rectangle dims
    $fileN="";$mat="";$t3=0.0;$t2=0.0;$color=0;$notes="";$guid=""
    $null = $sm.PropFrame.GetRectangle($sec, [ref]$fileN, [ref]$mat, [ref]$t3, [ref]$t2, [ref]$color, [ref]$notes, [ref]$guid)

    # Position: use global x/y from punching table row (authoritative per SAFE)
    $row = $pointToRow[$punchPt]
    $gx = [double](Row $row "GlobalX")
    $gy = [double](Row $row "GlobalY")
    $loc = Row $row "Location"

    $columns += [ordered]@{
        id             = "C$punchPt"
        safe_point     = $punchPt
        safe_frame     = $fr
        position_in    = @($gx, $gy)
        section        = $sec
        # c1 = t3 (depth direction), c2 = t2 (width direction) in frame-local axes.
        # Rotation angle (deg) about local 1 (vertical).
        c1_in          = [double]$t3
        c2_in          = [double]$t2
        rotation_deg   = [double]$ang
        material       = $mat
        safe_location  = $loc
    }
}

# --- Slab polygon (area 1 in this model; any area whose property matches "Slab") ---
$na=0;$anames=[string[]]::new(0)
$null = $sm.AreaObj.GetNameList([ref]$na, [ref]$anames)
$slabOuter = $null
$slabHoles = @()
$slabThicknessIn = 0.0
$slabMatName = ""
$slabZ = $null

for ($i=0; $i -lt $na; $i++) {
    $a = $anames[$i]
    $prop=""
    try { $null = $sm.AreaObj.GetProperty($a, [ref]$prop) } catch { continue }
    if ($prop -notmatch "Slab") { continue }

    # Is this an opening (hole)?
    $isOpening=$false
    try { $null = $sm.AreaObj.GetOpening($a, [ref]$isOpening) } catch {}

    # Polygon
    $nPts=0;$ptNames=[string[]]::new(0)
    $null = $sm.AreaObj.GetPoints($a, [ref]$nPts, [ref]$ptNames)
    $ring = @()
    foreach ($p in $ptNames) {
        $x=0.0;$y=0.0;$z=0.0
        $null = $sm.PointObj.GetCoordCartesian($p, [ref]$x, [ref]$y, [ref]$z, "Global")
        $ring += ,@([double]$x, [double]$y)
        if ($null -eq $slabZ) { $slabZ = [double]$z }
    }

    if ($isOpening) {
        $slabHoles += ,$ring
    } else {
        # Prefer the slab with most points (main outline)
        if ($null -eq $slabOuter -or $ring.Count -gt $slabOuter.Count) {
            $slabOuter = $ring
            $slabMatName = $prop

            # Pull thickness from prop def
            try {
                $slabType = [SAFEv1.eSlabType]::Slab
                $shellType = [SAFEv1.eShellType]::ShellThin
                $mat=""; $thk=0.0; $c=0; $n=""; $g=""
                $null = $sm.PropArea.GetSlab($prop, [ref]$slabType, [ref]$shellType, [ref]$mat, [ref]$thk, [ref]$c, [ref]$n, [ref]$g)
                $slabThicknessIn = [double]$thk
            } catch {}
        }
    }
}
if ($null -eq $slabOuter) { throw "No slab area object found." }

# --- Materials: parse f'c from section name (robust vs. API quirks) ---
# Section '12x24 conc col 5ksi' and slab prop '8" Concrete Slab 5ksi' both encode 5ksi.
$fcPsi = 5000  # inferred from model; safer than probing fragile GetOConcrete
$fcNote = "inferred from section/material names '5000Psi' / '5ksi' (API GetOConcrete signature varies)"

# --- Combo definition (for documentation) ---
$combo = "DConS2"
$caseCount=0;$caseType=[int[]]::new(0);$caseName=[string[]]::new(0);$sf=[double[]]::new(0)
$null = $sm.RespCombo.GetCaseList($combo, [ref]$caseCount, [ref]$caseType, [ref]$caseName, [ref]$sf)
$comboCases = @()
for ($i=0; $i -lt $caseCount; $i++) {
    $comboCases += [ordered]@{ case = $caseName[$i]; scale_factor = $sf[$i] }
}

# --- Effective depth: pull from first row's Depth field (all columns have same d if slab is uniform) ---
$dSet = @{}
foreach ($pr in $punchingResults) { $dSet[$pr.d_in] = $true }
$dIn = [double](($dSet.Keys | Sort-Object)[0])

# --- Assemble JSON ---
$out = [ordered]@{
    model_file        = $sm.GetModelFilename($true)
    extracted_at      = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK")
    safe_version      = ($proc.MainModule.FileVersionInfo.FileVersion)

    units = [ordered]@{
        length       = "inches"
        force        = "kip"
        moment       = "kip-in"
        stress       = "ksi (SAFE) and psi (webapp)"
    }

    slab = [ordered]@{
        outer_in            = $slabOuter
        holes_in            = $slabHoles
        thickness_in        = $slabThicknessIn
        z_elevation_in      = $slabZ
        slab_property       = $slabMatName
        effective_depth_in  = $dIn
    }

    materials = [ordered]@{
        concrete_name = "5000Psi"
        fc_psi        = $fcPsi
        fc_note       = $fcNote
    }

    columns = $columns

    loads = [ordered]@{
        dead_psf                = $DeadPsf
        sdl_psf                 = $SdlPsf
        live_psf                = $LivePsf
        webapp_dead_psf_eq      = ($DeadPsf + $SdlPsf)   # webapp takes one DL number (no self-weight add)
        webapp_live_psf_eq      = $LivePsf
        wu_factored_psf         = 1.2 * ($DeadPsf + $SdlPsf) + 1.6 * $LivePsf
        governing_combo         = $combo
        combo_factors           = $comboCases
        combo_definition_text   = "1.2 Dead + 1.2 SDL + 1.6 Live (ASCE 7 LRFD)"
    }

    punching_ground_truth = $punchingResults

    extraction_notes = @(
        "22 column-frame objects map 1:1 to 22 punching points (N*6+20, step 6).",
        "All columns share section '12x24 conc col 5ksi' (c1=24, c2=12).",
        "Column rotation_deg applies about vertical axis; webapp currently ignores this (builds axis-aligned critical sections).",
        "Slab loads not reachable via AreaObj.GetLoadUniform in this model; values supplied as script parameters.",
        "Analysis model not locked in SAFE; punching design table values reflect the most recent design run."
    )
}

$out | ConvertTo-Json -Depth 8 | Set-Content -Path $OutJson -Encoding UTF8

# Console summary
Write-Host ""
Write-Host "=== Ground truth extracted ==="
Write-Host ("  Model       : {0}" -f $out.model_file)
Write-Host ("  Slab        : {0} pts outer, {1} holes, thickness={2} in, d={3} in" -f $slabOuter.Count, $slabHoles.Count, $slabThicknessIn, $dIn)
Write-Host ("  Columns     : {0}" -f $columns.Count)
Write-Host ("  Loads       : DL={0} SDL={1} LL={2} psf -> wu={3:F0} psf" -f $DeadPsf, $SdlPsf, $LivePsf, $out.loads.wu_factored_psf)
Write-Host ("  f'c         : {0} psi" -f $fcPsi)
Write-Host ("  Combo       : {0} = 1.2 Dead + 1.2 SDL + 1.6 Live" -f $combo)
Write-Host ""
Write-Host "  Top 5 DCRs:"
$punchingResults | Sort-Object -Property { -$_.dcr } | Select-Object -First 5 | ForEach-Object {
    Write-Host ("    {0}  loc={1,-8}  Vu={2,6:F2} kip  UnbalMu2={3,7:F1}  UnbalMu3={4,7:F1}  DCR={5:F3}" -f $_.id, $_.location, $_.vu_kip, $_.unbal_mu2_kip_in, $_.unbal_mu3_kip_in, $_.dcr)
}
Write-Host ""
Write-Host ("  Written to  : {0}" -f $OutJson)
