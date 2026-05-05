param(
    [int]$SafePid,
    [double]$Tolerance = 0.05,
    [switch]$Apply,
    [switch]$InsertVertex,
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
Add-Type -Path (Resolve-SafeApiDll -Process $process)

$helper = [SAFEv1.cHelper](New-Object SAFEv1.Helper)
$api = $helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", $process.Id)
$sap = $api.SapModel

$areaCount = 0
$areaNames = [string[]]@()
$null = $sap.AreaObj.GetNameList([ref]$areaCount, [ref]$areaNames)

$pointCount = 0
$pointNames = [string[]]@()
$null = $sap.PointObj.GetNameList([ref]$pointCount, [ref]$pointNames)

$pointCoords = @{}
foreach ($p in $pointNames) {
    $x = 0.0; $y = 0.0; $z = 0.0
    $null = $sap.PointObj.GetCoordCartesian($p, [ref]$x, [ref]$y, [ref]$z, "Global")
    $pointCoords[$p] = @($x, $y, $z)
}

$findings = New-Object System.Collections.Generic.List[object]
$areaEdgeConstraint = @{}

foreach ($areaName in $areaNames) {
    $vCount = 0
    $verts = [string[]]@()
    try { $null = $sap.AreaObj.GetPoints($areaName, [ref]$vCount, [ref]$verts) } catch { continue }
    if ($vCount -lt 3) { continue }

    $vertSet = @{}
    foreach ($v in $verts) { $vertSet[$v] = $true }

    $ecOn = $false
    try { $null = $sap.AreaObj.GetEdgeConstraint($areaName, [ref]$ecOn) } catch { $ecOn = $false }
    $areaEdgeConstraint[$areaName] = $ecOn

    for ($i = 0; $i -lt $vCount; $i++) {
        $vA = $verts[$i]
        $vB = $verts[($i + 1) % $vCount]
        if (-not $pointCoords.ContainsKey($vA)) { continue }
        if (-not $pointCoords.ContainsKey($vB)) { continue }
        $cA = $pointCoords[$vA]
        $cB = $pointCoords[$vB]

        $ABx = $cB[0] - $cA[0]
        $ABy = $cB[1] - $cA[1]
        $ABz = $cB[2] - $cA[2]
        $abLen2 = $ABx*$ABx + $ABy*$ABy + $ABz*$ABz
        if ($abLen2 -lt 1e-9) { continue }

        foreach ($p in $pointNames) {
            if ($vertSet.ContainsKey($p)) { continue }
            $cP = $pointCoords[$p]
            $APx = $cP[0] - $cA[0]
            $APy = $cP[1] - $cA[1]
            $APz = $cP[2] - $cA[2]
            $t = ($APx*$ABx + $APy*$ABy + $APz*$ABz) / $abLen2
            if ($t -le 1e-4 -or $t -ge 1 - 1e-4) { continue }
            $cx = $cA[0] + $t*$ABx
            $cy = $cA[1] + $t*$ABy
            $cz = $cA[2] + $t*$ABz
            $dx = $cP[0] - $cx
            $dy = $cP[1] - $cy
            $dz = $cP[2] - $cz
            $d = [Math]::Sqrt($dx*$dx + $dy*$dy + $dz*$dz)
            if ($d -gt $Tolerance) { continue }

            $aLabel = ""; $aStory = ""
            try { $null = $sap.AreaObj.GetLabelFromName($areaName, [ref]$aLabel, [ref]$aStory) } catch {}
            $pLabel = ""; $pStory = ""
            try { $null = $sap.PointObj.GetLabelFromName($p, [ref]$pLabel, [ref]$pStory) } catch {}

            $findings.Add([pscustomobject]@{
                AreaName         = $areaName
                AreaLabel        = $aLabel
                AreaStory        = $aStory
                PointName        = $p
                PointLabel       = $pLabel
                PointStory       = $pStory
                EdgeFromVertex   = $vA
                EdgeToVertex     = $vB
                EdgeIndex        = $i
                ParamT           = [math]::Round($t, 5)
                DistanceFromEdge = [math]::Round($d, 5)
                EdgeConstraintOn = $ecOn
                PointCoord       = @([math]::Round($cP[0],4), [math]::Round($cP[1],4), [math]::Round($cP[2],4))
                EdgeFromCoord    = @([math]::Round($cA[0],4), [math]::Round($cA[1],4), [math]::Round($cA[2],4))
                EdgeToCoord      = @([math]::Round($cB[0],4), [math]::Round($cB[1],4), [math]::Round($cB[2],4))
            })
        }
    }
}

$affectedAreas = @($findings | Select-Object -ExpandProperty AreaName -Unique)

$applied = @()
if ($Apply -and $affectedAreas.Count -gt 0) {
    $null = $sap.SetModelIsLocked($false)
    foreach ($a in $affectedAreas) {
        $already = $areaEdgeConstraint[$a]
        $ret = $sap.AreaObj.SetEdgeConstraint($a, $true, [SAFEv1.eItemType]::Objects)
        $applied += [pscustomobject]@{
            AreaName               = $a
            EdgeConstraintWasOn    = $already
            ApiReturn              = $ret
        }
    }
}

function Snapshot-Area {
    param($sap, [string]$name)

    $snap = [ordered]@{}
    $snap.Name = $name

    $lbl = ""; $sty = ""
    try { $null = $sap.AreaObj.GetLabelFromName($name, [ref]$lbl, [ref]$sty) } catch {}
    $snap.Label = $lbl
    $snap.Story = $sty

    $vc = 0; $vs = [string[]]@()
    $null = $sap.AreaObj.GetPoints($name, [ref]$vc, [ref]$vs)
    $snap.Vertices = @($vs)

    $prop = ""
    try { $null = $sap.AreaObj.GetProperty($name, [ref]$prop) } catch {}
    $snap.Property = $prop

    $isOpening = $false
    try { $null = $sap.AreaObj.GetOpening($name, [ref]$isOpening) } catch {}
    $snap.IsOpening = $isOpening

    $ec = $false
    try { $null = $sap.AreaObj.GetEdgeConstraint($name, [ref]$ec) } catch {}
    $snap.EdgeConstraint = $ec

    $mods = New-Object 'double[]' 10
    try { $null = $sap.AreaObj.GetModifiers($name, [ref]$mods) } catch { $mods = $null }
    $snap.Modifiers = $mods

    $ang = 0.0; $adv = $false
    try { $null = $sap.AreaObj.GetLocalAxes($name, [ref]$ang, [ref]$adv) } catch {}
    $snap.LocalAxisAngle = $ang
    $snap.LocalAxisAdvanced = $adv

    $matOv = ""
    try { $null = $sap.AreaObj.GetMaterialOverwrite($name, [ref]$matOv) } catch {}
    $snap.MaterialOverwrite = $matOv

    $ng = 0; $grps = [string[]]@()
    try { $null = $sap.AreaObj.GetGroupAssign($name, [ref]$ng, [ref]$grps) } catch {}
    $snap.Groups = @($grps)

    $ulN = 0
    $ulArea = [string[]]@()
    $ulPat = [string[]]@()
    $ulCSys = [string[]]@()
    $ulDir = [int[]]@()
    $ulVal = [double[]]@()
    try {
        $null = $sap.AreaObj.GetLoadUniform($name, [ref]$ulN, [ref]$ulArea, [ref]$ulPat, [ref]$ulCSys, [ref]$ulDir, [ref]$ulVal, [SAFEv1.eItemType]::Objects)
    } catch { $ulN = 0 }
    $snap.UniformLoads = @()
    for ($i = 0; $i -lt $ulN; $i++) {
        if ($ulArea[$i] -ne $name) { continue }
        $snap.UniformLoads += [pscustomobject]@{
            Pattern = $ulPat[$i]; CSys = $ulCSys[$i]; Dir = $ulDir[$i]; Value = $ulVal[$i]
        }
    }

    $nCE = 0
    $cvType = [int[]]@(); $cvTen = [double[]]@(); $cvNPt = [int[]]@()
    $cvGx = [double[]]@(); $cvGy = [double[]]@(); $cvGz = [double[]]@()
    try {
        $null = $sap.AreaObj.GetCurvedEdges($name, [ref]$nCE, [ref]$cvType, [ref]$cvTen, [ref]$cvNPt, [ref]$cvGx, [ref]$cvGy, [ref]$cvGz)
    } catch { $nCE = 0 }
    $actualCurved = 0
    for ($ci = 0; $ci -lt $nCE; $ci++) {
        if ($ci -lt $cvType.Length -and $cvType[$ci] -ne 0) { $actualCurved++ }
    }
    $snap.CurvedEdgeCount = $actualCurved

    return $snap
}

function Restore-Area {
    param($sap, $snap, [string]$newName)

    $log = @()

    if (-not [string]::IsNullOrWhiteSpace($snap.Property)) {
        $ret = $sap.AreaObj.SetProperty($newName, $snap.Property, [SAFEv1.eItemType]::Objects)
        $log += "SetProperty=$ret"
    }
    $ret = $sap.AreaObj.SetOpening($newName, $snap.IsOpening, [SAFEv1.eItemType]::Objects)
    $log += "SetOpening=$ret"
    $ret = $sap.AreaObj.SetEdgeConstraint($newName, $snap.EdgeConstraint, [SAFEv1.eItemType]::Objects)
    $log += "SetEdgeConstraint=$ret"
    if ($null -ne $snap.Modifiers) {
        $mods = $snap.Modifiers
        $ret = $sap.AreaObj.SetModifiers($newName, [ref]$mods, [SAFEv1.eItemType]::Objects)
        $log += "SetModifiers=$ret"
    }
    if ($snap.LocalAxisAngle -ne 0.0) {
        $ret = $sap.AreaObj.SetLocalAxes($newName, $snap.LocalAxisAngle, [SAFEv1.eItemType]::Objects)
        $log += "SetLocalAxes=$ret"
    }
    if (-not [string]::IsNullOrWhiteSpace($snap.MaterialOverwrite)) {
        $ret = $sap.AreaObj.SetMaterialOverwrite($newName, $snap.MaterialOverwrite, [SAFEv1.eItemType]::Objects)
        $log += "SetMaterialOverwrite=$ret"
    }
    foreach ($g in $snap.Groups) {
        if ([string]::IsNullOrWhiteSpace($g)) { continue }
        $ret = $sap.AreaObj.SetGroupAssign($newName, $g, $false, [SAFEv1.eItemType]::Objects)
        $log += "SetGroupAssign($g)=$ret"
    }
    foreach ($ul in $snap.UniformLoads) {
        $ret = $sap.AreaObj.SetLoadUniform($newName, $ul.Pattern, $ul.Value, $ul.Dir, $false, $ul.CSys, [SAFEv1.eItemType]::Objects)
        $log += "SetLoadUniform($($ul.Pattern))=$ret"
    }
    return $log
}

$insertResults = @()
if ($InsertVertex -and $findings.Count -gt 0) {
    $null = $sap.SetModelIsLocked($false)

    $byArea = $findings | Group-Object AreaName
    foreach ($grp in $byArea) {
        $areaName = $grp.Name
        $areaFindings = @($grp.Group | Sort-Object -Property EdgeIndex, ParamT -Descending)

        $snap = Snapshot-Area -sap $sap -name $areaName

        if ($snap.CurvedEdgeCount -gt 0) {
            $insertResults += [pscustomobject]@{
                AreaName = $areaName
                Skipped  = $true
                Reason   = "area has curved edges; insert-vertex not supported"
            }
            continue
        }

        $newVerts = New-Object System.Collections.Generic.List[string]
        foreach ($v in $snap.Vertices) { $newVerts.Add($v) }

        foreach ($f in $areaFindings) {
            $insertAt = $f.EdgeIndex + 1
            if ($insertAt -gt $newVerts.Count) { $insertAt = $newVerts.Count }
            $newVerts.Insert($insertAt, $f.PointName)
        }

        $delRet = $sap.AreaObj.Delete($areaName, [SAFEv1.eItemType]::Objects)

        $newName = ""
        $ptArray = $newVerts.ToArray()
        $addRet = $sap.AreaObj.AddByPoint($newVerts.Count, [ref]$ptArray, [ref]$newName, $snap.Property, $snap.Label)

        $restoreLog = Restore-Area -sap $sap -snap $snap -newName $newName

        $insertResults += [pscustomobject]@{
            OldName       = $areaName
            Label         = $snap.Label
            Story         = $snap.Story
            InsertedPts   = @($areaFindings | Select-Object -ExpandProperty PointName)
            OldVertCount  = $snap.Vertices.Count
            NewVertCount  = $newVerts.Count
            DeleteRet     = $delRet
            AddRet        = $addRet
            NewName       = $newName
            RestoreLog    = $restoreLog
        }
    }
}

$summary = [pscustomobject]@{
    ProcessId       = $process.Id
    ModelTitle      = $process.MainWindowTitle
    Units           = $sap.GetPresentUnits()
    Tolerance       = $Tolerance
    AreaCount       = $areaCount
    PointCount      = $pointCount
    FindingCount    = $findings.Count
    AffectedAreas   = @($affectedAreas)
    Applied         = $Apply.IsPresent
    AppliedDetails  = $applied
    InsertVertex    = $InsertVertex.IsPresent
    InsertResults   = $insertResults
    Findings        = $findings
}

if ($AsJson) {
    $summary | ConvertTo-Json -Depth 10
}
else {
    Write-Output ("SAFE PID {0} | units {1} | areas {2} | points {3} | tol {4}" -f `
        $process.Id, $summary.Units, $areaCount, $pointCount, $Tolerance)
    Write-Output ""
    if ($findings.Count -eq 0) {
        Write-Output "No floating edge nodes found (no point lies on any area edge without being a vertex of that area)."
    } else {
        Write-Output ("Found {0} floating edge node(s) across {1} area(s):" -f `
            $findings.Count, @($affectedAreas).Count)
        foreach ($f in $findings) {
            Write-Output ""
            Write-Output ("  point {0} (label {1}) lies on area {2} (label {3}, story {4})" -f `
                $f.PointName, $f.PointLabel, $f.AreaName, $f.AreaLabel, $f.AreaStory)
            Write-Output ("     edge {0}->{1}  t={2}  perp-dist={3}  edgeConstraint={4}" -f `
                $f.EdgeFromVertex, $f.EdgeToVertex, $f.ParamT, $f.DistanceFromEdge, $f.EdgeConstraintOn)
            Write-Output ("     point  @ ({0:F3}, {1:F3}, {2:F3})" -f `
                $f.PointCoord[0], $f.PointCoord[1], $f.PointCoord[2])
            Write-Output ("     edge A @ ({0:F3}, {1:F3}, {2:F3})" -f `
                $f.EdgeFromCoord[0], $f.EdgeFromCoord[1], $f.EdgeFromCoord[2])
            Write-Output ("     edge B @ ({0:F3}, {1:F3}, {2:F3})" -f `
                $f.EdgeToCoord[0], $f.EdgeToCoord[1], $f.EdgeToCoord[2])
        }
    }

    if ($Apply -and @($applied).Count -gt 0) {
        Write-Output ""
        Write-Output "Applied edge constraint on affected areas:"
        foreach ($r in $applied) {
            Write-Output ("  area {0}  was={1}  ret={2}" -f $r.AreaName, $r.EdgeConstraintWasOn, $r.ApiReturn)
        }
    }
    if ($InsertVertex -and @($insertResults).Count -gt 0) {
        Write-Output ""
        Write-Output "Insert-vertex results:"
        foreach ($r in $insertResults) {
            if ($r.PSObject.Properties.Match('Skipped').Count -gt 0 -and $r.Skipped) {
                Write-Output ("  area {0}: SKIPPED ({1})" -f $r.AreaName, $r.Reason)
                continue
            }
            Write-Output ("  area old={0} label={1}: verts {2} -> {3}, inserted [{4}], del={5}, add={6}, newName={7}" -f `
                $r.OldName, $r.Label, $r.OldVertCount, $r.NewVertCount, ($r.InsertedPts -join ','), $r.DeleteRet, $r.AddRet, $r.NewName)
            foreach ($l in $r.RestoreLog) { Write-Output "    $l" }
        }
    }
    if ($findings.Count -gt 0 -and -not $Apply -and -not $InsertVertex) {
        Write-Output ""
        Write-Output "Dry run. Options:"
        Write-Output "  -Apply         enable edge constraint on affected areas (non-destructive)"
        Write-Output "  -InsertVertex  delete+recreate each affected area with the point inserted as a vertex (destructive; restores property/loads/modifiers/groups/edge-constraint/local-axes/opening/material overwrite; does NOT restore: GUID, area name, offsets3, pier/spandrel, curved edges)"
    }
}
