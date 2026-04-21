param([int]$SafePid)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$proc = Get-Process SAFE -ErrorAction SilentlyContinue | Where-Object { -not $SafePid -or $_.Id -eq $SafePid } | Select-Object -First 1
$dll = Join-Path (Split-Path -Parent $proc.Path) "SAFEv1.dll"
if (-not (Test-Path $dll)) { $dll = "C:\Program Files\Computers and Structures\SAFE 21\SAFEv1.dll" }
Add-Type -Path $dll
$api = ([SAFEv1.cHelper](New-Object SAFEv1.Helper)).GetObjectProcess("CSI.SAFE.API.ETABSObject", $proc.Id)
$sm = $api.SapModel
$null = $sm.SetPresentUnits(3)

# FrameObj methods
Write-Host "=== FrameObj methods ==="
($sm.FrameObj | Get-Member -MemberType Method | Where-Object { $_.Name -notmatch "^(get_|set_|add_|remove_|GetType|ToString|Equals|GetHashCode|CreateObjRef|InitializeLifetimeService|GetLifetimeService)" }).Name | Sort-Object

Write-Host ""
Write-Host "=== PropFrame methods ==="
($sm.PropFrame | Get-Member -MemberType Method | Where-Object { $_.Name -notmatch "^(get_|set_|add_|remove_|GetType|ToString|Equals|GetHashCode|CreateObjRef|InitializeLifetimeService|GetLifetimeService)" }).Name | Sort-Object

Write-Host ""
Write-Host "=== Frame → Point connectivity for punching points ==="
# Get frames
$nf = 0;$fnames = [string[]]::new(0)
$null = $sm.FrameObj.GetNameList([ref]$nf, [ref]$fnames)

# Get punching point IDs first
$v=0;$f=[string[]]::new(0);$nr=0;$d=[string[]]::new(0)
$null = $sm.DatabaseTables.GetTableForDisplayArray("Concrete Slab Design - Punching Shear Data", [ref]$f, "", [ref]$v, [ref]$f, [ref]$nr, [ref]$d)
$pIdx = [array]::IndexOf($f,"Point")
$gxIdx = [array]::IndexOf($f,"GlobalX")
$gyIdx = [array]::IndexOf($f,"GlobalY")
$locIdx = [array]::IndexOf($f,"Location")

$punchMap = @{}
for ($i=0; $i -lt $nr; $i++) {
    $p = $d[$i*$f.Length + $pIdx]
    $punchMap[$p] = [ordered]@{
        x = [double]$d[$i*$f.Length + $gxIdx]
        y = [double]$d[$i*$f.Length + $gyIdx]
        loc = $d[$i*$f.Length + $locIdx]
    }
}

# Walk frames, print which connect to punching points + their sections
foreach ($fr in $fnames[0..([Math]::Min(24,$fnames.Length-1))]) {
    $p1=""; $p2=""
    $null = $sm.FrameObj.GetPoints($fr, [ref]$p1, [ref]$p2)
    $sec=""; $sauto=""
    $null = $sm.FrameObj.GetSection($fr, [ref]$sec, [ref]$sauto)
    $hitPoint = if ($punchMap.Contains($p1)) { $p1 } elseif ($punchMap.Contains($p2)) { $p2 } else { "" }
    Write-Host ("Frame {0}: {1}-{2}  sec={3}  punch_pt={4}" -f $fr, $p1, $p2, $sec, $hitPoint)
}

Write-Host ""
Write-Host "=== Unique section names on the 22 frames ==="
$secs = @{}
foreach ($fr in $fnames) {
    $sec=""; $sauto=""
    $null = $sm.FrameObj.GetSection($fr, [ref]$sec, [ref]$sauto)
    $secs[$sec] = ($secs[$sec] + 1)
}
$secs.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" }

Write-Host ""
Write-Host "=== Section property details ==="
foreach ($s in $secs.Keys) {
    $fileName="";$mat="";$t3=0.0;$t2=0.0;$color=0;$notes="";$guid=""
    try {
        $null = $sm.PropFrame.GetRectangle($s, [ref]$fileName, [ref]$mat, [ref]$t3, [ref]$t2, [ref]$color, [ref]$notes, [ref]$guid)
        Write-Host ("  {0}: rect t3={1} (depth) t2={2} (width) mat={3}" -f $s, $t3, $t2, $mat)
    } catch {
        Write-Host ("  {0}: not rect - {1}" -f $s, $_.Exception.Message)
        # Try other shapes
        try {
            $d=0.0
            $null = $sm.PropFrame.GetCircle($s, [ref]$fileName, [ref]$mat, [ref]$d, [ref]$color, [ref]$notes, [ref]$guid)
            Write-Host ("  {0}: circle d={1} mat={2}" -f $s, $d, $mat)
        } catch {}
    }
}

Write-Host ""
Write-Host "=== GetAllTables (no args) ==="
try {
    $numTables = 0
    $tableKeyList = [string[]]::new(0)
    $null = $sm.DatabaseTables.GetAllTables([ref]$numTables, [ref]$tableKeyList)
    Write-Host "Total: $numTables"
    $tableKeyList | Where-Object { $_ -match "(?i)(punch|slab|concrete|area.*load|material|column|frame)" } | Sort-Object | ForEach-Object { Write-Host "  $_" }
} catch {
    Write-Host "GetAllTables err: $($_.Exception.Message)"
    # Find right signature
    $sigs = $sm.DatabaseTables.GetType().GetMethod("GetAllTables").GetParameters()
    Write-Host "Params: $($sigs | ForEach-Object { $_.ParameterType.Name + ' ' + $_.Name })"
}
