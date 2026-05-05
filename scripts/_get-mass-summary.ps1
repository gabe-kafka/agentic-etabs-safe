Add-Type -Path "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll"
$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", 50936)
$sm = $api.SapModel

$groupName = ""
$fk = ""
$tv = 0
$f = @()
$n = 0
$d = @()
$sm.DatabaseTables.GetTableForDisplayArray("Mass Summary by Story", [ref]$fk, $groupName, [ref]$tv, [ref]$f, [ref]$n, [ref]$d) | Out-Null
$nf = $f.Length
Write-Host "Columns: $($f -join ', ') | Records: $n"
for ($i = 0; $i -lt $n; $i++) {
    $line = ""
    for ($j = 0; $j -lt $nf; $j++) {
        $line += "$($f[$j])=$($d[$i * $nf + $j])  "
    }
    Write-Host $line
}

Write-Host ""
Write-Host "=== Seismic load pattern type ==="
$patType = 0
$ret = $sm.LoadPatterns.GetLoadType("Seismic", [ref]$patType)
Write-Host "Type enum: $patType (ret=$ret)"
# 5=Quake, 3=Wind, 1=Dead, 2=SuperDead, 4=Live

Write-Host ""
Write-Host "=== Seismic self-weight multiplier ==="
$swMult = 0.0
$sm.LoadPatterns.GetSelfWTMultiplier("Seismic", [ref]$swMult) | Out-Null
Write-Host "SelfWt multiplier: $swMult"
