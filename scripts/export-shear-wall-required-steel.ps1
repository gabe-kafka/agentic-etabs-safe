param(
    [int]$EtabsPid,

    [string]$OutputPath,

    [switch]$OpenWorkbook,

    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-Success {
    param(
        [int]$ReturnCode,
        [string]$Operation
    )

    if ($ReturnCode -ne 0) {
        throw "$Operation failed with return code $ReturnCode."
    }
}

function Get-EtabsProcess {
    param(
        [int]$RequestedPid
    )

    $processes = Get-Process ETABS -ErrorAction SilentlyContinue | Sort-Object Id
    if ($RequestedPid) {
        return $processes | Where-Object { $_.Id -eq $RequestedPid } | Select-Object -First 1
    }

    $withWindow = $processes | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($withWindow) {
        return $withWindow | Select-Object -First 1
    }

    return $processes | Select-Object -First 1
}

function Resolve-EtabsApiDll {
    param(
        [System.Diagnostics.Process]$Process
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($null -ne $Process) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($Process.Path)) {
                $candidates.Add((Join-Path -Path (Split-Path -Parent $Process.Path) -ChildPath "ETABSv1.dll"))
            }
        }
        catch {
        }
    }

    @(
        "C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files\Computers and Structures\ETABS 20\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 22\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 21\ETABSv1.dll",
        "C:\Program Files (x86)\Computers and Structures\ETABS 20\ETABSv1.dll"
    ) | ForEach-Object {
        $candidates.Add($_)
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    throw "ETABSv1.dll was not found in the running ETABS folder or a standard ETABS install path."
}

function Resolve-CanonicalModelPath {
    param(
        [string]$RawModelPath,
        [string]$ModelDirectory
    )

    if ([string]::IsNullOrWhiteSpace($RawModelPath)) {
        return $null
    }

    if ([System.IO.Path]::GetExtension($RawModelPath) -ieq ".EDB") {
        return $RawModelPath
    }

    if (-not [string]::IsNullOrWhiteSpace($ModelDirectory)) {
        $candidate = Join-Path -Path $ModelDirectory -ChildPath ("{0}.EDB" -f [System.IO.Path]::GetFileNameWithoutExtension($RawModelPath))
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $RawModelPath
}

function Get-CurrentModelPath {
    param(
        $Api
    )

    $rawModelPath = $Api.SapModel.GetModelFilename($true)
    $modelDirectory = $Api.SapModel.GetModelFilepath()
    return Resolve-CanonicalModelPath -RawModelPath $rawModelPath -ModelDirectory $modelDirectory
}

function Get-PresentUnits {
    param(
        $SapModel
    )

    $force = [ETABSv1.eForce]::lb
    $length = [ETABSv1.eLength]::inch
    $temperature = [ETABSv1.eTemperature]::F
    Assert-Success ($SapModel.GetPresentUnits_2([ref]$force, [ref]$length, [ref]$temperature)) "Get present units"

    return [pscustomobject]@{
        Force = $force.ToString()
        Length = $length.ToString()
        Temperature = $temperature.ToString()
    }
}

function Get-LengthUnitLabel {
    param(
        [string]$LengthUnit
    )

    switch -Regex ($LengthUnit) {
        "^inch$" { return "in" }
        "^ft$" { return "ft" }
        "^mm$" { return "mm" }
        "^cm$" { return "cm" }
        "^m$" { return "m" }
        default { return $LengthUnit }
    }
}

function Get-AreaUnitLabel {
    param(
        [string]$LengthUnit
    )

    return "{0}^2" -f (Get-LengthUnitLabel -LengthUnit $LengthUnit)
}

function Get-MomentUnitLabel {
    param(
        [string]$ForceUnit,
        [string]$LengthUnit
    )

    return "{0}-{1}" -f $ForceUnit, (Get-LengthUnitLabel -LengthUnit $LengthUnit)
}

function Get-LengthToFootFactor {
    param(
        [string]$LengthUnit
    )

    switch -Regex ($LengthUnit) {
        "^inch$" { return (1.0 / 12.0) }
        "^ft$" { return 1.0 }
        "^mm$" { return 0.0032808398950131233 }
        "^cm$" { return 0.032808398950131233 }
        "^m$" { return 3.2808398950131235 }
        default { throw "Unsupported length unit for k-ft conversion: $LengthUnit" }
    }
}

function Convert-LengthValueToFeet {
    param(
        [AllowNull()]
        [object]$Value,
        [string]$LengthUnit
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    return ([double]$Value) * (Get-LengthToFootFactor -LengthUnit $LengthUnit)
}

function Convert-AreaValueToSquareInches {
    param(
        [AllowNull()]
        [object]$Value,
        [string]$LengthUnit
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    $lengthFactor = Get-LengthToFootFactor -LengthUnit $LengthUnit
    return ([double]$Value) * [Math]::Pow($lengthFactor * 12.0, 2.0)
}

function Get-DefaultOutputPath {
    param(
        [string]$ModelPath
    )

    if (-not [string]::IsNullOrWhiteSpace($ModelPath)) {
        $directory = Split-Path -Parent $ModelPath
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($ModelPath)
        return Join-Path -Path $directory -ChildPath ("{0}-shear-wall-required-steel.xlsx" -f $stem)
    }

    return Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "etabs-shear-wall-required-steel.xlsx"
}

function Get-PierSortKey {
    param(
        [string]$PierLabel
    )

    if ($PierLabel -match "^([A-Za-z]+)(\d+)$") {
        $prefix = $Matches[1].ToUpperInvariant()
        $suffix = [int]$Matches[2]
        return "{0}|{1:D8}" -f $prefix, $suffix
    }

    return $PierLabel.ToUpperInvariant()
}

function Get-StorySortKey {
    param(
        [string]$StoryName
    )

    if ([string]::IsNullOrWhiteSpace($StoryName)) {
        return 500000
    }

    $normalized = $StoryName.Trim().ToLowerInvariant()
    if ($normalized -eq "roof") {
        return 200000
    }

    if ($normalized -eq "cellar") {
        return -1000
    }

    if ($normalized -match "^story\s*(\d+)$") {
        return [int]$Matches[1]
    }

    if ($normalized -match "^b(?:asement)?\s*(\d+)$") {
        return -1 * [int]$Matches[1]
    }

    return 100000
}

function Get-NormalizedPierLegLabel {
    param(
        [string]$PierLeg
    )

    if ([string]::IsNullOrWhiteSpace($PierLeg)) {
        return ""
    }

    $trimmed = $PierLeg.Trim()
    if ($trimmed -match "^(?:Top|Bottom)\s+(.+)$") {
        return $Matches[1].Trim()
    }

    return $trimmed
}

function Merge-DesignMessages {
    param(
        [string]$WarnMessage,
        [string]$ErrMessage
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($message in @($WarnMessage, $ErrMessage)) {
        if (-not [string]::IsNullOrWhiteSpace($message) -and $message -ne "No Message") {
            $parts.Add($message.Trim()) | Out-Null
        }
    }

    return ($parts -join " | ")
}

function Round-OrNull {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Digits = 3
    )

    if ($null -eq $Value) {
        return $null
    }

    return [Math]::Round([double]$Value, $Digits)
}

function Format-NumberWithCommasOrEmpty {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Digits = 3
    )

    if ($null -eq $Value -or $Value -eq "") {
        return ""
    }

    return ("{0:N$Digits}" -f ([double]$Value))
}

function Sanitize-SheetName {
    param(
        [string]$SheetName
    )

    $cleaned = [regex]::Replace($SheetName, "[\[\]\:\*\?/\\]", "_")
    $cleaned = $cleaned.Trim()
    if ([string]::IsNullOrWhiteSpace($cleaned)) {
        $cleaned = "Sheet"
    }

    if ($cleaned.Length -gt 31) {
        return $cleaned.Substring(0, 31)
    }

    return $cleaned
}

function Get-UniqueSheetNames {
    param(
        [object[]]$SheetDefinitions
    )

    $usedNames = @{}
    $sheetNames = New-Object System.Collections.Generic.List[string]

    foreach ($definition in $SheetDefinitions) {
        $baseName = Sanitize-SheetName -SheetName $definition.Name
        $candidate = $baseName
        $suffix = 1

        while ($usedNames.ContainsKey($candidate.ToUpperInvariant())) {
            $suffixText = "-$suffix"
            $maxBaseLength = [Math]::Max(1, 31 - $suffixText.Length)
            $truncatedBase = $baseName.Substring(0, [Math]::Min($baseName.Length, $maxBaseLength))
            $candidate = $truncatedBase + $suffixText
            $suffix++
        }

        $usedNames[$candidate.ToUpperInvariant()] = $true
        $sheetNames.Add($candidate) | Out-Null
    }

    return $sheetNames.ToArray()
}

function ConvertTo-ColumnName {
    param(
        [int]$Index
    )

    $name = ""
    $current = $Index
    while ($current -gt 0) {
        $current--
        $name = [char](65 + ($current % 26)) + $name
        $current = [int][Math]::Floor($current / 26)
    }

    return $name
}

function Escape-Xml {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Format-NumberInvariant {
    param(
        [object]$Value
    )

    return [Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-CellXml {
    param(
        [string]$CellReference,
        [AllowNull()]
        [object]$Value,
        [int]$StyleId
    )

    if ($null -ne $Value -and $Value -is [ValueType] -and $Value -isnot [bool]) {
        return '<c r="{0}" s="{1}"><v>{2}</v></c>' -f $CellReference, $StyleId, (Format-NumberInvariant -Value $Value)
    }

    return '<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>' -f $CellReference, $StyleId, (Escape-Xml -Value $Value)
}

function Get-WorksheetXml {
    param(
        [string[]]$Headers,
        [object[]]$Rows
    )

    $widths = New-Object System.Collections.Generic.List[int]
    for ($columnIndex = 0; $columnIndex -lt $Headers.Length; $columnIndex++) {
        $maxLength = ([string]$Headers[$columnIndex]).Length
        foreach ($row in $Rows) {
            $value = $row.Cells[$columnIndex]
            if ($null -eq $value) {
                $text = ""
            }
            elseif ($value -is [double] -or $value -is [float] -or $value -is [decimal]) {
                $text = ("{0:N3}" -f $value)
            }
            else {
                $text = [string]$value
            }

            if ($text.Length -gt $maxLength) {
                $maxLength = $text.Length
            }
        }

        $widths.Add([Math]::Min($maxLength + 2, 60)) | Out-Null
    }

    $columnXml = New-Object System.Collections.Generic.List[string]
    for ($index = 1; $index -le $Headers.Length; $index++) {
        $width = $widths[$index - 1]
        $columnXml.Add(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f $index, $width)) | Out-Null
    }

    $rowsXml = New-Object System.Collections.Generic.List[string]
    $headerCells = New-Object System.Collections.Generic.List[string]
    for ($columnIndex = 1; $columnIndex -le $Headers.Length; $columnIndex++) {
        $headerCells.Add((Get-CellXml -CellReference ("{0}1" -f (ConvertTo-ColumnName -Index $columnIndex)) -Value $Headers[$columnIndex - 1] -StyleId 1)) | Out-Null
    }
    $rowsXml.Add(('<row r="1">{0}</row>' -f ($headerCells -join ""))) | Out-Null

    for ($rowIndex = 0; $rowIndex -lt $Rows.Length; $rowIndex++) {
        $styleId = if ($Rows[$rowIndex].Highlight) { 2 } else { 0 }
        $cells = New-Object System.Collections.Generic.List[string]
        for ($columnIndex = 1; $columnIndex -le $Headers.Length; $columnIndex++) {
            $cells.Add((Get-CellXml -CellReference ("{0}{1}" -f (ConvertTo-ColumnName -Index $columnIndex), ($rowIndex + 2)) -Value $Rows[$rowIndex].Cells[$columnIndex - 1] -StyleId $styleId)) | Out-Null
        }
        $rowsXml.Add(('<row r="{0}">{1}</row>' -f ($rowIndex + 2), ($cells -join ""))) | Out-Null
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>$($columnXml -join "")</cols>
  <sheetData>$($rowsXml -join "")</sheetData>
</worksheet>
"@
}

function Get-WorkbookXml {
    param(
        [string[]]$SheetNames
    )

    $sheetXml = New-Object System.Collections.Generic.List[string]
    for ($index = 1; $index -le $SheetNames.Length; $index++) {
        $escapedSheetName = Escape-Xml -Value ($SheetNames[$index - 1])
        $sheetXml.Add(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f $escapedSheetName, $index)) | Out-Null
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>$($sheetXml -join "")</sheets>
</workbook>
"@
}

function Get-WorkbookRelsXml {
    param(
        [int]$SheetCount
    )

    $relationships = New-Object System.Collections.Generic.List[string]
    for ($index = 1; $index -le $SheetCount; $index++) {
        $relationships.Add(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $index)) | Out-Null
    }
    $relationships.Add(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' -f ($SheetCount + 1))) | Out-Null

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $($relationships -join "")
</Relationships>
"@
}

function Get-ContentTypesXml {
    param(
        [int]$SheetCount
    )

    $overrides = New-Object System.Collections.Generic.List[string]
    $overrides.Add('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>') | Out-Null
    $overrides.Add('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>') | Out-Null
    $overrides.Add('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>') | Out-Null
    $overrides.Add('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>') | Out-Null
    for ($index = 1; $index -le $SheetCount; $index++) {
        $overrides.Add(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $index)) | Out-Null
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  $($overrides -join "")
</Types>
"@
}

function Get-RootRelsXml {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@
}

function Get-StylesXml {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@
}

function Get-CoreXml {
    $timestamp = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$timestamp</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$timestamp</dcterms:modified>
</cp:coreProperties>
"@
}

function Get-AppXml {
    param(
        [string[]]$SheetNames
    )

    $titles = New-Object System.Collections.Generic.List[string]
    foreach ($sheetName in $SheetNames) {
        $titles.Add("<vt:lpstr>$(Escape-Xml -Value $sheetName)</vt:lpstr>") | Out-Null
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <TitlesOfParts><vt:vector size="$($SheetNames.Length)" baseType="lpstr">$($titles -join "")</vt:vector></TitlesOfParts>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>$($SheetNames.Length)</vt:i4></vt:variant></vt:vector></HeadingPairs>
</Properties>
"@
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    [System.IO.File]::WriteAllText($Path, $Content, ([System.Text.UTF8Encoding]::new($false)))
}

function Write-XlsxWorkbook {
    param(
        [string]$Path,
        [object[]]$SheetDefinitions
    )

    $sheetNames = Get-UniqueSheetNames -SheetDefinitions $SheetDefinitions
    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    $tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    $tempZip = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("{0}.zip" -f [System.IO.Path]::GetRandomFileName())

    try {
        New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "[Content_Types].xml") -Content (Get-ContentTypesXml -SheetCount $sheetNames.Length)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "_rels/.rels") -Content (Get-RootRelsXml)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "xl/workbook.xml") -Content (Get-WorkbookXml -SheetNames $sheetNames)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "xl/_rels/workbook.xml.rels") -Content (Get-WorkbookRelsXml -SheetCount $sheetNames.Length)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "xl/styles.xml") -Content (Get-StylesXml)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "docProps/core.xml") -Content (Get-CoreXml)
        Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath "docProps/app.xml") -Content (Get-AppXml -SheetNames $sheetNames)

        for ($index = 0; $index -lt $SheetDefinitions.Length; $index++) {
            Write-Utf8File -Path (Join-Path -Path $tempRoot -ChildPath ("xl/worksheets/sheet{0}.xml" -f ($index + 1))) -Content (Get-WorksheetXml -Headers $SheetDefinitions[$index].Headers -Rows $SheetDefinitions[$index].Rows)
        }

        if (Test-Path -LiteralPath $tempZip) {
            Remove-Item -LiteralPath $tempZip -Force
        }

        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }

        $pythonZipCode = @'
import os
import sys
import zipfile

root = sys.argv[1]
output_path = sys.argv[2]

with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
    for dirpath, _dirnames, filenames in os.walk(root):
        for filename in filenames:
            full_path = os.path.join(dirpath, filename)
            relative_path = os.path.relpath(full_path, root).replace("\\", "/")
            archive.write(full_path, relative_path)
'@

        $pythonZipCode | & python - $tempRoot $tempZip
        if ($LASTEXITCODE -ne 0) {
            throw "Python failed while packaging the XLSX workbook."
        }

        Move-Item -LiteralPath $tempZip -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }

        if (Test-Path -LiteralPath $tempZip) {
            Remove-Item -LiteralPath $tempZip -Force
        }
    }
}

function Get-PierSummaryRows {
    param(
        $SapModel
    )

    $summaryTableLookup = Get-PierDesignSummaryTableLookup -SapModel $SapModel
    $pierGeometryLookup = Get-PierGeometryTableLookup -SapModel $SapModel

    [string[]]$Story = @()
    [string[]]$PierLabel = @()
    [string[]]$Station = @()
    [string[]]$DesignType = @()
    [string[]]$PierSecType = @()
    [string[]]$EdgeBar = @()
    [string[]]$EndBar = @()
    [double[]]$BarSpacing = @()
    [double[]]$ReinfPercent = @()
    [double[]]$CurrPercent = @()
    [double[]]$DCRatio = @()
    [string[]]$PierLeg = @()
    [double[]]$LegX1 = @()
    [double[]]$LegY1 = @()
    [double[]]$LegX2 = @()
    [double[]]$LegY2 = @()
    [double[]]$EdgeLeft = @()
    [double[]]$EdgeRight = @()
    [double[]]$AsLeft = @()
    [double[]]$AsRight = @()
    [double[]]$ShearAv = @()
    [double[]]$StressCompLeft = @()
    [double[]]$StressCompRight = @()
    [double[]]$StressLimitLeft = @()
    [double[]]$StressLimitRight = @()
    [double[]]$CDepthLeft = @()
    [double[]]$CLimitLeft = @()
    [double[]]$CDepthRight = @()
    [double[]]$CLimitRight = @()
    [double[]]$InelasticRotDemand = @()
    [double[]]$InelasticRotCapacity = @()
    [double[]]$NormCompStress = @()
    [double[]]$NormCompStressLimit = @()
    [double[]]$CDepth = @()
    [double[]]$BZoneL = @()
    [double[]]$BZoneR = @()
    [double[]]$BZoneLength = @()
    [string[]]$WarnMsg = @()
    [string[]]$ErrMsg = @()

    Assert-Success (
        $SapModel.DesignShearWall.GetPierSummaryResults(
            [ref]$Story,
            [ref]$PierLabel,
            [ref]$Station,
            [ref]$DesignType,
            [ref]$PierSecType,
            [ref]$EdgeBar,
            [ref]$EndBar,
            [ref]$BarSpacing,
            [ref]$ReinfPercent,
            [ref]$CurrPercent,
            [ref]$DCRatio,
            [ref]$PierLeg,
            [ref]$LegX1,
            [ref]$LegY1,
            [ref]$LegX2,
            [ref]$LegY2,
            [ref]$EdgeLeft,
            [ref]$EdgeRight,
            [ref]$AsLeft,
            [ref]$AsRight,
            [ref]$ShearAv,
            [ref]$StressCompLeft,
            [ref]$StressCompRight,
            [ref]$StressLimitLeft,
            [ref]$StressLimitRight,
            [ref]$CDepthLeft,
            [ref]$CLimitLeft,
            [ref]$CDepthRight,
            [ref]$CLimitRight,
            [ref]$InelasticRotDemand,
            [ref]$InelasticRotCapacity,
            [ref]$NormCompStress,
            [ref]$NormCompStressLimit,
            [ref]$CDepth,
            [ref]$BZoneL,
            [ref]$BZoneR,
            [ref]$BZoneLength,
            [ref]$WarnMsg,
            [ref]$ErrMsg
        )
    ) "Get shear wall pier summary results"

    $rows = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $Story.Length; $index++) {
        $summaryKey = "{0}|{1}|{2}" -f $PierLabel[$index], $Story[$index], $Station[$index]
        $summaryRow = if ($summaryTableLookup.ContainsKey($summaryKey)) { $summaryTableLookup[$summaryKey] } else { $null }
        $geometryKey = "{0}|{1}" -f $PierLabel[$index], $Story[$index]
        $geometryRow = if ($pierGeometryLookup.ContainsKey($geometryKey)) { $pierGeometryLookup[$geometryKey] } else { $null }
        $rows.Add([pscustomobject]@{
            Story = $Story[$index]
            PierLabel = $PierLabel[$index]
            Station = $Station[$index]
            DesignType = $DesignType[$index]
            PierSecType = $PierSecType[$index]
            PierLeg = $PierLeg[$index]
            Length = if ($null -ne $summaryRow -and $null -ne $summaryRow.Length) { [double]$summaryRow.Length } elseif ($null -ne $geometryRow -and $null -ne $geometryRow.Length) { [double]$geometryRow.Length } else { $null }
            Thickness = if ($null -ne $summaryRow -and $null -ne $summaryRow.Thickness) { [double]$summaryRow.Thickness } elseif ($null -ne $geometryRow -and $null -ne $geometryRow.Thickness) { [double]$geometryRow.Thickness } else { $null }
            EdgeBar = $EdgeBar[$index]
            EndBar = $EndBar[$index]
            BarSpacing = [double]$BarSpacing[$index]
            ReinfPercent = [double]$ReinfPercent[$index]
            CurrPercent = [double]$CurrPercent[$index]
            AsLeft = [double]$AsLeft[$index]
            AsRight = [double]$AsRight[$index]
            ShearAv = [double]$ShearAv[$index]
            StressCompLeft = [double]$StressCompLeft[$index]
            StressCompRight = [double]$StressCompRight[$index]
            BZoneL = [double]$BZoneL[$index]
            BZoneR = [double]$BZoneR[$index]
            BZoneLength = [double]$BZoneLength[$index]
            WarnMsg = $WarnMsg[$index]
            ErrMsg = $ErrMsg[$index]
        }) | Out-Null
    }

    return $rows.ToArray()
}

function Build-PierReportRows {
    param(
        [object[]]$PierSummaryRows
    )

    $grouped = @{}
    foreach ($row in $PierSummaryRows) {
        $normalizedLeg = Get-NormalizedPierLegLabel -PierLeg $row.PierLeg
        $key = "{0}|{1}|{2}|{3}|{4}" -f $row.PierLabel, $row.Story, $normalizedLeg, $row.DesignType, $row.PierSecType
        if (-not $grouped.ContainsKey($key)) {
            $grouped[$key] = [ordered]@{
                PierLabel = $row.PierLabel
                Story = $row.Story
                PierLeg = $normalizedLeg
                DesignType = $row.DesignType
                PierSecType = $row.PierSecType
                Length = [Nullable[double]]$null
                TopEdgeBar = ""
                TopEndBar = ""
                TopBarSpacing = [Nullable[double]]$null
                TopReinfPercent = [Nullable[double]]$null
                TopCurrPercent = [Nullable[double]]$null
                TopAsLeft = [Nullable[double]]$null
                TopAsRight = [Nullable[double]]$null
                BottomAsLeft = [Nullable[double]]$null
                BottomAsRight = [Nullable[double]]$null
                TopShearAv = [Nullable[double]]$null
                BottomShearAv = [Nullable[double]]$null
                TopBZoneL = [Nullable[double]]$null
                TopBZoneR = [Nullable[double]]$null
                TopBZoneLength = [Nullable[double]]$null
                BottomEdgeBar = ""
                BottomEndBar = ""
                BottomBarSpacing = [Nullable[double]]$null
                BottomReinfPercent = [Nullable[double]]$null
                BottomCurrPercent = [Nullable[double]]$null
                BottomBZoneL = [Nullable[double]]$null
                BottomBZoneR = [Nullable[double]]$null
                BottomBZoneLength = [Nullable[double]]$null
                TopMessage = ""
                BottomMessage = ""
                HasIssue = $false
            }
        }

        $entry = $grouped[$key]
        $entry.Length = Get-MaxDoubleOrNull -Values @($entry.Length, $row.Length)
        $message = Merge-DesignMessages -WarnMessage $row.WarnMsg -ErrMessage $row.ErrMsg
        switch -Regex ($row.Station) {
            "^Top$" {
                $entry.TopEdgeBar = $row.EdgeBar
                $entry.TopEndBar = $row.EndBar
                $entry.TopBarSpacing = [Nullable[double]]$row.BarSpacing
                $entry.TopReinfPercent = [Nullable[double]]$row.ReinfPercent
                $entry.TopCurrPercent = [Nullable[double]]$row.CurrPercent
                $entry.TopAsLeft = [Nullable[double]]$row.AsLeft
                $entry.TopAsRight = [Nullable[double]]$row.AsRight
                $entry.TopShearAv = [Nullable[double]]$row.ShearAv
                $entry.TopBZoneL = [Nullable[double]]$row.BZoneL
                $entry.TopBZoneR = [Nullable[double]]$row.BZoneR
                $entry.TopBZoneLength = [Nullable[double]]$row.BZoneLength
                $entry.TopMessage = $message
                break
            }
            "^Bottom$" {
                $entry.BottomEdgeBar = $row.EdgeBar
                $entry.BottomEndBar = $row.EndBar
                $entry.BottomBarSpacing = [Nullable[double]]$row.BarSpacing
                $entry.BottomReinfPercent = [Nullable[double]]$row.ReinfPercent
                $entry.BottomCurrPercent = [Nullable[double]]$row.CurrPercent
                $entry.BottomAsLeft = [Nullable[double]]$row.AsLeft
                $entry.BottomAsRight = [Nullable[double]]$row.AsRight
                $entry.BottomShearAv = [Nullable[double]]$row.ShearAv
                $entry.BottomBZoneL = [Nullable[double]]$row.BZoneL
                $entry.BottomBZoneR = [Nullable[double]]$row.BZoneR
                $entry.BottomBZoneLength = [Nullable[double]]$row.BZoneLength
                $entry.BottomMessage = $message
                break
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($message)) {
            $entry.HasIssue = $true
        }
    }

    $result = $grouped.Values | Sort-Object `
        @{ Expression = { Get-PierSortKey -PierLabel $_.PierLabel } }, `
        @{ Expression = { Get-StorySortKey -StoryName $_.Story } }, `
        @{ Expression = { $_.PierLeg } }, `
        @{ Expression = { $_.DesignType } }, `
        @{ Expression = { $_.PierSecType } }

    return @($result | ForEach-Object { [pscustomobject]$_ })
}

function Get-MaxDoubleOrNull {
    param(
        [object[]]$Values
    )

    $valid = New-Object System.Collections.Generic.List[double]
    foreach ($value in $Values) {
        if ($null -ne $value -and $value -ne "") {
            $valid.Add([double]$value) | Out-Null
        }
    }

    if ($valid.Count -eq 0) {
        return $null
    }

    return [double](($valid | Measure-Object -Maximum).Maximum)
}

function Get-FirstNonEmptyString {
    param(
        [string[]]$Values
    )

    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Merge-MessageValues {
    param(
        [string[]]$Values
    )

    $unique = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace($value) -and -not $unique.Contains($value)) {
            $unique.Add($value) | Out-Null
        }
    }

    return ($unique -join " | ")
}

function Get-SelectedResultState {
    param(
        $SapModel
    )

    $setup = $SapModel.Results.Setup

    $caseNamesList = New-Object System.Collections.Generic.List[string]
    foreach ($caseType in [System.Enum]::GetValues([ETABSv1.eLoadCaseType])) {
        $caseCount = 0
        [string[]]$caseNames = @()
        $ret = $SapModel.LoadCases.GetNameList([ref]$caseCount, [ref]$caseNames, $caseType)
        if ($ret -eq 0 -and $null -ne $caseNames) {
            foreach ($caseName in $caseNames) {
                if (-not [string]::IsNullOrWhiteSpace($caseName) -and -not $caseNamesList.Contains($caseName)) {
                    $caseNamesList.Add($caseName) | Out-Null
                }
            }
        }
    }

    $comboCount = 0
    [string[]]$comboNames = @()
    Assert-Success ($SapModel.RespCombo.GetNameList([ref]$comboCount, [ref]$comboNames)) "Get load combination names"

    $selectedCases = @{}
    foreach ($caseName in $caseNamesList) {
        $selected = $false
        Assert-Success ($setup.GetCaseSelectedForOutput($caseName, [ref]$selected)) "Get selected output case '$caseName'"
        $selectedCases[$caseName] = [bool]$selected
    }

    $selectedCombos = @{}
    foreach ($comboName in $comboNames) {
        $selected = $false
        Assert-Success ($setup.GetComboSelectedForOutput($comboName, [ref]$selected)) "Get selected output combo '$comboName'"
        $selectedCombos[$comboName] = [bool]$selected
    }

    return [pscustomobject]@{
        Cases = $selectedCases
        Combos = $selectedCombos
    }
}

function Restore-SelectedResultState {
    param(
        $SapModel,
        [object]$State
    )

    $setup = $SapModel.Results.Setup
    Assert-Success ($setup.DeselectAllCasesAndCombosForOutput()) "Deselect output cases and combos"

    foreach ($caseName in $State.Cases.Keys) {
        if ($State.Cases[$caseName]) {
            Assert-Success ($setup.SetCaseSelectedForOutput($caseName, $true)) "Restore output case '$caseName'"
        }
    }

    foreach ($comboName in $State.Combos.Keys) {
        if ($State.Combos[$comboName]) {
            Assert-Success ($setup.SetComboSelectedForOutput($comboName, $true)) "Restore output combo '$comboName'"
        }
    }
}

function Get-LrfdEnvPierMomentLookup {
    param(
        $SapModel
    )

    $savedState = Get-SelectedResultState -SapModel $SapModel
    $setup = $SapModel.Results.Setup

    try {
        Assert-Success ($setup.DeselectAllCasesAndCombosForOutput()) "Deselect output cases and combos"
        Assert-Success ($setup.SetComboSelectedForOutput("LRFD-ENV", $true)) "Select LRFD-ENV for output"

        $numberResults = 0
        [string[]]$storyNames = @()
        [string[]]$pierNames = @()
        [string[]]$loadCases = @()
        [string[]]$locations = @()
        [double[]]$P = @()
        [double[]]$V2 = @()
        [double[]]$V3 = @()
        [double[]]$T = @()
        [double[]]$M2 = @()
        [double[]]$M3 = @()

        Assert-Success (
            $SapModel.Results.PierForce(
                [ref]$numberResults,
                [ref]$storyNames,
                [ref]$pierNames,
                [ref]$loadCases,
                [ref]$locations,
                [ref]$P,
                [ref]$V2,
                [ref]$V3,
                [ref]$T,
                [ref]$M2,
                [ref]$M3
            )
        ) "Get LRFD-ENV pier forces"

        $lookup = @{}
        for ($index = 0; $index -lt $numberResults; $index++) {
            $key = "{0}|{1}" -f $pierNames[$index], $storyNames[$index]
            if (-not $lookup.ContainsKey($key)) {
                $lookup[$key] = [ordered]@{
                    MaxM3 = [Nullable[double]]$null
                    MinM3 = [Nullable[double]]$null
                }
            }

            $m3Value = [double]$M3[$index]
            $entry = $lookup[$key]

            if ($null -eq $entry.MaxM3 -or $m3Value -gt $entry.MaxM3) {
                $entry.MaxM3 = [Nullable[double]]$m3Value
            }

            if ($null -eq $entry.MinM3 -or $m3Value -lt $entry.MinM3) {
                $entry.MinM3 = [Nullable[double]]$m3Value
            }
        }

        return $lookup
    }
    finally {
        Restore-SelectedResultState -SapModel $SapModel -State $savedState
    }
}

function Get-PierDesignSummaryTableLookup {
    param(
        $SapModel
    )

    $db = $SapModel.DatabaseTables
    [string[]]$fieldKeys = @()
    [string[]]$fields = @()
    [string[]]$data = @()
    $tableVersion = 0
    $numRecords = 0

    Assert-Success (
        $db.GetTableForDisplayArray(
            "Shear Wall Pier Design Summary - ACI 318-19",
            [ref]$fieldKeys,
            "All",
            [ref]$tableVersion,
            [ref]$fields,
            [ref]$numRecords,
            [ref]$data
        )
    ) "Get shear wall pier design summary table"

    $fieldCount = $fields.Length
    $lookup = @{}
    for ($recordIndex = 0; $recordIndex -lt $numRecords; $recordIndex++) {
        $offset = $recordIndex * $fieldCount
        $record = @{}
        for ($fieldIndex = 0; $fieldIndex -lt $fieldCount; $fieldIndex++) {
            $record[$fields[$fieldIndex]] = $data[$offset + $fieldIndex]
        }

        $key = "{0}|{1}|{2}" -f $record["Pier"], $record["Story"], $record["Station"]
        $lookup[$key] = [pscustomobject]@{
            Length = if ($null -ne $record["Length"] -and $record["Length"] -ne "") { [double]$record["Length"] } else { $null }
            Thickness = if ($null -ne $record["Thickness"] -and $record["Thickness"] -ne "") { [double]$record["Thickness"] } else { $null }
        }
    }

    return $lookup
}

function Get-PierGeometryTableLookup {
    param(
        $SapModel
    )

    $db = $SapModel.DatabaseTables
    [string[]]$fieldKeys = @()
    [string[]]$fields = @()
    [string[]]$data = @()
    $tableVersion = 0
    $numRecords = 0

    Assert-Success (
        $db.GetTableForDisplayArray(
            "Pier Section Properties",
            [ref]$fieldKeys,
            "All",
            [ref]$tableVersion,
            [ref]$fields,
            [ref]$numRecords,
            [ref]$data
        )
    ) "Get pier section properties table"

    $fieldCount = $fields.Length
    $lookup = @{}
    for ($recordIndex = 0; $recordIndex -lt $numRecords; $recordIndex++) {
        $offset = $recordIndex * $fieldCount
        $record = @{}
        for ($fieldIndex = 0; $fieldIndex -lt $fieldCount; $fieldIndex++) {
            $record[$fields[$fieldIndex]] = $data[$offset + $fieldIndex]
        }

        $widthValues = @($record["WidthBot"], $record["WidthTop"]) | Where-Object { $null -ne $_ -and $_ -ne "" } | ForEach-Object { [double]$_ }
        $thicknessValues = @($record["ThickBot"], $record["ThickTop"]) | Where-Object { $null -ne $_ -and $_ -ne "" } | ForEach-Object { [double]$_ }
        $key = "{0}|{1}" -f $record["Pier"], $record["Story"]
        $lookup[$key] = [pscustomobject]@{
            Length = if ($widthValues.Count -gt 0) { ($widthValues | Measure-Object -Maximum).Maximum } else { $null }
            Thickness = if ($thicknessValues.Count -gt 0) { ($thicknessValues | Measure-Object -Maximum).Maximum } else { $null }
        }
    }

    return $lookup
}

function Build-FloorSummaryRows {
    param(
        [object[]]$PierReportRows,
        [hashtable]$LrfdEnvPierMomentLookup
    )

    $grouped = @{}
    foreach ($row in $PierReportRows) {
        $key = "{0}|{1}" -f $row.PierLabel, $row.Story
        if (-not $grouped.ContainsKey($key)) {
            $grouped[$key] = [ordered]@{
                PierLabel = $row.PierLabel
                Story = $row.Story
                DesignType = ""
                PierSecType = ""
                WallLength = [Nullable[double]]$null
                EdgeBar = ""
                EndBar = ""
                BarSpacing = [Nullable[double]]$null
                RequiredReinfPercent = [Nullable[double]]$null
                CurrentReinfPercent = [Nullable[double]]$null
                RequiredSteelLeft = [Nullable[double]]$null
                RequiredSteelRight = [Nullable[double]]$null
                ShearAv = [Nullable[double]]$null
                BoundaryZoneLeft = [Nullable[double]]$null
                BoundaryZoneRight = [Nullable[double]]$null
                LrfdEnvMaxM3 = [Nullable[double]]$null
                LrfdEnvMinM3 = [Nullable[double]]$null
                Message = ""
                HasIssue = $false
            }
        }

        $entry = $grouped[$key]
        $momentKey = "{0}|{1}" -f $row.PierLabel, $row.Story
        $directMomentRow = if ($LrfdEnvPierMomentLookup.ContainsKey($momentKey)) { $LrfdEnvPierMomentLookup[$momentKey] } else { $null }
        $entry.DesignType = Get-FirstNonEmptyString -Values @($entry.DesignType, $row.DesignType)
        $entry.PierSecType = Get-FirstNonEmptyString -Values @($entry.PierSecType, $row.PierSecType)
        $entry.WallLength = Get-MaxDoubleOrNull -Values @($entry.WallLength, $row.Length)
        $entry.EdgeBar = Get-FirstNonEmptyString -Values @($entry.EdgeBar, $row.TopEdgeBar, $row.BottomEdgeBar)
        $entry.EndBar = Get-FirstNonEmptyString -Values @($entry.EndBar, $row.TopEndBar, $row.BottomEndBar)
        $entry.BarSpacing = Get-MaxDoubleOrNull -Values @($entry.BarSpacing, $row.TopBarSpacing, $row.BottomBarSpacing)
        $entry.RequiredReinfPercent = Get-MaxDoubleOrNull -Values @($entry.RequiredReinfPercent, $row.TopReinfPercent, $row.BottomReinfPercent)
        $entry.CurrentReinfPercent = Get-MaxDoubleOrNull -Values @($entry.CurrentReinfPercent, $row.TopCurrPercent, $row.BottomCurrPercent)
        $entry.RequiredSteelLeft = Get-MaxDoubleOrNull -Values @($entry.RequiredSteelLeft, $row.TopAsLeft, $row.BottomAsLeft)
        $entry.RequiredSteelRight = Get-MaxDoubleOrNull -Values @($entry.RequiredSteelRight, $row.TopAsRight, $row.BottomAsRight)
        $entry.ShearAv = Get-MaxDoubleOrNull -Values @($entry.ShearAv, $row.TopShearAv, $row.BottomShearAv)
        $entry.BoundaryZoneLeft = Get-MaxDoubleOrNull -Values @($entry.BoundaryZoneLeft, $row.TopBZoneL, $row.BottomBZoneL)
        $entry.BoundaryZoneRight = Get-MaxDoubleOrNull -Values @($entry.BoundaryZoneRight, $row.TopBZoneR, $row.BottomBZoneR)
        if ($null -ne $directMomentRow) {
            $entry.LrfdEnvMaxM3 = [Nullable[double]]$directMomentRow.MaxM3
            $entry.LrfdEnvMinM3 = [Nullable[double]]$directMomentRow.MinM3
        }

        $entry.Message = Merge-MessageValues -Values @($entry.Message, $row.TopMessage, $row.BottomMessage)

        if (-not [string]::IsNullOrWhiteSpace($row.TopMessage) -or -not [string]::IsNullOrWhiteSpace($row.BottomMessage)) {
            $entry.HasIssue = $true
        }
    }

    $result = $grouped.Values | Sort-Object `
        @{ Expression = { Get-PierSortKey -PierLabel $_.PierLabel } }, `
        @{ Expression = { Get-StorySortKey -StoryName $_.Story } }

    return @($result | ForEach-Object { [pscustomobject]$_ })
}

$process = Get-EtabsProcess -RequestedPid $EtabsPid
if ($null -eq $process) {
    throw "No running ETABS process was found."
}

$apiDllPath = Resolve-EtabsApiDll -Process $process
Add-Type -Path $apiDllPath

$helper = [ETABSv1.cHelper](New-Object ETABSv1.Helper)
$api = $helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", $process.Id)
$modelPath = Get-CurrentModelPath -Api $api
$units = Get-PresentUnits -SapModel $api.SapModel
$areaUnits = Get-AreaUnitLabel -LengthUnit $units.Length
$steelAreaUnits = "in^2"
$momentUnits = Get-MomentUnitLabel -ForceUnit $units.Force -LengthUnit $units.Length
$resolvedOutputPath = if ([string]::IsNullOrWhiteSpace($OutputPath)) { Get-DefaultOutputPath -ModelPath $modelPath } else { $OutputPath }

$rawRows = Get-PierSummaryRows -SapModel $api.SapModel
if (-not $rawRows -or $rawRows.Count -eq 0) {
    throw "No shear wall pier summary results were returned from the active model."
}

$reportRows = Build-PierReportRows -PierSummaryRows $rawRows
$lrfdEnvMomentLookup = Get-LrfdEnvPierMomentLookup -SapModel $api.SapModel
$floorRows = Build-FloorSummaryRows -PierReportRows $reportRows -LrfdEnvPierMomentLookup $lrfdEnvMomentLookup
$headers = @(
    "Pier",
    "Story",
    "DesignType",
    "WallLength_ft",
    ("AS_REQ_LEFT_{0}" -f $steelAreaUnits),
    ("AS_REQ_RIGHT_{0}" -f $steelAreaUnits),
    ("LRFDEnvMaxM3_{0}" -f $momentUnits),
    ("LRFDEnvMinM3_{0}" -f $momentUnits),
    ("BoundaryZoneLeft_{0}" -f (Get-LengthUnitLabel -LengthUnit $units.Length)),
    ("BoundaryZoneRight_{0}" -f (Get-LengthUnitLabel -LengthUnit $units.Length)),
    "Message"
)

$infoRows = @(
    [pscustomobject]@{ Cells = @("ModelPath", $modelPath); Highlight = $false },
    [pscustomobject]@{ Cells = @("ProcessId", $process.Id); Highlight = $false },
    [pscustomobject]@{ Cells = @("ApiDllPath", $apiDllPath); Highlight = $false },
    [pscustomobject]@{ Cells = @("ForceUnits", $units.Force); Highlight = $false },
    [pscustomobject]@{ Cells = @("LengthUnits", $units.Length); Highlight = $false },
    [pscustomobject]@{ Cells = @("AreaUnits", $areaUnits); Highlight = $false },
    [pscustomobject]@{ Cells = @("SteelAreaOutputUnits", $steelAreaUnits); Highlight = $false },
    [pscustomobject]@{ Cells = @("MomentUnits", $momentUnits); Highlight = $false },
    [pscustomobject]@{ Cells = @("RawStationRowCount", $rawRows.Count); Highlight = $false },
    [pscustomobject]@{ Cells = @("GroupedPierStoryCount", $reportRows.Count); Highlight = $false },
    [pscustomobject]@{ Cells = @("FloorSummaryRowCount", $floorRows.Count); Highlight = $false },
    [pscustomobject]@{ Cells = @("PierCount", (@($floorRows | Select-Object -ExpandProperty PierLabel -Unique)).Count); Highlight = $false },
    [pscustomobject]@{ Cells = @("ExportedAtLocal", (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")); Highlight = $false }
)

$allPierRows = foreach ($row in $floorRows) {
    [pscustomobject]@{
        Cells = @(
            $row.PierLabel,
            $row.Story,
            $row.DesignType,
            (Round-OrNull -Value (Convert-LengthValueToFeet -Value $row.WallLength -LengthUnit $units.Length)),
            (Round-OrNull -Value (Convert-AreaValueToSquareInches -Value $row.RequiredSteelLeft -LengthUnit $units.Length)),
            (Round-OrNull -Value (Convert-AreaValueToSquareInches -Value $row.RequiredSteelRight -LengthUnit $units.Length)),
            (Format-NumberWithCommasOrEmpty -Value $row.LrfdEnvMaxM3),
            (Format-NumberWithCommasOrEmpty -Value $row.LrfdEnvMinM3),
            (Round-OrNull -Value $row.BoundaryZoneLeft -Digits 4),
            (Round-OrNull -Value $row.BoundaryZoneRight -Digits 4),
            $row.Message
        )
        Highlight = [bool]$row.HasIssue
    }
}

$sheetDefinitions = New-Object System.Collections.Generic.List[object]
$sheetDefinitions.Add([pscustomobject]@{
    Name = "Info"
    Headers = @("Field", "Value")
    Rows = $infoRows
}) | Out-Null

$sheetDefinitions.Add([pscustomobject]@{
    Name = "All Piers"
    Headers = $headers
    Rows = @($allPierRows)
}) | Out-Null

$pierLabels = @($floorRows | Select-Object -ExpandProperty PierLabel -Unique | Sort-Object { Get-PierSortKey -PierLabel $_ })
foreach ($pierLabel in $pierLabels) {
    $pierRows = $floorRows | Where-Object { $_.PierLabel -eq $pierLabel }
    $sheetRows = foreach ($row in $pierRows) {
        [pscustomobject]@{
            Cells = @(
                $row.Story,
                $row.DesignType,
                (Round-OrNull -Value (Convert-LengthValueToFeet -Value $row.WallLength -LengthUnit $units.Length)),
                (Round-OrNull -Value (Convert-AreaValueToSquareInches -Value $row.RequiredSteelLeft -LengthUnit $units.Length)),
                (Round-OrNull -Value (Convert-AreaValueToSquareInches -Value $row.RequiredSteelRight -LengthUnit $units.Length)),
                (Format-NumberWithCommasOrEmpty -Value $row.LrfdEnvMaxM3),
                (Format-NumberWithCommasOrEmpty -Value $row.LrfdEnvMinM3),
                (Round-OrNull -Value $row.BoundaryZoneLeft -Digits 4),
                (Round-OrNull -Value $row.BoundaryZoneRight -Digits 4),
                $row.Message
            )
            Highlight = [bool]$row.HasIssue
        }
    }

    $sheetDefinitions.Add([pscustomobject]@{
        Name = $pierLabel
        Headers = @(
            "Story",
            "DesignType",
            "WallLength_ft",
            ("AS_REQ_LEFT_{0}" -f $steelAreaUnits),
            ("AS_REQ_RIGHT_{0}" -f $steelAreaUnits),
            ("LRFDEnvMaxM3_{0}" -f $momentUnits),
            ("LRFDEnvMinM3_{0}" -f $momentUnits),
            ("BoundaryZoneLeft_{0}" -f (Get-LengthUnitLabel -LengthUnit $units.Length)),
            ("BoundaryZoneRight_{0}" -f (Get-LengthUnitLabel -LengthUnit $units.Length)),
            "Message"
        )
        Rows = @($sheetRows)
    }) | Out-Null
}

Write-XlsxWorkbook -Path $resolvedOutputPath -SheetDefinitions ($sheetDefinitions.ToArray())

if ($OpenWorkbook) {
    Start-Process -FilePath $resolvedOutputPath | Out-Null
}

$result = [pscustomobject]@{
    ProcessId = $process.Id
    ModelPath = $modelPath
    OutputPath = $resolvedOutputPath
    ForceUnits = $units.Force
    LengthUnits = $units.Length
    AreaUnits = $areaUnits
    RawStationRowCount = $rawRows.Count
    GroupedPierStoryCount = $reportRows.Count
    FloorSummaryRowCount = $floorRows.Count
    PierCount = $pierLabels.Count
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
