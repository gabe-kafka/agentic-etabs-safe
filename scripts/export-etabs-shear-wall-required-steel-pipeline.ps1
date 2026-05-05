param(
    [int]$EtabsPid,

    [string]$OutputDirectory,

    [string]$OutputWorkbookPath,

    [string]$ExpectedModelPath,

    [string[]]$Pier = @(),

    [string[]]$Story = @(),

    [switch]$OpenOutputDirectory,

    [switch]$OpenWorkbook,

    [switch]$NoCsv,

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

function Resolve-ExistingPathOrNull {
    param(
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if (Test-Path -LiteralPath $PathValue) {
        return (Resolve-Path -LiteralPath $PathValue).Path
    }

    return $PathValue
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
        default { throw "Unsupported length unit: $LengthUnit" }
    }
}

function Convert-AreaToSquareInches {
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

function Convert-LengthToFeet {
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

function Round-OrNull {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Digits = 3
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    return [Math]::Round([double]$Value, $Digits)
}

function Normalize-DesignMessage {
    param(
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message) -or $Message -eq "No Message") {
        return ""
    }

    return $Message.Trim()
}

function Add-Message {
    param(
        [System.Collections.Generic.List[string]]$Messages,
        [string]$Message
    )

    $normalized = Normalize-DesignMessage -Message $Message
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    if (-not $Messages.Contains($normalized)) {
        $Messages.Add($normalized) | Out-Null
    }
}

function Test-NameFilter {
    param(
        [string]$Value,
        [string[]]$Filters
    )

    if ($null -eq $Filters -or @($Filters).Count -eq 0) {
        return $true
    }

    foreach ($filter in $Filters) {
        if ($Value -like $filter) {
            return $true
        }
    }

    return $false
}

function Get-PierSortKey {
    param(
        [string]$PierLabel
    )

    if ($PierLabel -match "^([A-Za-z]+)(\d+)$") {
        return "{0}|{1:D8}" -f $Matches[1].ToUpperInvariant(), [int]$Matches[2]
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

function New-DefaultOutputDirectory {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path -Path (Join-Path -Path (Get-Location) -ChildPath "out\etabs-shear-wall-required-steel") -ChildPath $timestamp
}

function Export-Rows {
    param(
        [string]$Path,
        [object[]]$Rows
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        [System.IO.File]::WriteAllText($Path, "", ([System.Text.UTF8Encoding]::new($false)))
        return
    }

    @($Rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
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
        foreach ($row in @($Rows)) {
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

        $widths.Add([Math]::Min($maxLength + 2, 80)) | Out-Null
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

    $rowArray = if ($null -eq $Rows) { @() } else { @($Rows) }
    for ($rowIndex = 0; $rowIndex -lt @($rowArray).Count; $rowIndex++) {
        $row = @($rowArray)[$rowIndex]
        $styleId = if ($row.Highlight) { 2 } else { 0 }
        $cells = New-Object System.Collections.Generic.List[string]
        for ($columnIndex = 1; $columnIndex -le $Headers.Length; $columnIndex++) {
            $cells.Add((Get-CellXml -CellReference ("{0}{1}" -f (ConvertTo-ColumnName -Index $columnIndex), ($rowIndex + 2)) -Value $row.Cells[$columnIndex - 1] -StyleId $styleId)) | Out-Null
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

function New-SheetRow {
    param(
        [object[]]$Cells,
        [bool]$Highlight = $false
    )

    return [pscustomobject]@{
        Cells = @($Cells)
        Highlight = $Highlight
    }
}

function Convert-ObjectsToSheetRows {
    param(
        [object[]]$Rows,
        [string[]]$Headers,
        [switch]$HighlightMessages
    )

    $sheetRows = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Rows) {
        return $sheetRows.ToArray()
    }

    foreach ($row in @($Rows)) {
        $cells = foreach ($header in $Headers) {
            $property = $row.PSObject.Properties[$header]
            if ($null -ne $property) {
                $property.Value
            }
            else {
                ""
            }
        }

        $highlight = $false
        if ($HighlightMessages) {
            foreach ($messageProperty in @("Messages", "WarnMsg", "ErrMsg")) {
                $property = $row.PSObject.Properties[$messageProperty]
                if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                    $highlight = $true
                    break
                }
            }
        }

        $sheetRows.Add((New-SheetRow -Cells $cells -Highlight:$highlight)) | Out-Null
    }

    return $sheetRows.ToArray()
}

function New-WorkbookSheetDefinitions {
    param(
        [object[]]$InfoRows,
        [object[]]$EnvelopeRows,
        [object[]]$StationRows,
        [object[]]$WarningRows
    )

    $sheetDefinitions = New-Object System.Collections.Generic.List[object]
    $infoHeaders = @("Field", "Value")
    $envelopeHeaders = @(
        "Pier",
        "Story",
        "RequiredSteelLeft_in2",
        "LeftControlStation",
        "LeftControlPierLeg",
        "RequiredSteelRight_in2",
        "RightControlStation",
        "RightControlPierLeg",
        "MaxBZoneL_ft",
        "MaxBZoneR_ft",
        "MaxDCRatio",
        "DesignTypes",
        "PierSecTypes",
        "Messages"
    )
    $pierHeaders = @(
        "Story",
        "RequiredSteelLeft_in2",
        "LeftControlStation",
        "LeftControlPierLeg",
        "RequiredSteelRight_in2",
        "RightControlStation",
        "RightControlPierLeg",
        "MaxBZoneL_ft",
        "MaxBZoneR_ft",
        "MaxDCRatio",
        "DesignTypes",
        "PierSecTypes",
        "Messages"
    )
    $stationHeaders = @(
        "Pier",
        "Story",
        "Station",
        "PierLeg",
        "DesignType",
        "PierSecType",
        "EdgeBar",
        "EndBar",
        "BarSpacing_SourceLengthUnits",
        "RequiredReinfPercent",
        "CurrentReinfPercent",
        "DCRatio",
        "AsLeft_SourceLengthSquared",
        "AsRight_SourceLengthSquared",
        "AsLeft_in2",
        "AsRight_in2",
        "ShearAv_SourceLengthSquaredPerLength",
        "EdgeLeft_ft",
        "EdgeRight_ft",
        "BZoneL_ft",
        "BZoneR_ft",
        "BZoneLength_ft",
        "StressCompLeft",
        "StressCompRight",
        "StressLimitLeft",
        "StressLimitRight",
        "CDepthLeft_ft",
        "CLimitLeft_ft",
        "CDepthRight_ft",
        "CLimitRight_ft",
        "WarnMsg",
        "ErrMsg"
    )

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "Info"
        Headers = $infoHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $InfoRows -Headers $infoHeaders
    }) | Out-Null

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "All Piers"
        Headers = $envelopeHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $EnvelopeRows -Headers $envelopeHeaders -HighlightMessages
    }) | Out-Null

    foreach ($pierLabel in @($EnvelopeRows | Select-Object -ExpandProperty Pier -Unique | Sort-Object { Get-PierSortKey -PierLabel $_ })) {
        $pierRows = @($EnvelopeRows | Where-Object { $_.Pier -eq $pierLabel })
        $sheetDefinitions.Add([pscustomobject]@{
            Name = $pierLabel
            Headers = $pierHeaders
            Rows = Convert-ObjectsToSheetRows -Rows $pierRows -Headers $pierHeaders -HighlightMessages
        }) | Out-Null
    }

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "Raw Station Results"
        Headers = $stationHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $StationRows -Headers $stationHeaders -HighlightMessages
    }) | Out-Null

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "Warnings"
        Headers = $stationHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $WarningRows -Headers $stationHeaders -HighlightMessages
    }) | Out-Null

    return $sheetDefinitions.ToArray()
}

function Get-ShearWallStationRows {
    param(
        $SapModel,
        [string]$LengthUnit,
        [string[]]$PierFilter,
        [string[]]$StoryFilter
    )

    [string[]]$storyNames = @()
    [string[]]$pierLabels = @()
    [string[]]$stations = @()
    [string[]]$designTypes = @()
    [string[]]$pierSecTypes = @()
    [string[]]$edgeBars = @()
    [string[]]$endBars = @()
    [double[]]$barSpacing = @()
    [double[]]$reinfPercent = @()
    [double[]]$currPercent = @()
    [double[]]$dcRatio = @()
    [string[]]$pierLegs = @()
    [double[]]$legX1 = @()
    [double[]]$legY1 = @()
    [double[]]$legX2 = @()
    [double[]]$legY2 = @()
    [double[]]$edgeLeft = @()
    [double[]]$edgeRight = @()
    [double[]]$asLeft = @()
    [double[]]$asRight = @()
    [double[]]$shearAv = @()
    [double[]]$stressCompLeft = @()
    [double[]]$stressCompRight = @()
    [double[]]$stressLimitLeft = @()
    [double[]]$stressLimitRight = @()
    [double[]]$cDepthLeft = @()
    [double[]]$cLimitLeft = @()
    [double[]]$cDepthRight = @()
    [double[]]$cLimitRight = @()
    [double[]]$inelasticRotDemand = @()
    [double[]]$inelasticRotCapacity = @()
    [double[]]$normCompStress = @()
    [double[]]$normCompStressLimit = @()
    [double[]]$cDepth = @()
    [double[]]$bZoneL = @()
    [double[]]$bZoneR = @()
    [double[]]$bZoneLength = @()
    [string[]]$warnMsg = @()
    [string[]]$errMsg = @()

    Assert-Success (
        $SapModel.DesignShearWall.GetPierSummaryResults(
            [ref]$storyNames,
            [ref]$pierLabels,
            [ref]$stations,
            [ref]$designTypes,
            [ref]$pierSecTypes,
            [ref]$edgeBars,
            [ref]$endBars,
            [ref]$barSpacing,
            [ref]$reinfPercent,
            [ref]$currPercent,
            [ref]$dcRatio,
            [ref]$pierLegs,
            [ref]$legX1,
            [ref]$legY1,
            [ref]$legX2,
            [ref]$legY2,
            [ref]$edgeLeft,
            [ref]$edgeRight,
            [ref]$asLeft,
            [ref]$asRight,
            [ref]$shearAv,
            [ref]$stressCompLeft,
            [ref]$stressCompRight,
            [ref]$stressLimitLeft,
            [ref]$stressLimitRight,
            [ref]$cDepthLeft,
            [ref]$cLimitLeft,
            [ref]$cDepthRight,
            [ref]$cLimitRight,
            [ref]$inelasticRotDemand,
            [ref]$inelasticRotCapacity,
            [ref]$normCompStress,
            [ref]$normCompStressLimit,
            [ref]$cDepth,
            [ref]$bZoneL,
            [ref]$bZoneR,
            [ref]$bZoneLength,
            [ref]$warnMsg,
            [ref]$errMsg
        )
    ) "Get shear wall pier summary results"

    $rows = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $storyNames.Length; $index++) {
        if (-not (Test-NameFilter -Value $pierLabels[$index] -Filters $PierFilter)) {
            continue
        }

        if (-not (Test-NameFilter -Value $storyNames[$index] -Filters $StoryFilter)) {
            continue
        }

        $warn = Normalize-DesignMessage -Message $warnMsg[$index]
        $err = Normalize-DesignMessage -Message $errMsg[$index]
        $rows.Add([pscustomobject]@{
            Pier = $pierLabels[$index]
            Story = $storyNames[$index]
            Station = $stations[$index]
            PierLeg = $pierLegs[$index]
            DesignType = $designTypes[$index]
            PierSecType = $pierSecTypes[$index]
            EdgeBar = $edgeBars[$index]
            EndBar = $endBars[$index]
            BarSpacing_SourceLengthUnits = (Round-OrNull -Value $barSpacing[$index])
            RequiredReinfPercent = (Round-OrNull -Value $reinfPercent[$index] -Digits 6)
            CurrentReinfPercent = (Round-OrNull -Value $currPercent[$index] -Digits 6)
            DCRatio = (Round-OrNull -Value $dcRatio[$index] -Digits 6)
            AsLeft_SourceLengthSquared = (Round-OrNull -Value $asLeft[$index] -Digits 6)
            AsRight_SourceLengthSquared = (Round-OrNull -Value $asRight[$index] -Digits 6)
            AsLeft_in2 = (Round-OrNull -Value (Convert-AreaToSquareInches -Value $asLeft[$index] -LengthUnit $LengthUnit))
            AsRight_in2 = (Round-OrNull -Value (Convert-AreaToSquareInches -Value $asRight[$index] -LengthUnit $LengthUnit))
            ShearAv_SourceLengthSquaredPerLength = (Round-OrNull -Value $shearAv[$index] -Digits 6)
            EdgeLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $edgeLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            EdgeRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $edgeRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneL_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneL[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneR_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneR[$index] -LengthUnit $LengthUnit) -Digits 4)
            BZoneLength_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $bZoneLength[$index] -LengthUnit $LengthUnit) -Digits 4)
            StressCompLeft = (Round-OrNull -Value $stressCompLeft[$index])
            StressCompRight = (Round-OrNull -Value $stressCompRight[$index])
            StressLimitLeft = (Round-OrNull -Value $stressLimitLeft[$index])
            StressLimitRight = (Round-OrNull -Value $stressLimitRight[$index])
            CDepthLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cDepthLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            CLimitLeft_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cLimitLeft[$index] -LengthUnit $LengthUnit) -Digits 4)
            CDepthRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cDepthRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            CLimitRight_ft = (Round-OrNull -Value (Convert-LengthToFeet -Value $cLimitRight[$index] -LengthUnit $LengthUnit) -Digits 4)
            WarnMsg = $warn
            ErrMsg = $err
        }) | Out-Null
    }

    return $rows.ToArray()
}

function Get-EnvelopeRows {
    param(
        [object[]]$StationRows
    )

    $entries = @{}
    foreach ($row in $StationRows) {
        $key = "{0}|{1}" -f $row.Pier, $row.Story
        if (-not $entries.ContainsKey($key)) {
            $entries[$key] = [ordered]@{
                Pier = $row.Pier
                Story = $row.Story
                RequiredSteelLeft_in2 = $null
                LeftControlStation = ""
                LeftControlPierLeg = ""
                RequiredSteelRight_in2 = $null
                RightControlStation = ""
                RightControlPierLeg = ""
                MaxBZoneL_ft = $null
                MaxBZoneR_ft = $null
                MaxDCRatio = $null
                DesignTypes = New-Object System.Collections.Generic.List[string]
                PierSecTypes = New-Object System.Collections.Generic.List[string]
                Messages = New-Object System.Collections.Generic.List[string]
            }
        }

        $entry = $entries[$key]
        if (-not [string]::IsNullOrWhiteSpace($row.DesignType) -and -not $entry["DesignTypes"].Contains($row.DesignType)) {
            $entry["DesignTypes"].Add($row.DesignType) | Out-Null
        }

        if (-not [string]::IsNullOrWhiteSpace($row.PierSecType) -and -not $entry["PierSecTypes"].Contains($row.PierSecType)) {
            $entry["PierSecTypes"].Add($row.PierSecType) | Out-Null
        }

        Add-Message -Messages $entry["Messages"] -Message $row.WarnMsg
        Add-Message -Messages $entry["Messages"] -Message $row.ErrMsg

        if ($null -ne $row.AsLeft_in2 -and ($null -eq $entry["RequiredSteelLeft_in2"] -or $row.AsLeft_in2 -gt $entry["RequiredSteelLeft_in2"])) {
            $entry["RequiredSteelLeft_in2"] = $row.AsLeft_in2
            $entry["LeftControlStation"] = $row.Station
            $entry["LeftControlPierLeg"] = $row.PierLeg
        }

        if ($null -ne $row.AsRight_in2 -and ($null -eq $entry["RequiredSteelRight_in2"] -or $row.AsRight_in2 -gt $entry["RequiredSteelRight_in2"])) {
            $entry["RequiredSteelRight_in2"] = $row.AsRight_in2
            $entry["RightControlStation"] = $row.Station
            $entry["RightControlPierLeg"] = $row.PierLeg
        }

        if ($null -ne $row.BZoneL_ft -and ($null -eq $entry["MaxBZoneL_ft"] -or $row.BZoneL_ft -gt $entry["MaxBZoneL_ft"])) {
            $entry["MaxBZoneL_ft"] = $row.BZoneL_ft
        }

        if ($null -ne $row.BZoneR_ft -and ($null -eq $entry["MaxBZoneR_ft"] -or $row.BZoneR_ft -gt $entry["MaxBZoneR_ft"])) {
            $entry["MaxBZoneR_ft"] = $row.BZoneR_ft
        }

        if ($null -ne $row.DCRatio -and ($null -eq $entry["MaxDCRatio"] -or $row.DCRatio -gt $entry["MaxDCRatio"])) {
            $entry["MaxDCRatio"] = $row.DCRatio
        }
    }

    $rows = foreach ($entry in $entries.Values) {
        [pscustomobject]@{
            Pier = $entry["Pier"]
            Story = $entry["Story"]
            RequiredSteelLeft_in2 = $entry["RequiredSteelLeft_in2"]
            LeftControlStation = $entry["LeftControlStation"]
            LeftControlPierLeg = $entry["LeftControlPierLeg"]
            RequiredSteelRight_in2 = $entry["RequiredSteelRight_in2"]
            RightControlStation = $entry["RightControlStation"]
            RightControlPierLeg = $entry["RightControlPierLeg"]
            MaxBZoneL_ft = $entry["MaxBZoneL_ft"]
            MaxBZoneR_ft = $entry["MaxBZoneR_ft"]
            MaxDCRatio = $entry["MaxDCRatio"]
            DesignTypes = ($entry["DesignTypes"] -join " | ")
            PierSecTypes = ($entry["PierSecTypes"] -join " | ")
            Messages = ($entry["Messages"] -join " | ")
        }
    }

    return @($rows | Sort-Object @{ Expression = { Get-PierSortKey -PierLabel $_.Pier } }, @{ Expression = { Get-StorySortKey -StoryName $_.Story } })
}

function Get-WarningRows {
    param(
        [object[]]$StationRows
    )

    return @($StationRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.WarnMsg) -or -not [string]::IsNullOrWhiteSpace($_.ErrMsg)
        } | Sort-Object @{ Expression = { Get-PierSortKey -PierLabel $_.Pier } }, @{ Expression = { Get-StorySortKey -StoryName $_.Story } }, Station)
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
$resolvedExpectedModelPath = Resolve-ExistingPathOrNull -PathValue $ExpectedModelPath
if (-not [string]::IsNullOrWhiteSpace($resolvedExpectedModelPath) -and -not [string]::Equals($modelPath, $resolvedExpectedModelPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The live ETABS model does not match ExpectedModelPath. Live: '$modelPath'. Expected: '$resolvedExpectedModelPath'."
}

$units = Get-PresentUnits -SapModel $api.SapModel
$resolvedOutputDirectory = if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory
}
elseif (-not [string]::IsNullOrWhiteSpace($OutputWorkbookPath)) {
    $workbookDirectory = Split-Path -Parent $OutputWorkbookPath
    if ([string]::IsNullOrWhiteSpace($workbookDirectory)) { Get-Location } else { $workbookDirectory }
}
else {
    New-DefaultOutputDirectory
}
New-Item -ItemType Directory -Force -Path $resolvedOutputDirectory | Out-Null
$resolvedOutputDirectory = (Resolve-Path -LiteralPath $resolvedOutputDirectory).Path
$resolvedOutputWorkbookPath = if ([string]::IsNullOrWhiteSpace($OutputWorkbookPath)) {
    Join-Path -Path $resolvedOutputDirectory -ChildPath "shear-wall-design-required-steel.xlsx"
}
else {
    $OutputWorkbookPath
}

$modelFile = if (-not [string]::IsNullOrWhiteSpace($modelPath) -and (Test-Path -LiteralPath $modelPath -PathType Leaf)) { Get-Item -LiteralPath $modelPath } else { $null }
$modelLastWriteTime = if ($null -ne $modelFile) { $modelFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
$modelFileLength = if ($null -ne $modelFile) { $modelFile.Length } else { "" }
$stationRows = Get-ShearWallStationRows -SapModel $api.SapModel -LengthUnit $units.Length -PierFilter $Pier -StoryFilter $Story
$stationRowCount = @($stationRows).Count
if (-not $stationRows -or $stationRowCount -eq 0) {
    throw "No shear wall pier summary results were returned from the active model for the selected filters."
}

$envelopeRows = Get-EnvelopeRows -StationRows $stationRows
$warningRows = Get-WarningRows -StationRows $stationRows
$envelopeRowCount = @($envelopeRows).Count
$warningRowCount = @($warningRows).Count
$infoRows = @(
    [pscustomobject]@{ Field = "ExportedAtLocal"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") },
    [pscustomobject]@{ Field = "ProcessId"; Value = $process.Id },
    [pscustomobject]@{ Field = "MainWindowTitle"; Value = $process.MainWindowTitle },
    [pscustomobject]@{ Field = "ModelPath"; Value = $modelPath },
    [pscustomobject]@{ Field = "ModelLastWriteTime"; Value = $modelLastWriteTime },
    [pscustomobject]@{ Field = "ModelFileLength"; Value = $modelFileLength },
    [pscustomobject]@{ Field = "ExpectedModelPath"; Value = $resolvedExpectedModelPath },
    [pscustomobject]@{ Field = "OutputWorkbookPath"; Value = $resolvedOutputWorkbookPath },
    [pscustomobject]@{ Field = "ApiDllPath"; Value = $apiDllPath },
    [pscustomobject]@{ Field = "ForceUnits"; Value = $units.Force },
    [pscustomobject]@{ Field = "LengthUnits"; Value = $units.Length },
    [pscustomobject]@{ Field = "OutputSteelAreaUnits"; Value = "in^2" },
    [pscustomobject]@{ Field = "OutputLengthUnits"; Value = "ft" },
    [pscustomobject]@{ Field = "PierFilter"; Value = ($Pier -join ", ") },
    [pscustomobject]@{ Field = "StoryFilter"; Value = ($Story -join ", ") },
    [pscustomobject]@{ Field = "RawStationRowCount"; Value = $stationRowCount },
    [pscustomobject]@{ Field = "StoryEnvelopeRowCount"; Value = $envelopeRowCount },
    [pscustomobject]@{ Field = "WarningRowCount"; Value = $warningRowCount },
    [pscustomobject]@{ Field = "Source"; Value = "SapModel.DesignShearWall.GetPierSummaryResults" },
    [pscustomobject]@{ Field = "EnvelopeRule"; Value = "Per pier/story, use max AsLeft and max AsRight across ETABS station rows; keep controlling station columns." }
)

$infoPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "info.csv"
$stationPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "raw-station-results.csv"
$envelopePath = Join-Path -Path $resolvedOutputDirectory -ChildPath "story-envelope.csv"
$warningPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "warnings.csv"

$sheetDefinitions = New-WorkbookSheetDefinitions -InfoRows $infoRows -EnvelopeRows $envelopeRows -StationRows $stationRows -WarningRows $warningRows
Write-XlsxWorkbook -Path $resolvedOutputWorkbookPath -SheetDefinitions $sheetDefinitions

if (-not $NoCsv) {
    Export-Rows -Path $infoPath -Rows $infoRows
    Export-Rows -Path $stationPath -Rows $stationRows
    Export-Rows -Path $envelopePath -Rows $envelopeRows
    Export-Rows -Path $warningPath -Rows $warningRows
}

if ($OpenOutputDirectory) {
    Start-Process -FilePath $resolvedOutputDirectory | Out-Null
}

if ($OpenWorkbook) {
    Start-Process -FilePath $resolvedOutputWorkbookPath | Out-Null
}

$result = [pscustomobject]@{
    ProcessId = $process.Id
    ModelPath = $modelPath
    OutputWorkbookPath = (Resolve-Path -LiteralPath $resolvedOutputWorkbookPath).Path
    OutputDirectory = $resolvedOutputDirectory
    InfoCsv = if ($NoCsv) { $null } else { $infoPath }
    RawStationResultsCsv = if ($NoCsv) { $null } else { $stationPath }
    StoryEnvelopeCsv = if ($NoCsv) { $null } else { $envelopePath }
    WarningsCsv = if ($NoCsv) { $null } else { $warningPath }
    ForceUnits = $units.Force
    LengthUnits = $units.Length
    RawStationRowCount = $stationRowCount
    StoryEnvelopeRowCount = $envelopeRowCount
    WarningRowCount = $warningRowCount
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
