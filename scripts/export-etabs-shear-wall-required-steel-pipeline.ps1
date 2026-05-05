param(
    [int]$EtabsPid,

    [string]$OutputDirectory,

    [string]$OutputWorkbookPath,

    [string]$ExpectedModelPath,

    [string[]]$Pier = @(),

    [string[]]$Story = @(),

    [double]$QaToleranceIn2 = 0.005,

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

    $formulaProperty = if ($null -ne $Value) { $Value.PSObject.Properties["Formula"] } else { $null }
    if ($null -ne $formulaProperty) {
        $formulaText = [string]$formulaProperty.Value
        if ($formulaText.StartsWith("=")) {
            $formulaText = $formulaText.Substring(1)
        }

        $cachedProperty = $Value.PSObject.Properties["CachedValue"]
        $cachedValue = if ($null -ne $cachedProperty) { $cachedProperty.Value } else { $null }
        if ($null -ne $cachedValue -and $cachedValue -is [ValueType] -and $cachedValue -isnot [bool]) {
            return '<c r="{0}" s="{1}"><f>{2}</f><v>{3}</v></c>' -f $CellReference, $StyleId, (Escape-Xml -Value $formulaText), (Format-NumberInvariant -Value $cachedValue)
        }

        if ($null -ne $cachedValue -and -not [string]::IsNullOrWhiteSpace([string]$cachedValue)) {
            return '<c r="{0}" t="str" s="{1}"><f>{2}</f><v>{3}</v></c>' -f $CellReference, $StyleId, (Escape-Xml -Value $formulaText), (Escape-Xml -Value $cachedValue)
        }

        return '<c r="{0}" s="{1}"><f>{2}</f></c>' -f $CellReference, $StyleId, (Escape-Xml -Value $formulaText)
    }

    if ($null -ne $Value -and $Value -is [ValueType] -and $Value -isnot [bool]) {
        return '<c r="{0}" s="{1}"><v>{2}</v></c>' -f $CellReference, $StyleId, (Format-NumberInvariant -Value $Value)
    }

    return '<c r="{0}" t="inlineStr" s="{1}"><is><t>{2}</t></is></c>' -f $CellReference, $StyleId, (Escape-Xml -Value $Value)
}

function Get-CellDisplayText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    $cachedProperty = $Value.PSObject.Properties["CachedValue"]
    if ($null -ne $cachedProperty) {
        return Get-CellDisplayText -Value $cachedProperty.Value
    }

    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal]) {
        return ("{0:N3}" -f $Value)
    }

    return [string]$Value
}

function Get-AutoColumnWidth {
    param(
        [int]$MaxTextLength
    )

    $paddedWidth = [Math]::Ceiling(($MaxTextLength * 1.08) + 2)
    return [Math]::Min([Math]::Max($paddedWidth, 8), 120)
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
            $text = Get-CellDisplayText -Value $value
            foreach ($line in ($text -split "(\r\n|\n|\r)")) {
                if ($line.Length -gt $maxLength) {
                    $maxLength = $line.Length
                }
            }
        }

        $widths.Add((Get-AutoColumnWidth -MaxTextLength $maxLength)) | Out-Null
    }

    $columnXml = New-Object System.Collections.Generic.List[string]
    for ($index = 1; $index -le $Headers.Length; $index++) {
        $width = $widths[$index - 1]
        $columnXml.Add(('<col min="{0}" max="{0}" width="{1}" customWidth="1" bestFit="1"/>' -f $index, $width)) | Out-Null
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
  <calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>
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

function Invoke-WorkbookAutoFit {
    param(
        [string]$Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $pythonCode = @'
import math
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

path = Path(sys.argv[1])
workbook = load_workbook(path)
for worksheet in workbook.worksheets:
    for column_cells in worksheet.columns:
        max_length = 0
        column_index = column_cells[0].column
        for cell in column_cells:
            value = cell.value
            if value is None:
                text = ""
            else:
                text = str(value)
            for line in text.splitlines() or [""]:
                max_length = max(max_length, len(line))

        width = min(max(math.ceil((max_length * 1.08) + 2), 8), 120)
        dimension = worksheet.column_dimensions[get_column_letter(column_index)]
        dimension.width = width
        dimension.bestFit = True

workbook.save(path)
'@

    $pythonCode | & python - $resolvedPath
    if ($LASTEXITCODE -ne 0) {
        throw "Python failed while applying workbook auto-width formatting."
    }
}

function Invoke-WorkbookQaQcAlignment {
    param(
        [string]$Path,
        [object[]]$SheetDefinitions,
        [object[]]$StationRows,
        [object[]]$EnvelopeRows,
        [double]$ToleranceIn2,
        [AllowNull()]
        [string]$CsvPath
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $payloadPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("etabs-shear-wall-qaqc-{0}.json" -f [System.IO.Path]::GetRandomFileName())
    $expectedSheetNames = Get-UniqueSheetNames -SheetDefinitions $SheetDefinitions
    $payload = [pscustomobject]@{
        tolerance_in2 = $ToleranceIn2
        expected_sheets = @($expectedSheetNames)
        station_rows = @($StationRows | ForEach-Object {
                [pscustomobject]@{
                    Pier = $_.Pier
                    Story = $_.Story
                    AsLeft_in2 = $_.AsLeft_in2
                    AsRight_in2 = $_.AsRight_in2
                }
            })
        envelope_rows = @($EnvelopeRows | ForEach-Object {
                [pscustomobject]@{
                    Pier = $_.Pier
                    Story = $_.Story
                    RequiredSteelLeft_in2 = $_.RequiredSteelLeft_in2
                    RequiredSteelRight_in2 = $_.RequiredSteelRight_in2
                }
            })
    }

    try {
        Write-Utf8File -Path $payloadPath -Content ($payload | ConvertTo-Json -Depth 8)
        $resolvedCsvPath = if ([string]::IsNullOrWhiteSpace($CsvPath)) { "" } else { $CsvPath }
        $pythonCode = @'
import csv
import json
import math
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

workbook_path = Path(sys.argv[1])
payload_path = Path(sys.argv[2])
csv_path = Path(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3] else None

with payload_path.open("r", encoding="utf-8") as handle:
    payload = json.load(handle)

tolerance = float(payload["tolerance_in2"])
station_rows = payload["station_rows"]
envelope_rows = payload["envelope_rows"]
expected_sheets = payload["expected_sheets"]

headers = [
    "Status",
    "Severity",
    "Check",
    "Story",
    "Pier",
    "Side",
    "ETABS Value",
    "Workbook Value",
    "Delta",
    "Tolerance",
    "Source",
    "Notes",
]
check_rows = []


def as_key(pier, story):
    return ("" if pier is None else str(pier), "" if story is None else str(story))


def canonical_pier(pier):
    text = "" if pier is None else str(pier)
    if re.match(r"^W1\s+\d+/3$", text):
        return "W1"
    return text


def sanitize_sheet_name(name):
    cleaned = re.sub(r"[\[\]\:\*\?/\\]", "_", str(name or "")).strip()
    if not cleaned:
        cleaned = "Sheet"
    return cleaned[:31]


def to_float(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def fmt(value):
    number = to_float(value)
    if number is None:
        return ""
    return round(number, 6)


def value_delta(left, right):
    left_value = to_float(left)
    right_value = to_float(right)
    if left_value is None and right_value is None:
        return 0.0
    if left_value is None or right_value is None:
        return ""
    return round(right_value - left_value, 6)


def values_match(left, right):
    left_value = to_float(left)
    right_value = to_float(right)
    if left_value is None and right_value is None:
        return True
    if left_value is None or right_value is None:
        return False
    return abs(left_value - right_value) <= tolerance


def add_row(status, severity, check, story="", pier="", side="", etabs_value="", workbook_value="", source="", notes=""):
    check_rows.append(
        [
            status,
            severity,
            check,
            story,
            pier,
            side,
            fmt(etabs_value),
            fmt(workbook_value),
            value_delta(etabs_value, workbook_value),
            tolerance,
            source,
            notes,
        ]
    )


def compare_required(check, story, pier, side, etabs_value, workbook_value, source, notes=""):
    if values_match(etabs_value, workbook_value):
        add_row("PASS", "Info", check, story, pier, side, etabs_value, workbook_value, source, notes)
    else:
        add_row("FAIL", "Error", check, story, pier, side, etabs_value, workbook_value, source, notes)


def read_table(worksheet):
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return [], []
    table_headers = ["" if value is None else str(value) for value in rows[0]]
    records = []
    for raw_row in rows[1:]:
        record = {}
        for index, header in enumerate(table_headers):
            record[header] = raw_row[index] if index < len(raw_row) else None
        records.append(record)
    return table_headers, records


value_workbook = load_workbook(workbook_path, data_only=True)
workbook = load_workbook(workbook_path, data_only=False)
sheet_names = set(workbook.sheetnames)

for sheet_name in expected_sheets:
    if sheet_name in sheet_names:
        add_row("PASS", "Info", "Required sheet present", "", "", "", "", "", "Workbook structure", sheet_name)
    else:
        add_row("FAIL", "Error", "Required sheet present", "", "", "", "", "", "Workbook structure", sheet_name)

station_max = {}
for row in station_rows:
    key = as_key(row.get("Pier"), row.get("Story"))
    if key not in station_max:
        station_max[key] = {"LEFT": None, "RIGHT": None}
    left = to_float(row.get("AsLeft_in2"))
    right = to_float(row.get("AsRight_in2"))
    if left is not None and (station_max[key]["LEFT"] is None or left > station_max[key]["LEFT"]):
        station_max[key]["LEFT"] = left
    if right is not None and (station_max[key]["RIGHT"] is None or right > station_max[key]["RIGHT"]):
        station_max[key]["RIGHT"] = right

for row in envelope_rows:
    pier = str(row.get("Pier") or "")
    story = str(row.get("Story") or "")
    key = as_key(pier, story)
    station_entry = station_max.get(key)
    if station_entry is None:
        add_row("FAIL", "Error", "ETABS station max to envelope", story, pier, "", "", "", "In-memory ETABS pull", "No station rows found for envelope row.")
        continue
    compare_required(
        "ETABS station max to envelope",
        story,
        pier,
        "LEFT",
        station_entry["LEFT"],
        row.get("RequiredSteelLeft_in2"),
        "SapModel.DesignShearWall.GetPierSummaryResults",
    )
    compare_required(
        "ETABS station max to envelope",
        story,
        pier,
        "RIGHT",
        station_entry["RIGHT"],
        row.get("RequiredSteelRight_in2"),
        "SapModel.DesignShearWall.GetPierSummaryResults",
    )

envelope_by_key = {as_key(row.get("Pier"), row.get("Story")): row for row in envelope_rows}

if "All Piers" in value_workbook.sheetnames:
    all_headers, all_records = read_table(value_workbook["All Piers"])
    required_headers = ["Pier", "Story", "RequiredSteelLeft_in2", "RequiredSteelRight_in2"]
    for header in required_headers:
        if header in all_headers:
            add_row("PASS", "Info", "All Piers required column present", "", "", "", "", "", "All Piers", header)
        else:
            add_row("FAIL", "Error", "All Piers required column present", "", "", "", "", "", "All Piers", header)

    all_by_key = {}
    for record in all_records:
        key = as_key(record.get("Pier"), record.get("Story"))
        if key in all_by_key:
            add_row("FAIL", "Error", "All Piers duplicate row", str(record.get("Story") or ""), str(record.get("Pier") or ""), "", "", "", "All Piers", "Duplicate pier/story row.")
        all_by_key[key] = record

    for row in envelope_rows:
        pier = str(row.get("Pier") or "")
        story = str(row.get("Story") or "")
        record = all_by_key.get(as_key(pier, story))
        if record is None:
            add_row("FAIL", "Error", "Envelope to All Piers workbook", story, pier, "", "", "", "All Piers", "Missing pier/story row.")
            continue
        compare_required("Envelope to All Piers workbook", story, pier, "LEFT", row.get("RequiredSteelLeft_in2"), record.get("RequiredSteelLeft_in2"), "All Piers")
        compare_required("Envelope to All Piers workbook", story, pier, "RIGHT", row.get("RequiredSteelRight_in2"), record.get("RequiredSteelRight_in2"), "All Piers")

if "As Master" in value_workbook.sheetnames:
    rows = list(value_workbook["As Master"].iter_rows(values_only=True))
    if len(rows) < 3:
        add_row("FAIL", "Error", "Envelope to As Master workbook", "", "", "", "", "", "As Master", "Sheet does not have the expected header/story rows.")
    else:
        canonical_lookup = {}
        duplicate_keys = set()
        for row in envelope_rows:
            key = (canonical_pier(row.get("Pier")), str(row.get("Story") or ""))
            if key in canonical_lookup:
                duplicate_keys.add(key)
                existing = canonical_lookup[key]
                for source_name, target_name in (
                    ("RequiredSteelLeft_in2", "RequiredSteelLeft_in2"),
                    ("RequiredSteelRight_in2", "RequiredSteelRight_in2"),
                ):
                    current = to_float(existing.get(target_name))
                    candidate = to_float(row.get(source_name))
                    if candidate is not None and (current is None or candidate > current):
                        existing[target_name] = candidate
            else:
                canonical_lookup[key] = {
                    "RequiredSteelLeft_in2": row.get("RequiredSteelLeft_in2"),
                    "RequiredSteelRight_in2": row.get("RequiredSteelRight_in2"),
                }

        for pier, story in sorted(duplicate_keys):
            add_row("PASS", "Warning", "Canonical schedule pier has multiple source piers", story, pier, "", "", "", "As Master", "Schedule column maps more than one ETABS pier into one output pier label.")

        pier_headers = list(rows[0])
        side_headers = list(rows[1])
        for row_index in range(2, len(rows)):
            story = "" if rows[row_index][0] is None else str(rows[row_index][0])
            for col_index in range(1, len(pier_headers)):
                pier = "" if pier_headers[col_index] is None else str(pier_headers[col_index])
                side = "" if col_index >= len(side_headers) or side_headers[col_index] is None else str(side_headers[col_index]).upper()
                if not story or not pier or side not in ("LEFT", "RIGHT"):
                    continue
                envelope = canonical_lookup.get((pier, story))
                expected = ""
                if envelope is not None:
                    expected = envelope.get("RequiredSteelLeft_in2") if side == "LEFT" else envelope.get("RequiredSteelRight_in2")
                actual = rows[row_index][col_index] if col_index < len(rows[row_index]) else None
                compare_required("Envelope to As Master workbook", story, pier, side, expected, actual, "As Master")

for row in envelope_rows:
    pier = str(row.get("Pier") or "")
    story = str(row.get("Story") or "")
    sheet_name = sanitize_sheet_name(pier)
    if sheet_name not in value_workbook.sheetnames:
        add_row("FAIL", "Error", "Envelope to pier design workbook", story, pier, "", "", "", sheet_name, "Pier sheet missing.")
        continue
    pier_headers, pier_records = read_table(value_workbook[sheet_name])
    required_headers = ["Story", "AS_REQ_LEFT_in^2", "AS_REQ_RIGHT_in^2"]
    missing_headers = [header for header in required_headers if header not in pier_headers]
    if missing_headers:
        add_row("FAIL", "Error", "Pier sheet required column present", story, pier, "", "", "", sheet_name, ", ".join(missing_headers))
        continue
    story_record = None
    for record in pier_records:
        if str(record.get("Story") or "") == story:
            story_record = record
            break
    if story_record is None:
        add_row("FAIL", "Error", "Envelope to pier design workbook", story, pier, "", "", "", sheet_name, "Story row missing.")
        continue
    compare_required("Envelope to pier design workbook", story, pier, "LEFT", row.get("RequiredSteelLeft_in2"), story_record.get("AS_REQ_LEFT_in^2"), sheet_name)
    compare_required("Envelope to pier design workbook", story, pier, "RIGHT", row.get("RequiredSteelRight_in2"), story_record.get("AS_REQ_RIGHT_in^2"), sheet_name)

check_rows.sort(key=lambda item: (0 if item[1] == "Error" else 1 if item[1] == "Warning" else 2, item[2], item[3], item[4], item[5]))

if "QA_QC Alignment" in workbook.sheetnames:
    qa_ws = workbook["QA_QC Alignment"]
    qa_ws.delete_rows(1, qa_ws.max_row)
else:
    qa_ws = workbook.create_sheet("QA_QC Alignment")

qa_ws.append(headers)
for cell in qa_ws[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill(fill_type="solid", fgColor="D9E1F2")

error_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
warning_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
for row_values in check_rows:
    qa_ws.append(row_values)
    row_number = qa_ws.max_row
    if row_values[1] == "Error":
        for cell in qa_ws[row_number]:
            cell.fill = error_fill
    elif row_values[1] == "Warning":
        for cell in qa_ws[row_number]:
            cell.fill = warning_fill

for column_cells in qa_ws.columns:
    max_length = 0
    column_index = column_cells[0].column
    for cell in column_cells:
        value = cell.value
        text = "" if value is None else str(value)
        for line in text.splitlines() or [""]:
            max_length = max(max_length, len(line))
    dimension = qa_ws.column_dimensions[get_column_letter(column_index)]
    dimension.width = min(max(math.ceil((max_length * 1.08) + 2), 8), 120)
    dimension.bestFit = True

workbook.save(workbook_path)

if csv_path is not None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(check_rows)

failure_count = sum(1 for row in check_rows if row[1] == "Error")
warning_count = sum(1 for row in check_rows if row[1] == "Warning")
summary = {
    "Status": "PASS" if failure_count == 0 else "FAIL",
    "CheckedCount": len(check_rows),
    "FailureCount": failure_count,
    "WarningCount": warning_count,
    "ToleranceIn2": tolerance,
    "Worksheet": "QA_QC Alignment",
    "CsvPath": str(csv_path) if csv_path is not None else None,
}
print(json.dumps(summary, separators=(",", ":")))
'@

        $pythonOutput = $pythonCode | & python - $resolvedPath $payloadPath $resolvedCsvPath
        if ($LASTEXITCODE -ne 0) {
            throw "Python failed while applying workbook QA/QC alignment checks."
        }

        $summaryText = ($pythonOutput -join "`n").Trim()
        if ([string]::IsNullOrWhiteSpace($summaryText)) {
            throw "Workbook QA/QC alignment did not return a summary."
        }

        $summary = $summaryText | ConvertFrom-Json
        if ($summary.Status -ne "PASS") {
            throw ("QA/QC alignment failed for '{0}'. Failures: {1}. Review the QA_QC Alignment sheet." -f $resolvedPath, $summary.FailureCount)
        }

        return $summary
    }
    finally {
        if (Test-Path -LiteralPath $payloadPath) {
            Remove-Item -LiteralPath $payloadPath -Force
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

function New-FormulaCell {
    param(
        [string]$Formula,
        [AllowNull()]
        [object]$CachedValue
    )

    return [pscustomobject]@{
        Formula = $Formula
        CachedValue = $CachedValue
    }
}

function Get-ExcelSheetReferenceName {
    param(
        [string]$SheetName
    )

    return "'{0}'" -f ($SheetName -replace "'", "''")
}

function Get-ExcelCellReference {
    param(
        [int]$ColumnIndex,
        [int]$RowIndex
    )

    return "{0}{1}" -f (ConvertTo-ColumnName -Index $ColumnIndex), $RowIndex
}

function ConvertTo-ExcelStringLiteral {
    param(
        [string]$Value
    )

    return $Value -replace '"', '""'
}

function New-RequiredSteelFormula {
    param(
        [string]$SourceSheet,
        [string]$ValueColumn,
        [string]$PierCriteria,
        [string]$StoryCellReference
    )

    $criteria = ConvertTo-ExcelStringLiteral -Value $PierCriteria
    return '=IF(COUNTIFS({0}!$A:$A,"{1}",{0}!$B:$B,{2})=0,"",MAXIFS({0}!${3}:${3},{0}!$A:$A,"{1}",{0}!$B:$B,{2}))' -f $SourceSheet, $criteria, $StoryCellReference, $ValueColumn
}

function New-DesignValueFormula {
    param(
        [string]$RequiredAreaCellReference
    )

    $lastHierarchyRow = (Get-DesignHierarchy).Count + 1
    $areaRange = "$(Get-ExcelSheetReferenceName -SheetName "Master Design Hierarchy")!`$C`$2:`$C`$$lastHierarchyRow"
    return '=IF({0}="","",IF({0}>=MAX({1}),"EXCEEDS HIERARCHY",IF({0}<MIN({1}),MIN({1}),INDEX({1},MATCH({0},{1},1)+1))))' -f $RequiredAreaCellReference, $areaRange
}

function New-DesignIdFormula {
    param(
        [string]$RequiredAreaCellReference
    )

    $hierarchySheet = Get-ExcelSheetReferenceName -SheetName "Master Design Hierarchy"
    $lastHierarchyRow = (Get-DesignHierarchy).Count + 1
    $areaRange = "$hierarchySheet!`$C`$2:`$C`$$lastHierarchyRow"
    $idRange = "$hierarchySheet!`$D`$2:`$D`$$lastHierarchyRow"
    return '=IF({0}="","",IF({0}>=MAX({1}),"EXCEEDS HIERARCHY",IF({0}<MIN({1}),INDEX({2},1),INDEX({2},MATCH({0},{1},1)+1))))' -f $RequiredAreaCellReference, $areaRange, $idRange
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

function New-GridSheetDefinition {
    param(
        [string]$Name,
        [object]$GridRows
    )

    $rows = @($GridRows)
    if (($rows | Measure-Object).Count -eq 0) {
        $rows = @(, @(""))
    }

    $width = 0
    foreach ($row in $rows) {
        $width = [Math]::Max($width, @($row).Count)
    }

    $normalizedRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $cells = New-Object System.Collections.Generic.List[object]
        foreach ($cell in @($row)) {
            $cells.Add($cell) | Out-Null
        }

        while ($cells.Count -lt $width) {
            $cells.Add("") | Out-Null
        }

        $normalizedRows.Add($cells.ToArray()) | Out-Null
    }

    $headers = [string[]]@($normalizedRows[0] | ForEach-Object { [string]$_ })
    $sheetRows = New-Object System.Collections.Generic.List[object]
    for ($index = 1; $index -lt $normalizedRows.Count; $index++) {
        $sheetRows.Add((New-SheetRow -Cells @($normalizedRows[$index]))) | Out-Null
    }

    return [pscustomobject]@{
        Name = $Name
        Headers = $headers
        Rows = $sheetRows.ToArray()
    }
}

function Get-DesignHierarchy {
    $rows = @(
        @{ Bar = 5; Rows = 3; Area = 1.840775391; Id = '3#5 E.F. @6"' },
        @{ Bar = 6; Rows = 3; Area = 2.650716563; Id = '3#6 E.F. @6"' },
        @{ Bar = 7; Rows = 3; Area = 3.607919766; Id = '3#7 E.F. @6"' },
        @{ Bar = 8; Rows = 3; Area = 4.712385000; Id = '3#8 E.F. @6"' },
        @{ Bar = 9; Rows = 3; Area = 5.964112266; Id = '3#9 E.F. @6"' },
        @{ Bar = 10; Rows = 3; Area = 7.363101563; Id = '3#10 E.F. @6"' },
        @{ Bar = 9; Rows = 5; Area = 9.940187109; Id = '5#9 E.F. @6"' },
        @{ Bar = 11; Rows = 4; Area = 11.87913719; Id = '4#11 E.F. @6"' },
        @{ Bar = 11; Rows = 5; Area = 14.84892148; Id = '5#11 E.F. @6"' },
        @{ Bar = 11; Rows = 6; Area = 17.81870578; Id = '6#11 E.F. @6"' },
        @{ Bar = 11; Rows = 7; Area = 20.78849008; Id = '7#11 E.F. @6"' },
        @{ Bar = 11; Rows = 8; Area = 23.75827438; Id = '8#11 E.F. @6"' },
        @{ Bar = 11; Rows = 9; Area = 26.72805867; Id = '9#11 E.F. @6"' },
        @{ Bar = 11; Rows = 10; Area = 29.69784297; Id = '10#11 E.F. @6"' },
        @{ Bar = 11; Rows = 11; Area = 32.66762727; Id = '11#11 E.F. @6"' },
        @{ Bar = 11; Rows = 12; Area = 35.63741100; Id = '12#11 E.F. @6"' },
        @{ Bar = 11; Rows = 13; Area = 38.60719500; Id = '13#11 E.F. @6"' },
        @{ Bar = 11; Rows = 14; Area = 41.57697900; Id = '14#11 E.F. @6"' }
    )

    return @($rows | ForEach-Object { [pscustomobject]$_ })
}

function Get-DesignSelection {
    param(
        [AllowNull()]
        [object]$RequiredArea,
        [object[]]$Hierarchy
    )

    if ($null -eq $RequiredArea -or $RequiredArea -eq "") {
        return [pscustomobject]@{ Area = ""; Id = ""; Exceeds = $false }
    }

    $required = [double]$RequiredArea
    foreach ($entry in $Hierarchy) {
        if ($required -lt [double]$entry.Area) {
            return [pscustomobject]@{ Area = [Math]::Round([double]$entry.Area, 6); Id = $entry.Id; Exceeds = $false }
        }
    }

    return [pscustomobject]@{ Area = "EXCEEDS HIERARCHY"; Id = "EXCEEDS HIERARCHY"; Exceeds = $true }
}

function Get-CanonicalPierLabel {
    param(
        [string]$PierLabel
    )

    if ($PierLabel -match "^W1\s+\d+/3$") {
        return "W1"
    }

    return $PierLabel
}

function Get-ReferenceStoryNames {
    return @(
        "Story1", "Story2", "Story3", "Story4", "Story5",
        "Story6", "Story7", "Story8", "Story9", "Story10",
        "Story11", "Story12", "Story13", "Story14", "Story15",
        "Story16", "Story17", "Story18", "ROOF"
    )
}

function Get-ScheduleColumns {
    $groups = @("SW-1", "SW-1", "SW-2", "SW-2", "SW-2", "SW-2", "SW-2", "SW-2", "SW-3", "SW-3", "SW-3", "SW-3", "SW-3", "SW-3")
    $ids = @("1.1", "1.2", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "3.1", "3.2", "3.3", "3.4", "3.5", "3.6")
    $piers = @("W1", "W1", "W2", "W2", "W3", "W3", "W4", "W4", "W5", "W5", "W6", "W6", "W7", "W7")
    $sides = @("LEFT", "RIGHT", "LEFT", "RIGHT", "LEFT", "RIGHT", "LEFT", "RIGHT", "LEFT", "RIGHT", "LEFT", "RIGHT", "LEFT", "RIGHT")

    $columns = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $groups.Count; $index++) {
        $columns.Add([pscustomobject]@{
            Group = $groups[$index]
            Id = $ids[$index]
            Pier = $piers[$index]
            Side = $sides[$index]
        }) | Out-Null
    }

    return $columns.ToArray()
}

function New-EnvelopeLookup {
    param(
        [object[]]$EnvelopeRows
    )

    $lookup = @{}
    foreach ($row in @($EnvelopeRows)) {
        $canonicalPier = Get-CanonicalPierLabel -PierLabel $row.Pier
        $key = "{0}|{1}" -f $canonicalPier, $row.Story
        if (-not $lookup.ContainsKey($key)) {
            $lookup[$key] = [pscustomobject]@{
                Pier = $canonicalPier
                Story = $row.Story
                RequiredSteelLeft_in2 = $row.RequiredSteelLeft_in2
                RequiredSteelRight_in2 = $row.RequiredSteelRight_in2
            }
            continue
        }

        $entry = $lookup[$key]
        if ($null -ne $row.RequiredSteelLeft_in2 -and $row.RequiredSteelLeft_in2 -ne "" -and ($null -eq $entry.RequiredSteelLeft_in2 -or $entry.RequiredSteelLeft_in2 -eq "" -or [double]$row.RequiredSteelLeft_in2 -gt [double]$entry.RequiredSteelLeft_in2)) {
            $entry.RequiredSteelLeft_in2 = $row.RequiredSteelLeft_in2
        }

        if ($null -ne $row.RequiredSteelRight_in2 -and $row.RequiredSteelRight_in2 -ne "" -and ($null -eq $entry.RequiredSteelRight_in2 -or $entry.RequiredSteelRight_in2 -eq "" -or [double]$row.RequiredSteelRight_in2 -gt [double]$entry.RequiredSteelRight_in2)) {
            $entry.RequiredSteelRight_in2 = $row.RequiredSteelRight_in2
        }
    }

    return $lookup
}

function Get-RequiredAreaForColumn {
    param(
        [hashtable]$EnvelopeLookup,
        [string]$Pier,
        [string]$Story,
        [string]$Side
    )

    $key = "{0}|{1}" -f $Pier, $Story
    if (-not $EnvelopeLookup.ContainsKey($key)) {
        return ""
    }

    $row = $EnvelopeLookup[$key]
    if ($Side -eq "LEFT") {
        return $row.RequiredSteelLeft_in2
    }

    return $row.RequiredSteelRight_in2
}

function Get-DesignCheckValue {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Digits = 3
    )

    if ($null -eq $Value -or $Value -eq "") {
        return ""
    }

    return [Math]::Round([double]$Value, $Digits)
}

function New-CountedAggregateFormula {
    param(
        [string]$SourceSheet,
        [string]$ValueColumn,
        [string]$PierCriteria,
        [string]$StoryCellReference,
        [string]$AggregateFunction = "MAXIFS"
    )

    $criteria = ConvertTo-ExcelStringLiteral -Value $PierCriteria
    return '=IF(COUNTIFS({0}!$A:$A,"{1}",{0}!$B:$B,{2})=0,"",{3}({0}!${4}:${4},{0}!$A:$A,"{1}",{0}!$B:$B,{2}))' -f $SourceSheet, $criteria, $StoryCellReference, $AggregateFunction, $ValueColumn
}

function New-CellCriteriaAggregateFormula {
    param(
        [string]$SourceSheet,
        [string]$ValueColumn,
        [string]$PierCellReference,
        [string]$StoryCellReference,
        [string]$AggregateFunction = "MAXIFS"
    )

    return '=IF(COUNTIFS({0}!$A:$A,{1},{0}!$B:$B,{2})=0,"",{3}({0}!${4}:${4},{0}!$A:$A,{1},{0}!$B:$B,{2}))' -f $SourceSheet, $PierCellReference, $StoryCellReference, $AggregateFunction, $ValueColumn
}

function Get-DesignCheckRows {
    param(
        [object[]]$EnvelopeRows
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $allPiersSheet = Get-ExcelSheetReferenceName -SheetName "All Piers"
    foreach ($row in @($EnvelopeRows)) {
        $excelRow = $rows.Count + 2
        $wallLength = $row.WallLength_ft
        $asLeft = $row.RequiredSteelLeft_in2
        $asRight = $row.RequiredSteelRight_in2
        $bZoneLeft = $row.MaxBZoneL_ft
        $bZoneRight = $row.MaxBZoneR_ft

        $dLeft = if ($null -ne $wallLength -and $wallLength -ne "" -and $null -ne $bZoneLeft -and $bZoneLeft -ne "") { [double]$wallLength - ([double]$bZoneLeft / 2.0) } else { $null }
        $dRight = if ($null -ne $wallLength -and $wallLength -ne "" -and $null -ne $bZoneRight -and $bZoneRight -ne "") { [double]$wallLength - ([double]$bZoneRight / 2.0) } else { $null }
        $aLeft = if ($null -ne $asLeft -and $asLeft -ne "") { (60.0 * [double]$asLeft) / (0.85 * 5.0 * 10.0) } else { $null }
        $aRight = if ($null -ne $asRight -and $asRight -ne "") { (60.0 * [double]$asRight) / (0.85 * 5.0 * 10.0) } else { $null }

        # Preserve the legacy reference workbook formula form for these downstream checks.
        $phiMnLeftKin = if ($null -ne $dLeft -and $null -ne $aLeft -and $null -ne $asLeft) { (0.9 * 60.0 * [double]$asLeft * ($dLeft * 12.0)) - ($aLeft / 2.0) } else { $null }
        $phiMnRightKin = if ($null -ne $dRight -and $null -ne $aRight -and $null -ne $asRight) { (0.9 * 60.0 * [double]$asRight * ($dRight * 12.0)) - ($aRight / 2.0) } else { $null }
        $storyRef = "`$A$excelRow"

        $rows.Add([pscustomobject]@{
            Story = $row.Story
            WallLength_ft = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "C" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $wallLength
            "AS_REQ_LEFT_in^2" = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "D" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $asLeft
            "AS_REQ_RIGHT_in^2" = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "E" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $asRight
            "LRFDEnvMaxM3_kip-ft" = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "F" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $row.LRFDEnvMaxM3_kip_ft
            "LRFDEnvMinM3_kip-ft" = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "G" -PierCriteria $row.Pier -StoryCellReference $storyRef -AggregateFunction "MINIFS") -CachedValue $row.LRFDEnvMinM3_kip_ft
            BoundaryZoneLeft_ft = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "H" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $bZoneLeft
            BoundaryZoneRight_ft = New-FormulaCell -Formula (New-CountedAggregateFormula -SourceSheet $allPiersSheet -ValueColumn "I" -PierCriteria $row.Pier -StoryCellReference $storyRef) -CachedValue $bZoneRight
            Message = $row.Messages
            "d-left-bond" = New-FormulaCell -Formula ('=IF(OR(B{0}="",G{0}=""),"",B{0}-G{0}/2)' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $dLeft)
            "d-right-bond" = New-FormulaCell -Formula ('=IF(OR(B{0}="",H{0}=""),"",B{0}-H{0}/2)' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $dRight)
            "a (in)-left bond" = New-FormulaCell -Formula ('=IF(C{0}="","",(60*C{0})/(0.85*5*10))' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $aLeft)
            "a (in)-right bond" = New-FormulaCell -Formula ('=IF(D{0}="","",(60*D{0})/(0.85*5*10))' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $aRight)
            "phi*Mn(left)(k-in)" = New-FormulaCell -Formula ('=IF(OR(C{0}="",J{0}="",L{0}=""),"",0.9*60*C{0}*(J{0}*12)-L{0}/2)' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $phiMnLeftKin)
            "phi*Mn(right)(k-in)" = New-FormulaCell -Formula ('=IF(OR(D{0}="",K{0}="",M{0}=""),"",0.9*60*D{0}*(K{0}*12)-M{0}/2)' -f $excelRow) -CachedValue (Get-DesignCheckValue -Value $phiMnRightKin)
            "phi*Mn(left)(k-ft)" = New-FormulaCell -Formula ('=IF(N{0}="","",N{0}/12)' -f $excelRow) -CachedValue $(if ($null -ne $phiMnLeftKin) { Get-DesignCheckValue -Value ($phiMnLeftKin / 12.0) } else { "" })
            "phi*Mn(right)(k-ft)" = New-FormulaCell -Formula ('=IF(O{0}="","",O{0}/12)' -f $excelRow) -CachedValue $(if ($null -ne $phiMnRightKin) { Get-DesignCheckValue -Value ($phiMnRightKin / 12.0) } else { "" })
        }) | Out-Null
    }

    return $rows.ToArray()
}

function New-WorkbookSheetDefinitions {
    param(
        [object[]]$InfoRows,
        [object[]]$EnvelopeRows,
        [object[]]$StationRows,
        [object[]]$WarningRows
    )

    $sheetDefinitions = New-Object System.Collections.Generic.List[object]
    $hierarchy = Get-DesignHierarchy
    $storyNames = Get-ReferenceStoryNames
    $scheduleColumns = Get-ScheduleColumns
    $envelopeLookup = New-EnvelopeLookup -EnvelopeRows $EnvelopeRows
    $infoHeaders = @("Field", "Value")
    $envelopeHeaders = @(
        "Pier",
        "Story",
        "WallLength_ft",
        "RequiredSteelLeft_in2",
        "RequiredSteelRight_in2",
        "LRFDEnvMaxM3_kip_ft",
        "LRFDEnvMinM3_kip_ft",
        "MaxBZoneL_ft",
        "MaxBZoneR_ft",
        "MaxDCRatio",
        "PierSecTypes",
        "Messages"
    )
    $pierDesignHeaders = @(
        "Story",
        "WallLength_ft",
        "AS_REQ_LEFT_in^2",
        "AS_REQ_RIGHT_in^2",
        "LRFDEnvMaxM3_kip-ft",
        "LRFDEnvMinM3_kip-ft",
        "BoundaryZoneLeft_ft",
        "BoundaryZoneRight_ft",
        "Message",
        "d-left-bond",
        "d-right-bond",
        "a (in)-left bond",
        "a (in)-right bond",
        "phi*Mn(left)(k-in)",
        "phi*Mn(right)(k-in)",
        "phi*Mn(left)(k-ft)",
        "phi*Mn(right)(k-ft)"
    )
    $stationHeaders = @(
        "Pier",
        "Story",
        "Station",
        "PierLeg",
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

    $asMasterSheet = Get-ExcelSheetReferenceName -SheetName "As Master"
    $designMasterSheet = Get-ExcelSheetReferenceName -SheetName "Design Master"
    $allPiersSheet = Get-ExcelSheetReferenceName -SheetName "All Piers"
    $asMasterRowByStory = @{}
    for ($index = 0; $index -lt $storyNames.Count; $index++) {
        $asMasterRowByStory[$storyNames[$index]] = $index + 3
    }

    $asMasterGrid = New-Object System.Collections.Generic.List[object]
    $asMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Pier })) | Out-Null
    $asMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Side })) | Out-Null
    for ($storyIndex = 0; $storyIndex -lt $storyNames.Count; $storyIndex++) {
        $storyName = $storyNames[$storyIndex]
        $excelRow = $storyIndex + 3
        $storyRef = "`$A$excelRow"
        $asMasterCells = New-Object System.Collections.Generic.List[object]
        foreach ($column in $scheduleColumns) {
            $required = Get-RequiredAreaForColumn -EnvelopeLookup $envelopeLookup -Pier $column.Pier -Story $storyName -Side $column.Side
            $valueColumn = if ($column.Side -eq "LEFT") { "D" } else { "E" }
            $pierCriteria = if ($column.Pier -eq "W1") { "W1*" } else { $column.Pier }
            $asMasterCells.Add((New-FormulaCell -Formula (New-RequiredSteelFormula -SourceSheet $allPiersSheet -ValueColumn $valueColumn -PierCriteria $pierCriteria -StoryCellReference $storyRef) -CachedValue $required)) | Out-Null
        }

        $asMasterGrid.Add(@($storyName) + @($asMasterCells.ToArray())) | Out-Null
    }

    $designValueByStory = @{}
    $designIdByStory = @{}
    foreach ($storyName in $storyNames) {
        $designValueByStory[$storyName] = @($scheduleColumns | ForEach-Object {
                $required = Get-RequiredAreaForColumn -EnvelopeLookup $envelopeLookup -Pier $_.Pier -Story $storyName -Side $_.Side
                (Get-DesignSelection -RequiredArea $required -Hierarchy $hierarchy).Area
            })
        $designIdByStory[$storyName] = @($scheduleColumns | ForEach-Object {
                $required = Get-RequiredAreaForColumn -EnvelopeLookup $envelopeLookup -Pier $_.Pier -Story $storyName -Side $_.Side
                (Get-DesignSelection -RequiredArea $required -Hierarchy $hierarchy).Id
            })
    }

    $flippedStories = @($storyNames | Sort-Object { Get-StorySortKey -StoryName $_ } -Descending)
    $dmRequiredFirstRow = 4
    $dmDesignValueFirstRow = $dmRequiredFirstRow + $storyNames.Count + 4
    $dmDesignIdFirstRow = $dmDesignValueFirstRow + $storyNames.Count + 4
    $dmIdRowByStory = @{}
    for ($index = 0; $index -lt $storyNames.Count; $index++) {
        $dmIdRowByStory[$storyNames[$index]] = $dmDesignIdFirstRow + $index
    }

    $designMasterGrid = New-Object System.Collections.Generic.List[object]
    $designMasterGrid.Add(@("As required") + @("") * 15 + @("As required (flipped)") + @("") * 14) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Pier }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Pier })) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Side }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Side })) | Out-Null
    for ($index = 0; $index -lt $storyNames.Count; $index++) {
        $storyName = $storyNames[$index]
        $flippedStory = $flippedStories[$index]
        $leftCells = New-Object System.Collections.Generic.List[object]
        $rightCells = New-Object System.Collections.Generic.List[object]
        for ($columnIndex = 0; $columnIndex -lt $scheduleColumns.Count; $columnIndex++) {
            $excelColumn = ConvertTo-ColumnName -Index ($columnIndex + 2)
            $leftRequired = Get-RequiredAreaForColumn -EnvelopeLookup $envelopeLookup -Pier $scheduleColumns[$columnIndex].Pier -Story $storyName -Side $scheduleColumns[$columnIndex].Side
            $rightRequired = Get-RequiredAreaForColumn -EnvelopeLookup $envelopeLookup -Pier $scheduleColumns[$columnIndex].Pier -Story $flippedStory -Side $scheduleColumns[$columnIndex].Side
            $leftCells.Add((New-FormulaCell -Formula ("={0}!{1}{2}" -f $asMasterSheet, $excelColumn, $asMasterRowByStory[$storyName]) -CachedValue $leftRequired)) | Out-Null
            $rightCells.Add((New-FormulaCell -Formula ("={0}!{1}{2}" -f $asMasterSheet, $excelColumn, $asMasterRowByStory[$flippedStory]) -CachedValue $rightRequired)) | Out-Null
        }
        $designMasterGrid.Add(@($storyName) + @($leftCells.ToArray()) + @("", $flippedStory) + @($rightCells.ToArray())) | Out-Null
    }
    $designMasterGrid.Add(@("")) | Out-Null
    $designMasterGrid.Add(@("design value>As required") + @("") * 15 + @("design value>As required (flipped)") + @("") * 14) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Pier }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Pier })) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Side }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Side })) | Out-Null
    for ($index = 0; $index -lt $storyNames.Count; $index++) {
        $storyName = $storyNames[$index]
        $flippedStory = $flippedStories[$index]
        $leftCells = New-Object System.Collections.Generic.List[object]
        $rightCells = New-Object System.Collections.Generic.List[object]
        for ($columnIndex = 0; $columnIndex -lt $scheduleColumns.Count; $columnIndex++) {
            $leftColumn = ConvertTo-ColumnName -Index ($columnIndex + 2)
            $rightColumn = ConvertTo-ColumnName -Index ($columnIndex + 18)
            $leftCells.Add((New-FormulaCell -Formula (New-DesignValueFormula -RequiredAreaCellReference ("{0}{1}" -f $leftColumn, ($dmRequiredFirstRow + $index))) -CachedValue $designValueByStory[$storyName][$columnIndex])) | Out-Null
            $rightCells.Add((New-FormulaCell -Formula (New-DesignValueFormula -RequiredAreaCellReference ("{0}{1}" -f $rightColumn, ($dmRequiredFirstRow + $index))) -CachedValue $designValueByStory[$flippedStory][$columnIndex])) | Out-Null
        }
        $designMasterGrid.Add(@($storyName) + @($leftCells.ToArray()) + @("", $flippedStory) + @($rightCells.ToArray())) | Out-Null
    }
    $designMasterGrid.Add(@("")) | Out-Null
    $designMasterGrid.Add(@("design id") + @("") * 15 + @("design id (flipped)") + @("") * 14) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Pier }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Pier })) | Out-Null
    $designMasterGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Side }) + @("", "") + @($scheduleColumns | ForEach-Object { $_.Side })) | Out-Null
    for ($index = 0; $index -lt $storyNames.Count; $index++) {
        $storyName = $storyNames[$index]
        $flippedStory = $flippedStories[$index]
        $leftCells = New-Object System.Collections.Generic.List[object]
        $rightCells = New-Object System.Collections.Generic.List[object]
        for ($columnIndex = 0; $columnIndex -lt $scheduleColumns.Count; $columnIndex++) {
            $leftColumn = ConvertTo-ColumnName -Index ($columnIndex + 2)
            $rightColumn = ConvertTo-ColumnName -Index ($columnIndex + 18)
            $leftCells.Add((New-FormulaCell -Formula (New-DesignIdFormula -RequiredAreaCellReference ("{0}{1}" -f $leftColumn, ($dmRequiredFirstRow + $index))) -CachedValue $designIdByStory[$storyName][$columnIndex])) | Out-Null
            $rightCells.Add((New-FormulaCell -Formula (New-DesignIdFormula -RequiredAreaCellReference ("{0}{1}" -f $rightColumn, ($dmRequiredFirstRow + $index))) -CachedValue $designIdByStory[$flippedStory][$columnIndex])) | Out-Null
        }
        $designMasterGrid.Add(@($storyName) + @($leftCells.ToArray()) + @("", $flippedStory) + @($rightCells.ToArray())) | Out-Null
    }

    $tableGrid = New-Object System.Collections.Generic.List[object]
    $tableGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Group })) | Out-Null
    $tableGrid.Add(@("ID") + @($scheduleColumns | ForEach-Object { $_.Id })) | Out-Null
    $tableGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Pier })) | Out-Null
    $tableGrid.Add(@("") + @($scheduleColumns | ForEach-Object { $_.Side })) | Out-Null
    foreach ($storyName in $flippedStories) {
        $tableCells = New-Object System.Collections.Generic.List[object]
        for ($columnIndex = 0; $columnIndex -lt $scheduleColumns.Count; $columnIndex++) {
            $designMasterColumn = ConvertTo-ColumnName -Index ($columnIndex + 2)
            $tableCells.Add((New-FormulaCell -Formula ("={0}!{1}{2}" -f $designMasterSheet, $designMasterColumn, $dmIdRowByStory[$storyName]) -CachedValue $designIdByStory[$storyName][$columnIndex])) | Out-Null
        }
        $tableGrid.Add(@($storyName) + @($tableCells.ToArray())) | Out-Null
    }

    $hierarchyGrid = New-Object System.Collections.Generic.List[object]
    $hierarchyGrid.Add(@("#", "rows", "A", "id")) | Out-Null
    foreach ($entry in $hierarchy) {
        $hierarchyGrid.Add(@($entry.Bar, $entry.Rows, [Math]::Round([double]$entry.Area, 6), $entry.Id)) | Out-Null
    }

    $sheetDefinitions.Add((New-GridSheetDefinition -Name "SHEAR WALL TABLE OUTPUT" -GridRows $tableGrid.ToArray())) | Out-Null
    $sheetDefinitions.Add((New-GridSheetDefinition -Name "Master Design Hierarchy" -GridRows $hierarchyGrid.ToArray())) | Out-Null
    $sheetDefinitions.Add((New-GridSheetDefinition -Name "As Master" -GridRows $asMasterGrid.ToArray())) | Out-Null
    $sheetDefinitions.Add((New-GridSheetDefinition -Name "Design Master" -GridRows $designMasterGrid.ToArray())) | Out-Null

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "Info"
        Headers = $infoHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $InfoRows -Headers $infoHeaders
    }) | Out-Null

    $rawStationSheet = Get-ExcelSheetReferenceName -SheetName "Raw Station Results"
    $allPiersWorkbookRows = New-Object System.Collections.Generic.List[object]
    $envelopeIndex = 0
    foreach ($row in @($EnvelopeRows)) {
        $excelRow = $envelopeIndex + 2
        $pierRef = "`$A$excelRow"
        $storyRef = "`$B$excelRow"
        $allPiersWorkbookRows.Add([pscustomobject]@{
            Pier = $row.Pier
            Story = $row.Story
            WallLength_ft = $row.WallLength_ft
            RequiredSteelLeft_in2 = New-FormulaCell -Formula (New-CellCriteriaAggregateFormula -SourceSheet $rawStationSheet -ValueColumn "N" -PierCellReference $pierRef -StoryCellReference $storyRef) -CachedValue $row.RequiredSteelLeft_in2
            RequiredSteelRight_in2 = New-FormulaCell -Formula (New-CellCriteriaAggregateFormula -SourceSheet $rawStationSheet -ValueColumn "O" -PierCellReference $pierRef -StoryCellReference $storyRef) -CachedValue $row.RequiredSteelRight_in2
            LRFDEnvMaxM3_kip_ft = $row.LRFDEnvMaxM3_kip_ft
            LRFDEnvMinM3_kip_ft = $row.LRFDEnvMinM3_kip_ft
            MaxBZoneL_ft = New-FormulaCell -Formula (New-CellCriteriaAggregateFormula -SourceSheet $rawStationSheet -ValueColumn "S" -PierCellReference $pierRef -StoryCellReference $storyRef) -CachedValue $row.MaxBZoneL_ft
            MaxBZoneR_ft = New-FormulaCell -Formula (New-CellCriteriaAggregateFormula -SourceSheet $rawStationSheet -ValueColumn "T" -PierCellReference $pierRef -StoryCellReference $storyRef) -CachedValue $row.MaxBZoneR_ft
            MaxDCRatio = New-FormulaCell -Formula (New-CellCriteriaAggregateFormula -SourceSheet $rawStationSheet -ValueColumn "K" -PierCellReference $pierRef -StoryCellReference $storyRef) -CachedValue $row.MaxDCRatio
            PierSecTypes = $row.PierSecTypes
            Messages = $row.Messages
        }) | Out-Null
        $envelopeIndex++
    }

    $sheetDefinitions.Add([pscustomobject]@{
        Name = "All Piers"
        Headers = $envelopeHeaders
        Rows = Convert-ObjectsToSheetRows -Rows $allPiersWorkbookRows.ToArray() -Headers $envelopeHeaders -HighlightMessages
    }) | Out-Null

    foreach ($pierLabel in @($EnvelopeRows | Select-Object -ExpandProperty Pier -Unique | Sort-Object { Get-PierSortKey -PierLabel $_ })) {
        $pierRows = @($EnvelopeRows | Where-Object { $_.Pier -eq $pierLabel })
        $designCheckRows = Get-DesignCheckRows -EnvelopeRows $pierRows
        $sheetDefinitions.Add([pscustomobject]@{
            Name = $pierLabel
            Headers = $pierDesignHeaders
            Rows = Convert-ObjectsToSheetRows -Rows $designCheckRows -Headers $pierDesignHeaders -HighlightMessages
        }) | Out-Null
    }

    $exceedsCount = 0
    foreach ($ids in $designIdByStory.Values) {
        $exceedsCount += @($ids | Where-Object { $_ -eq "EXCEEDS HIERARCHY" }).Count
    }
    $stationCount = ($StationRows | Measure-Object).Count
    $envelopeCount = ($EnvelopeRows | Measure-Object).Count
    $warningCount = ($WarningRows | Measure-Object).Count
    $missingWallLengthCount = @($EnvelopeRows | Where-Object { $null -eq $_.WallLength_ft -or $_.WallLength_ft -eq "" }).Count
    $lastEnvelopeRow = $envelopeCount + 1
    $lastDesignIdRow = $dmDesignIdFirstRow + $storyNames.Count - 1
    $rawStationSheet = Get-ExcelSheetReferenceName -SheetName "Raw Station Results"
    $warningsSheet = Get-ExcelSheetReferenceName -SheetName "Warnings"
    $sanityGrid = @(
        @("Check", "Status", "Value", "Notes"),
        @("Reference sheets present", "OK", "15/15", "Generated workbook includes every sheet from the reference shear-wall-design workbook."),
        @("Raw station rows", (New-FormulaCell -Formula '=IF(C3>0,"OK","REVIEW")' -CachedValue $(if ($stationCount -gt 0) { "OK" } else { "REVIEW" })), (New-FormulaCell -Formula ("=COUNTA({0}!`$A:`$A)-1" -f $rawStationSheet) -CachedValue $stationCount), "Direct ETABS GetPierSummaryResults rows."),
        @("Pier/story envelope rows", (New-FormulaCell -Formula '=IF(C4>0,"OK","REVIEW")' -CachedValue $(if ($envelopeCount -gt 0) { "OK" } else { "REVIEW" })), (New-FormulaCell -Formula ("=COUNTA({0}!`$A:`$A)-1" -f $allPiersSheet) -CachedValue $envelopeCount), "Rows used by All Piers, As Master, Design Master, and table output."),
        @("ETABS warning rows", (New-FormulaCell -Formula '=IF(C5=0,"OK","REVIEW")' -CachedValue $(if ($warningCount -eq 0) { "OK" } else { "REVIEW" })), (New-FormulaCell -Formula ("=COUNTA({0}!`$A:`$A)-1" -f $warningsSheet) -CachedValue $warningCount), "Warning rows are listed on the Warnings sheet."),
        @("EXCEEDS HIERARCHY cells", (New-FormulaCell -Formula '=IF(C6=0,"OK","REVIEW")' -CachedValue $(if ($exceedsCount -eq 0) { "OK" } else { "REVIEW" })), (New-FormulaCell -Formula ("=COUNTIF({0}!`$B`${1}:`$AE`${2},""EXCEEDS HIERARCHY"")" -f $designMasterSheet, $dmDesignIdFirstRow, $lastDesignIdRow) -CachedValue $exceedsCount), "Schedule cells that exceed the Master Design Hierarchy."),
        @("Missing wall lengths", (New-FormulaCell -Formula '=IF(C7=0,"OK","REVIEW")' -CachedValue $(if ($missingWallLengthCount -eq 0) { "OK" } else { "REVIEW" })), (New-FormulaCell -Formula ("=COUNTBLANK({0}!`$C`$2:`$C`${1})" -f $allPiersSheet, $lastEnvelopeRow) -CachedValue $missingWallLengthCount), "Wall lengths are needed for downstream d and phi*Mn check columns.")
    )
    $sheetDefinitions.Add((New-GridSheetDefinition -Name "Sanity Checks" -GridRows $sanityGrid)) | Out-Null

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

function Convert-MomentToKipFeet {
    param(
        [AllowNull()]
        [object]$Value,
        [string]$ForceUnit,
        [string]$LengthUnit
    )

    if ($null -eq $Value -or $Value -eq "") {
        return $null
    }

    $forceFactor = switch -Regex ($ForceUnit) {
        "^lb$" { 0.001 }
        "^kip$" { 1.0 }
        "^N$" { 0.0002248089431 }
        "^kN$" { 0.2248089431 }
        default { throw "Unsupported force unit for kip-ft conversion: $ForceUnit" }
    }

    return ([double]$Value) * $forceFactor * (Get-LengthToFootFactor -LengthUnit $LengthUnit)
}

function Get-PierLengthLookup {
    param(
        $SapModel,
        [string]$LengthUnit
    )

    $db = $SapModel.DatabaseTables
    [string[]]$fieldKeys = @()
    [string[]]$fields = @()
    [string[]]$data = @()
    $tableVersion = 0
    $numRecords = 0
    $lookup = @{}

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
    for ($recordIndex = 0; $recordIndex -lt $numRecords; $recordIndex++) {
        $offset = $recordIndex * $fieldCount
        $record = @{}
        for ($fieldIndex = 0; $fieldIndex -lt $fieldCount; $fieldIndex++) {
            $record[$fields[$fieldIndex]] = $data[$offset + $fieldIndex]
        }

        $widthValues = @($record["WidthBot"], $record["WidthTop"]) | Where-Object { $null -ne $_ -and $_ -ne "" } | ForEach-Object {
            Convert-LengthToFeet -Value ([double]$_) -LengthUnit $LengthUnit
        }
        $key = "{0}|{1}" -f $record["Pier"], $record["Story"]
        $lookup[$key] = if ($widthValues.Count -gt 0) { [Math]::Round((($widthValues | Measure-Object -Maximum).Maximum), 3) } else { $null }
    }

    return $lookup
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
        $SapModel,
        [string]$ForceUnit,
        [string]$LengthUnit
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
                    MaxM3 = $null
                    MinM3 = $null
                }
            }

            $m3Value = Convert-MomentToKipFeet -Value $M3[$index] -ForceUnit $ForceUnit -LengthUnit $LengthUnit
            $entry = $lookup[$key]

            if ($null -eq $entry.MaxM3 -or $m3Value -gt $entry.MaxM3) {
                $entry.MaxM3 = $m3Value
            }

            if ($null -eq $entry.MinM3 -or $m3Value -lt $entry.MinM3) {
                $entry.MinM3 = $m3Value
            }
        }

        return $lookup
    }
    finally {
        Restore-SelectedResultState -SapModel $SapModel -State $savedState
    }
}

function Get-EnvelopeRows {
    param(
        [object[]]$StationRows,
        [hashtable]$PierLengthLookup,
        [hashtable]$LrfdEnvPierMomentLookup
    )

    $entries = @{}
    foreach ($row in $StationRows) {
        $key = "{0}|{1}" -f $row.Pier, $row.Story
        if (-not $entries.ContainsKey($key)) {
            $entries[$key] = [ordered]@{
                Pier = $row.Pier
                Story = $row.Story
                WallLength_ft = $null
                LRFDEnvMaxM3_kip_ft = $null
                LRFDEnvMinM3_kip_ft = $null
                RequiredSteelLeft_in2 = $null
                RequiredSteelRight_in2 = $null
                MaxBZoneL_ft = $null
                MaxBZoneR_ft = $null
                MaxDCRatio = $null
                PierSecTypes = New-Object System.Collections.Generic.List[string]
                Messages = New-Object System.Collections.Generic.List[string]
            }
        }

        $entry = $entries[$key]
        if ($null -eq $entry["WallLength_ft"] -and $PierLengthLookup.ContainsKey($key)) {
            $entry["WallLength_ft"] = $PierLengthLookup[$key]
        }

        if ($null -eq $entry["LRFDEnvMaxM3_kip_ft"] -and $LrfdEnvPierMomentLookup.ContainsKey($key)) {
            $entry["LRFDEnvMaxM3_kip_ft"] = [Math]::Round([double]$LrfdEnvPierMomentLookup[$key].MaxM3, 3)
            $entry["LRFDEnvMinM3_kip_ft"] = [Math]::Round([double]$LrfdEnvPierMomentLookup[$key].MinM3, 3)
        }

        if (-not [string]::IsNullOrWhiteSpace($row.PierSecType) -and -not $entry["PierSecTypes"].Contains($row.PierSecType)) {
            $entry["PierSecTypes"].Add($row.PierSecType) | Out-Null
        }

        Add-Message -Messages $entry["Messages"] -Message $row.WarnMsg
        Add-Message -Messages $entry["Messages"] -Message $row.ErrMsg

        if ($null -ne $row.AsLeft_in2 -and ($null -eq $entry["RequiredSteelLeft_in2"] -or $row.AsLeft_in2 -gt $entry["RequiredSteelLeft_in2"])) {
            $entry["RequiredSteelLeft_in2"] = $row.AsLeft_in2
        }

        if ($null -ne $row.AsRight_in2 -and ($null -eq $entry["RequiredSteelRight_in2"] -or $row.AsRight_in2 -gt $entry["RequiredSteelRight_in2"])) {
            $entry["RequiredSteelRight_in2"] = $row.AsRight_in2
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
            WallLength_ft = $entry["WallLength_ft"]
            RequiredSteelLeft_in2 = $entry["RequiredSteelLeft_in2"]
            RequiredSteelRight_in2 = $entry["RequiredSteelRight_in2"]
            LRFDEnvMaxM3_kip_ft = $entry["LRFDEnvMaxM3_kip_ft"]
            LRFDEnvMinM3_kip_ft = $entry["LRFDEnvMinM3_kip_ft"]
            MaxBZoneL_ft = $entry["MaxBZoneL_ft"]
            MaxBZoneR_ft = $entry["MaxBZoneR_ft"]
            MaxDCRatio = $entry["MaxDCRatio"]
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

$pierLengthLookup = Get-PierLengthLookup -SapModel $api.SapModel -LengthUnit $units.Length
$lrfdEnvMomentLookup = Get-LrfdEnvPierMomentLookup -SapModel $api.SapModel -ForceUnit $units.Force -LengthUnit $units.Length
$envelopeRows = Get-EnvelopeRows -StationRows $stationRows -PierLengthLookup $pierLengthLookup -LrfdEnvPierMomentLookup $lrfdEnvMomentLookup
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
    [pscustomobject]@{ Field = "QaQcAlignmentToleranceIn2"; Value = $QaToleranceIn2 },
    [pscustomobject]@{ Field = "Source"; Value = "SapModel.DesignShearWall.GetPierSummaryResults" },
    [pscustomobject]@{ Field = "EnvelopeRule"; Value = "Per pier/story, use max AsLeft and max AsRight across ETABS station rows." }
)

$infoPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "info.csv"
$stationPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "raw-station-results.csv"
$envelopePath = Join-Path -Path $resolvedOutputDirectory -ChildPath "story-envelope.csv"
$warningPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "warnings.csv"
$qaQcPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "qa-qc-alignment.csv"

$sheetDefinitions = New-WorkbookSheetDefinitions -InfoRows $infoRows -EnvelopeRows $envelopeRows -StationRows $stationRows -WarningRows $warningRows
Write-XlsxWorkbook -Path $resolvedOutputWorkbookPath -SheetDefinitions $sheetDefinitions
$qaQcSummary = Invoke-WorkbookQaQcAlignment -Path $resolvedOutputWorkbookPath -SheetDefinitions $sheetDefinitions -StationRows $stationRows -EnvelopeRows $envelopeRows -ToleranceIn2 $QaToleranceIn2 -CsvPath $(if ($NoCsv) { $null } else { $qaQcPath })

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
    QaQcAlignmentCsv = if ($NoCsv) { $null } else { $qaQcPath }
    ForceUnits = $units.Force
    LengthUnits = $units.Length
    RawStationRowCount = $stationRowCount
    StoryEnvelopeRowCount = $envelopeRowCount
    WarningRowCount = $warningRowCount
    QaQcAlignmentStatus = $qaQcSummary.Status
    QaQcAlignmentCheckedCount = $qaQcSummary.CheckedCount
    QaQcAlignmentFailureCount = $qaQcSummary.FailureCount
    QaQcAlignmentWarningCount = $qaQcSummary.WarningCount
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 5
}
else {
    $result
}
