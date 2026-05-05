function Resolve-InstabilityKnowledgePath {
    param(
        [string]$KnowledgePath
    )

    if ([string]::IsNullOrWhiteSpace($KnowledgePath)) {
        $defaultRoot = Join-Path -Path ([Environment]::GetFolderPath("LocalApplicationData")) -ChildPath "AgenticEtabsSafe"
        return (Join-Path -Path $defaultRoot -ChildPath "instability-knowledge.json")
    }

    if (Test-Path -LiteralPath $KnowledgePath -PathType Container) {
        $resolvedDirectory = (Resolve-Path -LiteralPath $KnowledgePath).Path
        return (Join-Path -Path $resolvedDirectory -ChildPath "instability-knowledge.json")
    }

    if (Test-Path -LiteralPath $KnowledgePath -PathType Leaf) {
        return (Resolve-Path -LiteralPath $KnowledgePath).Path
    }

    return $KnowledgePath
}

function New-InstabilityKnowledgeStore {
    [pscustomobject]@{
        SchemaVersion = 1
        Observations  = @()
        Resolutions   = @()
    }
}

function Import-InstabilityKnowledgeStore {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return (New-InstabilityKnowledgeStore)
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    }
    catch {
        return (New-InstabilityKnowledgeStore)
    }

    [pscustomobject]@{
        SchemaVersion = if ($raw.SchemaVersion) { [int]$raw.SchemaVersion } else { 1 }
        Observations  = @($raw.Observations)
        Resolutions   = @($raw.Resolutions)
    }
}

function Export-InstabilityKnowledgeStore {
    param(
        $Store,
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Store | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path
}

function New-InstabilityFinding {
    param(
        [string]$Severity,
        [string]$Category,
        [string]$Signature,
        [string]$Detail,
        $Evidence = $null
    )

    [pscustomobject]@{
        Severity  = $Severity
        Category  = $Category
        Signature = $Signature
        Detail    = $Detail
        Evidence  = $Evidence
    }
}

function Get-InstabilitySeverityRank {
    param(
        [string]$Severity
    )

    $severityValue = if ($null -eq $Severity) { "" } else { $Severity }

    switch ($severityValue.ToUpperInvariant()) {
        "ERROR" { return 3 }
        "WARN" { return 2 }
        "INFO" { return 1 }
        default { return 0 }
    }
}

function Get-InstabilityHighestSeverity {
    param(
        [object[]]$Findings
    )

    $highestRank = 0
    $highestSeverity = "INFO"

    foreach ($finding in @($Findings)) {
        $rank = Get-InstabilitySeverityRank -Severity $finding.Severity
        if ($rank -gt $highestRank) {
            $highestRank = $rank
            $highestSeverity = $finding.Severity
        }
    }

    return $highestSeverity
}

function Get-InstabilityFindingKey {
    param(
        $Finding
    )

    $signature = if ($Finding.PSObject.Properties.Name -contains "Signature" -and -not [string]::IsNullOrWhiteSpace([string]$Finding.Signature)) {
        [string]$Finding.Signature
    }
    else {
        "general"
    }

    return ("{0}|{1}" -f [string]$Finding.Category, $signature)
}

function Get-InstabilityFindingKeys {
    param(
        [object[]]$Findings
    )

    return @(
        @($Findings) |
            ForEach-Object { Get-InstabilityFindingKey -Finding $_ } |
            Sort-Object -Unique
    )
}

function Add-InstabilityObservation {
    param(
        $Store,
        [string]$Platform,
        [string]$ModelPath,
        [object[]]$Findings
    )

    $entry = [pscustomobject]@{
        Timestamp       = (Get-Date).ToString("s")
        Platform        = $Platform
        ModelPath       = $ModelPath
        HighestSeverity = (Get-InstabilityHighestSeverity -Findings $Findings)
        FindingCount    = @($Findings).Count
        FindingKeys     = @(Get-InstabilityFindingKeys -Findings $Findings)
        Categories      = @(@($Findings | ForEach-Object { $_.Category }) | Sort-Object -Unique)
    }

    $Store.Observations = @(@($Store.Observations) + $entry)
    if ($Store.Observations.Count -gt 200) {
        $Store.Observations = @($Store.Observations | Select-Object -Last 200)
    }
}

function Add-InstabilityResolution {
    param(
        $Store,
        [string]$Platform,
        [string]$ModelPath,
        [object[]]$Findings,
        [string]$Resolution,
        [string[]]$Tags
    )

    if ([string]::IsNullOrWhiteSpace($Resolution)) {
        return
    }

    $cleanTags = @(
        @($Tags) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $entry = [pscustomobject]@{
        Timestamp   = (Get-Date).ToString("s")
        Platform    = $Platform
        ModelPath   = $ModelPath
        FindingKeys = @(Get-InstabilityFindingKeys -Findings $Findings)
        Resolution  = $Resolution.Trim()
        Tags        = $cleanTags
    }

    $Store.Resolutions = @(@($Store.Resolutions) + $entry)
    if ($Store.Resolutions.Count -gt 200) {
        $Store.Resolutions = @($Store.Resolutions | Select-Object -Last 200)
    }
}

function Get-InstabilityRecurringPatterns {
    param(
        $Store,
        [object[]]$Findings
    )

    $currentKeys = @(Get-InstabilityFindingKeys -Findings $Findings)
    $patterns = @()

    foreach ($key in $currentKeys) {
        $seenInRuns = @(
            @($Store.Observations) |
                Where-Object { @($_.FindingKeys) -contains $key }
        ).Count

        if ($seenInRuns -gt 0) {
            $parts = $key -split "\|", 2
            $patterns += [pscustomobject]@{
                Key        = $key
                Category   = $parts[0]
                Signature  = if ($parts.Count -gt 1) { $parts[1] } else { "general" }
                SeenInRuns = $seenInRuns
            }
        }
    }

    return @(
        $patterns |
            Sort-Object -Property @{ Expression = "SeenInRuns"; Descending = $true }, @{ Expression = "Key"; Descending = $false }
    )
}

function Get-InstabilitySuggestedResolutions {
    param(
        $Store,
        [string]$Platform,
        [object[]]$Findings
    )

    $currentKeys = @(Get-InstabilityFindingKeys -Findings $Findings)
    if ($currentKeys.Count -eq 0) {
        return @()
    }

    $suggestions = @()

    foreach ($resolution in @($Store.Resolutions)) {
        $resolutionKeys = @($resolution.FindingKeys)
        if ($resolutionKeys.Count -eq 0) {
            continue
        }

        $matchingKeys = @($currentKeys | Where-Object { $resolutionKeys -contains $_ })
        if ($matchingKeys.Count -eq 0) {
            continue
        }

        $denominator = [Math]::Max($currentKeys.Count, $resolutionKeys.Count)
        if ($denominator -le 0) {
            $denominator = 1
        }

        $score = [double]$matchingKeys.Count / [double]$denominator
        if ([string]::Equals($resolution.Platform, $Platform, [System.StringComparison]::OrdinalIgnoreCase)) {
            $score += 0.25
        }

        $suggestions += [pscustomobject]@{
            Score              = [Math]::Round($score, 3)
            Platform           = $resolution.Platform
            Timestamp          = $resolution.Timestamp
            Resolution         = $resolution.Resolution
            Tags               = @($resolution.Tags)
            MatchingFindingKeys = $matchingKeys
        }
    }

    return @(
        $suggestions |
            Sort-Object -Property @{ Expression = "Score"; Descending = $true }, @{ Expression = "Timestamp"; Descending = $true } |
            Select-Object -First 5
    )
}

function New-InstabilityLearningSummary {
    param(
        $Store,
        [string]$KnowledgePath,
        [string]$Platform,
        [object[]]$Findings
    )

    [pscustomobject]@{
        KnowledgePath        = $KnowledgePath
        ObservationCount     = @($Store.Observations).Count
        ResolutionCount      = @($Store.Resolutions).Count
        RecurringPatterns    = @(Get-InstabilityRecurringPatterns -Store $Store -Findings $Findings)
        SuggestedResolutions = @(Get-InstabilitySuggestedResolutions -Store $Store -Platform $Platform -Findings $Findings)
    }
}
