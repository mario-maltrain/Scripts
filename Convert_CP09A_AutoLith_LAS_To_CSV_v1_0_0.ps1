# Classification: Internal
# Script: Convert_CP09A_AutoLith_LAS_To_CSV_v1_0_0.ps1
# Purpose: Convert governed AutoLith confidence LAS files to per-hole CSV files with a unified schema, QA/QC and evidence.
# Safety: AuditOnly is the default. LAS sources are read-only. No delete, move, rename or overwrite.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$StagingRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\01_Processed_Exports\GEOAI-EXPORTS-20261002-225517\13_Composites\AUTOLITH-LAS-STAGING-20261003-013700',

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-AUTOLITH-CSV-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-AUTOLITH-CSV-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [ValidateRange(1, 100000)]
    [int]$ExpectedFiles = 150,

    [Parameter()]
    [switch]$CommitRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Convert_CP09A_AutoLith_LAS_To_CSV_v1_0_0.ps1'
$ScriptVersion = '1.0.0'
$Classification = 'Internal'
$ProgressId = 1400
$ExitCode = 0
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$NumberStyle = [System.Globalization.NumberStyles]::Float
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$StandardCurves = @('DEPT', 'ABX1_PROB', 'ABX2_PROB', 'QFP1_PROB', 'QFP2_PROB', 'IND_PROB', 'H-CONFIDENCE')
$OutputHeader = 'HOLE_ID,DEPT,ABX1_PROB,ABX2_PROB,QFP1_PROB,QFP2_PROB,IND_PROB,H-CONFIDENCE'
$ExpectedSchemas = [ordered]@{
    'DEPT|ABX1_PROB|ABX2_PROB|IND_PROB|H-CONFIDENCE' = 61
    'DEPT|ABX1_PROB|QFP1_PROB|H-CONFIDENCE' = 52
    'DEPT|ABX2_PROB|QFP2_PROB|ABX1_PROB|QFP1_PROB|H-CONFIDENCE' = 32
    'DEPT|ABX2_PROB|QFP2_PROB|H-CONFIDENCE' = 5
}

function Get-LasHeaderValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Line)
    $dot = $Line.IndexOf('.')
    if ($dot -lt 1) { return $null }
    $colon = $Line.LastIndexOf(':')
    if ($colon -lt $dot) { $colon = $Line.Length }
    $mnemonic = $Line.Substring(0, $dot).Trim().ToUpperInvariant()
    $body = $Line.Substring($dot + 1, $colon - $dot - 1)
    $value = ''
    if ($body.Length -gt 0 -and -not [char]::IsWhiteSpace($body[0])) {
        $parts = @($body.Trim() -split '\s+', 2)
        if ($parts.Count -gt 1) { $value = $parts[1].Trim() }
    }
    else {
        $value = $body.Trim()
    }
    return [pscustomobject]@{ Mnemonic = $mnemonic; Value = $value }
}

function Get-HeaderNumber {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][hashtable]$Header,
        [Parameter(Mandatory)][string]$Key
    )
    if (-not $Header.ContainsKey($Key)) { return $null }
    $parsed = 0.0
    if ([double]::TryParse([string]$Header[$Key], $NumberStyle, $Invariant, [ref]$parsed)) { return $parsed }
    return $null
}

function Invoke-LasConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter()][AllowEmptyString()][string]$OutputPath = '',
        [Parameter(Mandatory)][string]$SourceHash
    )

    $holeMatch = [regex]::Match($File.Name, '^(SPD\d+)_', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $holeId = if ($holeMatch.Success) { $holeMatch.Groups[1].Value.ToUpperInvariant() } else { '' }
    $section = ''
    $curves = [System.Collections.Generic.List[string]]::new()
    $well = @{}
    $columnIndex = @{}
    $rows = 0L
    $tokenErrors = 0L
    $parseErrors = 0L
    $nullCells = 0L
    $probOutOfRange = 0L
    $confOutOfRange = 0L
    $nonMonotonic = 0L
    $stepGaps = 0L
    $firstDepth = $null
    $lastDepth = $null
    $nullValue = -999.25
    $step = 0.01
    $strt = $null
    $stop = $null
    $writer = $null
    $reader = [System.IO.StreamReader]::new($File.FullName)

    try {
        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
            $writer = [System.IO.StreamWriter]::new($OutputPath, $false, $Utf8NoBom)
            $writer.WriteLine($OutputHeader)
        }

        while ($null -ne ($line = $reader.ReadLine())) {
            $trimmed = $line.Trim()

            if ($section -ne 'A') {
                if ($trimmed.StartsWith('~')) {
                    $section = if ($trimmed.Length -gt 1) { $trimmed.Substring(1, 1).ToUpperInvariant() } else { '' }
                    if ($section -eq 'A') {
                        $headerNull = Get-HeaderNumber -Header $well -Key 'NULL'
                        if ($null -ne $headerNull) { $nullValue = $headerNull }
                        $headerStep = Get-HeaderNumber -Header $well -Key 'STEP'
                        if ($null -ne $headerStep) { $step = [math]::Abs($headerStep) }
                        $strt = Get-HeaderNumber -Header $well -Key 'STRT'
                        $stop = Get-HeaderNumber -Header $well -Key 'STOP'
                        for ($c = 0; $c -lt $curves.Count; $c++) { $columnIndex[$curves[$c]] = $c }
                    }
                    continue
                }
                if ($trimmed.Length -eq 0 -or $trimmed.StartsWith('#')) { continue }
                if ($section -eq 'W') {
                    $headerItem = Get-LasHeaderValue -Line $trimmed
                    if ($null -ne $headerItem) { $well[$headerItem.Mnemonic] = $headerItem.Value }
                }
                elseif ($section -eq 'C') {
                    $dot = $trimmed.IndexOf('.')
                    if ($dot -gt 0) { $curves.Add($trimmed.Substring(0, $dot).Trim().ToUpperInvariant()) }
                }
                continue
            }

            if ($trimmed.Length -eq 0 -or $trimmed.StartsWith('#')) { continue }
            $tokens = $trimmed -split '\s+'
            if ($tokens.Count -ne $curves.Count) { $tokenErrors++; continue }

            $values = [string[]]::new($StandardCurves.Count)
            $rowValid = $true
            $depthValue = $null

            for ($k = 0; $k -lt $StandardCurves.Count; $k++) {
                $curveName = $StandardCurves[$k]
                if (-not $columnIndex.ContainsKey($curveName)) { $values[$k] = ''; continue }
                $raw = $tokens[$columnIndex[$curveName]]
                $number = 0.0
                if (-not [double]::TryParse($raw, $NumberStyle, $Invariant, [ref]$number)) { $parseErrors++; $rowValid = $false; break }
                if ([math]::Abs($number - $nullValue) -lt 0.000001) { $values[$k] = ''; $nullCells++; continue }
                if ($curveName -eq 'DEPT') { $depthValue = $number }
                elseif ($number -lt -0.000000001 -or $number -gt 1.000000001) {
                    if ($curveName -eq 'H-CONFIDENCE') { $confOutOfRange++ } else { $probOutOfRange++ }
                }
                $values[$k] = $raw
            }

            if (-not $rowValid) { continue }
            if ($null -eq $depthValue) { $parseErrors++; continue }
            if ($null -ne $lastDepth) {
                if ($depthValue -le $lastDepth) { $nonMonotonic++ }
                elseif ($step -gt 0 -and [math]::Abs(($depthValue - $lastDepth) - $step) -gt 0.0005) { $stepGaps++ }
            }
            if ($null -eq $firstDepth) { $firstDepth = $depthValue }
            $lastDepth = $depthValue
            $rows++
            if ($null -ne $writer) { $writer.WriteLine($holeId + ',' + ($values -join ',')) }
        }
    }
    finally {
        $reader.Dispose()
        if ($null -ne $writer) { $writer.Dispose() }
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $uwi = if ($well.ContainsKey('UWI')) { [string]$well['UWI'] } else { '' }

    if ([string]::IsNullOrEmpty($holeId)) { $issues.Add('HOLE_ID_NOT_IN_FILENAME') }
    elseif (-not $uwi.ToUpperInvariant().EndsWith($holeId)) { $issues.Add('UWI_MISMATCH') }
    if ($well.ContainsKey('WRAP') -and ([string]$well['WRAP']).ToUpperInvariant() -eq 'YES') { $issues.Add('WRAPPED_LAS_NOT_SUPPORTED') }
    if (-not $curves.Contains('DEPT')) { $issues.Add('DEPT_MISSING') }
    $unknownCurves = @($curves | Where-Object { $StandardCurves -notcontains $_ })
    if ($unknownCurves.Count -gt 0) { $issues.Add('UNKNOWN_CURVES=' + ($unknownCurves -join '|')) }
    if (@($curves | Select-Object -Unique).Count -ne $curves.Count) { $issues.Add('DUPLICATE_CURVES') }
    if ($rows -eq 0) { $issues.Add('NO_DATA_ROWS') }
    if ($tokenErrors -gt 0) { $issues.Add('TOKEN_COUNT_ERRORS=' + $tokenErrors) }
    if ($parseErrors -gt 0) { $issues.Add('PARSE_ERRORS=' + $parseErrors) }
    if ($nonMonotonic -gt 0) { $issues.Add('NON_MONOTONIC_DEPTH=' + $nonMonotonic) }
    if ($stepGaps -gt 0) { $warnings.Add('STEP_GAPS=' + $stepGaps) }
    if ($probOutOfRange -gt 0) { $warnings.Add('PROBABILITY_OUT_OF_RANGE=' + $probOutOfRange) }
    if ($confOutOfRange -gt 0) { $warnings.Add('CONFIDENCE_OUT_OF_RANGE=' + $confOutOfRange) }
    if ($null -ne $strt -and $null -ne $firstDepth -and [math]::Abs($firstDepth - $strt) -gt 0.0005) { $warnings.Add('FIRST_DEPTH_NOT_STRT') }
    if ($null -ne $stop -and $null -ne $lastDepth -and [math]::Abs($lastDepth - $stop) -gt 0.0005) { $warnings.Add('LAST_DEPTH_NOT_STOP') }

    return [pscustomobject]@{
        File = $File.Name
        HoleId = $holeId
        Uwi = $uwi
        Signature = ($curves -join '|')
        CurveCount = $curves.Count
        Rows = $rows
        FirstDepth = $firstDepth
        LastDepth = $lastDepth
        Strt = $strt
        Stop = $stop
        Step = $step
        NullValue = $nullValue
        NullCells = $nullCells
        TokenErrors = $tokenErrors
        ParseErrors = $parseErrors
        NonMonotonic = $nonMonotonic
        StepGaps = $stepGaps
        ProbabilityOutOfRange = $probOutOfRange
        ConfidenceOutOfRange = $confOutOfRange
        Status = if ($issues.Count -gt 0) { 'FAIL' } elseif ($warnings.Count -gt 0) { 'WARN' } else { 'PASS' }
        Issues = ($issues -join ';')
        Warnings = ($warnings -join ';')
        SourceSHA256 = $SourceHash
        OutputFile = $OutputPath
        OutputSHA256 = ''
    }
}

try {
    Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Status 'Validating context' -PercentComplete 2

    if (-not (Test-Path -LiteralPath $StagingRoot -PathType Container)) { throw "StagingRoot does not exist: $StagingRoot" }
    $ResolvedStaging = (Get-Item -LiteralPath $StagingRoot -ErrorAction Stop).FullName
    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) { throw 'OutputRoot is required when CommitRecords is used.' }

    $LasFiles = @(Get-ChildItem -LiteralPath $ResolvedStaging -File -Recurse -ErrorAction Stop | Where-Object { $_.Extension -eq '.las' } | Sort-Object Name)
    if ($LasFiles.Count -eq 0) { throw "No LAS files found under: $ResolvedStaging" }

    Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Status 'Hashing source LAS files' -PercentComplete 5
    $HashBefore = @{}
    foreach ($lasFile in $LasFiles) { $HashBefore[$lasFile.FullName] = (Get-FileHash -LiteralPath $lasFile.FullName -Algorithm SHA256).Hash }

    $RegisterPath = Join-Path $ResolvedStaging 'AutoLith_LAS_Register.csv'
    $RegisterStatus = 'NOT_FOUND'
    $RegisterMismatches = 0
    if (Test-Path -LiteralPath $RegisterPath -PathType Leaf) {
        $RegisterRows = @(Import-Csv -LiteralPath $RegisterPath)
        $RegisterByName = @{}
        foreach ($registerRow in $RegisterRows) { $RegisterByName[$registerRow.Name] = $registerRow.SHA256 }
        foreach ($lasFile in $LasFiles) {
            if (-not $RegisterByName.ContainsKey($lasFile.Name) -or $RegisterByName[$lasFile.Name] -ne $HashBefore[$lasFile.FullName]) { $RegisterMismatches++ }
        }
        if ($RegisterRows.Count -ne $LasFiles.Count) { $RegisterMismatches++ }
        $RegisterStatus = if ($RegisterMismatches -eq 0) { 'PASS' } else { 'FAIL' }
    }

    $RunOutput = ''
    $CsvFolder = ''
    if (-not $AuditOnly) {
        $RunOutput = Join-Path ([System.IO.Path]::GetFullPath($OutputRoot)) $RunId
        if (Test-Path -LiteralPath $RunOutput) { throw "Overwrite is prohibited. Run output already exists: $RunOutput" }
        if (-not $PSCmdlet.ShouldProcess($RunOutput, 'Create AutoLith CSV conversion outputs')) { return }
        $CsvFolder = Join-Path $RunOutput 'CSV'
        New-Item -ItemType Directory -Path $CsvFolder -ErrorAction Stop | Out-Null
    }

    $Records = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $LasFiles.Count; $i++) {
        $lasFile = $LasFiles[$i]
        $percent = 10 + [int][math]::Floor($i / $LasFiles.Count / 0.0125)
        Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Status ('{0} of {1}: {2}' -f ($i + 1), $LasFiles.Count, $lasFile.Name) -PercentComplete $percent
        $outputPath = ''
        if (-not $AuditOnly) { $outputPath = Join-Path $CsvFolder ($lasFile.BaseName + '.csv') }
        $Records.Add((Invoke-LasConversion -File $lasFile -OutputPath $outputPath -SourceHash $HashBefore[$lasFile.FullName]))
    }

    Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Status 'Running QA/QC' -PercentComplete 92

    $DuplicateGroups = @($Records | Where-Object { -not [string]::IsNullOrEmpty($_.HoleId) } | Group-Object HoleId | Where-Object { $_.Count -gt 1 })
    foreach ($group in $DuplicateGroups) {
        foreach ($record in $group.Group) {
            $record.Status = 'FAIL'
            $record.Issues = (@($record.Issues, 'DUPLICATE_HOLE_ID') | Where-Object { -not [string]::IsNullOrEmpty($_) }) -join ';'
        }
    }

    $SchemaRows = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($Records | Group-Object Signature)) {
        $expected = if ($ExpectedSchemas.Contains($group.Name)) { [int]$ExpectedSchemas[$group.Name] } else { 0 }
        $SchemaRows.Add([pscustomobject]@{ Signature = $group.Name; Files = $group.Count; ExpectedFiles = $expected; Status = if ($group.Count -eq $expected) { 'PASS' } else { 'REVIEW' } })
    }
    foreach ($expectedSignature in $ExpectedSchemas.Keys) {
        if (@($SchemaRows | Where-Object { $_.Signature -eq $expectedSignature }).Count -eq 0) {
            $SchemaRows.Add([pscustomobject]@{ Signature = $expectedSignature; Files = 0; ExpectedFiles = [int]$ExpectedSchemas[$expectedSignature]; Status = 'REVIEW' })
        }
    }

    $ChangedSources = 0
    foreach ($lasFile in $LasFiles) {
        if ((Get-FileHash -LiteralPath $lasFile.FullName -Algorithm SHA256).Hash -ne $HashBefore[$lasFile.FullName]) { $ChangedSources++ }
    }

    $FailFiles = @($Records | Where-Object { $_.Status -eq 'FAIL' })
    $WarnFiles = @($Records | Where-Object { $_.Status -eq 'WARN' })
    $ReviewSchemas = @($SchemaRows | Where-Object { $_.Status -ne 'PASS' })
    $TotalRows = [long](($Records | Measure-Object -Property Rows -Sum).Sum)
    $TotalNullCells = [long](($Records | Measure-Object -Property NullCells -Sum).Sum)
    $FileCountPass = ($LasFiles.Count -eq $ExpectedFiles)

    $Blocking = $FailFiles.Count + $ChangedSources
    if (-not $FileCountPass) { $Blocking++ }
    if ($RegisterStatus -eq 'FAIL') { $Blocking++ }
    $Decision = if ($Blocking -gt 0) { 'NO_GO' } elseif ($WarnFiles.Count -gt 0 -or $ReviewSchemas.Count -gt 0 -or $RegisterStatus -eq 'NOT_FOUND') { 'PASS_WITH_WARNINGS' } else { 'PASS' }

    $QaRows = @(
        [pscustomobject]@{ Control = 'StagingRoot'; Status = 'PASS'; Evidence = $ResolvedStaging }
        [pscustomobject]@{ Control = 'LasFileCount'; Status = if ($FileCountPass) { 'PASS' } else { 'FAIL' }; Evidence = ('Found={0}; Expected={1}' -f $LasFiles.Count, $ExpectedFiles) }
        [pscustomobject]@{ Control = 'LasRegisterIntegrity'; Status = $RegisterStatus; Evidence = ('Mismatches={0}' -f $RegisterMismatches) }
        [pscustomobject]@{ Control = 'SchemaDistribution'; Status = if ($ReviewSchemas.Count -eq 0) { 'PASS' } else { 'REVIEW' }; Evidence = ('Signatures={0}' -f $SchemaRows.Count) }
        [pscustomobject]@{ Control = 'FileBlockingIssues'; Status = if ($FailFiles.Count -eq 0) { 'PASS' } else { 'FAIL' }; Evidence = ('FailFiles={0}' -f $FailFiles.Count) }
        [pscustomobject]@{ Control = 'DuplicateHoleId'; Status = if ($DuplicateGroups.Count -eq 0) { 'PASS' } else { 'FAIL' }; Evidence = ('DuplicateGroups={0}' -f $DuplicateGroups.Count) }
        [pscustomobject]@{ Control = 'FileWarnings'; Status = if ($WarnFiles.Count -eq 0) { 'PASS' } else { 'WARN' }; Evidence = ('WarnFiles={0}' -f $WarnFiles.Count) }
        [pscustomobject]@{ Control = 'NullHandling'; Status = 'PASS'; Evidence = ('LAS NULL converted to empty cells={0}' -f $TotalNullCells) }
        [pscustomobject]@{ Control = 'NoRoundingNoNormalization'; Status = 'PASS'; Evidence = 'Source numeric tokens written unchanged' }
        [pscustomobject]@{ Control = 'SourceImmutability'; Status = if ($ChangedSources -eq 0) { 'PASS' } else { 'FAIL' }; Evidence = ('ChangedSources={0}' -f $ChangedSources) }
    )

    Write-Host ''
    Write-Host ('CP09A AUTOLITH LAS TO CSV v{0}' -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' }))
    Write-Host ('LAS FILES: {0}' -f $LasFiles.Count)
    Write-Host ('DATA ROWS: {0}' -f $TotalRows)
    Write-Host ('LAS REGISTER: {0}' -f $RegisterStatus)
    Write-Host ''
    $SchemaRows | Format-Table Signature, Files, ExpectedFiles, Status -AutoSize -Wrap
    $QaRows | Format-Table Control, Status, Evidence -AutoSize -Wrap
    if ($FailFiles.Count -gt 0) { $FailFiles | Select-Object -First 25 File, Issues | Format-Table -AutoSize -Wrap }
    if ($WarnFiles.Count -gt 0) { $WarnFiles | Select-Object -First 25 File, Warnings | Format-Table -AutoSize -Wrap }

    if (-not $AuditOnly) {
        Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Status 'Writing registers and evidence' -PercentComplete 95
        foreach ($record in $Records) {
            if (-not [string]::IsNullOrEmpty($record.OutputFile) -and (Test-Path -LiteralPath $record.OutputFile -PathType Leaf)) {
                $record.OutputSHA256 = (Get-FileHash -LiteralPath $record.OutputFile -Algorithm SHA256).Hash
            }
        }
        $OutputCsvCount = @(Get-ChildItem -LiteralPath $CsvFolder -File | Where-Object { $_.Extension -eq '.csv' }).Count

        $Records | Export-Csv -LiteralPath (Join-Path $RunOutput 'AutoLith_File_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
        $SchemaRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'AutoLith_Schema_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
        $QaRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'AutoLith_QAQC_Register.csv') -NoTypeInformation -Encoding utf8NoBOM

        $RunRecord = [ordered]@{
            Classification = $Classification
            ScriptName = $ScriptName
            ScriptVersion = $ScriptVersion
            RunId = $RunId
            StagingRoot = $ResolvedStaging
            OutputRoot = $RunOutput
            CsvFolder = $CsvFolder
            LasFiles = $LasFiles.Count
            ExpectedFiles = $ExpectedFiles
            OutputCsvFiles = $OutputCsvCount
            DataRows = $TotalRows
            NullCells = $TotalNullCells
            OutputColumns = $OutputHeader
            NullPolicy = 'LAS NULL value written as empty cell'
            Normalization = $false
            Rounding = $false
            SourceModified = ($ChangedSources -gt 0)
            Decision = $Decision
            CreatedAt = (Get-Date).ToString('o')
        }
        ($RunRecord | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath (Join-Path $RunOutput 'Run_Register.json') -Encoding utf8NoBOM

        $Readme = @(
            '# CP09A AutoLith LAS to CSV'
            ''
            ('Classification: {0}' -f $Classification)
            ('Run ID: {0}' -f $RunId)
            ('Decision: {0}' -f $Decision)
            ''
            'One CSV per LAS file in the CSV folder.'
            ('Columns: {0}' -f $OutputHeader)
            'Curves absent from a LAS schema are written as empty cells.'
            'LAS NULL values are written as empty cells.'
            'Numeric values are copied without rounding or normalization.'
            'Source LAS files are read-only and verified unchanged by SHA256.'
        ) -join [Environment]::NewLine
        $Readme | Set-Content -LiteralPath (Join-Path $RunOutput 'README.md') -Encoding utf8NoBOM

        $ChangeLog = @(
            '# Changelog'
            ''
            '## v1.0.0'
            '- Native PowerShell LAS 2.0 reader for the four AutoLith confidence schemas.'
            '- Unified per-hole CSV output with HOLE_ID derived from the LAS file name and checked against UWI.'
            '- QA/QC for token counts, parsing, depth monotonicity, step gaps, value ranges, duplicates and schema distribution.'
            '- AuditOnly default, CommitRecords mode, manifest and SHA256 evidence.'
        ) -join [Environment]::NewLine
        $ChangeLog | Set-Content -LiteralPath (Join-Path $RunOutput 'CHANGELOG.md') -Encoding utf8NoBOM

        $ManifestItems = @(
            Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Sort-Object FullName | ForEach-Object {
                [pscustomobject]@{
                    RelativePath = [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName)
                    Length = [long]$_.Length
                    SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                }
            }
        )
        $Manifest = [ordered]@{ Classification = $Classification; RunId = $RunId; FileCount = $ManifestItems.Count; Files = $ManifestItems }
        ($Manifest | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath (Join-Path $RunOutput 'Manifest.json') -Encoding utf8NoBOM

        $HashLines = @(
            Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object FullName | ForEach-Object {
                '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName)
            }
        )
        ($HashLines -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Encoding ascii

        foreach ($hashLine in $HashLines) {
            $parts = $hashLine -split '\s{2}', 2
            $actual = (Get-FileHash -LiteralPath (Join-Path $RunOutput $parts[1]) -Algorithm SHA256).Hash
            if ($actual -ne $parts[0]) { throw ('SHA256 verification failed: {0}' -f $parts[1]) }
        }

        Write-Host ('OUTPUT CSV FILES: {0}' -f $OutputCsvCount)
        Write-Host ('CSV FOLDER: {0}' -f $CsvFolder) -ForegroundColor Green
        Write-Host ('RECORDS: {0}' -f $RunOutput) -ForegroundColor Green
    }
    else {
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
    }

    $DecisionColor = if ($Decision -eq 'PASS') { 'Green' } elseif ($Decision -eq 'PASS_WITH_WARNINGS') { 'Yellow' } else { 'Red' }
    Write-Host ('DECISION: {0}' -f $Decision) -ForegroundColor $DecisionColor
    if ($Decision -eq 'NO_GO') { $ExitCode = 1 }
}
catch {
    Write-Host ''
    Write-Host ('FAIL: {0}' -f $_.Exception.Message) -ForegroundColor Red
    $ExitCode = 1
}
finally {
    Write-Progress -Id $ProgressId -Activity 'CP09A AutoLith LAS to CSV' -Completed
}

exit $ExitCode
