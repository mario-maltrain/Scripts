#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$InputPath = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\01_LAS_to_CSV',
    [string]$OutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\13_Delimiter_Normalization',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [int]$SampleLines = 50,
    [switch]$AuditOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.0.0'
$Classification = 'Internal'
$RunId = 'CP09A-CSV-DELIMITER-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$StagingRoot = Join-Path $OutputRoot ('.staging-' + $RunId)
$FinalRoot = Join-Path $OutputRoot $RunId
$DataRoot = Join-Path $StagingRoot 'CSV'
$RecordRoot = Join-Path $StagingRoot '_Delimiter_Record'
$LogPath = Join-Path $RecordRoot 'Execution.log'
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\CSV_Delimiter'
$InstalledPath = Join-Path $InstalledFolder 'Normalize_CP09A_CSV_Delimiter_v1_0_0.ps1'
$SourceScriptPath = $MyInvocation.MyCommand.Path

# Candidate delimiters evaluated by the detector. Tab is represented by the real tab character.
$CandidateDelimiters = @(',', ';', "`t", '|')

function Write-RunLog {
    param([string]$Message,[ValidateSet('INFO','PASS','WARN','ERROR')][string]$Level='INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $line
    if (Test-Path -LiteralPath $RecordRoot -PathType Container) { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 }
}

function Test-ScriptSyntax {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    return ($null -eq $errors -or $errors.Count -eq 0)
}

function Get-DelimiterLabel {
    param([string]$Delimiter)
    switch ($Delimiter) {
        "`t" { return 'TAB' }
        ','  { return 'COMMA' }
        ';'  { return 'SEMICOLON' }
        '|'  { return 'PIPE' }
        default { return $Delimiter }
    }
}

# Quote-aware field split for a single physical CSV record line.
# Respects RFC 4180 double-quote enclosure and the "" escape sequence.
# Returns the array of field values with surrounding enclosure quotes removed.
function Split-CsvLine {
    param([string]$Line,[string]$Delimiter)
    $fields = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $i = 0
    $length = $Line.Length
    while ($i -lt $length) {
        $char = $Line[$i]
        if ($inQuotes) {
            if ($char -eq '"') {
                if ($i + 1 -lt $length -and $Line[$i + 1] -eq '"') {
                    [void]$current.Append('"')
                    $i += 2
                    continue
                }
                $inQuotes = $false
                $i++
                continue
            }
            [void]$current.Append($char)
            $i++
            continue
        }
        if ($char -eq '"') {
            $inQuotes = $true
            $i++
            continue
        }
        if ($char -eq $Delimiter) {
            [void]$fields.Add($current.ToString())
            $current = New-Object System.Text.StringBuilder
            $i++
            continue
        }
        [void]$current.Append($char)
        $i++
    }
    [void]$fields.Add($current.ToString())
    return ,$fields.ToArray()
}

# Reads up to $MaxLines physical records from a file, joining records that span
# multiple physical lines because of an unterminated quoted field.
function Read-SampleRecords {
    param([string]$Path,[int]$MaxLines)
    $records = New-Object System.Collections.Generic.List[string]
    $pending = $null
    $reader = [System.IO.File]::OpenText($Path)
    try {
        while ($records.Count -lt $MaxLines) {
            $line = $reader.ReadLine()
            if ($null -eq $line) { break }
            if ($null -eq $pending) { $pending = $line } else { $pending = $pending + "`n" + $line }
            $quoteCount = ([regex]::Matches($pending,'"')).Count
            if (($quoteCount % 2) -eq 0) {
                [void]$records.Add($pending)
                $pending = $null
            }
        }
    }
    finally {
        $reader.Dispose()
    }
    if ($null -ne $pending) { [void]$records.Add($pending) }
    return ,$records.ToArray()
}

# Reusable detector. Scores each candidate delimiter by the consistency of its
# field count across the sampled records and returns the most consistent one.
# This is the shared robust replacement for the fragile inline 'Delimiter'
# function in Cross_Export_Comparison_Engine (which reads only the first line
# via -TotalCount 1 and miscounts delimiters inside quoted fields).
function Get-CsvDelimiter {
    param(
        [string]$Path,
        [int]$MaxLines = 50,
        [string[]]$Candidates = @(',', ';', "`t", '|')
    )
    $records = @(Read-SampleRecords -Path $Path -MaxLines $MaxLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($records.Count -eq 0) {
        return [pscustomobject]@{
            Delimiter = ','; DelimiterLabel = 'COMMA'; FieldCount = 0; SampledRecords = 0
            Consistency = 0.0; ConsistentRecords = 0; Confident = $false
            Note = 'No non-empty records sampled; defaulted to comma.'
        }
    }
    $best = $null
    foreach ($candidate in $Candidates) {
        $counts = @($records | ForEach-Object { (Split-CsvLine $_ $candidate).Count })
        $groups = @($counts | Group-Object | Sort-Object Count -Descending)
        $modeCount = [int]$groups[0].Name
        $modeFreq = [int]$groups[0].Count
        $consistency = $modeFreq / [double]$records.Count
        # A real delimiter must produce more than one field in the dominant mode.
        if ($modeCount -le 1) { continue }
        $score = [pscustomobject]@{
            Delimiter = $candidate; DelimiterLabel = (Get-DelimiterLabel $candidate)
            FieldCount = $modeCount; SampledRecords = $records.Count
            Consistency = [math]::Round($consistency,4); ConsistentRecords = $modeFreq
        }
        if ($null -eq $best) { $best = $score; continue }
        # Prefer higher consistency, then higher stable field count.
        if ($score.Consistency -gt $best.Consistency -or
            ($score.Consistency -eq $best.Consistency -and $score.FieldCount -gt $best.FieldCount)) {
            $best = $score
        }
    }
    if ($null -eq $best) {
        return [pscustomobject]@{
            Delimiter = ','; DelimiterLabel = 'COMMA'; FieldCount = 1; SampledRecords = $records.Count
            Consistency = 0.0; ConsistentRecords = 0; Confident = $false
            Note = 'No candidate delimiter produced multiple fields; defaulted to comma (single-column file).'
        }
    }
    return [pscustomobject]@{
        Delimiter = $best.Delimiter; DelimiterLabel = $best.DelimiterLabel; FieldCount = $best.FieldCount
        SampledRecords = $best.SampledRecords; Consistency = $best.Consistency; ConsistentRecords = $best.ConsistentRecords
        Confident = ($best.Consistency -ge 0.95); Note = 'Delimiter detected by multi-line quote-aware consistency scoring.'
    }
}

function Get-SafeCsvField {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + $Value.Replace('"','""') + '"'
}

# Rewrites a CSV into canonical comma-delimited, fully quoted form using the
# detected delimiter. Written via staging with a no-overwrite guard.
function Convert-OneCsvFile {
    param([string]$SourcePath,[string]$TargetPath,[string]$Delimiter,[int]$MaxLines)
    $parent = Split-Path -Path $TargetPath -Parent
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    if (Test-Path -LiteralPath $TargetPath) { throw ('No-overwrite safety block: {0}' -f $TargetPath) }
    $utf8 = New-Object System.Text.UTF8Encoding($true)
    $writer = New-Object System.IO.StreamWriter($TargetPath,$false,$utf8)
    $reader = [System.IO.File]::OpenText($SourcePath)
    $recordCount = 0
    $fieldCount = -1
    $pending = $null
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($null -eq $pending) { $pending = $line } else { $pending = $pending + "`n" + $line }
            $quoteCount = ([regex]::Matches($pending,'"')).Count
            if (($quoteCount % 2) -ne 0) { continue }
            $fields = Split-CsvLine $pending $Delimiter
            $pending = $null
            if ($recordCount -eq 0) { $fieldCount = $fields.Count }
            $writer.WriteLine((($fields | ForEach-Object { Get-SafeCsvField $_ }) -join ','))
            $recordCount++
        }
        if ($null -ne $pending) { throw ('Unterminated quoted field at end of file: {0}' -f $SourcePath) }
    }
    finally {
        $reader.Dispose()
        $writer.Dispose()
    }
    if ($recordCount -eq 0) { throw ('No records normalized from {0}' -f $SourcePath) }
    [pscustomobject]@{
        SourceCSV=$SourcePath;OutputCSV=$TargetPath;SourceDelimiter=(Get-DelimiterLabel $Delimiter)
        HeaderFieldCount=$fieldCount;Records=$recordCount
        SourceBytes=(Get-Item -LiteralPath $SourcePath).Length;OutputBytes=(Get-Item -LiteralPath $TargetPath).Length
        SourceSHA256=(Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
        OutputSHA256=(Get-FileHash -LiteralPath $TargetPath -Algorithm SHA256).Hash;Status='PASS'
    }
}

function Install-CurrentScript {
    if (-not (Test-Path -LiteralPath $SourceScriptPath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-ScriptSyntax $SourceScriptPath)) { throw 'Current script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    Copy-Item -LiteralPath $SourceScriptPath -Destination $InstalledPath -Force
    if (-not (Test-ScriptSyntax $InstalledPath)) { throw 'Installed script failed parser validation.' }
}

try {
    if (-not (Test-Path -LiteralPath $InputPath)) { throw ('Input path not found: {0}' -f $InputPath) }
    if ($SampleLines -lt 1) { throw 'SampleLines must be a positive integer.' }

    $inputIsFolder = Test-Path -LiteralPath $InputPath -PathType Container
    if ($inputIsFolder) {
        $csvFiles = @(Get-ChildItem -LiteralPath $InputPath -Recurse -File | Where-Object { $_.Extension -ieq '.csv' } | Sort-Object FullName)
    } else {
        $csvFiles = @(Get-Item -LiteralPath $InputPath)
    }
    if ($csvFiles.Count -eq 0) { throw 'No CSV files found at the input path.' }

    Write-Host ''
    Write-Host ('{0} | CP09A CSV DELIMITER NORMALIZER v{1}' -f $Classification,$Version) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('INPUT: {0}' -f $InputPath)
    Write-Host ('SAMPLE LINES: {0}' -f $SampleLines)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'NORMALIZE' }))

    $detections = New-Object System.Collections.ArrayList
    foreach ($file in $csvFiles) {
        $d = Get-CsvDelimiter -Path $file.FullName -MaxLines $SampleLines -Candidates $CandidateDelimiters
        [void]$detections.Add([pscustomobject]@{
            File=$file.FullName;Name=$file.Name;Delimiter=$d.DelimiterLabel;DelimiterChar=$d.Delimiter;FieldCount=$d.FieldCount
            SampledRecords=$d.SampledRecords;Consistency=$d.Consistency;Confident=$d.Confident;Note=$d.Note
        })
    }
    $detectionReport = @($detections | Select-Object File,Name,Delimiter,FieldCount,SampledRecords,Consistency,Confident,Note)
    $detectionReport | Format-Table Name,Delimiter,FieldCount,SampledRecords,Consistency,Confident -AutoSize

    if ($AuditOnly) {
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED.' -ForegroundColor Green
        exit 0
    }

    if (Test-Path -LiteralPath $FinalRoot) { throw ('Final output already exists: {0}' -f $FinalRoot) }
    if (Test-Path -LiteralPath $StagingRoot) { throw ('Staging output already exists: {0}' -f $StagingRoot) }
    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null
    Write-RunLog ('BHP INVENT | CP09A CSV Delimiter Normalizer | Version {0} | Classification {1}' -f $Version,$Classification)
    Write-RunLog ('RUN_ID: {0}' -f $RunId)

    Install-CurrentScript
    Write-RunLog ('Installed script: {0}' -f $InstalledPath) 'PASS'

    $results = New-Object System.Collections.ArrayList
    $failed = New-Object System.Collections.ArrayList
    $baseFolder = if ($inputIsFolder) { (Get-Item -LiteralPath $InputPath).FullName } else { (Get-Item -LiteralPath $InputPath).DirectoryName }
    foreach ($file in $csvFiles) {
        $detection = @($detections | Where-Object { $_.File -eq $file.FullName })[0]
        $relative = if ($inputIsFolder) { $file.FullName.Substring($baseFolder.Length).TrimStart('\','/') } else { $file.Name }
        $targetPath = Join-Path $DataRoot $relative
        try {
            [void]$results.Add((Convert-OneCsvFile $file.FullName $targetPath $detection.DelimiterChar $SampleLines))
        }
        catch {
            [void]$failed.Add([pscustomobject]@{SourceCSV=$file.FullName;Error=$_.Exception.Message;Status='FAIL'})
            break
        }
    }

    $detectionReport | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Delimiter_Detection.csv') -NoTypeInformation -Encoding UTF8
    if ($failed.Count -gt 0) {
        $failed | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Normalization_Failures.csv') -NoTypeInformation -Encoding UTF8
        throw ('Normalization stopped after failure in {0}' -f $failed[0].SourceCSV)
    }
    if ($results.Count -ne $csvFiles.Count) { throw 'Normalized file count does not match source CSV count.' }
    $results | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Normalization_Register.csv') -NoTypeInformation -Encoding UTF8

    $summary = [ordered]@{
        Title='BHP INVENT | CP09A CSV Delimiter Normalization';Classification=$Classification;Version=$Version;RunId=$RunId
        InputPath=$InputPath;OutputRoot=$FinalRoot;SampleLines=$SampleLines;CSVFileCount=$csvFiles.Count
        NormalizedFileCount=$results.Count;OverwriteCount=0;Status='PASS'
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.txt') -Encoding UTF8
    $manifestFiles = @(Get-ChildItem -LiteralPath $RecordRoot -File | Where-Object Name -ne 'SHA256SUMS.txt')
    $manifest = foreach ($file in $manifestFiles) { $hash=Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256; '{0} *{1}' -f $hash.Hash,$file.Name }
    $manifest | Set-Content -LiteralPath (Join-Path $RecordRoot 'SHA256SUMS.txt') -Encoding ASCII

    Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot
    if (-not (Test-Path -LiteralPath $FinalRoot -PathType Container)) { throw 'Final promotion failed.' }

    Write-Host ''
    Write-Host 'CSV DELIMITER NORMALIZATION: PASS' -ForegroundColor Green
    Write-Host ('FILES: {0}/{1}' -f $results.Count,$csvFiles.Count) -ForegroundColor Green
    Write-Host ('OUTPUT: {0}' -f $FinalRoot) -ForegroundColor Cyan
    Write-Host ''
    exit 0
}
catch {
    try {
        if (-not (Test-Path -LiteralPath $RecordRoot -PathType Container)) { New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null }
        Write-RunLog $_.Exception.Message 'ERROR'
        @('MESSAGE: '+$_.Exception.Message,'SCRIPT: '+$_.InvocationInfo.ScriptName,'LINE: '+$_.InvocationInfo.ScriptLineNumber,'POSITION: '+$_.InvocationInfo.PositionMessage,'',($_ | Out-String)) | Set-Content -LiteralPath (Join-Path $RecordRoot 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: '+$_.Exception.Message) -ForegroundColor Red }
    exit 1
}
