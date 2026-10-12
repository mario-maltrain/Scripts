#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\00_Consolidacion_BD_GeoAI\CP09A-GEOAI-CONSOLIDATION-20261001-134313\Exports_2026-09-28\LAS LEP 2026-09-28',
    [string]$OutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\01_LAS_to_CSV',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.0.0'
$Classification = 'Internal'
$RunId = 'CP09A-LAS-TO-CSV-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$StagingRoot = Join-Path $OutputRoot ('.staging-' + $RunId)
$FinalRoot = Join-Path $OutputRoot $RunId
$DataRoot = Join-Path $StagingRoot 'CSV'
$RecordRoot = Join-Path $StagingRoot '_Conversion_Record'
$LogPath = Join-Path $RecordRoot 'Execution.log'
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\LAS_to_CSV'
$InstalledPath = Join-Path $InstalledFolder 'Convert_CP09A_LAS_to_CSV_v1_0_0.ps1'
$SourceScriptPath = $MyInvocation.MyCommand.Path

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

function Get-SafeCsvField {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + $Value.Replace('"','""') + '"'
}

function Get-UniqueColumnNames {
    param([string[]]$Names)
    $used = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($raw in $Names) {
        $name = if ([string]::IsNullOrWhiteSpace($raw)) { 'COLUMN' } else { $raw.Trim() }
        $name = [regex]::Replace($name,'[^A-Za-z0-9_\-]','_')
        if ([string]::IsNullOrWhiteSpace($name)) { $name = 'COLUMN' }
        $base = $name
        $suffix = 1
        while ($used.ContainsKey($name.ToUpperInvariant())) {
            $suffix++
            $name = '{0}_{1}' -f $base,$suffix
        }
        $used[$name.ToUpperInvariant()] = $true
        $result.Add($name)
    }
    return $result.ToArray()
}

function Read-LasDescriptor {
    param([string]$Path)
    $curveNames = New-Object System.Collections.Generic.List[string]
    $curveUnits = New-Object System.Collections.Generic.List[string]
    $section = ''
    $wrap = 'NO'
    $nullValue = ''
    $asciiFound = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('~')) {
            $sectionCode = $trimmed.Substring(1).Trim().ToUpperInvariant()
            if ($sectionCode.StartsWith('C')) { $section = 'CURVE'; continue }
            if ($sectionCode.StartsWith('V')) { $section = 'VERSION'; continue }
            if ($sectionCode.StartsWith('W')) { $section = 'WELL'; continue }
            if ($sectionCode.StartsWith('A')) { $asciiFound = $true; break }
            $section = 'OTHER'
            continue
        }
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) { continue }
        if ($section -eq 'VERSION' -and $trimmed -match '^\s*WRAP\s*\.\s*\S*\s+([^\s:]+)') { $wrap = $Matches[1].Trim().ToUpperInvariant() }
        if ($section -eq 'WELL' -and $trimmed -match '^\s*NULL\s*\.\s*\S*\s+([^\s:]+)') { $nullValue = $Matches[1].Trim() }
        if ($section -eq 'CURVE') {
            $left = ($trimmed -split ':',2)[0].Trim()
            if ($left -match '^([^\.\s]+)\s*\.\s*([^\s]*)') {
                $curveNames.Add($Matches[1].Trim())
                $curveUnits.Add($Matches[2].Trim())
            }
        }
    }
    if (-not $asciiFound) { throw ('ASCII section not found in {0}' -f $Path) }
    if ($curveNames.Count -eq 0) { throw ('No curve definitions found in {0}' -f $Path) }
    [pscustomobject]@{CurveNames=(Get-UniqueColumnNames $curveNames.ToArray());CurveUnits=$curveUnits.ToArray();Wrap=$wrap;NullValue=$nullValue;ColumnCount=$curveNames.Count}
}

function Convert-OneLasFile {
    param([string]$LasPath,[string]$CsvPath)
    $descriptor = Read-LasDescriptor $LasPath
    if ($descriptor.Wrap -notin @('NO','YES')) { throw ('Unsupported WRAP value {0} in {1}' -f $descriptor.Wrap,$LasPath) }
    $parent = Split-Path -Path $CsvPath -Parent
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    if (Test-Path -LiteralPath $CsvPath) { throw ('No-overwrite safety block: {0}' -f $CsvPath) }

    $utf8 = New-Object System.Text.UTF8Encoding($true)
    $writer = New-Object System.IO.StreamWriter($CsvPath,$false,$utf8)
    $reader = [System.IO.File]::OpenText($LasPath)
    $dataMode = $false
    $buffer = New-Object System.Collections.Generic.List[string]
    $rowCount = 0
    $invalidRows = 0
    try {
        $header = ($descriptor.CurveNames | ForEach-Object { Get-SafeCsvField $_ }) -join ','
        $writer.WriteLine($header)
        while ($null -ne ($line = $reader.ReadLine())) {
            $trimmed = $line.Trim()
            if (-not $dataMode) {
                if ($trimmed -match '^~A(?:SCII)?\b') { $dataMode = $true }
                continue
            }
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) { continue }
            if ($trimmed.StartsWith('~')) { break }
            $tokens = @($trimmed -split '\s+' | Where-Object { $_ -ne '' })
            foreach ($token in $tokens) { $buffer.Add($token) }
            while ($buffer.Count -ge $descriptor.ColumnCount) {
                $row = [string[]]::new($descriptor.ColumnCount)
                for ($i=0; $i -lt $descriptor.ColumnCount; $i++) { $row[$i] = $buffer[0]; $buffer.RemoveAt(0) }
                $writer.WriteLine((($row | ForEach-Object { Get-SafeCsvField $_ }) -join ','))
                $rowCount++
            }
        }
        if ($buffer.Count -gt 0) { $invalidRows = 1 }
    }
    finally {
        $reader.Dispose()
        $writer.Dispose()
    }
    if ($invalidRows -gt 0) { throw ('Incomplete final data row in {0}' -f $LasPath) }
    if ($rowCount -eq 0) { throw ('No data rows converted from {0}' -f $LasPath) }
    $csvLineCount = 0
    $checkReader = [System.IO.File]::OpenText($CsvPath)
    try { while ($null -ne $checkReader.ReadLine()) { $csvLineCount++ } } finally { $checkReader.Dispose() }
    if ($csvLineCount -ne ($rowCount + 1)) { throw ('CSV line-count validation failed for {0}' -f $CsvPath) }
    [pscustomobject]@{
        SourceLAS=$LasPath;OutputCSV=$CsvPath;Wrap=$descriptor.Wrap;NullValue=$descriptor.NullValue
        ColumnCount=$descriptor.ColumnCount;Columns=($descriptor.CurveNames -join '|');Units=($descriptor.CurveUnits -join '|')
        DataRows=$rowCount;CSVLines=$csvLineCount;SourceBytes=(Get-Item -LiteralPath $LasPath).Length
        OutputBytes=(Get-Item -LiteralPath $CsvPath).Length;SourceSHA256=(Get-FileHash -LiteralPath $LasPath -Algorithm SHA256).Hash
        OutputSHA256=(Get-FileHash -LiteralPath $CsvPath -Algorithm SHA256).Hash;Status='PASS'
    }
}

function Install-CurrentScript {
    if (-not (Test-Path -LiteralPath $SourceScriptPath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-ScriptSyntax $SourceScriptPath)) { throw 'Current script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    Copy-Item -LiteralPath $SourceScriptPath -Destination $InstalledPath -Force
    if (-not (Test-ScriptSyntax $InstalledPath)) { throw 'Installed script failed parser validation.' }
    foreach ($folder in @($InstalledFolder,(Join-Path $env:USERPROFILE 'Downloads'))) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        Get-ChildItem -LiteralPath $folder -Filter 'Convert_CP09A_LAS_to_CSV_v*.ps1' -File -ErrorAction SilentlyContinue | Where-Object {
            -not $_.FullName.Equals($InstalledPath,[System.StringComparison]::OrdinalIgnoreCase) -and
            -not $_.FullName.Equals($SourceScriptPath,[System.StringComparison]::OrdinalIgnoreCase)
        } | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force; Write-RunLog ('Removed obsolete converter: {0}' -f $_.FullName) 'WARN' }
    }
}

try {
    if (-not (Test-Path -LiteralPath $InputRoot -PathType Container)) { throw ('Input folder not found: {0}' -f $InputRoot) }
    if (Test-Path -LiteralPath $FinalRoot) { throw ('Final output already exists: {0}' -f $FinalRoot) }
    if (Test-Path -LiteralPath $StagingRoot) { throw ('Staging output already exists: {0}' -f $StagingRoot) }
    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null
    Write-RunLog ('BHP INVENT | CP09A LAS to CSV | Version {0} | Classification {1}' -f $Version,$Classification)
    Write-RunLog ('RUN_ID: {0}' -f $RunId)
    Write-RunLog ('INPUT: {0}' -f $InputRoot)
    Write-RunLog ('OUTPUT: {0}' -f $FinalRoot)

    Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Status 'Installing and validating script' -PercentComplete 3
    Install-CurrentScript
    Write-RunLog ('Installed script: {0}' -f $InstalledPath) 'PASS'

    $lasFiles = @(Get-ChildItem -LiteralPath $InputRoot -Recurse -File | Where-Object { $_.Extension -ieq '.las' } | Sort-Object FullName)
    if ($lasFiles.Count -eq 0) { throw 'No LAS files found.' }
    $sourceRelativeNames = @($lasFiles | ForEach-Object { $_.FullName.Substring($InputRoot.Length).TrimStart('\') })
    $duplicateOutputs = @($sourceRelativeNames | ForEach-Object { [System.IO.Path]::ChangeExtension($_,'.csv').ToLowerInvariant() } | Group-Object | Where-Object Count -gt 1)
    if ($duplicateOutputs.Count -gt 0) { throw 'Duplicate target CSV paths detected.' }

    $results = New-Object System.Collections.ArrayList
    $failed = New-Object System.Collections.ArrayList
    for ($index=0; $index -lt $lasFiles.Count; $index++) {
        $las = $lasFiles[$index]
        $relative = $las.FullName.Substring($InputRoot.Length).TrimStart('\')
        $csvRelative = [System.IO.Path]::ChangeExtension($relative,'.csv')
        $csvPath = Join-Path $DataRoot $csvRelative
        $percent = 5 + [int](80 * (($index + 1) / [double]$lasFiles.Count))
        Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Status 'Converting LAS without overwrite' -CurrentOperation ('File {0} of {1}: {2}' -f ($index+1),$lasFiles.Count,$las.Name) -PercentComplete $percent
        try { [void]$results.Add((Convert-OneLasFile $las.FullName $csvPath)) }
        catch { [void]$failed.Add([pscustomobject]@{SourceLAS=$las.FullName;Error=$_.Exception.Message;Status='FAIL'}); break }
    }
    if ($failed.Count -gt 0) {
        $failed | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Conversion_Failures.csv') -NoTypeInformation -Encoding UTF8
        throw ('Conversion stopped after failure in {0}' -f $failed[0].SourceLAS)
    }
    if ($results.Count -ne $lasFiles.Count) { throw 'Converted file count does not match LAS source count.' }

    Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Status 'Validating converted outputs' -PercentComplete 88
    $results | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Conversion_Register.csv') -NoTypeInformation -Encoding UTF8
    $outputFiles = @(Get-ChildItem -LiteralPath $DataRoot -Recurse -File -Filter '*.csv')
    if ($outputFiles.Count -ne $lasFiles.Count) { throw 'Output CSV count validation failed.' }
    if (@($results | Where-Object Status -ne 'PASS').Count -gt 0) { throw 'One or more conversion records failed.' }
    if (@($outputFiles | Where-Object Length -le 0).Count -gt 0) { throw 'One or more output CSV files are empty.' }

    $totalRows = [int64](($results | Measure-Object DataRows -Sum).Sum)
    $totalSourceBytes = [int64](($results | Measure-Object SourceBytes -Sum).Sum)
    $totalOutputBytes = [int64](($results | Measure-Object OutputBytes -Sum).Sum)
    $schemaSummary = @($results | Group-Object Columns | ForEach-Object { [pscustomobject]@{Columns=$_.Name;FileCount=$_.Count;ColumnCount=$_.Group[0].ColumnCount;Units=$_.Group[0].Units} })
    $schemaSummary | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Schema_Summary.csv') -NoTypeInformation -Encoding UTF8

    $summary = [ordered]@{
        Title='BHP INVENT | CP09A LAS to CSV Conversion';Classification=$Classification;Version=$Version;RunId=$RunId
        InputRoot=$InputRoot;OutputRoot=$FinalRoot;LASFileCount=$lasFiles.Count;CSVFileCount=$outputFiles.Count
        ConvertedDataRows=$totalRows;SourceBytes=$totalSourceBytes;OutputBytes=$totalOutputBytes
        FailedFiles=0;OverwriteCount=0;SchemaCount=$schemaSummary.Count;Status='PASS_REAL_TEST'
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.txt') -Encoding UTF8
    $manifestFiles = @(Get-ChildItem -LiteralPath $RecordRoot -File | Where-Object Name -ne 'SHA256SUMS.txt')
    $manifest = foreach ($file in $manifestFiles) { $hash=Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256; '{0} *{1}' -f $hash.Hash,$file.Name }
    $manifest | Set-Content -LiteralPath (Join-Path $RecordRoot 'SHA256SUMS.txt') -Encoding ASCII

    Write-RunLog 'RESULT REVIEW: PASS_REAL_TEST' 'PASS'
    Write-RunLog ('LAS FILES: {0}' -f $lasFiles.Count) 'PASS'
    Write-RunLog ('CSV FILES: {0}' -f $outputFiles.Count) 'PASS'
    Write-RunLog ('DATA ROWS: {0}' -f $totalRows) 'PASS'
    Write-RunLog 'OVERWRITES: 0' 'PASS'

    Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Status 'Promoting completed run' -PercentComplete 98
    Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot
    if (-not (Test-Path -LiteralPath $FinalRoot -PathType Container)) { throw 'Final promotion failed.' }
    Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Completed
    Write-Host ''
    Write-Host 'LAS TO CSV: PASS_REAL_TEST' -ForegroundColor Green
    Write-Host ('FILES: {0}/{1}' -f $outputFiles.Count,$lasFiles.Count) -ForegroundColor Green
    Write-Host 'OVERWRITES: 0' -ForegroundColor Green
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
    Write-Progress -Id 1 -Activity 'CP09A LAS to CSV' -Completed
    exit 1
}
