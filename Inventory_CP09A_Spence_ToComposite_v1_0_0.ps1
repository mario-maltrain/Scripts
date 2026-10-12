# Classification: Internal
# Script: Inventory_CP09A_Spence_ToComposite_v1_0_0.ps1
# Purpose: Read-only inventory of a source folder (default Spence 00_TO_COMPOSITE), column profiling and draft variable catalog.
# Safety: AuditOnly is the default. Sources are never modified. Records are written only with -CommitRecords to a new RUN_ID folder.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SourceRoot = 'C:\Python\GeoAi\00_Base_Datos\Spence\00_DataBase\00_TO_COMPOSITE',

    [Parameter()]
    [ValidateSet('SPENCE', 'GEOAI_LEP', 'GEOAI_AUTOLITH', 'GEOAI_AUTOLOGGER', 'VENDOR')]
    [string]$SourceId = 'SPENCE',

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-SOURCE-INVENTORY-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-SOURCE-INVENTORY-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [ValidateRange(10, 1000000)]
    [int]$MaxProfileRows = 5000,

    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$MaxDistinctValues = 50,

    [Parameter()]
    [switch]$CommitRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Inventory_CP09A_Spence_ToComposite_v1_0_0.ps1'
$ScriptVersion = '1.0.0'
$Classification = 'Internal'
$CatalogVersion = '0.1.0'
$ProgressId = 1600
$ExitCode = 0
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$NumberStyle = [System.Globalization.NumberStyles]::Float
$TextExtensions = @('.csv', '.txt', '.tsv')
$NullTokens = @('', 'NA', 'N/A', 'NULL', 'NAN', '-999', '-999.25', '-9999')

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-Delimiter {
    param([Parameter(Mandatory)][string]$Path)
    $reader = [System.IO.StreamReader]::new($Path, $true)
    try { $line = $reader.ReadLine() } finally { $reader.Dispose() }
    if ($null -eq $line) { return ',' }
    $candidates = @(
        [pscustomobject]@{ Delimiter = ','; Count = @($line.ToCharArray() | Where-Object { $_ -eq [char]44 }).Count }
        [pscustomobject]@{ Delimiter = ';'; Count = @($line.ToCharArray() | Where-Object { $_ -eq [char]59 }).Count }
        [pscustomobject]@{ Delimiter = [string][char]9; Count = @($line.ToCharArray() | Where-Object { $_ -eq [char]9 }).Count }
        [pscustomobject]@{ Delimiter = '|'; Count = @($line.ToCharArray() | Where-Object { $_ -eq [char]124 }).Count }
    )
    $best = @($candidates | Sort-Object Count -Descending)[0]
    if ($best.Count -eq 0) { return ',' }
    return $best.Delimiter
}

function Get-DelimiterName {
    param([string]$Delimiter)
    switch ($Delimiter) { ',' { 'COMMA' } ';' { 'SEMICOLON' } '|' { 'PIPE' } default { 'TAB' } }
}

function Get-ColumnRole {
    param([Parameter(Mandatory)][string]$Name)
    $n = $Name.Trim().ToUpperInvariant()
    if ($n -match '^(HOLE_?ID|HOLEID|BHID|DHID|UWI|HOLE)$') { return 'ID' }
    if ($n -match '^(FROM|FROM_M|DEPTH_FROM|MFROM|DEPTHFROM)$') { return 'FROM' }
    if ($n -match '^(TO|TO_M|DEPTH_TO|MTO|DEPTHTO)$') { return 'TO' }
    if ($n -match '^(DEPT|DEPTH|MD|DEPTH_M)$') { return 'DEPTH' }
    if ($n -match '(EXPORT|RUN_ID|SOURCE_FILE|DATE|TIME|COMMENT|REMARK|USER|LOGGED|_DESC$|DESCRIPTION|PROJECT|COMPANY)') { return 'METADATA' }
    return 'ATTRIBUTE'
}

function Get-CanonicalName {
    param([string]$Prefix, [string]$Name)
    $clean = [regex]::Replace($Name.Trim().ToUpperInvariant(), '[^A-Z0-9]+', '_').Trim('_')
    if ($clean.Length -eq 0) { $clean = 'UNNAMED' }
    return ('{0}_{1}' -f $Prefix, $clean)
}

function Invoke-TableProfile {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    $delimiter = Get-Delimiter -Path $File.FullName
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($File.FullName, [System.Text.Encoding]::UTF8, $true)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters([string[]]@($delimiter))
    $parser.HasFieldsEnclosedInQuotes = $true
    $parser.TrimWhiteSpace = $false

    $header = @()
    $rows = 0L
    $badRows = 0L
    $firstRow = ''
    $stats = @()
    try {
        if ($parser.EndOfData) { throw 'Empty file.' }
        $header = @($parser.ReadFields() | ForEach-Object { $_.Trim() })
        for ($c = 0; $c -lt $header.Count; $c++) {
            $stats += [pscustomobject]@{
                Name = $header[$c]
                NonNull = 0L
                Nulls = 0L
                Numeric = 0L
                Min = [double]::PositiveInfinity
                Max = [double]::NegativeInfinity
                Distinct = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
                Sample = ''
            }
        }
        while (-not $parser.EndOfData) {
            $fields = $null
            try { $fields = $parser.ReadFields() }
            catch { $badRows++; continue }
            if ($null -eq $fields) { continue }
            $rows++
            if ($fields.Count -ne $header.Count) { $badRows++; continue }
            if ($rows -eq 1) { $firstRow = $fields -join ' | ' }
            if ($rows -gt $MaxProfileRows) { continue }
            for ($c = 0; $c -lt $header.Count; $c++) {
                $value = $fields[$c].Trim()
                $s = $stats[$c]
                if ($NullTokens -contains $value.ToUpperInvariant()) { $s.Nulls++; continue }
                $s.NonNull++
                if ($s.Sample.Length -eq 0) { $s.Sample = $value }
                if ($s.Distinct.Count -le $MaxDistinctValues) { [void]$s.Distinct.Add($value) }
                $number = 0.0
                if ([double]::TryParse($value, $NumberStyle, $Invariant, [ref]$number)) {
                    $s.Numeric++
                    if ($number -lt $s.Min) { $s.Min = $number }
                    if ($number -gt $s.Max) { $s.Max = $number }
                }
            }
        }
    }
    finally {
        $parser.Dispose()
    }

    $roles = @($header | ForEach-Object { Get-ColumnRole -Name $_ })
    $isInterval = ($roles -contains 'FROM') -and ($roles -contains 'TO')
    $isPoint = (-not $isInterval) -and ($roles -contains 'DEPTH')
    $tableType = if ($isInterval) { 'INTERVAL' } elseif ($isPoint) { 'POINT' } else { 'UNKNOWN' }

    $columns = [System.Collections.Generic.List[object]]::new()
    for ($c = 0; $c -lt $header.Count; $c++) {
        $s = $stats[$c]
        $numericRatio = if ($s.NonNull -gt 0) { $s.Numeric / $s.NonNull } else { 0.0 }
        $dataType = if ($s.NonNull -eq 0) { 'EMPTY' } elseif ($numericRatio -ge 0.98) { 'NUMERIC' } else { 'TEXT' }
        $role = $roles[$c]
        if ($role -eq 'ATTRIBUTE') { $role = if ($dataType -eq 'NUMERIC') { 'VALUE' } elseif ($dataType -eq 'TEXT') { 'CATEGORY' } else { 'METADATA' } }
        $method = switch ($role) {
            'VALUE' { if ($tableType -eq 'INTERVAL') { 'LENGTH_WEIGHTED_MEAN' } else { 'MEAN' } }
            'CATEGORY' { 'DOMINANT_CODE' }
            default { 'NONE' }
        }
        $columns.Add([pscustomobject]@{
            File = $File.Name
            TableType = $tableType
            Ordinal = $c + 1
            SourceColumn = $s.Name
            InferredRole = $role
            InferredDataType = $dataType
            SuggestedMethod = $method
            ProfiledRows = [math]::Min($rows, $MaxProfileRows)
            NonNull = $s.NonNull
            Nulls = $s.Nulls
            NumericRatio = [math]::Round($numericRatio, 4)
            Min = $(if ($s.Numeric -gt 0) { $s.Min.ToString('R', $Invariant) } else { '' })
            Max = $(if ($s.Numeric -gt 0) { $s.Max.ToString('R', $Invariant) } else { '' })
            DistinctValues = $(if ($s.Distinct.Count -gt $MaxDistinctValues) { ('>{0}' -f $MaxDistinctValues) } else { [string]$s.Distinct.Count })
            SampleValue = $s.Sample
        })
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    if (@($header | Group-Object { $_.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 }).Count -gt 0) { $issues.Add('DUPLICATE_HEADERS') }
    if (@($header | Where-Object { $_.Length -eq 0 }).Count -gt 0) { $issues.Add('EMPTY_HEADER') }
    if (-not ($roles -contains 'ID')) { $issues.Add('NO_HOLE_ID_COLUMN') }
    if ($tableType -eq 'UNKNOWN') { $issues.Add('NO_DEPTH_OR_INTERVAL_COLUMNS') }
    if ($badRows -gt 0) { $issues.Add(('MALFORMED_ROWS={0}' -f $badRows)) }

    return [pscustomobject]@{
        Delimiter = (Get-DelimiterName -Delimiter $delimiter)
        Header = ($header -join '|')
        ColumnCount = $header.Count
        Rows = $rows
        BadRows = $badRows
        TableType = $tableType
        FirstRow = $firstRow
        Issues = ($issues -join ';')
        Columns = $columns
    }
}

try {
    $Started = Get-Date
    Write-Progress -Id $ProgressId -Activity 'CP09A source inventory' -Status 'Validating source root' -PercentComplete 2

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) { throw "SourceRoot does not exist: $SourceRoot" }
    $ResolvedSource = (Get-Item -LiteralPath $SourceRoot).FullName
    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) { throw 'OutputRoot is required when CommitRecords is used.' }
    if (-not $AuditOnly) {
        $ResolvedOutput = [System.IO.Path]::GetFullPath($OutputRoot)
        if ($ResolvedOutput.StartsWith($ResolvedSource, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'OutputRoot cannot be inside SourceRoot.' }
    }

    $Files = @(Get-ChildItem -LiteralPath $ResolvedSource -File -Recurse -Force | Sort-Object FullName)
    if ($Files.Count -eq 0) { throw "No files found under: $ResolvedSource" }

    $HashBefore = @{}
    $FileRows = [System.Collections.Generic.List[object]]::new()
    $ColumnRows = [System.Collections.Generic.List[object]]::new()
    $ZipRows = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $Files.Count; $i++) {
        $file = $Files[$i]
        Write-Progress -Id $ProgressId -Activity 'CP09A source inventory' -Status ('{0} of {1}: {2}' -f ($i + 1), $Files.Count, $file.Name) -PercentComplete ([math]::Min(90, 5 + [int][math]::Floor($i / $Files.Count / 0.0118)))
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $HashBefore[$file.FullName] = $hash
        $extension = $file.Extension.ToLowerInvariant()
        $relative = [System.IO.Path]::GetRelativePath($ResolvedSource, $file.FullName)
        $record = [ordered]@{
            SourceId = $SourceId
            RelativePath = $relative
            Extension = $extension
            SizeMB = [math]::Round($file.Length / 1MB, 3)
            LastWriteTime = $file.LastWriteTime.ToString('s')
            SHA256 = $hash
            Kind = 'OTHER'
            Delimiter = ''
            TableType = ''
            Columns = ''
            Rows = ''
            BadRows = ''
            Header = ''
            FirstRow = ''
            Status = 'INFO'
            Issues = ''
        }
        try {
            if ($TextExtensions -contains $extension) {
                $tableProfile = Invoke-TableProfile -File $file
                $record.Kind = 'TABLE'
                $record.Delimiter = $tableProfile.Delimiter
                $record.TableType = $tableProfile.TableType
                $record.Columns = $tableProfile.ColumnCount
                $record.Rows = $tableProfile.Rows
                $record.BadRows = $tableProfile.BadRows
                $record.Header = $tableProfile.Header
                $record.FirstRow = $tableProfile.FirstRow
                $record.Issues = $tableProfile.Issues
                $record.Status = if ($tableProfile.Issues.Length -eq 0) { 'PASS' } else { 'REVIEW' }
                foreach ($column in $tableProfile.Columns) { $ColumnRows.Add($column) }
            }
            elseif ($extension -eq '.zip') {
                $record.Kind = 'ARCHIVE'
                $archive = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
                try {
                    foreach ($entry in $archive.Entries) {
                        if ($entry.FullName.EndsWith('/')) { continue }
                        $ZipRows.Add([pscustomobject]@{ Archive = $relative; Entry = $entry.FullName; Extension = [System.IO.Path]::GetExtension($entry.FullName).ToLowerInvariant(); Length = $entry.Length })
                    }
                    $record.Rows = @($archive.Entries | Where-Object { -not $_.FullName.EndsWith('/') }).Count
                }
                finally { $archive.Dispose() }
                $record.Status = 'REVIEW'
                $record.Issues = 'ARCHIVE_REQUIRES_VENDOR_INTAKE'
            }
            elseif ($extension -in @('.xlsx', '.xls', '.xlsm')) {
                $record.Kind = 'WORKBOOK'
                $record.Status = 'REVIEW'
                $record.Issues = 'WORKBOOK_REQUIRES_CSV_EXPORT'
            }
            elseif ($extension -eq '.las') {
                $record.Kind = 'LAS'
                $record.Status = 'REVIEW'
                $record.Issues = 'LAS_REQUIRES_CONVERSION'
            }
        }
        catch {
            $record.Status = 'FAIL'
            $record.Issues = ('READ_ERROR: {0}' -f $_.Exception.Message)
        }
        $FileRows.Add([pscustomobject]$record)
    }

    Write-Progress -Id $ProgressId -Activity 'CP09A source inventory' -Status 'Building draft catalog' -PercentComplete 92

    $prefix = switch ($SourceId) { 'SPENCE' { 'SP' } 'GEOAI_LEP' { 'LEP' } 'GEOAI_AUTOLITH' { 'AL' } 'GEOAI_AUTOLOGGER' { 'ALG' } default { 'VND' } }
    $Catalog = [System.Collections.Generic.List[object]]::new()
    foreach ($column in $ColumnRows) {
        $tableName = [System.IO.Path]::GetFileNameWithoutExtension($column.File)
        $enabled = if ($column.InferredRole -in @('ID', 'FROM', 'TO', 'DEPTH', 'VALUE', 'CATEGORY')) { 'Y' } else { 'N' }
        $required = if ($column.InferredRole -in @('ID', 'FROM', 'TO', 'DEPTH')) { 'Y' } else { 'N' }
        $Catalog.Add([pscustomobject]@{
            CatalogVersion = $CatalogVersion
            VariableId = ('{0}-{1:000}-{2:00}' -f $prefix, ($Catalog.Count + 1), $column.Ordinal)
            SourceId = $SourceId
            Table = $tableName
            SourceColumn = $column.SourceColumn
            CanonicalName = $(if ($column.InferredRole -in @('ID', 'FROM', 'TO', 'DEPTH')) { @{ ID = 'HOLE_ID'; FROM = 'FROM_M'; TO = 'TO_M'; DEPTH = 'DEPTH_M' }[$column.InferredRole] } else { Get-CanonicalName -Prefix $prefix -Name $column.SourceColumn })
            Role = $column.InferredRole
            DataType = $column.InferredDataType
            Unit = ''
            CompositeMethod = $column.SuggestedMethod
            ArgmaxGroup = ''
            NullValues = '-999.25|-999|empty'
            ValidMin = $column.Min
            ValidMax = $column.Max
            Required = $required
            Enabled = $enabled
            Status = 'DRAFT'
            AddedInVersion = $CatalogVersion
            RetiredInVersion = ''
            Owner = 'Requires Validation'
            Notes = 'Auto-inferred from profile; validate role, unit, method and range'
        })
    }

    $Changed = 0
    foreach ($path in $HashBefore.Keys) { if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $HashBefore[$path]) { $Changed++ } }

    $Tables = @($FileRows | Where-Object { $_.Kind -eq 'TABLE' })
    $Failures = @($FileRows | Where-Object { $_.Status -eq 'FAIL' })
    $Reviews = @($FileRows | Where-Object { $_.Status -eq 'REVIEW' })
    $Decision = if ($Failures.Count -gt 0 -or $Changed -gt 0) { 'NO_GO' } elseif ($Reviews.Count -gt 0 -or $Tables.Count -eq 0) { 'REVIEW_REQUIRED' } else { 'PASS' }

    Write-Host ''
    Write-Host ('CP09A SOURCE INVENTORY v{0}' -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('SOURCE: {0}  [{1}]' -f $ResolvedSource, $SourceId)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' }))
    Write-Host ('FILES: {0} | TABLES: {1} | ARCHIVES: {2} | COLUMNS: {3} | ZIP ENTRIES: {4}' -f $Files.Count, $Tables.Count, @($FileRows | Where-Object { $_.Kind -eq 'ARCHIVE' }).Count, $ColumnRows.Count, $ZipRows.Count)
    Write-Host ''
    $FileRows | Format-Table RelativePath, Kind, Delimiter, TableType, Columns, Rows, Status, Issues -AutoSize -Wrap
    if ($Tables.Count -gt 0) {
        foreach ($table in $Tables) {
            Write-Host ''
            Write-Host ('TABLE : {0}' -f $table.RelativePath) -ForegroundColor Cyan
            Write-Host ('HEADER: {0}' -f $table.Header)
            Write-Host ('ROW 1 : {0}' -f $table.FirstRow)
        }
        Write-Host ''
        $ColumnRows | Format-Table File, SourceColumn, InferredRole, InferredDataType, SuggestedMethod, NonNull, Nulls, Min, Max, DistinctValues -AutoSize -Wrap
    }
    if ($ZipRows.Count -gt 0) { $ZipRows | Select-Object -First 40 | Format-Table Archive, Entry, Length -AutoSize -Wrap }

    if (-not $AuditOnly) {
        $RunOutput = Join-Path $ResolvedOutput $RunId
        if (Test-Path -LiteralPath $RunOutput) { throw "Overwrite is prohibited. Run output already exists: $RunOutput" }
        if ($PSCmdlet.ShouldProcess($RunOutput, 'Create inventory records')) {
            New-Item -ItemType Directory -Path $RunOutput | Out-Null
            $FileRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Source_File_Inventory.csv') -NoTypeInformation -Encoding utf8NoBOM
            $ColumnRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Source_Column_Profile.csv') -NoTypeInformation -Encoding utf8NoBOM
            $Catalog | Export-Csv -LiteralPath (Join-Path $RunOutput ('Variable_Catalog_DRAFT_v{0}.csv' -f $CatalogVersion)) -NoTypeInformation -Encoding utf8NoBOM
            if ($ZipRows.Count -gt 0) { $ZipRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Archive_Entry_Inventory.csv') -NoTypeInformation -Encoding utf8NoBOM }
            ([ordered]@{ Classification = $Classification; ScriptName = $ScriptName; ScriptVersion = $ScriptVersion; RunId = $RunId; SourceId = $SourceId; SourceRoot = $ResolvedSource; Files = $Files.Count; Tables = $Tables.Count; Columns = $ColumnRows.Count; CatalogVersion = $CatalogVersion; SourceModified = ($Changed -gt 0); Decision = $Decision; ElapsedSeconds = [math]::Round(((Get-Date) - $Started).TotalSeconds, 1); CreatedAt = (Get-Date).ToString('o') } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $RunOutput 'Run_Register.json') -Encoding utf8NoBOM
            $hashLines = @(Get-ChildItem -LiteralPath $RunOutput -File | Sort-Object Name | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, $_.Name })
            ($hashLines -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Encoding ascii
            Write-Host ('RECORDS: {0}' -f $RunOutput) -ForegroundColor Green
        }
    }
    else {
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
    }

    Write-Host ('SOURCE UNCHANGED: {0}' -f ($Changed -eq 0))
    $color = if ($Decision -eq 'PASS') { 'Green' } elseif ($Decision -eq 'REVIEW_REQUIRED') { 'Yellow' } else { 'Red' }
    Write-Host ('DECISION: {0}' -f $Decision) -ForegroundColor $color
    if ($Decision -eq 'NO_GO') { $ExitCode = 1 }
}
catch {
    Write-Host ''
    Write-Host ('FAIL: {0}' -f $_.Exception.Message) -ForegroundColor Red
    $ExitCode = 1
}
finally {
    Write-Progress -Id $ProgressId -Activity 'CP09A source inventory' -Completed
}

exit $ExitCode
