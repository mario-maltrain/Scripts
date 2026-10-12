# Classification: Internal
# Script: Run_CP09A_Spence_Composite_v1_0_1.ps1
# Purpose: Composite Spence geology and assay interval tables at 1m, 3m and 6m, filtered to the vendor-analysed drillholes.
# Rules: Only interval tables with HOLEID, FROM and TO are composited. Long free text is excluded.
#        Assay numeric values use length-weighted means. Geology codes use the length-dominant code.
# Safety: AuditOnly is the default. Sources are read-only and verified by SHA256. No delete, move, rename or overwrite.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SourceRoot = 'C:\Python\GeoAi\00_Base_Datos\Spence\00_DataBase\00_TO_COMPOSITE',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$HoleListPath = (Join-Path $env:USERPROFILE 'Downloads\Vendor_Hole_List.csv'),

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-SPENCE-COMPOSITE-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-SPENCE-COMPOSITE-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [ValidateRange(0.0, 1.0)]
    [double]$MinimumCoverage = 0.8,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int]$MaxCodeLength = 15,

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$MaxDistinctCodes = 200,

    [Parameter()]
    [AllowEmptyString()]
    [string]$SupplementaryCollarPath = '',

    [Parameter()]
    [switch]$StopOnMissingCollar,

    [Parameter()]
    [switch]$CommitRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Run_CP09A_Spence_Composite_v1_0_1.ps1'
$ScriptVersion = '1.0.1'
$Classification = 'Internal'
$ProgressId = 1700
$ExitCode = 0
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$NumberStyle = [System.Globalization.NumberStyles]::Float
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$Scales = [double[]]@(1.0, 3.0, 6.0)
$ScaleNames = @('1m', '3m', '6m')
$NullTokens = @('', '-99', '-999', '-999.25', '-9999', 'NA', 'N/A', 'NULL', 'NAN')
$SkipColumns = @('PROJECTCODE', 'SAMPLEID')
$HoleNames = @('HOLEID', 'HOLE_ID')
$FreeTextPattern = '(?i)(obs|desc|comment|observ|remark)'
$Tiny = 0.000000001

Add-Type -AssemblyName Microsoft.VisualBasic
[System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
$Windows1252 = [System.Text.Encoding]::GetEncoding(1252)
$StrictUtf8 = [System.Text.UTF8Encoding]::new($false, $true)

function Get-Product {
    param([double]$A, [double]$B)
    return [math]::FusedMultiplyAdd($A, $B, 0.0)
}

function Format-Number {
    param([double]$Value)
    if ([double]::IsNaN($Value)) { return '' }
    return $Value.ToString('R', $Invariant)
}

function ConvertTo-CsvField {
    param([AllowEmptyString()][string]$Text)
    if ($Text.IndexOf(',') -ge 0 -or $Text.IndexOf([char]34) -ge 0) {
        return ([string][char]34) + $Text.Replace([string][char]34, ([string][char]34) + ([string][char]34)) + ([string][char]34)
    }
    return $Text
}

function Test-NullToken {
    param([AllowEmptyString()][string]$Text)
    return ($NullTokens -contains $Text.Trim().ToUpperInvariant())
}

function ConvertTo-NumberOrNaN {
    param([AllowEmptyString()][string]$Text)
    $clean = $Text.Trim()
    if (Test-NullToken -Text $clean) { return [double]::NaN }
    $number = 0.0
    if ([double]::TryParse($clean, $NumberStyle, $Invariant, [ref]$number)) { return $number }
    return [double]::NaN
}

function Get-FileEncoding {
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) { return [pscustomobject]@{ Name = 'UTF-8-BOM'; Encoding = [System.Text.Encoding]::UTF8 } }
    try {
        [void]$StrictUtf8.GetString($bytes)
        return [pscustomobject]@{ Name = 'UTF-8'; Encoding = [System.Text.Encoding]::UTF8 }
    }
    catch {
        return [pscustomobject]@{ Name = 'WINDOWS-1252'; Encoding = $Windows1252 }
    }
}

function Find-ColumnIndex {
    param([string[]]$Header, [string[]]$Names)
    for ($c = 0; $c -lt $Header.Count; $c++) {
        if ($Names -contains $Header[$c].ToUpperInvariant()) { return $c }
    }
    return -1
}

function Read-SpenceTable {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][System.Text.Encoding]$Encoding,
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]]$HoleSet
    )
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($File.FullName, $Encoding)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters([string[]]@(','))
    $parser.HasFieldsEnclosedInQuotes = $true
    $rows = [System.Collections.Generic.List[string[]]]::new()
    $total = 0L
    $malformed = 0L
    $filteredOut = 0L
    $malformedLines = [System.Collections.Generic.List[string]]::new()
    $holesFound = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        if ($parser.EndOfData) { throw ('Empty file: {0}' -f $File.Name) }
        $header = @($parser.ReadFields() | ForEach-Object { $_.Trim() })
        $holeIndex = Find-ColumnIndex -Header $header -Names $HoleNames
        $fromIndex = Find-ColumnIndex -Header $header -Names @('FROM')
        $toIndex = Find-ColumnIndex -Header $header -Names @('TO')
        while (-not $parser.EndOfData) {
            $lineNumber = $parser.LineNumber
            $fields = $null
            try { $fields = $parser.ReadFields() }
            catch { $malformed++; if ($malformedLines.Count -lt 50) { $malformedLines.Add(('{0}|PARSE_ERROR' -f $lineNumber)) }; continue }
            if ($null -eq $fields) { continue }
            $total++
            if ($fields.Count -ne $header.Count) { $malformed++; if ($malformedLines.Count -lt 50) { $malformedLines.Add(('{0}|FIELDS={1}' -f $lineNumber, $fields.Count)) }; continue }
            if ($holeIndex -lt 0) { continue }
            $hole = $fields[$holeIndex].Trim().ToUpperInvariant()
            if (-not $HoleSet.Contains($hole)) { $filteredOut++; continue }
            [void]$holesFound.Add($hole)
            $rows.Add($fields)
        }
    }
    finally {
        $parser.Dispose()
    }
    $tableType = if ($holeIndex -ge 0 -and $fromIndex -ge 0 -and $toIndex -ge 0) { 'INTERVAL' } elseif ($holeIndex -ge 0) { 'REFERENCE' } else { 'UNKNOWN' }
    return [pscustomobject]@{
        Header = $header; HoleIndex = $holeIndex; FromIndex = $fromIndex; ToIndex = $toIndex
        TableType = $tableType; Rows = $rows; TotalRows = $total; Malformed = $malformed
        FilteredOut = $filteredOut; MalformedLines = $malformedLines; HolesFound = $holesFound
    }
}

function Get-ColumnClassification {
    param([Parameter(Mandatory)][object]$Table, [Parameter(Mandatory)][string]$TableName)
    $isAssay = ($TableName -match '(?i)ASSAY')
    $result = [System.Collections.Generic.List[object]]::new()
    for ($c = 0; $c -lt $Table.Header.Count; $c++) {
        if ($c -eq $Table.HoleIndex -or $c -eq $Table.FromIndex -or $c -eq $Table.ToIndex) { continue }
        $name = $Table.Header[$c]
        if ($SkipColumns -contains $name.ToUpperInvariant()) {
            $result.Add([pscustomobject]@{ Table = $TableName; Column = $name; Index = $c; Class = 'EXCLUDED_ID_METADATA'; Method = 'NONE'; NonNull = 0; Numeric = 0; Distinct = 0; MaxLength = 0; Reason = 'Identifier or metadata column' })
            continue
        }
        $nonNull = 0L; $numeric = 0L; $integers = 0L; $maxLength = 0; $hasSpace = $false
        $distinct = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        foreach ($row in $Table.Rows) {
            $value = $row[$c].Trim()
            if (Test-NullToken -Text $value) { continue }
            $nonNull++
            if ($value.Length -gt $maxLength) { $maxLength = $value.Length }
            if ($value.IndexOf(' ') -ge 0) { $hasSpace = $true }
            if ($distinct.Count -le $MaxDistinctCodes) { [void]$distinct.Add($value) }
            $number = 0.0
            if ([double]::TryParse($value, $NumberStyle, $Invariant, [ref]$number)) {
                $numeric++
                if ([math]::Abs($number - [math]::Round($number)) -lt $Tiny) { $integers++ }
            }
        }
        $class = ''; $method = 'NONE'; $reason = ''
        if ($nonNull -eq 0) { $class = 'EXCLUDED_EMPTY'; $reason = 'No non-null values for vendor holes' }
        elseif ($name -match $FreeTextPattern) { $class = 'EXCLUDED_TEXT'; $reason = 'Free-text column name pattern' }
        elseif ($numeric -ge [math]::Ceiling($nonNull - $nonNull / 50.0)) {
            if ((-not $isAssay) -and $integers -eq $numeric -and $distinct.Count -le 20) { $class = 'CATEGORY_ORDINAL'; $method = 'DOMINANT_CODE'; $reason = 'Integer geology code with 20 or fewer values' }
            else { $class = 'NUMERIC'; $method = 'LENGTH_WEIGHTED_MEAN'; $reason = 'Numeric values' }
        }
        elseif ($hasSpace -or $maxLength -gt $MaxCodeLength -or $distinct.Count -gt $MaxDistinctCodes) { $class = 'EXCLUDED_TEXT'; $reason = ('Long or free text: maxLength={0}; spaces={1}; distinct={2}' -f $maxLength, $hasSpace, $distinct.Count) }
        else { $class = 'CATEGORY'; $method = 'DOMINANT_CODE'; $reason = 'Short geology code' }
        $result.Add([pscustomobject]@{ Table = $TableName; Column = $name; Index = $c; Class = $class; Method = $method; NonNull = $nonNull; Numeric = $numeric; Distinct = $distinct.Count; MaxLength = $maxLength; Reason = $reason })
    }
    return $result
}

function Invoke-TableComposite {
    param(
        [Parameter(Mandatory)][object]$Table,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$NumericColumns,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$CategoryColumns,
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][string]$SourceFile
    )
    $nNum = $NumericColumns.Count
    $nCat = $CategoryColumns.Count
    $stores = @{}
    $invalid = 0L
    $byHole = @{}
    foreach ($row in $Table.Rows) {
        $hole = $row[$Table.HoleIndex].Trim().ToUpperInvariant()
        $f = ConvertTo-NumberOrNaN -Text $row[$Table.FromIndex]
        $t = ConvertTo-NumberOrNaN -Text $row[$Table.ToIndex]
        if ([double]::IsNaN($f) -or [double]::IsNaN($t) -or $t -le $f) { $invalid++; continue }
        if (-not $byHole.ContainsKey($hole)) { $byHole[$hole] = [System.Collections.Generic.List[double[]]]::new() }
        $byHole[$hole].Add([double[]]@($f, $t))
        if (-not $stores.ContainsKey($hole)) {
            $scaleStores = [object[]]::new($Scales.Count)
            for ($s = 0; $s -lt $Scales.Count; $s++) { $scaleStores[$s] = [System.Collections.Generic.Dictionary[long, object]]::new() }
            $stores[$hole] = $scaleStores
        }
        $numValues = [double[]]::new($nNum)
        for ($k = 0; $k -lt $nNum; $k++) { $numValues[$k] = ConvertTo-NumberOrNaN -Text $row[$NumericColumns[$k].Index] }
        $catValues = [string[]]::new($nCat)
        for ($k = 0; $k -lt $nCat; $k++) {
            $code = $row[$CategoryColumns[$k].Index].Trim()
            $catValues[$k] = if (Test-NullToken -Text $code) { '' } else { $code }
        }
        for ($s = 0; $s -lt $Scales.Count; $s++) {
            $scale = $Scales[$s]
            $first = [long][math]::Floor($f / $scale + $Tiny)
            $last = [long][math]::Floor(($t - $Tiny) / $scale)
            for ($bin = $first; $bin -le $last; $bin++) {
                $binFrom = Get-Product -A ([double]$bin) -B $scale
                $binTo = Get-Product -A ([double]($bin + 1)) -B $scale
                $overlap = [math]::Min($t, $binTo) - [math]::Max($f, $binFrom)
                if ($overlap -le $Tiny) { continue }
                $acc = $null
                if (-not $stores[$hole][$s].TryGetValue($bin, [ref]$acc)) {
                    $cats = [object[]]::new($nCat)
                    for ($k = 0; $k -lt $nCat; $k++) { $cats[$k] = @{} }
                    $acc = [pscustomobject]@{ Covered = 0.0; Sum = [double[]]::new($nNum); Len = [double[]]::new($nNum); Cats = $cats }
                    $stores[$hole][$s][$bin] = $acc
                }
                $acc.Covered += $overlap
                for ($k = 0; $k -lt $nNum; $k++) {
                    if (-not [double]::IsNaN($numValues[$k])) {
                        $acc.Sum[$k] += (Get-Product -A $numValues[$k] -B $overlap)
                        $acc.Len[$k] += $overlap
                    }
                }
                for ($k = 0; $k -lt $nCat; $k++) {
                    if ($catValues[$k].Length -eq 0) { continue }
                    $h = $acc.Cats[$k]
                    if ($h.ContainsKey($catValues[$k])) { $h[$catValues[$k]] += $overlap } else { $h[$catValues[$k]] = $overlap }
                }
            }
        }
    }

    $overlaps = 0L
    foreach ($hole in $byHole.Keys) {
        $sorted = @($byHole[$hole] | Sort-Object { $_[0] })
        for ($k = 1; $k -lt $sorted.Count; $k++) { if ($sorted[$k][0] -lt $sorted[$k - 1][1] - 0.000001) { $overlaps++ } }
    }

    $header = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @('DATA_TYPE', 'TABLE', 'HOLE_ID', 'FROM', 'TO', 'CENTRE', 'SCALE_M', 'COVERED_LENGTH', 'COVERAGE_RATIO', 'PARTIAL_COMPOSITE', 'STATUS')) { $header.Add($name) }
    foreach ($col in $NumericColumns) { $header.Add(($col.Column + '_LWM')); $header.Add(($col.Column + '_LEN')) }
    foreach ($col in $CategoryColumns) { $header.Add(($col.Column + '_DOM')); $header.Add(($col.Column + '_FRAC')); $header.Add(($col.Column + '_TIE')) }
    $header.Add('SOURCE_FILE')

    $results = [System.Collections.Generic.List[object]]::new()
    for ($s = 0; $s -lt $Scales.Count; $s++) {
        $scale = $Scales[$s]
        $outputPath = Join-Path $OutputFolder ('SPENCE_{0}_Composite_{1}.csv' -f $TableName, $ScaleNames[$s])
        $rows = 0L; $pass = 0L; $partial = 0L; $ties = 0L
        $writer = [System.IO.StreamWriter]::new($outputPath, $false, $Utf8NoBom)
        try {
            $writer.WriteLine(($header -join ','))
            foreach ($hole in @($stores.Keys | Sort-Object)) {
                $dictionary = $stores[$hole][$s]
                foreach ($bin in @($dictionary.Keys | Sort-Object)) {
                    $acc = $dictionary[$bin]
                    $from = Get-Product -A ([double]$bin) -B $scale
                    $to = Get-Product -A ([double]($bin + 1)) -B $scale
                    $coverage = $acc.Covered / $scale
                    $status = if ($coverage -ge $MinimumCoverage) { 'PASS' } else { 'PARTIAL' }
                    $parts = [System.Collections.Generic.List[string]]::new()
                    $parts.Add('SPENCE'); $parts.Add($TableName); $parts.Add($hole)
                    $parts.Add((Format-Number $from)); $parts.Add((Format-Number $to)); $parts.Add((Format-Number (($from + $to) / 2)))
                    $parts.Add((Format-Number $scale)); $parts.Add((Format-Number $acc.Covered)); $parts.Add((Format-Number $coverage))
                    $parts.Add($(if ($coverage -lt 1.0) { 'True' } else { 'False' })); $parts.Add($status)
                    for ($k = 0; $k -lt $nNum; $k++) {
                        $mean = if ($acc.Len[$k] -gt 0) { $acc.Sum[$k] / $acc.Len[$k] } else { [double]::NaN }
                        $parts.Add((Format-Number $mean)); $parts.Add((Format-Number $acc.Len[$k]))
                    }
                    for ($k = 0; $k -lt $nCat; $k++) {
                        $h = $acc.Cats[$k]
                        $bestCode = ''; $bestLength = -1.0; $tie = $false; $codedLength = 0.0
                        foreach ($code in @($h.Keys | Sort-Object)) {
                            $length = [double]$h[$code]
                            $codedLength += $length
                            if ($length -gt $bestLength + $Tiny) { $bestCode = $code; $bestLength = $length; $tie = $false }
                            elseif ([math]::Abs($length - $bestLength) -le $Tiny) { $tie = $true }
                        }
                        if ($bestLength -lt 0) { $parts.Add(''); $parts.Add(''); $parts.Add('False') }
                        else {
                            $parts.Add((ConvertTo-CsvField -Text $bestCode)); $parts.Add((Format-Number ($bestLength / $acc.Covered))); $parts.Add($(if ($tie) { 'True' } else { 'False' }))
                            if ($tie) { $ties++ }
                        }
                    }
                    $parts.Add((ConvertTo-CsvField -Text $SourceFile))
                    $writer.WriteLine(($parts -join ','))
                    $rows++
                    if ($status -eq 'PASS') { $pass++ } else { $partial++ }
                }
            }
        }
        finally {
            $writer.Dispose()
        }
        $results.Add([pscustomobject]@{ Table = $TableName; Scale = $ScaleNames[$s]; Holes = $stores.Count; Rows = $rows; PassRows = $pass; PartialRows = $partial; CategoryTies = $ties; OutputFile = $outputPath; SHA256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash })
    }
    return [pscustomobject]@{ Results = $results; InvalidIntervals = $invalid; Overlaps = $overlaps }
}

try {
    $Started = Get-Date
    Write-Progress -Id $ProgressId -Activity 'CP09A Spence composite' -Status 'Validating inputs' -PercentComplete 2

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) { throw "SourceRoot does not exist: $SourceRoot" }
    if (-not (Test-Path -LiteralPath $HoleListPath -PathType Leaf)) { throw "Vendor hole list does not exist: $HoleListPath" }
    $ResolvedSource = (Get-Item -LiteralPath $SourceRoot).FullName
    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) { throw 'OutputRoot is required when CommitRecords is used.' }
    if (-not $AuditOnly) {
        $ResolvedOutput = [System.IO.Path]::GetFullPath($OutputRoot)
        if ($ResolvedOutput.StartsWith($ResolvedSource, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'OutputRoot cannot be inside SourceRoot.' }
    }

    $HoleListRows = @(Import-Csv -LiteralPath $HoleListPath)
    if ($HoleListRows.Count -eq 0 -or -not ($HoleListRows[0].PSObject.Properties.Name -contains 'HOLE_ID')) { throw 'Vendor hole list must contain a HOLE_ID column.' }
    $HoleSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $HoleListRows) { $id = ([string]$row.HOLE_ID).Trim().ToUpperInvariant(); if ($id.Length -gt 0) { [void]$HoleSet.Add($id) } }
    $HoleListHash = (Get-FileHash -LiteralPath $HoleListPath -Algorithm SHA256).Hash

    $Files = @(Get-ChildItem -LiteralPath $ResolvedSource -File | Where-Object { $_.Extension -eq '.csv' } | Sort-Object Name)
    if ($Files.Count -eq 0) { throw "No CSV files found in: $ResolvedSource" }
    $HashBefore = @{}
    foreach ($file in $Files) { $HashBefore[$file.FullName] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }

    $TableRecords = [System.Collections.Generic.List[object]]::new()
    $ColumnRecords = [System.Collections.Generic.List[object]]::new()
    $MalformedRecords = [System.Collections.Generic.List[object]]::new()
    $Tables = [ordered]@{}
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $file = $Files[$i]
        Write-Progress -Id $ProgressId -Activity 'CP09A Spence composite' -Status ('Reading {0} ({1} of {2})' -f $file.Name, ($i + 1), $Files.Count) -PercentComplete ([math]::Min(60, 5 + [int][math]::Floor($i / $Files.Count / 0.0182)))
        $encoding = Get-FileEncoding -Path $file.FullName
        $table = Read-SpenceTable -File $file -Encoding $encoding.Encoding -HoleSet $HoleSet
        $tableName = [regex]::Replace([System.IO.Path]::GetFileNameWithoutExtension($file.Name), '^\d+_?', '').ToUpperInvariant()
        foreach ($line in $table.MalformedLines) { $MalformedRecords.Add([pscustomobject]@{ File = $file.Name; Line = $line.Split('|')[0]; Issue = $line.Split('|')[1] }) }
        $classes = @()
        if ($table.TableType -eq 'INTERVAL') {
            $classes = @(Get-ColumnClassification -Table $table -TableName $tableName)
            foreach ($c in $classes) { $ColumnRecords.Add($c) }
        }
        $numericCols = @($classes | Where-Object { $_.Class -eq 'NUMERIC' })
        $categoryCols = @($classes | Where-Object { $_.Class -in @('CATEGORY', 'CATEGORY_ORDINAL') })
        $composite = ($table.TableType -eq 'INTERVAL') -and (($numericCols.Count + $categoryCols.Count) -gt 0) -and ($table.Rows.Count -gt 0)
        $TableRecords.Add([pscustomobject]@{
            File = $file.Name; TableName = $tableName; Encoding = $encoding.Name; TableType = $table.TableType
            TotalRows = $table.TotalRows; VendorRows = $table.Rows.Count; FilteredOut = $table.FilteredOut; Malformed = $table.Malformed
            VendorHolesFound = $table.HolesFound.Count; NumericVars = $numericCols.Count; CategoryVars = $categoryCols.Count
            ExcludedVars = @($classes | Where-Object { $_.Class.StartsWith('EXCLUDED_') }).Count
            Composite = $(if ($composite) { 'Y' } else { 'N' }); SHA256 = $HashBefore[$file.FullName]
        })
        $Tables[$file.FullName] = [pscustomobject]@{ File = $file; Name = $tableName; Table = $table; Numeric = $numericCols; Category = $categoryCols; Composite = $composite }
    }

    $CollarTable = @($Tables.Values | Where-Object { $_.Name -match 'COLLAR' } | Select-Object -First 1)
    $MissingInCollar = @()
    if ($CollarTable.Count -eq 1) { $MissingInCollar = @($HoleSet | Where-Object { -not $CollarTable[0].Table.HolesFound.Contains($_) } | Sort-Object) }

    $SupplementaryStatus = 'NOT_PROVIDED'
    $SupplementaryHash = ''
    $SupplementaryDepth = @{}
    if (-not [string]::IsNullOrWhiteSpace($SupplementaryCollarPath)) {
        if (-not (Test-Path -LiteralPath $SupplementaryCollarPath -PathType Leaf)) { throw "SupplementaryCollarPath does not exist: $SupplementaryCollarPath" }
        $supFile = Get-Item -LiteralPath $SupplementaryCollarPath
        $SupplementaryHash = (Get-FileHash -LiteralPath $supFile.FullName -Algorithm SHA256).Hash
        $supEncoding = Get-FileEncoding -Path $supFile.FullName
        $supTable = Read-SpenceTable -File $supFile -Encoding $supEncoding.Encoding -HoleSet $HoleSet
        $depthIndex = Find-ColumnIndex -Header $supTable.Header -Names @('DEPTH')
        foreach ($row in $supTable.Rows) {
            $hid = $row[$supTable.HoleIndex].Trim().ToUpperInvariant()
            $SupplementaryDepth[$hid] = $(if ($depthIndex -ge 0) { $row[$depthIndex].Trim() } else { '' })
        }
        $SupplementaryStatus = 'READ'
    }

    $MissingRegister = [System.Collections.Generic.List[object]]::new()
    foreach ($hid in $MissingInCollar) {
        $presentIn = @($Tables.Values | Where-Object { $_.Table.HolesFound.Contains($hid) } | ForEach-Object { $_.Name } | Sort-Object)
        $maxTo = [double]::NaN
        foreach ($entry in @($Tables.Values | Where-Object { $_.Table.TableType -eq 'INTERVAL' -and $_.Table.HolesFound.Contains($hid) })) {
            foreach ($row in $entry.Table.Rows) {
                if ($row[$entry.Table.HoleIndex].Trim().ToUpperInvariant() -ne $hid) { continue }
                $toValue = ConvertTo-NumberOrNaN -Text $row[$entry.Table.ToIndex]
                if (-not [double]::IsNaN($toValue) -and ([double]::IsNaN($maxTo) -or $toValue -gt $maxTo)) { $maxTo = $toValue }
            }
        }
        $inSup = $SupplementaryDepth.ContainsKey($hid)
        $MissingRegister.Add([pscustomobject]@{
            HOLE_ID = $hid; InSpenceCollar = 'N'; InSupplementaryCollar = $(if ($SupplementaryStatus -eq 'READ') { $(if ($inSup) { 'Y' } else { 'N' }) } else { 'NOT_CHECKED' })
            SupplementaryDepth = $(if ($inSup) { $SupplementaryDepth[$hid] } else { '' }); MaxLoggedTo = (Format-Number $maxTo)
            PresentInTables = ($presentIn -join '|'); CompositeImpact = $(if ($presentIn.Count -gt 0) { 'COMPOSITED_ONLY_IN_LISTED_TABLES' } else { 'NOT_COMPOSITED' })
            Status = 'NOT_IN_SPENCE_COLLAR'; Action = 'Validate with Spence database team before COMPARE'
        })
    }

    Write-Host ''
    Write-Host ('CP09A SPENCE COMPOSITE v{0}' -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('SOURCE: {0}' -f $ResolvedSource)
    Write-Host ('VENDOR HOLES: {0} (list SHA256 {1})' -f $HoleSet.Count, $HoleListHash)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' }))
    Write-Host ('VENDOR HOLES MISSING IN COLLAR: {0} {1}' -f $MissingInCollar.Count, ($MissingInCollar -join ', '))
    if ($MissingInCollar.Count -gt 0) {
        Write-Host ''
        Write-Host '================================================================' -ForegroundColor Yellow
        Write-Host ('WARNING: {0} VENDOR HOLES ARE NOT IN THE SPENCE COLLAR' -f $MissingInCollar.Count) -ForegroundColor Yellow
        Write-Host 'Probable cause: new drillholes not yet loaded in the Spence database.' -ForegroundColor Yellow
        Write-Host 'Spence composites will be incomplete for these holes. Validate with the database team before COMPARE.' -ForegroundColor Yellow
        Write-Host ('SUPPLEMENTARY COLLAR: {0}' -f $SupplementaryStatus) -ForegroundColor Yellow
        Write-Host '================================================================' -ForegroundColor Yellow
        $MissingRegister | Format-Table HOLE_ID, InSupplementaryCollar, SupplementaryDepth, MaxLoggedTo, PresentInTables -AutoSize -Wrap
        if ($StopOnMissingCollar) { throw ('StopOnMissingCollar is set and {0} vendor holes are missing in the Spence collar.' -f $MissingInCollar.Count) }
    }
    Write-Host ''
    $TableRecords | Format-Table TableName, Encoding, TableType, TotalRows, VendorRows, Malformed, VendorHolesFound, NumericVars, CategoryVars, ExcludedVars, Composite -AutoSize
    $ColumnRecords | Where-Object { $_.Class -ne 'NUMERIC' } | Format-Table Table, Column, Class, Distinct, MaxLength, Reason -AutoSize -Wrap

    if ($AuditOnly) {
        $changed = 0
        foreach ($path in $HashBefore.Keys) { if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $HashBefore[$path]) { $changed++ } }
        Write-Host ('SOURCE UNCHANGED: {0}' -f ($changed -eq 0))
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
        Write-Host ('READINESS: {0}' -f $(if (@($TableRecords | Where-Object { $_.Composite -eq 'Y' }).Count -gt 0) { 'READY_FOR_EXECUTION' } else { 'NO_GO' })) -ForegroundColor Green
        return
    }

    $RunOutput = Join-Path $ResolvedOutput $RunId
    if (Test-Path -LiteralPath $RunOutput) { throw "Overwrite is prohibited. Run output already exists: $RunOutput" }
    if (-not $PSCmdlet.ShouldProcess($RunOutput, 'Create Spence composite outputs')) { return }
    $CompositeFolder = Join-Path $RunOutput 'Composites'
    New-Item -ItemType Directory -Path $CompositeFolder -Force | Out-Null

    $JobRecords = [System.Collections.Generic.List[object]]::new()
    $IntegrityRecords = [System.Collections.Generic.List[object]]::new()
    $toRun = @($Tables.Values | Where-Object { $_.Composite })
    for ($i = 0; $i -lt $toRun.Count; $i++) {
        $entry = $toRun[$i]
        Write-Progress -Id $ProgressId -Activity 'CP09A Spence composite' -Status ('Compositing {0} ({1} of {2})' -f $entry.Name, ($i + 1), $toRun.Count) -PercentComplete ([math]::Min(92, 62 + [int][math]::Floor($i / $toRun.Count / 0.0334)))
        $result = Invoke-TableComposite -Table $entry.Table -TableName $entry.Name -NumericColumns $entry.Numeric -CategoryColumns $entry.Category -OutputFolder $CompositeFolder -SourceFile $entry.File.Name
        foreach ($r in $result.Results) { $JobRecords.Add($r) }
        $IntegrityRecords.Add([pscustomobject]@{ Table = $entry.Name; InvalidIntervals = $result.InvalidIntervals; Overlaps = $result.Overlaps })
    }

    Write-Progress -Id $ProgressId -Activity 'CP09A Spence composite' -Status 'QA/QC and evidence' -PercentComplete 94
    $changed = 0
    foreach ($path in $HashBefore.Keys) { if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $HashBefore[$path]) { $changed++ } }
    $totalMalformed = [long](($TableRecords | Measure-Object -Property Malformed -Sum).Sum)
    $totalInvalid = [long](($IntegrityRecords | Measure-Object -Property InvalidIntervals -Sum).Sum)
    $totalOverlaps = [long](($IntegrityRecords | Measure-Object -Property Overlaps -Sum).Sum)
    $QaRows = @(
        [pscustomobject]@{ Control = 'VendorHoleList'; Status = $(if ($HoleSet.Count -gt 0) { 'PASS' } else { 'FAIL' }); Evidence = ('Holes={0}; SHA256={1}' -f $HoleSet.Count, $HoleListHash) }
        [pscustomobject]@{ Control = 'VendorHolesInCollar'; Status = $(if ($MissingInCollar.Count -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Missing={0} {1}' -f $MissingInCollar.Count, ($MissingInCollar -join ' ')) }
        [pscustomobject]@{ Control = 'MissingHolesInSupplementaryCollar'; Status = $(if ($MissingInCollar.Count -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Supplementary={0}; Found={1} of {2}; SHA256={3}' -f $SupplementaryStatus, @($MissingRegister | Where-Object { $_.InSupplementaryCollar -eq 'Y' }).Count, $MissingInCollar.Count, $SupplementaryHash) }
        [pscustomobject]@{ Control = 'CompositeJobs'; Status = $(if ($JobRecords.Count -gt 0) { 'PASS' } else { 'FAIL' }); Evidence = ('Tables={0}; Jobs={1}' -f $toRun.Count, $JobRecords.Count) }
        [pscustomobject]@{ Control = 'MalformedRows'; Status = $(if ($totalMalformed -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Malformed rows excluded={0}' -f $totalMalformed) }
        [pscustomobject]@{ Control = 'InvalidIntervals'; Status = $(if ($totalInvalid -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Invalid FROM/TO excluded={0}' -f $totalInvalid) }
        [pscustomobject]@{ Control = 'OverlappingIntervals'; Status = $(if ($totalOverlaps -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Overlaps={0}' -f $totalOverlaps) }
        [pscustomobject]@{ Control = 'FreeTextExcluded'; Status = 'PASS'; Evidence = ('Excluded text columns={0}' -f @($ColumnRecords | Where-Object { $_.Class -eq 'EXCLUDED_TEXT' }).Count) }
        [pscustomobject]@{ Control = 'SourceImmutability'; Status = $(if ($changed -eq 0) { 'PASS' } else { 'FAIL' }); Evidence = ('ChangedSources={0}' -f $changed) }
        [pscustomobject]@{ Control = 'AssayUnits'; Status = 'REVIEW'; Evidence = 'AU, MO and ICP units require validation before reporting' }
    )
    $Decision = if (@($QaRows | Where-Object { $_.Status -eq 'FAIL' }).Count -gt 0) { 'NO_GO' } elseif (@($QaRows | Where-Object { $_.Status -eq 'REVIEW' }).Count -gt 0) { 'PASS_WITH_WARNINGS' } else { 'PASS' }

    $TableRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Table_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
    $ColumnRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Column_Classification.csv') -NoTypeInformation -Encoding utf8NoBOM
    $JobRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Composite_Job_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
    $IntegrityRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Interval_Integrity.csv') -NoTypeInformation -Encoding utf8NoBOM
    $QaRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_QAQC_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
    if ($MissingRegister.Count -gt 0) { $MissingRegister | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Missing_Vendor_Holes.csv') -NoTypeInformation -Encoding utf8NoBOM }
    if ($MalformedRecords.Count -gt 0) { $MalformedRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Spence_Malformed_Rows.csv') -NoTypeInformation -Encoding utf8NoBOM }

    $RunRecord = [ordered]@{
        Classification = $Classification; ScriptName = $ScriptName; ScriptVersion = $ScriptVersion; RunId = $RunId
        SourceRoot = $ResolvedSource; HoleListPath = $HoleListPath; HoleListSHA256 = $HoleListHash; VendorHoles = $HoleSet.Count
        OutputRoot = $RunOutput; Scales = $ScaleNames; MinimumCoverage = $MinimumCoverage
        GridRule = 'Bins aligned to multiples of the scale from depth zero'
        NumericRule = 'Length-weighted mean over non-null values'
        CategoryRule = 'Length-dominant code; null and -99 excluded; ties flagged'
        TextRule = ('Excluded when name matches Obs/Desc/Comment, contains spaces, length above {0} or more than {1} distinct values' -f $MaxCodeLength, $MaxDistinctCodes)
        Decision = $Decision; ElapsedSeconds = [math]::Round(((Get-Date) - $Started).TotalSeconds, 1); CreatedAt = (Get-Date).ToString('o')
    }
    ($RunRecord | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath (Join-Path $RunOutput 'Run_Register.json') -Encoding utf8NoBOM
    (@('# CP09A Spence Composite', '', ('Classification: {0}' -f $Classification), ('Run ID: {0}' -f $RunId), ('Decision: {0}' -f $Decision), '', 'Only vendor-analysed drillholes are composited.', 'Interval tables with HOLEID, FROM and TO only. Free text excluded.', 'Assays: length-weighted mean. Geology: length-dominant code.', 'Source files are read-only and verified unchanged by SHA256.') -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'README.md') -Encoding utf8NoBOM
    (@('# Changelog', '', '## v1.0.1', '- Fixed empty NumericColumns binding for geology-only tables.', '- Pre-composite warning and register for vendor holes missing in the Spence collar.', '- Optional supplementary collar check and StopOnMissingCollar switch.', '- MaxCodeLength default raised to 15.', '', '## v1.0.0', '- First Spence compositor filtered to vendor holes, with column classification, free-text exclusion and QA/QC.') -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'CHANGELOG.md') -Encoding utf8NoBOM

    $ManifestItems = @(Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Sort-Object FullName | ForEach-Object { [pscustomobject]@{ RelativePath = [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName); Length = [long]$_.Length; SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
    ([ordered]@{ Classification = $Classification; RunId = $RunId; FileCount = $ManifestItems.Count; Files = $ManifestItems } | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath (Join-Path $RunOutput 'Manifest.json') -Encoding utf8NoBOM
    $HashLines = @(Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object FullName | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName) })
    ($HashLines -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Encoding ascii

    Write-Host ''
    $JobRecords | Format-Table Table, Scale, Holes, Rows, PassRows, PartialRows, CategoryTies -AutoSize
    $QaRows | Format-Table Control, Status, Evidence -AutoSize -Wrap
    Write-Host ('OUTPUT: {0}' -f $RunOutput) -ForegroundColor Green
    Write-Host ('ELAPSED SECONDS: {0}' -f $RunRecord.ElapsedSeconds)
    $color = if ($Decision -eq 'PASS') { 'Green' } elseif ($Decision -eq 'PASS_WITH_WARNINGS') { 'Yellow' } else { 'Red' }
    Write-Host ('DECISION: {0}' -f $Decision) -ForegroundColor $color
    if ($Decision -eq 'NO_GO') { $ExitCode = 1 }
}
catch {
    Write-Host ''
    Write-Host ('FAIL: {0}' -f $_.Exception.Message) -ForegroundColor Red
    $ExitCode = 1
}
finally {
    Write-Progress -Id $ProgressId -Activity 'CP09A Spence composite' -Completed
}

exit $ExitCode
