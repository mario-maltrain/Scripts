# Classification: Internal
# Script: Run_CP09A_Composite_Jobs_v1_0_0.ps1
# Purpose: Execute the 12 governed CP09A composite jobs (4 sources x 1m, 3m, 6m) from an initialization run.
# Safety: AuditOnly is the default. Sources are read-only and verified by SHA256. No delete, move, rename or overwrite.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$InitRunRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\01_Processed_Exports\GEOAI-EXPORTS-20261002-225517\13_Composites\CP09A-COMPOSITE-20261003-014710',

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-COMPOSITE-EXEC-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-COMPOSITE-EXEC-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [ValidateRange(0.0, 1.0)]
    [double]$MinimumCoverage = 0.8,

    [Parameter()]
    [ValidateRange(0.0001, 1.0)]
    [double]$SampleStep = 0.01,

    [Parameter()]
    [AllowEmptyString()]
    [string]$LepBaselineRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\02_Composites_PowerShell\CP09A-PS-LEP-COMPOSITE-20261001-175040\Composites',

    [Parameter()]
    [switch]$CommitRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Run_CP09A_Composite_Jobs_v1_0_0.ps1'
$ScriptVersion = '1.0.0'
$Classification = 'Internal'
$ProgressId = 1500
$ExitCode = 0
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$NumberStyle = [System.Globalization.NumberStyles]::Float
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$QuoteTrim = [char[]]@([char]34, [char]32)
$LineFeed = [string][char]10
$Epsilon = 0.000001
$NullValue = -999.25
$ExpectedSources = @('SPENCE_MASTER', 'LAS_LEP', 'LEP_CONFIDENCE', 'AUTOLITH_CONFIDENCE')
$LepValues = @('GT20', 'GT30', 'GT50')
$AutoLithValues = @('ABX1_PROB', 'ABX2_PROB', 'QFP1_PROB', 'QFP2_PROB', 'IND_PROB', 'H-CONFIDENCE')
$AutoLithClassCount = 5
$SpenceSkipColumns = @('HOLEID', 'FROM', 'TO', 'EXPORT DATE', 'EXPORT_ID', 'SOURCE_FILE', 'RUN_ID')

function Get-Product {
    param([double]$A, [double]$B)
    return [math]::FusedMultiplyAdd($A, $B, 0.0)
}

function Format-Number {
    param([double]$Value)
    if ([double]::IsNaN($Value)) { return '' }
    return $Value.ToString('R', $Invariant)
}

function ConvertTo-NumberOrNaN {
    param([AllowEmptyString()][string]$Text)
    $clean = $Text.Trim($QuoteTrim)
    if ($clean.Length -eq 0) { return [double]::NaN }
    $number = 0.0
    if (-not [double]::TryParse($clean, $NumberStyle, $Invariant, [ref]$number)) { return [double]::NaN }
    if ([math]::Abs($number - $NullValue) -lt $Epsilon) { return [double]::NaN }
    return $number
}

function Test-SourceFileExcluded {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    if ($File.Extension -ne '.csv') { return $true }
    if ($File.Name -match '__DUP_SHA256_EXACT|__COPY_\d+|\bCopiar\b') { return $true }
    if ($File.Name -match '(?i)(^|_)composite(_|\b)|_[136]m\.csv$') { return $true }
    return $false
}

function Get-SourceFiles {
    param([Parameter(Mandatory)][string]$SourcePath)
    if (Test-Path -LiteralPath $SourcePath -PathType Leaf) { return @(Get-Item -LiteralPath $SourcePath) }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) { return @() }
    return @(
        Get-ChildItem -LiteralPath $SourcePath -File -Recurse |
            Where-Object { ($_.Extension -eq '.csv') -and (-not (Test-SourceFileExcluded -File $_)) } |
            Sort-Object FullName -Unique |
            Sort-Object FullName
    )
}

function Get-FileSetHash {
    param([Parameter(Mandatory)][object[]]$HashRows)
    $material = @($HashRows | Sort-Object FullName | ForEach-Object { '{0}|{1}' -f $_.FullName, $_.Hash }) -join $LineFeed
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.Convert]::ToHexString($algorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($material))) }
    finally { $algorithm.Dispose() }
}

function Get-HeaderColumns {
    param([Parameter(Mandatory)][string]$Path)
    $reader = [System.IO.StreamReader]::new($Path)
    try { $line = $reader.ReadLine() } finally { $reader.Dispose() }
    if ($null -eq $line) { return @() }
    return @($line.Split(',') | ForEach-Object { $_.Trim($QuoteTrim) })
}

function New-ScaleStores {
    param([int]$Count)
    $stores = [object[]]::new($Count)
    for ($s = 0; $s -lt $Count; $s++) { $stores[$s] = [System.Collections.Generic.Dictionary[long, double[]]]::new() }
    return ,$stores
}

function Read-PointFile {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string[]]$ValueNames,
        [Parameter(Mandatory)][double[]]$Scales,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.Dictionary[string, object]]$Store,
        [Parameter(Mandatory)][AllowEmptyCollection()][hashtable]$SourceByHole,
        [Parameter()][AllowEmptyString()][string]$HoleColumn = '',
        [Parameter()][AllowEmptyString()][string]$FileHoleColumn = ''
    )

    $reader = [System.IO.StreamReader]::new($File.FullName)
    $rows = 0L
    $bad = 0L
    try {
        $headerLine = $reader.ReadLine()
        if ($null -eq $headerLine) { throw ('Empty file: {0}' -f $File.Name) }
        $header = @($headerLine.Split(',') | ForEach-Object { $_.Trim($QuoteTrim).ToUpperInvariant() })
        $depthIndex = [array]::IndexOf($header, 'DEPT')
        if ($depthIndex -lt 0) { throw ('DEPT column missing: {0}' -f $File.Name) }
        $valueIndex = [int[]]::new($ValueNames.Count)
        for ($v = 0; $v -lt $ValueNames.Count; $v++) { $valueIndex[$v] = [array]::IndexOf($header, $ValueNames[$v].ToUpperInvariant()) }
        $holeIndex = if ([string]::IsNullOrEmpty($HoleColumn)) { -1 } else { [array]::IndexOf($header, $HoleColumn) }
        $fileHoleIndex = if ([string]::IsNullOrEmpty($FileHoleColumn)) { -1 } else { [array]::IndexOf($header, $FileHoleColumn) }
        if ((-not [string]::IsNullOrEmpty($HoleColumn)) -and $holeIndex -lt 0) { throw ('{0} column missing: {1}' -f $HoleColumn, $File.Name) }
        if ((-not [string]::IsNullOrEmpty($FileHoleColumn)) -and $fileHoleIndex -lt 0) { throw ('{0} column missing: {1}' -f $FileHoleColumn, $File.Name) }

        $fixedHole = ''
        if ($holeIndex -lt 0 -and $fileHoleIndex -lt 0) {
            $match = [regex]::Match($File.Name, '^(SPD\d+)_', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $match.Success) { throw ('HOLE_ID cannot be derived from file name: {0}' -f $File.Name) }
            $fixedHole = $match.Groups[1].Value.ToUpperInvariant()
        }

        $width = 3 + $ValueNames.Count + $ValueNames.Count
        $values = [double[]]::new($ValueNames.Count)
        $currentHole = ''
        $currentKeyText = $null
        $currentStores = $null
        $currentSource = $File.Name

        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.Length -eq 0) { continue }
            $tokens = $line.Split(',')
            if ($tokens.Count -ne $header.Count) { $bad++; continue }
            $depth = ConvertTo-NumberOrNaN -Text $tokens[$depthIndex]
            if ([double]::IsNaN($depth)) { $bad++; continue }

            if ($holeIndex -ge 0) { $keyText = $tokens[$holeIndex].Trim($QuoteTrim) }
            elseif ($fileHoleIndex -ge 0) { $keyText = $tokens[$fileHoleIndex].Trim($QuoteTrim) }
            else { $keyText = $fixedHole }

            if ($keyText -ne $currentKeyText) {
                $currentKeyText = $keyText
                if ($fileHoleIndex -ge 0) {
                    $match = [regex]::Match($keyText, '^(SPD\d+)_', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if (-not $match.Success) { $currentHole = ''; $currentStores = $null; $bad++; continue }
                    $currentHole = $match.Groups[1].Value.ToUpperInvariant()
                    $currentSource = $keyText
                }
                else {
                    $currentHole = $keyText.ToUpperInvariant()
                }
                $stores = $null
                if (-not $Store.TryGetValue($currentHole, [ref]$stores)) {
                    $stores = New-ScaleStores -Count $Scales.Count
                    $Store[$currentHole] = $stores
                }
                $currentStores = $stores
                $SourceByHole[$currentHole] = $currentSource
            }
            if ($null -eq $currentStores) { $bad++; continue }

            for ($v = 0; $v -lt $ValueNames.Count; $v++) {
                if ($valueIndex[$v] -lt 0) { $values[$v] = [double]::NaN }
                else { $values[$v] = ConvertTo-NumberOrNaN -Text $tokens[$valueIndex[$v]] }
            }

            for ($s = 0; $s -lt $Scales.Count; $s++) {
                $bin = [long][math]::Floor($depth / $Scales[$s] + 0.000000001)
                $dictionary = $currentStores[$s]
                $acc = $null
                if (-not $dictionary.TryGetValue($bin, [ref]$acc)) {
                    $acc = [double[]]::new($width)
                    $acc[1] = $depth
                    $acc[2] = $depth
                    $dictionary[$bin] = $acc
                }
                $acc[0] += 1
                if ($depth -lt $acc[1]) { $acc[1] = $depth }
                if ($depth -gt $acc[2]) { $acc[2] = $depth }
                for ($v = 0; $v -lt $ValueNames.Count; $v++) {
                    if (-not [double]::IsNaN($values[$v])) {
                        $acc[3 + $v + $v] += $values[$v]
                        $acc[4 + $v + $v] += 1
                    }
                }
            }
            $rows++
        }
    }
    finally {
        $reader.Dispose()
    }
    return [pscustomobject]@{ Rows = $rows; BadRows = $bad }
}

function Write-PointComposite {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.Dictionary[string, object]]$Store,
        [Parameter(Mandatory)][double]$Scale,
        [Parameter(Mandatory)][int]$ScaleIndex,
        [Parameter(Mandatory)][string]$DataType,
        [Parameter(Mandatory)][string[]]$ValueNames,
        [Parameter(Mandatory)][AllowEmptyCollection()][hashtable]$SourceByHole,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter()][int]$ClassCount = 0
    )

    $expected = [math]::Round($Scale / $SampleStep)
    $headerParts = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @('DATA_TYPE', 'HOLE_ID', 'FROM', 'TO', 'CENTRE', 'SCALE_M', 'N_SAMPLES', 'EXPECTED_SAMPLES', 'SAMPLE_COVERAGE_RATIO', 'PARTIAL_COMPOSITE', 'OBSERVED_MIN_DEPTH', 'OBSERVED_MAX_DEPTH')) { $headerParts.Add($name) }
    foreach ($name in $ValueNames) { $headerParts.Add(($name.Replace('-', '_')) + '_MEAN') }
    if ($ClassCount -gt 0) { foreach ($name in @('PROB_SUM', 'ARGMAX_CLASS', 'ARGMAX_PROB', 'ARGMAX_TIE')) { $headerParts.Add($name) } }
    $headerParts.Add('STATUS')
    $headerParts.Add('SOURCE_FILE')

    $rows = 0L
    $pass = 0L
    $partial = 0L
    $noProb = 0L
    $ties = 0L
    $writer = [System.IO.StreamWriter]::new($OutputPath, $false, $Utf8NoBom)
    try {
        $writer.WriteLine(($headerParts -join ','))
        foreach ($hole in @($Store.Keys | Sort-Object)) {
            $dictionary = $Store[$hole][$ScaleIndex]
            foreach ($bin in @($dictionary.Keys | Sort-Object)) {
                $acc = $dictionary[$bin]
                $from = Get-Product -A ([double]$bin) -B $Scale
                $to = Get-Product -A ([double]($bin + 1)) -B $Scale
                $coverage = $acc[0] / $expected
                $isPartial = ($coverage -lt 1.0)
                $parts = [System.Collections.Generic.List[string]]::new()
                $parts.Add($DataType); $parts.Add($hole)
                $parts.Add((Format-Number $from)); $parts.Add((Format-Number $to)); $parts.Add((Format-Number (($from + $to) / 2)))
                $parts.Add((Format-Number $Scale)); $parts.Add([string][long]$acc[0]); $parts.Add([string][long]$expected)
                $parts.Add((Format-Number $coverage)); $parts.Add($(if ($isPartial) { 'True' } else { 'False' }))
                $parts.Add((Format-Number $acc[1])); $parts.Add((Format-Number $acc[2]))
                $means = [double[]]::new($ValueNames.Count)
                for ($v = 0; $v -lt $ValueNames.Count; $v++) {
                    $n = $acc[4 + $v + $v]
                    $means[$v] = if ($n -gt 0) { $acc[3 + $v + $v] / $n } else { [double]::NaN }
                    $parts.Add((Format-Number $means[$v]))
                }
                $status = if ($coverage -ge $MinimumCoverage) { 'PASS' } else { 'PARTIAL' }
                if ($ClassCount -gt 0) {
                    $best = -1
                    $bestValue = [double]::NegativeInfinity
                    $tie = $false
                    $sum = 0.0
                    $present = 0
                    for ($c = 0; $c -lt $ClassCount; $c++) {
                        if ([double]::IsNaN($means[$c])) { continue }
                        $present++
                        $sum += $means[$c]
                        if ($means[$c] -gt $bestValue + $Epsilon) { $best = $c; $bestValue = $means[$c]; $tie = $false }
                        elseif ([math]::Abs($means[$c] - $bestValue) -le $Epsilon) { $tie = $true }
                    }
                    if ($present -eq 0) {
                        $parts.Add(''); $parts.Add('NO_PROBABILITY_INFORMATION'); $parts.Add(''); $parts.Add('False')
                        $status = 'NO_PROBABILITY_INFORMATION'
                        $noProb++
                    }
                    else {
                        $parts.Add((Format-Number $sum)); $parts.Add($ValueNames[$best].Replace('_PROB', '')); $parts.Add((Format-Number $bestValue)); $parts.Add($(if ($tie) { 'True' } else { 'False' }))
                        if ($tie) { $ties++ }
                    }
                }
                $parts.Add($status)
                $parts.Add([string]$SourceByHole[$hole])
                $writer.WriteLine(($parts -join ','))
                $rows++
                if ($status -eq 'PASS') { $pass++ } elseif ($status -eq 'PARTIAL') { $partial++ }
            }
        }
    }
    finally {
        $writer.Dispose()
    }
    return [pscustomobject]@{ Rows = $rows; Holes = $Store.Count; PassRows = $pass; PartialRows = $partial; NoProbabilityRows = $noProb; Ties = $ties }
}

function Compare-PointStores {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.Dictionary[string, object]]$Left,
        [Parameter(Mandatory)][System.Collections.Generic.Dictionary[string, object]]$Right,
        [Parameter(Mandatory)][int]$ScaleCount
    )
    $differences = 0L
    $holes = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Left.Keys))
    foreach ($key in $Right.Keys) { [void]$holes.Add($key) }
    foreach ($hole in $holes) {
        if (-not ($Left.ContainsKey($hole) -and $Right.ContainsKey($hole))) { $differences++; continue }
        for ($s = 0; $s -lt $ScaleCount; $s++) {
            $a = $Left[$hole][$s]
            $b = $Right[$hole][$s]
            if ($a.Count -ne $b.Count) { $differences++; continue }
            foreach ($bin in $a.Keys) {
                $other = $null
                if (-not $b.TryGetValue($bin, [ref]$other)) { $differences++; continue }
                $mine = $a[$bin]
                for ($k = 0; $k -lt $mine.Count; $k++) {
                    if ([math]::Abs($mine[$k] - $other[$k]) -gt (0.000000001 + [math]::Abs($mine[$k]) / 1000000000.0)) { $differences++; break }
                }
            }
        }
    }
    return $differences
}

function Invoke-SpenceComposite {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][double[]]$Scales,
        [Parameter(Mandatory)][string[]]$ScaleNames,
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][bool]$Write
    )

    $tables = [ordered]@{}
    $invalid = 0L
    $overlaps = 0L
    $intervals = 0L
    foreach ($file in $Files) {
        $columns = @(Get-HeaderColumns -Path $file.FullName)
        $attribute = @($columns | Where-Object { ($SpenceSkipColumns -notcontains $_.ToUpperInvariant()) -and ($_ -notmatch '(?i)_Desc$') } | Select-Object -First 1)
        if ($attribute.Count -ne 1) { throw ('Spence attribute column not resolved: {0}' -f $file.Name) }
        $attributeName = $attribute[0]
        $holeStores = [System.Collections.Generic.Dictionary[string, object]]::new()
        $holeIntervals = @{}
        foreach ($row in @(Import-Csv -LiteralPath $file.FullName)) {
            $hole = ([string]$row.holeid).Trim().ToUpperInvariant()
            $f = ConvertTo-NumberOrNaN -Text ([string]$row.From)
            $t = ConvertTo-NumberOrNaN -Text ([string]$row.To)
            if ([string]::IsNullOrEmpty($hole) -or [double]::IsNaN($f) -or [double]::IsNaN($t) -or $t -le $f) { $invalid++; continue }
            $code = ([string]$row.$attributeName).Trim()
            if ($code.Length -eq 0) { $code = '(BLANK)' }
            $intervals++
            if (-not $holeIntervals.ContainsKey($hole)) { $holeIntervals[$hole] = [System.Collections.Generic.List[double[]]]::new() }
            $holeIntervals[$hole].Add([double[]]@($f, $t))
            $stores = $null
            if (-not $holeStores.TryGetValue($hole, [ref]$stores)) {
                $stores = [object[]]::new($Scales.Count)
                for ($s = 0; $s -lt $Scales.Count; $s++) { $stores[$s] = [System.Collections.Generic.Dictionary[long, hashtable]]::new() }
                $holeStores[$hole] = $stores
            }
            for ($s = 0; $s -lt $Scales.Count; $s++) {
                $scale = $Scales[$s]
                $first = [long][math]::Floor($f / $scale + 0.000000001)
                $last = [long][math]::Floor(($t - 0.000000001) / $scale)
                for ($bin = $first; $bin -le $last; $bin++) {
                    $binFrom = Get-Product -A ([double]$bin) -B $scale
                    $binTo = Get-Product -A ([double]($bin + 1)) -B $scale
                    $overlap = [math]::Min($t, $binTo) - [math]::Max($f, $binFrom)
                    if ($overlap -le 0.000000001) { continue }
                    $codes = $null
                    if (-not $stores[$s].TryGetValue($bin, [ref]$codes)) { $codes = @{}; $stores[$s][$bin] = $codes }
                    if ($codes.ContainsKey($code)) { $codes[$code] += $overlap } else { $codes[$code] = $overlap }
                }
            }
        }
        foreach ($hole in $holeIntervals.Keys) {
            $sorted = @($holeIntervals[$hole] | Sort-Object { $_[0] })
            for ($k = 1; $k -lt $sorted.Count; $k++) { if ($sorted[$k][0] -lt $sorted[$k - 1][1] - 0.000001) { $overlaps++ } }
        }
        $tables[$attributeName] = [pscustomobject]@{ File = $file.Name; Stores = $holeStores }
    }

    $results = [System.Collections.Generic.List[object]]::new()
    for ($s = 0; $s -lt $Scales.Count; $s++) {
        $scale = $Scales[$s]
        $rows = 0L; $pass = 0L; $partial = 0L; $ties = 0L
        $outputPath = Join-Path $OutputFolder ('SPENCE_MASTER_Composite_{0}.csv' -f $ScaleNames[$s])
        $writer = $null
        if ($Write) {
            $writer = [System.IO.StreamWriter]::new($outputPath, $false, $Utf8NoBom)
            $writer.WriteLine('DATA_TYPE,TABLE,HOLE_ID,FROM,TO,CENTRE,SCALE_M,COVERED_LENGTH,COVERAGE_RATIO,PARTIAL_COMPOSITE,DOMINANT_CODE,DOMINANT_LENGTH,DOMINANT_FRACTION,N_CODES,DOMINANT_TIE,STATUS,SOURCE_FILE')
        }
        try {
            foreach ($tableName in $tables.Keys) {
                $table = $tables[$tableName]
                foreach ($hole in @($table.Stores.Keys | Sort-Object)) {
                    $dictionary = $table.Stores[$hole][$s]
                    foreach ($bin in @($dictionary.Keys | Sort-Object)) {
                        $codes = $dictionary[$bin]
                        $covered = 0.0
                        $bestCode = ''
                        $bestLength = -1.0
                        $tie = $false
                        foreach ($code in @($codes.Keys | Sort-Object)) {
                            $length = [double]$codes[$code]
                            $covered += $length
                            if ($length -gt $bestLength + 0.000000001) { $bestCode = $code; $bestLength = $length; $tie = $false }
                            elseif ([math]::Abs($length - $bestLength) -le 0.000000001) { $tie = $true }
                        }
                        $from = Get-Product -A ([double]$bin) -B $scale
                        $to = Get-Product -A ([double]($bin + 1)) -B $scale
                        $coverage = $covered / $scale
                        $status = if ($coverage -ge $MinimumCoverage) { 'PASS' } else { 'PARTIAL' }
                        if ($tie) { $ties++ }
                        if ($status -eq 'PASS') { $pass++ } else { $partial++ }
                        $rows++
                        if ($null -ne $writer) {
                            $parts = @('SPENCE_MASTER', $tableName, $hole, (Format-Number $from), (Format-Number $to), (Format-Number (($from + $to) / 2)), (Format-Number $scale), (Format-Number $covered), (Format-Number $coverage), $(if ($coverage -lt 1.0) { 'True' } else { 'False' }), $bestCode, (Format-Number $bestLength), (Format-Number ($bestLength / $covered)), [string]$codes.Count, $(if ($tie) { 'True' } else { 'False' }), $status, $table.File)
                            $writer.WriteLine(($parts -join ','))
                        }
                    }
                }
            }
        }
        finally {
            if ($null -ne $writer) { $writer.Dispose() }
        }
        $results.Add([pscustomobject]@{ Scale = $ScaleNames[$s]; Rows = $rows; PassRows = $pass; PartialRows = $partial; Ties = $ties; OutputFile = $(if ($Write) { $outputPath } else { '' }) })
    }
    return [pscustomobject]@{ Tables = @($tables.Keys); Intervals = $intervals; InvalidIntervals = $invalid; Overlaps = $overlaps; Results = $results }
}

try {
    $Started = Get-Date
    Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status 'Validating initialization run' -PercentComplete 2

    if (-not (Test-Path -LiteralPath $InitRunRoot -PathType Container)) { throw "InitRunRoot does not exist: $InitRunRoot" }
    $ResolvedInit = (Get-Item -LiteralPath $InitRunRoot).FullName
    $SourceRegisterPath = Join-Path $ResolvedInit 'Composite_Source_Register.csv'
    $JobPlanPath = Join-Path $ResolvedInit 'Composite_Job_Plan.csv'
    foreach ($required in @($SourceRegisterPath, $JobPlanPath)) { if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "Required initialization record missing: $required" } }

    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) { throw 'OutputRoot is required when CommitRecords is used.' }

    $SourceRegister = @(Import-Csv -LiteralPath $SourceRegisterPath)
    $JobPlan = @(Import-Csv -LiteralPath $JobPlanPath)
    $ReadyJobs = @($JobPlan | Where-Object { $_.Status -eq 'READY' })
    if ($JobPlan.Count -ne 12 -or $ReadyJobs.Count -ne 12) { throw ('Initialization plan is not 12 READY jobs. Jobs={0}; Ready={1}' -f $JobPlan.Count, $ReadyJobs.Count) }
    foreach ($id in $ExpectedSources) { if (@($SourceRegister | Where-Object { $_.SourceId -eq $id -and $_.Status -eq 'PASS' }).Count -ne 1) { throw "Source not PASS in register: $id" } }

    $ScaleNames = @($JobPlan | Select-Object -ExpandProperty Scale -Unique | Sort-Object { [double]($_.TrimEnd('m')) })
    $Scales = [double[]]@($ScaleNames | ForEach-Object { [double]::Parse($_.TrimEnd('m'), $Invariant) })

    $Sources = [ordered]@{}
    $HashBefore = @{}
    $SourceChecks = [System.Collections.Generic.List[object]]::new()
    $step = 0
    foreach ($id in $ExpectedSources) {
        $step++
        Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status ('Verifying source {0}' -f $id) -PercentComplete (2 + $step + $step)
        $entry = @($SourceRegister | Where-Object { $_.SourceId -eq $id })[0]
        $files = @(Get-SourceFiles -SourcePath $entry.SourcePath)
        $hashRows = @($files | ForEach-Object { [pscustomobject]@{ FullName = $_.FullName; Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
        foreach ($hashRow in $hashRows) { $HashBefore[$hashRow.FullName] = $hashRow.Hash }
        $actual = if ($entry.SelectionMode -eq 'FILE' -and $files.Count -eq 1) { $hashRows[0].Hash } else { Get-FileSetHash -HashRows $hashRows }
        $columns = if ($files.Count -gt 0) { (Get-HeaderColumns -Path $files[0].FullName) -join '|' } else { '' }
        $hashOk = ($actual -eq $entry.SHA256)
        $countOk = ($files.Count -eq [int]$entry.FileCount)
        $SourceChecks.Add([pscustomobject]@{ SourceId = $id; Files = $files.Count; RegisteredFiles = [int]$entry.FileCount; SHA256Match = $hashOk; Status = $(if ($hashOk -and $countOk) { 'PASS' } else { 'FAIL' }); FirstHeader = $columns })
        $Sources[$id] = [pscustomobject]@{ Entry = $entry; Files = $files; Bytes = [long](($files | Measure-Object -Property Length -Sum).Sum) }
    }

    $SchemaChecks = @(
        [pscustomobject]@{ SourceId = 'SPENCE_MASTER'; Required = 'holeid|From|To'; Status = $(if (@($Sources['SPENCE_MASTER'].Files | Where-Object { $h = @(Get-HeaderColumns -Path $_.FullName); -not (($h -contains 'holeid') -and ($h -contains 'From') -and ($h -contains 'To')) }).Count -eq 0) { 'PASS' } else { 'FAIL' }) }
        [pscustomobject]@{ SourceId = 'LAS_LEP'; Required = 'DEPT|GT20|GT30|GT50'; Status = $(if (((Get-HeaderColumns -Path $Sources['LAS_LEP'].Files[0].FullName) -join '|') -eq 'DEPT|GT20|GT30|GT50') { 'PASS' } else { 'FAIL' }) }
        [pscustomobject]@{ SourceId = 'LEP_CONFIDENCE'; Required = 'DEPT|GT20|GT30|GT50|SOURCE_FILE'; Status = $(if (@('DEPT', 'GT20', 'GT30', 'GT50', 'SOURCE_FILE' | Where-Object { (Get-HeaderColumns -Path $Sources['LEP_CONFIDENCE'].Files[0].FullName) -notcontains $_ }).Count -eq 0) { 'PASS' } else { 'FAIL' }) }
        [pscustomobject]@{ SourceId = 'AUTOLITH_CONFIDENCE'; Required = 'HOLE_ID|DEPT|ABX1_PROB|ABX2_PROB|QFP1_PROB|QFP2_PROB|IND_PROB|H-CONFIDENCE'; Status = $(if (((Get-HeaderColumns -Path $Sources['AUTOLITH_CONFIDENCE'].Files[0].FullName) -join '|') -eq 'HOLE_ID|DEPT|ABX1_PROB|ABX2_PROB|QFP1_PROB|QFP2_PROB|IND_PROB|H-CONFIDENCE') { 'PASS' } else { 'FAIL' }) }
    )

    $PreBlocking = @($SourceChecks | Where-Object { $_.Status -ne 'PASS' }).Count + @($SchemaChecks | Where-Object { $_.Status -ne 'PASS' }).Count
    $TotalMB = [math]::Round((($Sources.Values | Measure-Object -Property Bytes -Sum).Sum) / 1MB, 1)

    Write-Host ''
    Write-Host ('CP09A COMPOSITE JOBS v{0}' -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('INIT RUN: {0}' -f $ResolvedInit)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' }))
    Write-Host ('SCALES: {0}' -f ($ScaleNames -join ', '))
    Write-Host ('MINIMUM COVERAGE: {0}' -f $MinimumCoverage)
    Write-Host ('INPUT VOLUME MB: {0}' -f $TotalMB)
    Write-Host ''
    $SourceChecks | Format-Table SourceId, Files, RegisteredFiles, SHA256Match, Status -AutoSize
    $SchemaChecks | Format-Table SourceId, Required, Status -AutoSize -Wrap

    if ($PreBlocking -gt 0) { throw ('Pre-execution validation failed with {0} blocking checks.' -f $PreBlocking) }

    if ($AuditOnly) {
        Write-Host ('PROPOSED RUN FOLDER NAME: {0} (created under -OutputRoot only with -CommitRecords)' -f $RunId)
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
        Write-Host 'READINESS: READY_FOR_EXECUTION' -ForegroundColor Green
        return
    }

    $RunOutput = Join-Path ([System.IO.Path]::GetFullPath($OutputRoot)) $RunId
    if (Test-Path -LiteralPath $RunOutput) { throw "Overwrite is prohibited. Run output already exists: $RunOutput" }
    if (-not $PSCmdlet.ShouldProcess($RunOutput, 'Create composite outputs')) { return }
    $CompositeRoot = Join-Path $RunOutput 'Composites'
    New-Item -ItemType Directory -Path $CompositeRoot | Out-Null
    foreach ($id in $ExpectedSources) { New-Item -ItemType Directory -Path (Join-Path $CompositeRoot $id) | Out-Null }

    $JobRecords = [System.Collections.Generic.List[object]]::new()
    $Stats = @{}

    $pointSources = @(
        [pscustomobject]@{ Id = 'LAS_LEP'; DataType = 'LEP'; Values = $LepValues; HoleColumn = ''; FileHoleColumn = ''; Classes = 0 }
        [pscustomobject]@{ Id = 'LEP_CONFIDENCE'; DataType = 'LEP'; Values = $LepValues; HoleColumn = ''; FileHoleColumn = 'SOURCE_FILE'; Classes = 0 }
        [pscustomobject]@{ Id = 'AUTOLITH_CONFIDENCE'; DataType = 'AUTOLITH'; Values = $AutoLithValues; HoleColumn = 'HOLE_ID'; FileHoleColumn = ''; Classes = $AutoLithClassCount }
    )
    $PointStores = @{}
    $sourceNumber = 0
    foreach ($ps in $pointSources) {
        $sourceNumber++
        $jobStart = Get-Date
        $store = [System.Collections.Generic.Dictionary[string, object]]::new()
        $sourceByHole = @{}
        $rows = 0L
        $bad = 0L
        $files = $Sources[$ps.Id].Files
        for ($i = 0; $i -lt $files.Count; $i++) {
            Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status ('{0}: file {1} of {2}' -f $ps.Id, ($i + 1), $files.Count) -PercentComplete ([math]::Min(90, 10 + $sourceNumber + $sourceNumber + [int]($i / [math]::Max(1, $files.Count) / 0.05)))
            $read = Read-PointFile -File $files[$i] -ValueNames $ps.Values -Scales $Scales -Store $store -SourceByHole $sourceByHole -HoleColumn $ps.HoleColumn -FileHoleColumn $ps.FileHoleColumn
            $rows += $read.Rows
            $bad += $read.BadRows
        }
        $PointStores[$ps.Id] = $store
        $Stats[$ps.Id] = [pscustomobject]@{ Rows = $rows; BadRows = $bad; Holes = $store.Count }
        for ($s = 0; $s -lt $Scales.Count; $s++) {
            $outputPath = Join-Path (Join-Path $CompositeRoot $ps.Id) ('{0}_Composite_{1}.csv' -f $ps.Id, $ScaleNames[$s])
            $result = Write-PointComposite -Store $store -Scale $Scales[$s] -ScaleIndex $s -DataType $ps.DataType -ValueNames $ps.Values -SourceByHole $sourceByHole -OutputPath $outputPath -ClassCount $ps.Classes
            $JobRecords.Add([pscustomobject]@{ JobId = ('{0}-{1}' -f $ps.Id, $ScaleNames[$s]); SourceId = $ps.Id; Scale = $ScaleNames[$s]; InputRows = $rows; RejectedRows = $bad; Holes = $result.Holes; Rows = $result.Rows; PassRows = $result.PassRows; PartialRows = $result.PartialRows; NoProbabilityRows = $result.NoProbabilityRows; Ties = $result.Ties; Seconds = [math]::Round(((Get-Date) - $jobStart).TotalSeconds, 1); Status = 'COMPLETED'; OutputFile = $outputPath; SHA256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash })
        }
    }

    Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status 'SPENCE_MASTER interval compositing' -PercentComplete 91
    $spenceStart = Get-Date
    $spence = Invoke-SpenceComposite -Files $Sources['SPENCE_MASTER'].Files -Scales $Scales -ScaleNames $ScaleNames -OutputFolder (Join-Path $CompositeRoot 'SPENCE_MASTER') -Write $true
    foreach ($result in $spence.Results) {
        $JobRecords.Add([pscustomobject]@{ JobId = ('SPENCE_MASTER-{0}' -f $result.Scale); SourceId = 'SPENCE_MASTER'; Scale = $result.Scale; InputRows = $spence.Intervals; RejectedRows = $spence.InvalidIntervals; Holes = ''; Rows = $result.Rows; PassRows = $result.PassRows; PartialRows = $result.PartialRows; NoProbabilityRows = 0; Ties = $result.Ties; Seconds = [math]::Round(((Get-Date) - $spenceStart).TotalSeconds, 1); Status = 'COMPLETED'; OutputFile = $result.OutputFile; SHA256 = (Get-FileHash -LiteralPath $result.OutputFile -Algorithm SHA256).Hash })
    }

    Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status 'Running QA/QC' -PercentComplete 94
    $lepDifferences = Compare-PointStores -Left $PointStores['LAS_LEP'] -Right $PointStores['LEP_CONFIDENCE'] -ScaleCount $Scales.Count

    $baselineRows = [System.Collections.Generic.List[string]]::new()
    $baselineStatus = 'NOT_AVAILABLE'
    if (-not [string]::IsNullOrWhiteSpace($LepBaselineRoot) -and (Test-Path -LiteralPath $LepBaselineRoot -PathType Container)) {
        $baselineStatus = 'PASS'
        foreach ($scaleName in $ScaleNames) {
            $baseFile = Join-Path (Join-Path $LepBaselineRoot $scaleName) ('LEP_Composite_{0}.csv' -f $scaleName)
            $job = @($JobRecords | Where-Object { $_.JobId -eq ('LAS_LEP-{0}' -f $scaleName) })[0]
            if (Test-Path -LiteralPath $baseFile -PathType Leaf) {
                $baseCount = 0L
                $reader = [System.IO.StreamReader]::new($baseFile)
                try { [void]$reader.ReadLine(); while ($null -ne $reader.ReadLine()) { $baseCount++ } } finally { $reader.Dispose() }
                $baselineRows.Add(('{0}: baseline={1}; new={2}' -f $scaleName, $baseCount, $job.Rows))
                if ($baseCount -ne $job.Rows) { $baselineStatus = 'REVIEW' }
            }
            else { $baselineRows.Add(('{0}: baseline file missing' -f $scaleName)); $baselineStatus = 'REVIEW' }
        }
    }

    $changed = 0
    foreach ($path in $HashBefore.Keys) { if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $HashBefore[$path]) { $changed++ } }

    $autoJobs = @($JobRecords | Where-Object { $_.SourceId -eq 'AUTOLITH_CONFIDENCE' })
    $QaRows = @(
        [pscustomobject]@{ Control = 'JobsCompleted'; Status = $(if ($JobRecords.Count -eq 12) { 'PASS' } else { 'FAIL' }); Evidence = ('Jobs={0}' -f $JobRecords.Count) }
        [pscustomobject]@{ Control = 'LasLepHoles'; Status = $(if ($Stats['LAS_LEP'].Holes -eq 154) { 'PASS' } else { 'REVIEW' }); Evidence = ('Holes={0}; Rows={1}; Rejected={2}' -f $Stats['LAS_LEP'].Holes, $Stats['LAS_LEP'].Rows, $Stats['LAS_LEP'].BadRows) }
        [pscustomobject]@{ Control = 'LepConfidenceEqualsLasLep'; Status = $(if ($lepDifferences -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('BinDifferences={0}' -f $lepDifferences) }
        [pscustomobject]@{ Control = 'LepBaselineRegression'; Status = $baselineStatus; Evidence = ($baselineRows -join '; ') }
        [pscustomobject]@{ Control = 'AutoLithHoles'; Status = $(if ($Stats['AUTOLITH_CONFIDENCE'].Holes -eq 150) { 'PASS' } else { 'REVIEW' }); Evidence = ('Holes={0}; Rows={1}; Rejected={2}' -f $Stats['AUTOLITH_CONFIDENCE'].Holes, $Stats['AUTOLITH_CONFIDENCE'].Rows, $Stats['AUTOLITH_CONFIDENCE'].BadRows) }
        [pscustomobject]@{ Control = 'AutoLithArgmaxRawMeans'; Status = 'PASS'; Evidence = ('No normalization before ARGMAX; ties={0}; noProbabilityBins={1}' -f (($autoJobs | Measure-Object -Property Ties -Sum).Sum), (($autoJobs | Measure-Object -Property NoProbabilityRows -Sum).Sum)) }
        [pscustomobject]@{ Control = 'SpenceIntervals'; Status = $(if ($spence.InvalidIntervals -eq 0 -and $spence.Overlaps -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Tables={0}; Intervals={1}; Invalid={2}; Overlaps={3}' -f ($spence.Tables -join '|'), $spence.Intervals, $spence.InvalidIntervals, $spence.Overlaps) }
        [pscustomobject]@{ Control = 'RejectedPointRows'; Status = $(if ((($Stats.Values | Measure-Object -Property BadRows -Sum).Sum) -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Rejected={0}' -f (($Stats.Values | Measure-Object -Property BadRows -Sum).Sum)) }
        [pscustomobject]@{ Control = 'SourceImmutability'; Status = $(if ($changed -eq 0) { 'PASS' } else { 'FAIL' }); Evidence = ('ChangedSources={0}' -f $changed) }
        [pscustomobject]@{ Control = 'CrossExportSource'; Status = 'REVIEW'; Evidence = 'AUTOLITH_CONFIDENCE comes from Exports_13_07_2026; approval required before COMPARE' }
    )

    $blocking = @($QaRows | Where-Object { $_.Status -eq 'FAIL' }).Count
    $reviews = @($QaRows | Where-Object { $_.Status -eq 'REVIEW' }).Count
    $Decision = if ($blocking -gt 0) { 'NO_GO' } elseif ($reviews -gt 0) { 'PASS_WITH_WARNINGS' } else { 'PASS' }

    Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Status 'Writing evidence' -PercentComplete 97
    $JobRecords | Export-Csv -LiteralPath (Join-Path $RunOutput 'Composite_Job_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
    $QaRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Composite_QAQC_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
    $SourceChecks | Export-Csv -LiteralPath (Join-Path $RunOutput 'Composite_Source_Verification.csv') -NoTypeInformation -Encoding utf8NoBOM

    $RunRecord = [ordered]@{
        Classification = $Classification
        ScriptName = $ScriptName
        ScriptVersion = $ScriptVersion
        RunId = $RunId
        InitRunRoot = $ResolvedInit
        OutputRoot = $RunOutput
        Scales = $ScaleNames
        MinimumCoverage = $MinimumCoverage
        SampleStep = $SampleStep
        GridRule = 'Bins aligned to multiples of the scale from depth zero'
        LepRule = 'Arithmetic mean of GT20, GT30, GT50 per bin; no ARGMAX; no renormalization'
        AutoLithRule = 'Arithmetic mean of each class probability per bin over available samples; ARGMAX on raw means; no normalization; epsilon 0.000001'
        SpenceRule = 'Length-weighted dominant code per bin and table'
        NullPolicy = 'Empty cells and -999.25 excluded from means; missing means written as empty cells'
        Decision = $Decision
        ElapsedSeconds = [math]::Round(((Get-Date) - $Started).TotalSeconds, 1)
        CreatedAt = (Get-Date).ToString('o')
    }
    ($RunRecord | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $RunOutput 'Run_Register.json') -Encoding utf8NoBOM

    $Readme = @(
        '# CP09A Composite Jobs'
        ''
        ('Classification: {0}' -f $Classification)
        ('Run ID: {0}' -f $RunId)
        ('Decision: {0}' -f $Decision)
        ''
        'Composites are written per source under Composites\<SourceId>.'
        'Bins are aligned to multiples of the scale from depth zero.'
        ('STATUS is PASS when coverage is at least {0}; otherwise PARTIAL.' -f $MinimumCoverage)
        'LEP sources use means only. AutoLith uses ARGMAX on raw mean probabilities. Spence uses the length-weighted dominant code.'
        'Source files are read-only and verified unchanged by SHA256.'
    ) -join [Environment]::NewLine
    $Readme | Set-Content -LiteralPath (Join-Path $RunOutput 'README.md') -Encoding utf8NoBOM

    $ChangeLog = @('# Changelog', '', '## v1.0.0', '- First execution engine for the 12 governed CP09A composite jobs.', '- LEP equivalence check between LAS_LEP and LEP_CONFIDENCE.', '- LEP regression check against the 2026-10-01 baseline.') -join [Environment]::NewLine
    $ChangeLog | Set-Content -LiteralPath (Join-Path $RunOutput 'CHANGELOG.md') -Encoding utf8NoBOM

    $ManifestItems = @(Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Sort-Object FullName | ForEach-Object { [pscustomobject]@{ RelativePath = [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName); Length = [long]$_.Length; SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
    ([ordered]@{ Classification = $Classification; RunId = $RunId; FileCount = $ManifestItems.Count; Files = $ManifestItems } | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $RunOutput 'Manifest.json') -Encoding utf8NoBOM

    $HashLines = @(Get-ChildItem -LiteralPath $RunOutput -File -Recurse | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object FullName | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, [System.IO.Path]::GetRelativePath($RunOutput, $_.FullName) })
    ($HashLines -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Encoding ascii

    Write-Host ''
    $JobRecords | Format-Table JobId, Holes, Rows, PassRows, PartialRows, NoProbabilityRows, Ties, Seconds -AutoSize
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
    Write-Progress -Id $ProgressId -Activity 'CP09A composite jobs' -Completed
}

exit $ExitCode
