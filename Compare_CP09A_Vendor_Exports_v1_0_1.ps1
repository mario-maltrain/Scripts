# Classification: Internal
# Script: Compare_CP09A_Vendor_Exports_v1_0_1.ps1
# Purpose: Vendor export delta review. Compares consecutive vendor exports by drillhole in metres, FROM/TO and data values,
#          writes a reviewable change register, an append-only change ledger and a visual HTML report.
# Safety: AuditOnly is the default. Sources are read-only and verified by SHA256. No delete, move, rename or overwrite.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExportsRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\01_Processed_Exports\GEOAI-EXPORTS-20261002-225517',

    [Parameter()]
    [ValidateCount(2, 50)]
    [string[]]$ExportIds = @('Exports_2026-09-26', 'Exports_2026-09-28'),

    [Parameter()]
    [AllowEmptyString()]
    [string]$HoleListPath = (Join-Path $env:USERPROFILE 'Downloads\Vendor_Hole_List.csv'),

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-VENDOR-DELTA-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-VENDOR-DELTA-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [ValidateRange(0.0, 1.0)]
    [double]$ValueTolerance = 0.000001,

    [Parameter()]
    [ValidateRange(0.0, 10.0)]
    [double]$MeterTolerance = 0.001,

    [Parameter()]
    [ValidatePattern('^[A-Za-z0-9_\-]+\.csv$')]
    [string]$LedgerName = 'Vendor_Change_Ledger.csv',

    [Parameter()]
    [switch]$CommitRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Compare_CP09A_Vendor_Exports_v1_0_1.ps1'
$ScriptVersion = '1.0.1'
$Classification = 'Internal'
$ProgressId = 1800
$ExitCode = 0
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$NumberStyle = [System.Globalization.NumberStyles]::Float
$QuoteTrim = [char[]]@([char]34, [char]32)
$NullTokens = @('', '-99', '-999', '-999.25', '-9999', 'NA', 'N/A', 'NULL', 'NAN')
$MetaColumns = @('HOLE_ID', 'HOLEID', 'EXPORT_ID', 'SOURCE_FILE', 'RUN_ID', 'EXPORT DATE')
$Datasets = @(
    [pscustomobject]@{ DatasetId = 'GEOAI_LEP'; RelativePath = '05_Ready_For_Compositing\02_LEP_CONFIDENCE'; Kind = 'POINT' }
    [pscustomobject]@{ DatasetId = 'GEOAI_AUTOLITH'; RelativePath = '05_Ready_For_Compositing\01_AUTOLITH_CONFIDENCE'; Kind = 'POINT' }
    [pscustomobject]@{ DatasetId = 'GEOAI_AUTOLOGGER'; RelativePath = '06_Consolidated_Database\03_SPENCE_MASTER'; Kind = 'INTERVAL' }
)

Add-Type -AssemblyName Microsoft.VisualBasic

$script:Changes = [System.Collections.Generic.List[object]]::new()
$script:Summaries = [System.Collections.Generic.List[object]]::new()
$script:FileRegister = [System.Collections.Generic.List[object]]::new()
$script:HashBefore = @{}
$script:VendorHoles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:PairBase = ''
$script:PairCand = ''

function Get-Product { param([double]$A, [double]$B) return [math]::FusedMultiplyAdd($A, $B, 0.0) }
function Format-M { param([double]$Value) if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return '' } return ([math]::Round($Value, 3)).ToString($Invariant) }
function Format-V { param([double]$Value) if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return '' } return $Value.ToString('R', $Invariant) }

function ConvertTo-NumberOrNaN {
    param([AllowEmptyString()][string]$Text)
    $clean = $Text.Trim($QuoteTrim)
    if ($NullTokens -contains $clean.ToUpperInvariant()) { return [double]::NaN }
    $number = 0.0
    if ([double]::TryParse($clean, $NumberStyle, $Invariant, [ref]$number)) { return $number }
    return [double]::NaN
}

function Get-SourceHash {
    param([Parameter(Mandatory)][string]$Path)
    if (-not $script:HashBefore.ContainsKey($Path)) { $script:HashBefore[$Path] = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
    return $script:HashBefore[$Path]
}

function Get-AllCsvCount {
    param([AllowEmptyString()][string]$Folder)
    if ([string]::IsNullOrEmpty($Folder) -or -not (Test-Path -LiteralPath $Folder -PathType Container)) { return -1 }
    return @(Get-ChildItem -LiteralPath $Folder -File -Recurse | Where-Object { $_.Extension -eq '.csv' }).Count
}

function Get-EligibleFiles {
    param([AllowEmptyString()][string]$Folder)
    if ([string]::IsNullOrEmpty($Folder) -or -not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    return @(
        Get-ChildItem -LiteralPath $Folder -File -Recurse |
            Where-Object { ($_.Extension -eq '.csv') -and ($_.Name -notmatch '__DUP_SHA256_EXACT|__COPY_\d+|\bCopiar\b') } |
            Sort-Object FullName
    )
}

function Add-Change {
    param(
        [string]$DatasetId, [string]$Table, [string]$HoleId, [string]$ChangeType, [ValidateSet('INFO', 'WARN', 'MAJOR')][string]$Severity,
        [string]$Metric, [string]$BaselineValue, [string]$CandidateValue, [string]$Delta, [string]$Detail
    )
    $script:Changes.Add([pscustomobject]@{
        ChangeId = ''; RunId = $RunId; DetectedAt = (Get-Date).ToString('s'); BaselineExport = $script:PairBase; CandidateExport = $script:PairCand
        DatasetId = $DatasetId; Table = $Table; HoleId = $HoleId; InVendorList = $(if ($script:VendorHoles.Contains($HoleId)) { 'Y' } else { 'N' })
        ChangeType = $ChangeType; Severity = $Severity; Metric = $Metric; BaselineValue = $BaselineValue; CandidateValue = $CandidateValue
        Delta = $Delta; Detail = $Detail; ReviewStatus = 'OPEN'; ReviewDecision = ''; Reviewer = ''; ReviewDate = ''; ReviewNotes = ''
    })
}

function Add-Summary {
    param([string]$DatasetId, [string]$Table, [string]$HoleId, [string]$Status,
        [double]$BaseFrom, [double]$BaseTo, [double]$BaseMeters, [long]$BaseUnits,
        [double]$CandFrom, [double]$CandTo, [double]$CandMeters, [long]$CandUnits, [int]$ChangeCount, [string]$MaxSeverity)
    $script:Summaries.Add([pscustomobject]@{
        BaselineExport = $script:PairBase; CandidateExport = $script:PairCand; DatasetId = $DatasetId; Table = $Table; HoleId = $HoleId
        InVendorList = $(if ($script:VendorHoles.Contains($HoleId)) { 'Y' } else { 'N' }); Status = $Status
        BaselineFrom = (Format-M $BaseFrom); BaselineTo = (Format-M $BaseTo); BaselineMeters = (Format-M $BaseMeters); BaselineUnits = $BaseUnits
        CandidateFrom = (Format-M $CandFrom); CandidateTo = (Format-M $CandTo); CandidateMeters = (Format-M $CandMeters); CandidateUnits = $CandUnits
        DeltaMeters = $(if ([double]::IsNaN($BaseMeters) -or [double]::IsNaN($CandMeters)) { '' } else { Format-M ($CandMeters - $BaseMeters) })
        ChangeCount = $ChangeCount; MaxSeverity = $MaxSeverity
    })
}

function Get-MaxSeverity {
    param([int]$StartIndex)
    $max = 'NONE'
    for ($k = $StartIndex; $k -lt $script:Changes.Count; $k++) {
        $s = $script:Changes[$k].Severity
        if ($s -eq 'MAJOR') { return 'MAJOR' }
        if ($s -eq 'WARN') { $max = 'WARN' } elseif ($s -eq 'INFO' -and $max -eq 'NONE') { $max = 'INFO' }
    }
    return $max
}

function Read-PointFile {
    param([Parameter(Mandatory)][string]$Path)
    $reader = [System.IO.StreamReader]::new($Path)
    try {
        $headerLine = $reader.ReadLine()
        if ($null -eq $headerLine) { throw ('Empty file: {0}' -f $Path) }
        $header = @($headerLine.Split(',') | ForEach-Object { $_.Trim($QuoteTrim) })
        $upper = @($header | ForEach-Object { $_.ToUpperInvariant() })
        $depthIndex = [array]::IndexOf($upper, 'DEPT')
        if ($depthIndex -lt 0) { $depthIndex = [array]::IndexOf($upper, 'DEPTH') }
        if ($depthIndex -lt 0) { throw ('DEPT column missing: {0}' -f $Path) }
        $valueIndex = [System.Collections.Generic.List[int]]::new()
        $valueNames = [System.Collections.Generic.List[string]]::new()
        for ($c = 0; $c -lt $header.Count; $c++) {
            if ($c -ne $depthIndex -and $MetaColumns -notcontains $upper[$c]) { $valueIndex.Add($c); $valueNames.Add($upper[$c]) }
        }
        $data = [System.Collections.Generic.Dictionary[long, double[]]]::new()
        $min = [double]::PositiveInfinity; $max = [double]::NegativeInfinity; $bad = 0L; $dup = 0L
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.Length -eq 0) { continue }
            $tokens = $line.Split(',')
            if ($tokens.Count -ne $header.Count) { $bad++; continue }
            $depth = ConvertTo-NumberOrNaN -Text $tokens[$depthIndex]
            if ([double]::IsNaN($depth)) { $bad++; continue }
            $key = [long][math]::Round((Get-Product -A $depth -B 1000.0))
            if ($data.ContainsKey($key)) { $dup++; continue }
            $values = [double[]]::new($valueIndex.Count)
            for ($k = 0; $k -lt $valueIndex.Count; $k++) { $values[$k] = ConvertTo-NumberOrNaN -Text $tokens[$valueIndex[$k]] }
            $data.Add($key, $values)
            if ($depth -lt $min) { $min = $depth }
            if ($depth -gt $max) { $max = $depth }
        }
    }
    finally { $reader.Dispose() }
    $count = $data.Count
    if ($count -eq 0) { $min = [double]::NaN; $max = [double]::NaN }
    return [pscustomobject]@{ ValueNames = $valueNames.ToArray(); Data = $data; Min = $min; Max = $max; Meters = $(if ($count -gt 0) { $max - $min } else { [double]::NaN }); Count = [long]$count; BadRows = $bad; DuplicateDepths = $dup; Step = $(if ($count -gt 1) { ($max - $min) / ($count - 1) } else { 0.0 }) }
}

function Compare-PointDataset {
    param([Parameter(Mandatory)][object]$Dataset, [AllowEmptyString()][string]$BaseFolder, [AllowEmptyString()][string]$CandFolder)
    $maps = @()
    foreach ($folder in @($BaseFolder, $CandFolder)) {
        $map = @{}
        foreach ($file in @(Get-EligibleFiles -Folder $folder)) {
            $m = [regex]::Match($file.Name, '^(SPD\d+)_', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $m.Success) { $script:FileRegister.Add([pscustomobject]@{ Export = $(if ($folder -eq $BaseFolder) { $script:PairBase } else { $script:PairCand }); DatasetId = $Dataset.DatasetId; File = $file.Name; HoleId = ''; SHA256 = (Get-SourceHash -Path $file.FullName); Issue = 'HOLE_ID_NOT_IN_FILE_NAME' }); continue }
            $hole = $m.Groups[1].Value.ToUpperInvariant()
            if ($map.ContainsKey($hole)) { $script:FileRegister.Add([pscustomobject]@{ Export = $(if ($folder -eq $BaseFolder) { $script:PairBase } else { $script:PairCand }); DatasetId = $Dataset.DatasetId; File = $file.Name; HoleId = $hole; SHA256 = (Get-SourceHash -Path $file.FullName); Issue = 'DUPLICATE_HOLE_FILE_IGNORED' }); continue }
            $map[$hole] = $file
        }
        $maps += ,$map
    }
    $baseMap = $maps[0]; $candMap = $maps[1]
    $holes = @(@($baseMap.Keys) + @($candMap.Keys) | Sort-Object -Unique)
    for ($i = 0; $i -lt $holes.Count; $i++) {
        $hole = $holes[$i]
        Write-Progress -Id $ProgressId -Activity ('Vendor delta {0} vs {1}' -f $script:PairBase, $script:PairCand) -Status ('{0}: {1} ({2} of {3})' -f $Dataset.DatasetId, $hole, ($i + 1), $holes.Count) -PercentComplete ([math]::Min(99, [int][math]::Floor($i / [math]::Max(1, $holes.Count) / 0.01)))
        $start = $script:Changes.Count
        $bFile = $(if ($baseMap.ContainsKey($hole)) { $baseMap[$hole] } else { $null })
        $cFile = $(if ($candMap.ContainsKey($hole)) { $candMap[$hole] } else { $null })
        $bHash = $(if ($null -ne $bFile) { Get-SourceHash -Path $bFile.FullName } else { '' })
        $cHash = $(if ($null -ne $cFile) { Get-SourceHash -Path $cFile.FullName } else { '' })
        if ($null -ne $bFile) { $script:FileRegister.Add([pscustomobject]@{ Export = $script:PairBase; DatasetId = $Dataset.DatasetId; File = $bFile.Name; HoleId = $hole; SHA256 = $bHash; Issue = '' }) }
        if ($null -ne $cFile) { $script:FileRegister.Add([pscustomobject]@{ Export = $script:PairCand; DatasetId = $Dataset.DatasetId; File = $cFile.Name; HoleId = $hole; SHA256 = $cHash; Issue = '' }) }

        if ($null -ne $bFile -and $null -ne $cFile -and $bHash -eq $cHash) {
            $p = Read-PointFile -Path $bFile.FullName
            Add-Summary -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -Status 'IDENTICAL' -BaseFrom $p.Min -BaseTo $p.Max -BaseMeters $p.Meters -BaseUnits $p.Count -CandFrom $p.Min -CandTo $p.Max -CandMeters $p.Meters -CandUnits $p.Count -ChangeCount 0 -MaxSeverity 'NONE'
            continue
        }
        $pb = $(if ($null -ne $bFile) { Read-PointFile -Path $bFile.FullName } else { $null })
        $pc = $(if ($null -ne $cFile) { Read-PointFile -Path $cFile.FullName } else { $null })
        if ($null -eq $pc) {
            Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'HOLE_REMOVED' -Severity 'MAJOR' -Metric 'LengthM' -BaselineValue (Format-M $pb.Meters) -CandidateValue '' -Delta (Format-M (0 - $pb.Meters)) -Detail ('Samples={0}' -f $pb.Count)
            Add-Summary -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -Status 'REMOVED' -BaseFrom $pb.Min -BaseTo $pb.Max -BaseMeters $pb.Meters -BaseUnits $pb.Count -CandFrom ([double]::NaN) -CandTo ([double]::NaN) -CandMeters ([double]::NaN) -CandUnits 0 -ChangeCount 1 -MaxSeverity 'MAJOR'
            continue
        }
        if ($null -eq $pb) {
            Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'HOLE_ADDED' -Severity 'MAJOR' -Metric 'LengthM' -BaselineValue '' -CandidateValue (Format-M $pc.Meters) -Delta (Format-M $pc.Meters) -Detail ('Samples={0}' -f $pc.Count)
            Add-Summary -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -Status 'ADDED' -BaseFrom ([double]::NaN) -BaseTo ([double]::NaN) -BaseMeters ([double]::NaN) -BaseUnits 0 -CandFrom $pc.Min -CandTo $pc.Max -CandMeters $pc.Meters -CandUnits $pc.Count -ChangeCount 1 -MaxSeverity 'MAJOR'
            continue
        }
        foreach ($name in @($pb.ValueNames | Where-Object { $pc.ValueNames -notcontains $_ })) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'COLUMN_REMOVED' -Severity 'MAJOR' -Metric $name -BaselineValue 'present' -CandidateValue 'absent' -Delta '' -Detail '' }
        foreach ($name in @($pc.ValueNames | Where-Object { $pb.ValueNames -notcontains $_ })) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'COLUMN_ADDED' -Severity 'MAJOR' -Metric $name -BaselineValue 'absent' -CandidateValue 'present' -Delta '' -Detail '' }
        if ([math]::Abs($pc.Min - $pb.Min) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'FROM_CHANGED' -Severity 'MAJOR' -Metric 'FromM' -BaselineValue (Format-M $pb.Min) -CandidateValue (Format-M $pc.Min) -Delta (Format-M ($pc.Min - $pb.Min)) -Detail '' }
        if ([math]::Abs($pc.Max - $pb.Max) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'TO_CHANGED' -Severity 'MAJOR' -Metric 'ToM' -BaselineValue (Format-M $pb.Max) -CandidateValue (Format-M $pc.Max) -Delta (Format-M ($pc.Max - $pb.Max)) -Detail '' }
        if ([math]::Abs($pc.Meters - $pb.Meters) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'LENGTH_CHANGED' -Severity 'MAJOR' -Metric 'LengthM' -BaselineValue (Format-M $pb.Meters) -CandidateValue (Format-M $pc.Meters) -Delta (Format-M ($pc.Meters - $pb.Meters)) -Detail '' }
        $added = 0L; $removed = 0L
        foreach ($key in $pc.Data.Keys) { if (-not $pb.Data.ContainsKey($key)) { $added++ } }
        foreach ($key in $pb.Data.Keys) { if (-not $pc.Data.ContainsKey($key)) { $removed++ } }
        if ($added -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'SAMPLES_ADDED' -Severity 'WARN' -Metric 'Samples' -BaselineValue ([string]$pb.Count) -CandidateValue ([string]$pc.Count) -Delta ([string]$added) -Detail ('ApproxMeters={0}' -f (Format-M (Get-Product -A $added -B $pc.Step))) }
        if ($removed -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'SAMPLES_REMOVED' -Severity 'WARN' -Metric 'Samples' -BaselineValue ([string]$pb.Count) -CandidateValue ([string]$pc.Count) -Delta ([string](0 - $removed)) -Detail ('ApproxMeters={0}' -f (Format-M (Get-Product -A $removed -B $pb.Step))) }
        $common = @($pb.ValueNames | Where-Object { $pc.ValueNames -contains $_ })
        foreach ($name in $common) {
            $ib = [array]::IndexOf($pb.ValueNames, $name)
            $ic = [array]::IndexOf($pc.ValueNames, $name)
            $changed = 0L; $compared = 0L; $maxAbs = 0.0
            foreach ($key in $pb.Data.Keys) {
                $vc = $null
                if (-not $pc.Data.TryGetValue($key, [ref]$vc)) { continue }
                $compared++
                $a = $pb.Data[$key][$ib]; $b = $vc[$ic]
                $aNaN = [double]::IsNaN($a); $bNaN = [double]::IsNaN($b)
                if ($aNaN -and $bNaN) { continue }
                if ($aNaN -ne $bNaN) { $changed++; continue }
                $d = [math]::Abs($b - $a)
                if ($d -gt $ValueTolerance) { $changed++; if ($d -gt $maxAbs) { $maxAbs = $d } }
            }
            if ($changed -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -ChangeType 'VALUE_CHANGED' -Severity 'WARN' -Metric $name -BaselineValue ('compared={0}' -f $compared) -CandidateValue ('changed={0}' -f $changed) -Delta (Format-V $maxAbs) -Detail 'Delta is the maximum absolute difference' }
        }
        $count = $script:Changes.Count - $start
        Add-Summary -DatasetId $Dataset.DatasetId -Table 'POINT' -HoleId $hole -Status $(if ($count -gt 0) { 'CHANGED' } else { 'CONTENT_EQUAL' }) -BaseFrom $pb.Min -BaseTo $pb.Max -BaseMeters $pb.Meters -BaseUnits $pb.Count -CandFrom $pc.Min -CandTo $pc.Max -CandMeters $pc.Meters -CandUnits $pc.Count -ChangeCount $count -MaxSeverity (Get-MaxSeverity -StartIndex $start)
    }
}

function Read-IntervalFile {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($File.FullName, [System.Text.Encoding]::UTF8)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters([string[]]@(','))
    $parser.HasFieldsEnclosedInQuotes = $true
    $holes = @{}; $bad = 0L
    try {
        if ($parser.EndOfData) { throw ('Empty file: {0}' -f $File.Name) }
        $header = @($parser.ReadFields() | ForEach-Object { $_.Trim() })
        $upper = @($header | ForEach-Object { $_.ToUpperInvariant() })
        $hi = [array]::IndexOf($upper, 'HOLEID'); if ($hi -lt 0) { $hi = [array]::IndexOf($upper, 'HOLE_ID') }
        $fi = [array]::IndexOf($upper, 'FROM'); $ti = [array]::IndexOf($upper, 'TO')
        if ($hi -lt 0 -or $fi -lt 0 -or $ti -lt 0) { throw ('HOLEID, FROM or TO missing: {0}' -f $File.Name) }
        $ai = -1
        for ($c = 0; $c -lt $header.Count; $c++) {
            if ($c -eq $hi -or $c -eq $fi -or $c -eq $ti) { continue }
            if ($MetaColumns -contains $upper[$c] -or $upper[$c].EndsWith('_DESC')) { continue }
            $ai = $c; break
        }
        if ($ai -lt 0) { throw ('Attribute column not resolved: {0}' -f $File.Name) }
        while (-not $parser.EndOfData) {
            $fields = $null
            try { $fields = $parser.ReadFields() } catch { $bad++; continue }
            if ($null -eq $fields -or $fields.Count -ne $header.Count) { $bad++; continue }
            $hole = $fields[$hi].Trim().ToUpperInvariant()
            $f = ConvertTo-NumberOrNaN -Text $fields[$fi]; $t = ConvertTo-NumberOrNaN -Text $fields[$ti]
            if ($hole.Length -eq 0 -or [double]::IsNaN($f) -or [double]::IsNaN($t) -or $t -le $f) { $bad++; continue }
            if (-not $holes.ContainsKey($hole)) { $holes[$hole] = [System.Collections.Generic.List[object]]::new() }
            $holes[$hole].Add([pscustomobject]@{ From = $f; To = $t; Code = $fields[$ai].Trim() })
        }
    }
    finally { $parser.Dispose() }
    return [pscustomobject]@{ Attribute = $header[$ai]; Holes = $holes; BadRows = $bad; File = $File }
}

function Get-IntervalStats {
    param([AllowNull()][object]$List)
    if ($null -eq $List -or $List.Count -eq 0) { return [pscustomobject]@{ From = [double]::NaN; To = [double]::NaN; Meters = [double]::NaN; Count = 0L; ByKey = @{}; ByCode = @{} } }
    $min = [double]::PositiveInfinity; $max = [double]::NegativeInfinity; $meters = 0.0; $byKey = @{}; $byCode = @{}
    foreach ($iv in $List) {
        $len = $iv.To - $iv.From
        $meters += $len
        if ($iv.From -lt $min) { $min = $iv.From }
        if ($iv.To -gt $max) { $max = $iv.To }
        $key = '{0}|{1}' -f $iv.From.ToString('F3', $Invariant), $iv.To.ToString('F3', $Invariant)
        $byKey[$key] = $iv
        if ($byCode.ContainsKey($iv.Code)) { $byCode[$iv.Code] += $len } else { $byCode[$iv.Code] = $len }
    }
    return [pscustomobject]@{ From = $min; To = $max; Meters = $meters; Count = [long]$List.Count; ByKey = $byKey; ByCode = $byCode }
}

function Compare-IntervalDataset {
    param([Parameter(Mandatory)][object]$Dataset, [AllowEmptyString()][string]$BaseFolder, [AllowEmptyString()][string]$CandFolder)
    $sides = @()
    foreach ($folder in @($BaseFolder, $CandFolder)) {
        $tables = @{}
        foreach ($file in @(Get-EligibleFiles -Folder $folder)) {
            $hash = Get-SourceHash -Path $file.FullName
            $t = Read-IntervalFile -File $file
            $script:FileRegister.Add([pscustomobject]@{ Export = $(if ($folder -eq $BaseFolder) { $script:PairBase } else { $script:PairCand }); DatasetId = $Dataset.DatasetId; File = $file.Name; HoleId = ''; SHA256 = $hash; Issue = $(if ($t.BadRows -gt 0) { 'BAD_ROWS={0}' -f $t.BadRows } else { '' }) })
            $t | Add-Member -NotePropertyName Hash -NotePropertyValue $hash
            $tables[$t.Attribute] = $t
        }
        $sides += ,$tables
    }
    $baseTables = $sides[0]; $candTables = $sides[1]
    $attributes = @(@($baseTables.Keys) + @($candTables.Keys) | Sort-Object -Unique)
    foreach ($attr in $attributes) {
        $bt = $(if ($baseTables.ContainsKey($attr)) { $baseTables[$attr] } else { $null })
        $ct = $(if ($candTables.ContainsKey($attr)) { $candTables[$attr] } else { $null })
        if ($null -eq $bt) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId '' -ChangeType 'TABLE_ADDED' -Severity 'MAJOR' -Metric 'Table' -BaselineValue 'absent' -CandidateValue 'present' -Delta '' -Detail $ct.File.Name }
        if ($null -eq $ct) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId '' -ChangeType 'TABLE_REMOVED' -Severity 'MAJOR' -Metric 'Table' -BaselineValue 'present' -CandidateValue 'absent' -Delta '' -Detail $bt.File.Name }
        $identical = ($null -ne $bt -and $null -ne $ct -and $bt.Hash -eq $ct.Hash)
        $bHoles = $(if ($null -ne $bt) { $bt.Holes } else { @{} })
        $cHoles = $(if ($null -ne $ct) { $ct.Holes } else { @{} })
        $holes = @(@($bHoles.Keys) + @($cHoles.Keys) | Sort-Object -Unique)
        foreach ($hole in $holes) {
            $start = $script:Changes.Count
            $sb = Get-IntervalStats -List $(if ($bHoles.ContainsKey($hole)) { $bHoles[$hole] } else { $null })
            $sc = Get-IntervalStats -List $(if ($cHoles.ContainsKey($hole)) { $cHoles[$hole] } else { $null })
            if ($identical) {
                Add-Summary -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -Status 'IDENTICAL' -BaseFrom $sb.From -BaseTo $sb.To -BaseMeters $sb.Meters -BaseUnits $sb.Count -CandFrom $sb.From -CandTo $sb.To -CandMeters $sb.Meters -CandUnits $sb.Count -ChangeCount 0 -MaxSeverity 'NONE'
                continue
            }
            if ($sc.Count -eq 0) {
                Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'HOLE_REMOVED' -Severity 'MAJOR' -Metric 'Meters' -BaselineValue (Format-M $sb.Meters) -CandidateValue '' -Delta (Format-M (0 - $sb.Meters)) -Detail ('Intervals={0}' -f $sb.Count)
                Add-Summary -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -Status 'REMOVED' -BaseFrom $sb.From -BaseTo $sb.To -BaseMeters $sb.Meters -BaseUnits $sb.Count -CandFrom ([double]::NaN) -CandTo ([double]::NaN) -CandMeters ([double]::NaN) -CandUnits 0 -ChangeCount 1 -MaxSeverity 'MAJOR'
                continue
            }
            if ($sb.Count -eq 0) {
                Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'HOLE_ADDED' -Severity 'MAJOR' -Metric 'Meters' -BaselineValue '' -CandidateValue (Format-M $sc.Meters) -Delta (Format-M $sc.Meters) -Detail ('Intervals={0}' -f $sc.Count)
                Add-Summary -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -Status 'ADDED' -BaseFrom ([double]::NaN) -BaseTo ([double]::NaN) -BaseMeters ([double]::NaN) -BaseUnits 0 -CandFrom $sc.From -CandTo $sc.To -CandMeters $sc.Meters -CandUnits $sc.Count -ChangeCount 1 -MaxSeverity 'MAJOR'
                continue
            }
            if ([math]::Abs($sc.From - $sb.From) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'FROM_CHANGED' -Severity 'MAJOR' -Metric 'FromM' -BaselineValue (Format-M $sb.From) -CandidateValue (Format-M $sc.From) -Delta (Format-M ($sc.From - $sb.From)) -Detail '' }
            if ([math]::Abs($sc.To - $sb.To) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'TO_CHANGED' -Severity 'MAJOR' -Metric 'ToM' -BaselineValue (Format-M $sb.To) -CandidateValue (Format-M $sc.To) -Delta (Format-M ($sc.To - $sb.To)) -Detail '' }
            if ([math]::Abs($sc.Meters - $sb.Meters) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'METERS_CHANGED' -Severity 'MAJOR' -Metric 'Meters' -BaselineValue (Format-M $sb.Meters) -CandidateValue (Format-M $sc.Meters) -Delta (Format-M ($sc.Meters - $sb.Meters)) -Detail '' }
            $addedN = 0; $addedM = 0.0; $removedN = 0; $removedM = 0.0; $codeN = 0; $codeM = 0.0
            foreach ($key in $sc.ByKey.Keys) { if (-not $sb.ByKey.ContainsKey($key)) { $addedN++; $addedM += ($sc.ByKey[$key].To - $sc.ByKey[$key].From) } }
            foreach ($key in $sb.ByKey.Keys) {
                if (-not $sc.ByKey.ContainsKey($key)) { $removedN++; $removedM += ($sb.ByKey[$key].To - $sb.ByKey[$key].From); continue }
                if ($sb.ByKey[$key].Code -ne $sc.ByKey[$key].Code) { $codeN++; $codeM += ($sb.ByKey[$key].To - $sb.ByKey[$key].From) }
            }
            if ($addedN -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'INTERVALS_ADDED' -Severity 'WARN' -Metric 'Intervals' -BaselineValue ([string]$sb.Count) -CandidateValue ([string]$sc.Count) -Delta ([string]$addedN) -Detail ('Meters={0}' -f (Format-M $addedM)) }
            if ($removedN -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'INTERVALS_REMOVED' -Severity 'WARN' -Metric 'Intervals' -BaselineValue ([string]$sb.Count) -CandidateValue ([string]$sc.Count) -Delta ([string](0 - $removedN)) -Detail ('Meters={0}' -f (Format-M $removedM)) }
            if ($codeN -gt 0) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'CODE_CHANGED' -Severity 'MAJOR' -Metric 'SameFromTo' -BaselineValue ('intervals={0}' -f $codeN) -CandidateValue '' -Delta (Format-M $codeM) -Detail 'Delta is the metres with a different code on identical FROM/TO' }
            $codes = @(@($sb.ByCode.Keys) + @($sc.ByCode.Keys) | Sort-Object -Unique)
            foreach ($code in $codes) {
                $mb = $(if ($sb.ByCode.ContainsKey($code)) { [double]$sb.ByCode[$code] } else { 0.0 })
                $mc = $(if ($sc.ByCode.ContainsKey($code)) { [double]$sc.ByCode[$code] } else { 0.0 })
                if ([math]::Abs($mc - $mb) -gt $MeterTolerance) { Add-Change -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -ChangeType 'METERS_BY_CODE_CHANGED' -Severity 'WARN' -Metric $code -BaselineValue (Format-M $mb) -CandidateValue (Format-M $mc) -Delta (Format-M ($mc - $mb)) -Detail '' }
            }
            $count = $script:Changes.Count - $start
            Add-Summary -DatasetId $Dataset.DatasetId -Table $attr -HoleId $hole -Status $(if ($count -gt 0) { 'CHANGED' } else { 'CONTENT_EQUAL' }) -BaseFrom $sb.From -BaseTo $sb.To -BaseMeters $sb.Meters -BaseUnits $sb.Count -CandFrom $sc.From -CandTo $sc.To -CandMeters $sc.Meters -CandUnits $sc.Count -ChangeCount $count -MaxSeverity (Get-MaxSeverity -StartIndex $start)
        }
    }
}

function ConvertTo-HtmlTable {
    param([object[]]$Rows, [string[]]$Columns)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table><tr>')
    foreach ($c in $Columns) { [void]$sb.Append('<th>' + [System.Net.WebUtility]::HtmlEncode($c) + '</th>') }
    [void]$sb.Append('</tr>')
    foreach ($r in $Rows) {
        $cls = ''
        if ($r.PSObject.Properties.Name -contains 'MaxSeverity') { $cls = [string]$r.MaxSeverity }
        elseif ($r.PSObject.Properties.Name -contains 'Severity') { $cls = [string]$r.Severity }
        [void]$sb.Append('<tr class="' + $cls + '">')
        foreach ($c in $Columns) { [void]$sb.Append('<td>' + [System.Net.WebUtility]::HtmlEncode([string]$r.$c) + '</td>') }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

try {
    $Started = Get-Date
    if (-not (Test-Path -LiteralPath $ExportsRoot -PathType Container)) { throw "ExportsRoot does not exist: $ExportsRoot" }
    $ResolvedRoot = (Get-Item -LiteralPath $ExportsRoot).FullName
    foreach ($id in $ExportIds) { if (-not (Test-Path -LiteralPath (Join-Path $ResolvedRoot $id) -PathType Container)) { throw "Export folder does not exist: $id" } }
    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) { throw 'OutputRoot is required when CommitRecords is used.' }
    $HoleListHash = ''
    if (-not [string]::IsNullOrWhiteSpace($HoleListPath) -and (Test-Path -LiteralPath $HoleListPath -PathType Leaf)) {
        foreach ($row in @(Import-Csv -LiteralPath $HoleListPath)) { $hid = ([string]$row.HOLE_ID).Trim().ToUpperInvariant(); if ($hid.Length -gt 0) { [void]$script:VendorHoles.Add($hid) } }
        $HoleListHash = (Get-FileHash -LiteralPath $HoleListPath -Algorithm SHA256).Hash
    }

    $DatasetRows = [System.Collections.Generic.List[object]]::new()
    for ($p = 1; $p -lt $ExportIds.Count; $p++) {
        $script:PairBase = $ExportIds[$p - 1]
        $script:PairCand = $ExportIds[$p]
        foreach ($ds in $Datasets) {
            $bf = Join-Path (Join-Path $ResolvedRoot $script:PairBase) $ds.RelativePath
            $cf = Join-Path (Join-Path $ResolvedRoot $script:PairCand) $ds.RelativePath
            $bn = @(Get-EligibleFiles -Folder $bf).Count
            $cn = @(Get-EligibleFiles -Folder $cf).Count
            $ba = Get-AllCsvCount -Folder $bf
            $ca = Get-AllCsvCount -Folder $cf
            $notComparable = (($bn -eq 0 -and $ba -gt 0) -or ($cn -eq 0 -and $ca -gt 0))
            $dsStatus = if ($bn -eq 0 -and $cn -eq 0) { 'NOT_PRESENT' } elseif ($notComparable) { 'NOT_COMPARABLE' } else { 'COMPARED' }
            $DatasetRows.Add([pscustomobject]@{ BaselineExport = $script:PairBase; CandidateExport = $script:PairCand; DatasetId = $ds.DatasetId; Kind = $ds.Kind; BaselineFiles = $bn; CandidateFiles = $cn; BaselineCsvAll = $ba; CandidateCsvAll = $ca; Status = $dsStatus })
            if ($bn -eq 0 -and $cn -eq 0) { continue }
            if ($notComparable) {
                Add-Change -DatasetId $ds.DatasetId -Table '' -HoleId '' -ChangeType 'DATASET_NOT_COMPARABLE' -Severity 'WARN' -Metric 'Files' -BaselineValue ('eligible={0}; all={1}' -f $bn, $ba) -CandidateValue ('eligible={0}; all={1}' -f $cn, $ca) -Delta '' -Detail ('Folder has CSV files but none eligible (duplicates or copies only): {0}' -f $ds.RelativePath)
                continue
            }
            if ($bn -eq 0) { Add-Change -DatasetId $ds.DatasetId -Table '' -HoleId '' -ChangeType 'DATASET_ADDED' -Severity 'MAJOR' -Metric 'Files' -BaselineValue '0' -CandidateValue ([string]$cn) -Delta ([string]$cn) -Detail $ds.RelativePath }
            if ($cn -eq 0) { Add-Change -DatasetId $ds.DatasetId -Table '' -HoleId '' -ChangeType 'DATASET_REMOVED' -Severity 'MAJOR' -Metric 'Files' -BaselineValue ([string]$bn) -CandidateValue '0' -Delta ([string](0 - $bn)) -Detail $ds.RelativePath }
            if ($ds.Kind -eq 'POINT') { Compare-PointDataset -Dataset $ds -BaseFolder $bf -CandFolder $cf }
            else { Compare-IntervalDataset -Dataset $ds -BaseFolder $bf -CandFolder $cf }
        }
    }

    for ($k = 0; $k -lt $script:Changes.Count; $k++) { $script:Changes[$k].ChangeId = ('{0}-{1:000000}' -f $RunId, ($k + 1)) }

    $changed = 0
    foreach ($path in $script:HashBefore.Keys) { if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $script:HashBefore[$path]) { $changed++ } }

    $DatasetSummary = @(
        $script:Summaries | Group-Object BaselineExport, CandidateExport, DatasetId, Table | ForEach-Object {
            $g = @($_.Group)
            $bm = [double](($g | Where-Object { $_.BaselineMeters -ne '' } | ForEach-Object { [double]::Parse($_.BaselineMeters, $Invariant) } | Measure-Object -Sum).Sum)
            $cm = [double](($g | Where-Object { $_.CandidateMeters -ne '' } | ForEach-Object { [double]::Parse($_.CandidateMeters, $Invariant) } | Measure-Object -Sum).Sum)
            [pscustomobject]@{
                BaselineExport = $g[0].BaselineExport; CandidateExport = $g[0].CandidateExport; DatasetId = $g[0].DatasetId; Table = $g[0].Table
                Holes = $g.Count; Identical = @($g | Where-Object { $_.Status -in @('IDENTICAL', 'CONTENT_EQUAL') }).Count
                Changed = @($g | Where-Object { $_.Status -eq 'CHANGED' }).Count; Added = @($g | Where-Object { $_.Status -eq 'ADDED' }).Count; Removed = @($g | Where-Object { $_.Status -eq 'REMOVED' }).Count
                BaselineMeters = (Format-M $bm); CandidateMeters = (Format-M $cm); DeltaMeters = (Format-M ($cm - $bm))
                MaxSeverity = $(if (@($g | Where-Object { $_.MaxSeverity -eq 'MAJOR' }).Count -gt 0) { 'MAJOR' } elseif (@($g | Where-Object { $_.MaxSeverity -eq 'WARN' }).Count -gt 0) { 'WARN' } else { 'NONE' })
            }
        }
    )
    $ChangeTypeSummary = @($script:Changes | Group-Object DatasetId, ChangeType, Severity | ForEach-Object { [pscustomobject]@{ DatasetId = $_.Group[0].DatasetId; ChangeType = $_.Group[0].ChangeType; Severity = $_.Group[0].Severity; Count = $_.Count; VendorHoles = @($_.Group | Where-Object { $_.InVendorList -eq 'Y' } | Select-Object -ExpandProperty HoleId -Unique).Count } } | Sort-Object DatasetId, ChangeType)
    $Major = @($script:Changes | Where-Object { $_.Severity -eq 'MAJOR' }).Count
    $Warn = @($script:Changes | Where-Object { $_.Severity -eq 'WARN' }).Count
    $QaRows = @(
        [pscustomobject]@{ Control = 'ExportsCompared'; Status = 'PASS'; Evidence = ($ExportIds -join ' > ') }
        [pscustomobject]@{ Control = 'VendorHoleList'; Status = $(if ($script:VendorHoles.Count -gt 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Holes={0}; SHA256={1}' -f $script:VendorHoles.Count, $HoleListHash) }
        [pscustomobject]@{ Control = 'SourceImmutability'; Status = $(if ($changed -eq 0) { 'PASS' } else { 'FAIL' }); Evidence = ('FilesRead={0}; Changed={1}' -f $script:HashBefore.Count, $changed) }
        [pscustomobject]@{ Control = 'FileIssues'; Status = $(if (@($script:FileRegister | Where-Object { $_.Issue -ne '' }).Count -eq 0) { 'PASS' } else { 'REVIEW' }); Evidence = ('Issues={0}' -f @($script:FileRegister | Where-Object { $_.Issue -ne '' }).Count) }
        [pscustomobject]@{ Control = 'ChangeRegister'; Status = $(if ($Major -gt 0) { 'REVIEW' } else { 'PASS' }); Evidence = ('Changes={0}; MAJOR={1}; WARN={2}' -f $script:Changes.Count, $Major, $Warn) }
    )
    $Decision = if ($changed -gt 0) { 'NO_GO' } elseif ($Major -gt 0) { 'CHANGES_DETECTED_REVIEW_REQUIRED' } elseif ($script:Changes.Count -gt 0) { 'CHANGES_DETECTED_MINOR' } else { 'NO_CHANGES' }

    Write-Progress -Id $ProgressId -Activity 'Vendor delta' -Completed
    Write-Host ''
    Write-Host ('CP09A VENDOR EXPORT DELTA v{0}' -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host ('RUN_ID: {0}' -f $RunId)
    Write-Host ('EXPORTS: {0}' -f ($ExportIds -join ' > '))
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' }))
    Write-Host ''
    $DatasetRows | Format-Table BaselineExport, CandidateExport, DatasetId, BaselineFiles, CandidateFiles, Status -AutoSize
    $DatasetSummary | Format-Table DatasetId, Table, Holes, Identical, Changed, Added, Removed, BaselineMeters, CandidateMeters, DeltaMeters, MaxSeverity -AutoSize -Wrap
    $ChangeTypeSummary | Format-Table DatasetId, ChangeType, Severity, Count, VendorHoles -AutoSize
    $QaRows | Format-Table Control, Status, Evidence -AutoSize -Wrap

    if ($AuditOnly) {
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
    }
    else {
        $ResolvedOutput = [System.IO.Path]::GetFullPath($OutputRoot)
        $RunOutput = Join-Path $ResolvedOutput $RunId
        if (Test-Path -LiteralPath $RunOutput) { throw "Overwrite is prohibited. Run output already exists: $RunOutput" }
        if ($PSCmdlet.ShouldProcess($RunOutput, 'Create vendor delta records')) {
            New-Item -ItemType Directory -Path $RunOutput -Force | Out-Null
            $script:Changes | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Change_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
            $script:Summaries | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Hole_Delta_Summary.csv') -NoTypeInformation -Encoding utf8NoBOM
            $DatasetSummary | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Dataset_Delta_Summary.csv') -NoTypeInformation -Encoding utf8NoBOM
            $ChangeTypeSummary | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Change_Type_Summary.csv') -NoTypeInformation -Encoding utf8NoBOM
            $script:FileRegister | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Delta_File_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
            $QaRows | Export-Csv -LiteralPath (Join-Path $RunOutput 'Vendor_Delta_QAQC.csv') -NoTypeInformation -Encoding utf8NoBOM

            $LedgerPath = Join-Path $ResolvedOutput $LedgerName
            $IndexPath = Join-Path $ResolvedOutput ([System.IO.Path]::GetFileNameWithoutExtension($LedgerName) + '_Index.csv')
            $fpMaterial = (@($script:FileRegister | Sort-Object Export, DatasetId, File | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Export, $_.DatasetId, $_.File, $_.SHA256 }) -join '#') + ('#tol={0}|{1}' -f $ValueTolerance, $MeterTolerance)
            $fpAlg = [System.Security.Cryptography.SHA256]::Create()
            try { $Fingerprint = [System.Convert]::ToHexString($fpAlg.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($fpMaterial))) } finally { $fpAlg.Dispose() }
            $AlreadyLogged = @()
            if (Test-Path -LiteralPath $IndexPath -PathType Leaf) { $AlreadyLogged = @(Import-Csv -LiteralPath $IndexPath | Where-Object { $_.Fingerprint -eq $Fingerprint }) }
            $LedgerStatus = 'APPENDED'
            if ($AlreadyLogged.Count -gt 0) {
                $LedgerStatus = ('SKIPPED_DUPLICATE_OF_{0}' -f $AlreadyLogged[0].RunId)
                Write-Host ('LEDGER: same exports and same file hashes already logged by {0}. Ledger not modified.' -f $AlreadyLogged[0].RunId) -ForegroundColor Yellow
            }
            elseif ($script:Changes.Count -gt 0) {
                if (Test-Path -LiteralPath $LedgerPath -PathType Leaf) {
                    $existingHeader = (Get-Content -LiteralPath $LedgerPath -TotalCount 1)
                    $expectedHeader = (@($script:Changes[0].PSObject.Properties.Name | ForEach-Object { ([string][char]34) + $_ + ([string][char]34) }) -join ',')
                    if ($existingHeader -ne $expectedHeader) { throw "Ledger header mismatch. Ledger not modified: $LedgerPath" }
                    $script:Changes | Export-Csv -LiteralPath $LedgerPath -NoTypeInformation -Encoding utf8NoBOM -Append
                }
                else { $script:Changes | Export-Csv -LiteralPath $LedgerPath -NoTypeInformation -Encoding utf8NoBOM }
            }
            if ($AlreadyLogged.Count -eq 0) {
                @([pscustomobject]@{ Fingerprint = $Fingerprint; RunId = $RunId; Exports = ($ExportIds -join '>'); Changes = $script:Changes.Count; LoggedAt = (Get-Date).ToString('s') }) | Export-Csv -LiteralPath $IndexPath -NoTypeInformation -Encoding utf8NoBOM -Append
            }

            $topHoles = @($script:Summaries | Where-Object { $_.Status -in @('CHANGED', 'ADDED', 'REMOVED') } | Sort-Object @{ Expression = { if ($_.DeltaMeters -eq '') { 1000000.0 } else { [math]::Abs([double]::Parse($_.DeltaMeters, $Invariant)) } }; Descending = $true }, @{ Expression = 'ChangeCount'; Descending = $true } | Select-Object -First 100)
            $html = [System.Text.StringBuilder]::new()
            [void]$html.Append('<!DOCTYPE html><html><head><meta charset="utf-8"><title>Vendor Export Delta</title><style>body{font-family:Arial;font-size:12px;background:#0A1A20;color:#E6F2F2;margin:24px}h1{color:#28CFC5}h2{color:#12a39d;border-bottom:1px solid #1E4D55}table{border-collapse:collapse;margin-bottom:18px}th{background:#0A5C5F;color:#fff;padding:4px 8px}td{border:1px solid #1E4D55;padding:3px 8px}tr.MAJOR td{background:#4a1d27}tr.WARN td{background:#4a3d10}.k{display:inline-block;background:#0F262D;border:1px solid #1E4D55;padding:10px 16px;margin:0 10px 10px 0}.k b{font-size:18px;color:#28CFC5}</style></head><body>')
            [void]$html.Append('<h1>CP09A Vendor Export Delta</h1><p>Classification: Internal | Run ID: ' + $RunId + ' | Exports: ' + [System.Net.WebUtility]::HtmlEncode(($ExportIds -join ' > ')) + ' | Decision: <b>' + $Decision + '</b></p>')
            [void]$html.Append('<div class="k">Changes<br><b>' + $script:Changes.Count + '</b></div><div class="k">MAJOR<br><b>' + $Major + '</b></div><div class="k">WARN<br><b>' + $Warn + '</b></div><div class="k">Vendor holes<br><b>' + $script:VendorHoles.Count + '</b></div>')
            [void]$html.Append('<h2>Dataset summary (metres and holes)</h2>' + (ConvertTo-HtmlTable -Rows $DatasetSummary -Columns @('BaselineExport', 'CandidateExport', 'DatasetId', 'Table', 'Holes', 'Identical', 'Changed', 'Added', 'Removed', 'BaselineMeters', 'CandidateMeters', 'DeltaMeters', 'MaxSeverity')))
            [void]$html.Append('<h2>Change types</h2>' + (ConvertTo-HtmlTable -Rows $ChangeTypeSummary -Columns @('DatasetId', 'ChangeType', 'Severity', 'Count', 'VendorHoles')))
            [void]$html.Append('<h2>Top changed holes</h2>' + (ConvertTo-HtmlTable -Rows $topHoles -Columns @('DatasetId', 'Table', 'HoleId', 'InVendorList', 'Status', 'BaselineFrom', 'BaselineTo', 'CandidateFrom', 'CandidateTo', 'BaselineMeters', 'CandidateMeters', 'DeltaMeters', 'ChangeCount', 'MaxSeverity')))
            [void]$html.Append('<h2>Review rule</h2><p>Every MAJOR change must be reviewed in Vendor_Change_Register.csv (ReviewStatus OPEN to ACCEPTED or REJECTED) before compositing or comparing the candidate export. Every composite and analysis must cite the ExportId and this Run ID.</p></body></html>')
            $html.ToString() | Set-Content -LiteralPath (Join-Path $RunOutput 'Vendor_Delta_Report.html') -Encoding utf8NoBOM

            ([ordered]@{ Classification = $Classification; ScriptName = $ScriptName; ScriptVersion = $ScriptVersion; RunId = $RunId; ExportsRoot = $ResolvedRoot; Exports = $ExportIds; HoleListSHA256 = $HoleListHash; ValueTolerance = $ValueTolerance; MeterTolerance = $MeterTolerance; Changes = $script:Changes.Count; Major = $Major; Warn = $Warn; LedgerPath = $LedgerPath; LedgerStatus = $LedgerStatus; Fingerprint = $Fingerprint; SourceModified = ($changed -gt 0); Decision = $Decision; ElapsedSeconds = [math]::Round(((Get-Date) - $Started).TotalSeconds, 1); CreatedAt = (Get-Date).ToString('o') } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $RunOutput 'Run_Register.json') -Encoding utf8NoBOM
            (@('# CP09A Vendor Export Delta', '', ('Classification: {0}' -f $Classification), ('Run ID: {0}' -f $RunId), ('Decision: {0}' -f $Decision), '', 'Open Vendor_Delta_Report.html for the visual summary.', 'Review MAJOR rows in Vendor_Change_Register.csv before compositing.', ('Cumulative append-only ledger: {0}' -f $LedgerPath)) -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'README.md') -Encoding utf8NoBOM
            $items = @(Get-ChildItem -LiteralPath $RunOutput -File | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Length = [long]$_.Length; SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
            ([ordered]@{ Classification = $Classification; RunId = $RunId; FileCount = $items.Count; Files = $items } | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $RunOutput 'Manifest.json') -Encoding utf8NoBOM
            (@(Get-ChildItem -LiteralPath $RunOutput -File | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object Name | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, $_.Name }) -join [Environment]::NewLine) | Set-Content -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Encoding ascii
            Write-Host ('REPORT: {0}' -f (Join-Path $RunOutput 'Vendor_Delta_Report.html')) -ForegroundColor Green
            Write-Host ('LEDGER: {0}' -f $LedgerPath) -ForegroundColor Green
        }
    }
    Write-Host ('ELAPSED SECONDS: {0}' -f [math]::Round(((Get-Date) - $Started).TotalSeconds, 1))
    $color = if ($Decision -eq 'NO_CHANGES') { 'Green' } elseif ($Decision -eq 'NO_GO') { 'Red' } else { 'Yellow' }
    Write-Host ('DECISION: {0}' -f $Decision) -ForegroundColor $color
    if ($Decision -eq 'NO_GO') { $ExitCode = 1 }
}
catch {
    Write-Host ''
    Write-Host ('FAIL: {0}' -f $_.Exception.Message) -ForegroundColor Red
    $ExitCode = 1
}
finally {
    Write-Progress -Id $ProgressId -Activity 'Vendor delta' -Completed
}

exit $ExitCode
