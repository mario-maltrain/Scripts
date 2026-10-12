#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$InputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\01_LAS_to_CSV\CP09A-LAS-TO-CSV-20261001-140002\CSV',
    [string]$OutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\02_Composites_PowerShell',
    [double]$MinimumCoverage = 0.80,
    [double[]]$Scales = @(1.0, 3.0, 6.0)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.0.2'
$Classification = 'Internal'
$DataType = 'LEP'
$RunId = 'CP09A-PS-LEP-COMPOSITE-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$StagingRoot = Join-Path $OutputRoot ('.staging-' + $RunId)
$FinalRoot = Join-Path $OutputRoot $RunId
$DataRoot = Join-Path $StagingRoot 'Composites'
$RecordRoot = Join-Path $StagingRoot '_Composite_Record'
$LogPath = Join-Path $RecordRoot 'Execution.log'
$FailureRegisterPath = Join-Path $RecordRoot 'Consolidated_Failure_Real_Test_Register.csv'
$Writers = @{}
$NullSentinels = @(-999.25, -999.0, -9999.0)
$RequiredLepColumns = @('DEPT','GT20','GT30','GT50')
$OutputColumns = @(
    'DATA_TYPE','HOLE_ID','FROM','TO','CENTRE','SCALE_M','N_SAMPLES','EXPECTED_SAMPLES',
    'SAMPLE_COVERAGE_RATIO','PARTIAL_COMPOSITE','OBSERVED_MIN_DEPTH','OBSERVED_MAX_DEPTH',
    'GT20_MEAN','GT30_MEAN','GT50_MEAN','STATUS','SOURCE_FILE'
)

function Write-RunLog {
    param([string]$Message,[ValidateSet('INFO','PASS','WARN','ERROR')][string]$Level='INFO')
    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $Line
    if (Test-Path -LiteralPath $RecordRoot -PathType Container) {
        Add-Content -LiteralPath $LogPath -Value $Line -Encoding utf8
    }
}

function Add-TestRecord {
    param([string]$Artifact,[string]$Stage,[string]$Status,[string]$Evidence,[string]$Disposition)
    [pscustomobject]@{
        Timestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Artifact=$Artifact; Stage=$Stage
        Status=$Status; Evidence=$Evidence; Disposition=$Disposition
        Classification=$Classification; RunId=$RunId
    } | Export-Csv -LiteralPath $FailureRegisterPath -Append -NoTypeInformation -Encoding utf8
}

function Convert-ToDouble {
    param([AllowNull()][string]$Value)
    $Number=0.0
    if ([double]::TryParse($Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$Number)) { return $Number }
    return [double]::NaN
}

function Convert-ToCsvField {
    param($Value)
    if ($null -eq $Value) { return '' }
    $Text=[string]$Value
    if ($Text -match '[,\r\n"]') { return '"' + $Text.Replace('"','""') + '"' }
    return $Text
}

function Format-Number {
    param([double]$Number)
    if ([double]::IsNaN($Number)) { return '' }
    return $Number.ToString('G17',[Globalization.CultureInfo]::InvariantCulture)
}

function Get-Delimiter {
    param([string]$Header)
    $Semicolons=0; $Commas=0
    foreach ($Character in $Header.ToCharArray()) {
        if ($Character -eq ';') { $Semicolons++ }
        elseif ($Character -eq ',') { $Commas++ }
    }
    if ($Semicolons -gt $Commas) { return ';' }
    return ','
}

function Get-MedianPositiveStep {
    param([object[]]$Rows)
    $Differences=[System.Collections.Generic.List[double]]::new()
    for ($Index=1; $Index -lt $Rows.Length; $Index++) {
        $Difference=[double]$Rows[$Index].Depth-[double]$Rows[$Index-1].Depth
        if ($Difference -gt 0) { $Differences.Add($Difference) }
    }
    if ($Differences.Count -eq 0) { throw 'No positive depth differences were found.' }
    $Sorted=[double[]]@($Differences | Sort-Object)
    return $Sorted[[int][math]::Floor($Sorted.Length/2)]
}

function Test-IsNullValue {
    param([double]$Value)
    if ([double]::IsNaN($Value)) { return $true }
    foreach ($Sentinel in $NullSentinels) {
        if ([math]::Abs($Value-$Sentinel) -le 1e-9) { return $true }
    }
    return $false
}

function Get-Mean {
    param([object[]]$Rows,[int]$ValueIndex)
    $Sum=0.0; $Count=0
    foreach ($Row in $Rows) {
        $Value=[double]$Row.Values[$ValueIndex]
        if (-not (Test-IsNullValue $Value)) { $Sum+=$Value; $Count++ }
    }
    if ($Count -eq 0) { return [double]::NaN }
    return $Sum/$Count
}

function Get-LepIdentity {
    param([IO.FileInfo]$File,[string[]]$Columns)
    $Normalized=[string[]]@($Columns | ForEach-Object { $_.Trim().ToUpperInvariant() })
    $SchemaMatch=($Normalized.Length -eq 4)
    foreach ($Required in $RequiredLepColumns) {
        if ($Required -notin $Normalized) { $SchemaMatch=$false }
    }
    $NameMatch=($File.BaseName -match '(?i)SPENCELEP|(^|[_-])LEP([_-]|$)')
    if ($SchemaMatch -and $NameMatch) {
        return [pscustomobject]@{DataType='LEP';Identification='NAME_AND_SCHEMA';Schema=($Normalized -join '|')}
    }
    if ($SchemaMatch -and -not $NameMatch) {
        throw "Ambiguous LEP file: schema matched but filename lacks LEP identifier: $($File.FullName)"
    }
    if ($NameMatch -and -not $SchemaMatch) {
        throw "Invalid LEP file: filename indicates LEP but schema is incompatible: $($File.FullName); schema=$($Normalized -join '|')"
    }
    throw "Unsupported data type. File is not identified as LEP and will not be composited: $($File.FullName); schema=$($Normalized -join '|')"
}

function Test-CompositeFile {
    param([string]$Path,[double]$ExpectedScale)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "QA/QC missing output: $Path" }
    $Reader=[IO.File]::OpenText($Path)
    $RowsChecked=[int64]0
    try {
        $Header=$Reader.ReadLine()
        $HeaderColumns=[string[]]$Header.Split(',')
        if ($HeaderColumns.Length -ne $OutputColumns.Length) { throw "QA/QC header count mismatch: $Path" }
        for ($Index=0; $Index -lt $OutputColumns.Length; $Index++) {
            if ($HeaderColumns[$Index] -ne $OutputColumns[$Index]) { throw "QA/QC header mismatch at index $Index in $Path" }
        }
        while ($null -ne ($Line=$Reader.ReadLine())) {
            if ([string]::IsNullOrWhiteSpace($Line)) { continue }
            $Parts=[string[]]$Line.Split(',')
            if ($Parts.Length -ne $OutputColumns.Length) { throw "QA/QC row column-count mismatch in $Path at data row $($RowsChecked+1)" }
            if ($Parts[0] -ne 'LEP') { throw "QA/QC DATA_TYPE mismatch in $Path" }
            $From=Convert-ToDouble $Parts[2]; $To=Convert-ToDouble $Parts[3]; $Centre=Convert-ToDouble $Parts[4]; $Scale=Convert-ToDouble $Parts[5]
            $Samples=[int]$Parts[6]; $Expected=[int]$Parts[7]; $Coverage=Convert-ToDouble $Parts[8]
            if ([math]::Abs($Scale-$ExpectedScale) -gt 1e-9) { throw "QA/QC scale mismatch in $Path" }
            if ([math]::Abs(($To-$From)-$Scale) -gt 1e-8) { throw "QA/QC interval length mismatch in $Path" }
            if ([math]::Abs($Centre-(($From+$To)/2.0)) -gt 1e-8) { throw "QA/QC centre mismatch in $Path" }
            $ExpectedCoverage=[math]::Min(1.0,$Samples/[double]$Expected)
            if ([math]::Abs($Coverage-$ExpectedCoverage) -gt 1e-10) { throw "QA/QC coverage mismatch in $Path" }
            $ExpectedStatus=if ($Samples -eq 0) {'NO_INFORMATION'} elseif ($Coverage -lt $MinimumCoverage) {'PARTIAL'} else {'PASS'}
            $ExpectedPartial=if ($ExpectedStatus -eq 'PARTIAL') {'True'} else {'False'}
            if ($Parts[15] -ne $ExpectedStatus) { throw "QA/QC STATUS mismatch in $Path" }
            if ($Parts[9] -ne $ExpectedPartial) { throw "QA/QC PARTIAL_COMPOSITE mismatch in $Path" }
            $RowsChecked++
        }
    }
    finally { $Reader.Dispose() }
    if ($RowsChecked -eq 0) { throw "QA/QC output has no data rows: $Path" }
    return $RowsChecked
}

try {
    if (-not (Test-Path -LiteralPath $InputRoot -PathType Container)) { throw "Input folder not found: $InputRoot" }
    if ($MinimumCoverage -le 0 -or $MinimumCoverage -gt 1) { throw 'MinimumCoverage must be within (0,1].' }
    $ScaleCheck=[double[]]@($Scales | Sort-Object -Unique)
    if ($ScaleCheck.Length -ne 3 -or 1.0 -notin $ScaleCheck -or 3.0 -notin $ScaleCheck -or 6.0 -notin $ScaleCheck) { throw 'Scales must be exactly 1, 3 and 6 metres.' }
    if (Test-Path -LiteralPath $FinalRoot) { throw "Final output already exists: $FinalRoot" }
    if (Test-Path -LiteralPath $StagingRoot) { throw "Staging output already exists: $StagingRoot" }

    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null
    Write-RunLog "CP09A PowerShell LEP Compositor | Version $Version | Classification $Classification"
    Write-RunLog "RUN_ID: $RunId"
    Write-RunLog "INPUT: $InputRoot"
    Write-RunLog "OUTPUT: $FinalRoot"

    Add-TestRecord 'Run_CP09A_PowerShell_Compositor_v1_0_1.ps1' 'LEP_COMPOSITING' 'QAQC_FAIL' 'Column alignment, coverage and LEP argmax conceptual defect' 'DO_NOT_USE'

    $Files=[IO.FileInfo[]]@(Get-ChildItem -LiteralPath $InputRoot -Filter '*.csv' -File -Recurse | Sort-Object FullName)
    if ($Files.Length -eq 0) { throw 'No CSV files found.' }

    $OutputPaths=@{}; $ScaleStats=@{}
    foreach ($Scale in $Scales) {
        $Label='{0}m' -f [int]$Scale
        $Folder=Join-Path $DataRoot $Label
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        $Path=Join-Path $Folder ("LEP_Composite_$Label.csv")
        if (Test-Path -LiteralPath $Path) { throw "No-overwrite safety block: $Path" }
        $OutputPaths[$Scale]=$Path
        $ScaleStats[$Scale]=[ordered]@{Rows=[int64]0;Pass=[int64]0;Partial=[int64]0;NoInformation=[int64]0;Holes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)}
    }

    $Register=[Collections.Generic.List[object]]::new()
    $TotalInputRows=[int64]0

    for ($FileIndex=0; $FileIndex -lt $Files.Length; $FileIndex++) {
        $File=$Files[$FileIndex]
        $Percent=5+[int](85*(($FileIndex+1)/[double]$Files.Length))
        Write-Progress -Id 1 -Activity 'CP09A LEP composites 1 m, 3 m and 6 m' -Status 'Processing identified LEP files' -CurrentOperation ('File {0} of {1}: {2}' -f ($FileIndex+1),$Files.Length,$File.Name) -PercentComplete $Percent

        $Header=[IO.File]::ReadLines($File.FullName) | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($Header)) { throw "Empty header: $($File.FullName)" }
        $Delimiter=Get-Delimiter $Header
        $Columns=[string[]]@($Header.Split($Delimiter) | ForEach-Object { $_.Trim('"',' ') })
        $Identity=Get-LepIdentity $File $Columns
        $Upper=[string[]]@($Columns | ForEach-Object { $_.ToUpperInvariant() })
        $DeptIndex=[array]::IndexOf($Upper,'DEPT'); $Gt20Index=[array]::IndexOf($Upper,'GT20'); $Gt30Index=[array]::IndexOf($Upper,'GT30'); $Gt50Index=[array]::IndexOf($Upper,'GT50')

        $HoleId=$File.BaseName -replace '_SPENCE.*$',''
        $RowsList=[Collections.Generic.List[object]]::new()
        $Reader=[IO.File]::OpenText($File.FullName)
        try {
            $null=$Reader.ReadLine()
            while ($null -ne ($Line=$Reader.ReadLine())) {
                if ([string]::IsNullOrWhiteSpace($Line)) { continue }
                $Parts=[string[]]$Line.Split($Delimiter)
                if ($Parts.Length -ne $Columns.Length) { throw "Input column-count mismatch: $($File.FullName)" }
                $Depth=Convert-ToDouble $Parts[$DeptIndex].Trim('"')
                if ([double]::IsNaN($Depth)) { continue }
                $Values=[double[]]::new(3)
                $Values[0]=Convert-ToDouble $Parts[$Gt20Index].Trim('"')
                $Values[1]=Convert-ToDouble $Parts[$Gt30Index].Trim('"')
                $Values[2]=Convert-ToDouble $Parts[$Gt50Index].Trim('"')
                $RowsList.Add([pscustomobject]@{Depth=$Depth;Values=$Values})
                $TotalInputRows++
            }
        }
        finally { $Reader.Dispose() }

        $Rows=[object[]]@($RowsList | Sort-Object Depth)
        if ($Rows.Length -lt 2) { throw "Fewer than two valid LEP rows: $($File.FullName)" }
        $Step=Get-MedianPositiveStep $Rows
        if ($Step -le 0) { throw "Invalid depth step: $($File.FullName)" }

        foreach ($Scale in $Scales) {
            if (-not $Writers.ContainsKey($Scale)) {
                $Writer=[IO.StreamWriter]::new($OutputPaths[$Scale],$false,[Text.UTF8Encoding]::new($true))
                $Writer.WriteLine($OutputColumns -join ',')
                $Writers[$Scale]=$Writer
            }
            $Groups=@{}
            foreach ($Row in $Rows) {
                $From=[math]::Floor(([double]$Row.Depth+1e-10)/$Scale)*$Scale
                $Key=$From.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
                if (-not $Groups.ContainsKey($Key)) { $Groups[$Key]=[Collections.Generic.List[object]]::new() }
                $Groups[$Key].Add($Row)
            }
            $Bins=[double[]]@($Groups.Keys | ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::InvariantCulture)} | Sort-Object)
            foreach ($From in $Bins) {
                $Key=$From.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
                $GroupRows=[object[]]@($Groups[$Key])
                $To=$From+$Scale; $Centre=($From+$To)/2.0
                $Samples=$GroupRows.Length
                $Expected=[math]::Max(1,[int][math]::Round($Scale/$Step))
                $Coverage=[math]::Min(1.0,$Samples/[double]$Expected)
                $Partial=($Samples -gt 0 -and $Coverage -lt $MinimumCoverage)
                $Status=if ($Samples -eq 0) {'NO_INFORMATION'} elseif ($Partial) {'PARTIAL'} else {'PASS'}
                $Mean20=Get-Mean $GroupRows 0; $Mean30=Get-Mean $GroupRows 1; $Mean50=Get-Mean $GroupRows 2
                $ObservedMin=[double]($GroupRows | Measure-Object Depth -Minimum).Minimum
                $ObservedMax=[double]($GroupRows | Measure-Object Depth -Maximum).Maximum

                $Fields=[string[]]@(
                    'LEP',(Convert-ToCsvField $HoleId),(Format-Number $From),(Format-Number $To),(Format-Number $Centre),(Format-Number $Scale),
                    [string]$Samples,[string]$Expected,(Format-Number $Coverage),[string]$Partial,(Format-Number $ObservedMin),(Format-Number $ObservedMax),
                    (Format-Number $Mean20),(Format-Number $Mean30),(Format-Number $Mean50),$Status,(Convert-ToCsvField $File.FullName)
                )
                if ($Fields.Length -ne $OutputColumns.Length) { throw "Pre-write column-count mismatch: $($File.FullName) scale=$Scale" }
                $Writers[$Scale].WriteLine($Fields -join ',')
                $ScaleStats[$Scale].Rows++
                if ($Status -eq 'PASS') {$ScaleStats[$Scale].Pass++} elseif ($Status -eq 'PARTIAL') {$ScaleStats[$Scale].Partial++} else {$ScaleStats[$Scale].NoInformation++}
                $null=$ScaleStats[$Scale].Holes.Add($HoleId)
            }
            $Register.Add([pscustomobject]@{SourceFile=$File.FullName;HoleId=$HoleId;DataType=$Identity.DataType;Identification=$Identity.Identification;ScaleM=$Scale;InputRows=$Rows.Length;OutputRows=$Groups.Count;Delimiter=$Delimiter;SourceStepM=$Step;Schema=$Identity.Schema;Status='PASS'})
        }
    }

    foreach ($Writer in $Writers.Values) {$Writer.Flush();$Writer.Dispose()}
    $Writers.Clear()
    $Register | Export-Csv -LiteralPath (Join-Path $RecordRoot 'LEP_Composite_Register.csv') -NoTypeInformation -Encoding utf8

    $QaRows=[Collections.Generic.List[object]]::new()
    foreach ($Scale in $Scales) {
        $Checked=Test-CompositeFile $OutputPaths[$Scale] $Scale
        $QaRows.Add([pscustomobject]@{ScaleM=$Scale;RowsChecked=$Checked;ColumnAlignment='PASS';Coverage='PASS';Status='PASS';NoArgmax='PASS';DataType='LEP';Result='PASS'})
    }
    $QaRows | Export-Csv -LiteralPath (Join-Path $RecordRoot 'LEP_QAQC.csv') -NoTypeInformation -Encoding utf8

    $ScaleSummary=foreach ($Scale in $Scales) {
        $Path=$OutputPaths[$Scale]
        [pscustomobject]@{ScaleM=$Scale;Rows=$ScaleStats[$Scale].Rows;Holes=$ScaleStats[$Scale].Holes.Count;PassRows=$ScaleStats[$Scale].Pass;PartialRows=$ScaleStats[$Scale].Partial;NoInformationRows=$ScaleStats[$Scale].NoInformation;OutputFile=$Path;SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
    }
    $ScaleSummary | Export-Csv -LiteralPath (Join-Path $RecordRoot 'LEP_Scale_Summary.csv') -NoTypeInformation -Encoding utf8

    $Summary=[ordered]@{Classification=$Classification;Version=$Version;RunId=$RunId;DataType='LEP';IdentificationRule='SPENCELEP/LEP filename AND exact DEPT|GT20|GT30|GT50 schema';InputRoot=$InputRoot;InputFiles=$Files.Length;InputRows=$TotalInputRows;ScalesM='1,3,6';MinimumCoverage=$MinimumCoverage;Renormalization='DISABLED';Argmax='DISABLED';OverwriteCount=0;QAQC='PASS';Status='PASS_REAL_TEST'}
    $Summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.json') -Encoding utf8
    $Summary.GetEnumerator() | ForEach-Object {'{0}: {1}' -f $_.Key,$_.Value} | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.txt') -Encoding utf8
    Add-TestRecord 'Run_CP09A_PowerShell_LEP_Compositor_v1_0_2.ps1' 'LEP_COMPOSITING' 'PASS_REAL_TEST' $RunId 'ACTIVE_VALIDATED_BASELINE'

    $Manifest=Get-ChildItem -LiteralPath $StagingRoot -File -Recurse | Where-Object Name -ne 'SHA256SUMS.txt' | ForEach-Object {'{0} *{1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash,$_.FullName.Substring($StagingRoot.Length).TrimStart('\')}
    $Manifest | Set-Content -LiteralPath (Join-Path $RecordRoot 'SHA256SUMS.txt') -Encoding ascii
    Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot
    Write-Progress -Id 1 -Activity 'CP09A LEP composites' -Completed
    Write-Host ''
    Write-Host 'LEP COMPOSITES: PASS_REAL_TEST' -ForegroundColor Green
    Write-Host 'IDENTIFICATION: LEP NAME + SCHEMA' -ForegroundColor Green
    Write-Host 'SCALES: 1 m, 3 m, 6 m' -ForegroundColor Green
    Write-Host 'ARGMAX: DISABLED' -ForegroundColor Green
    Write-Host 'OVERWRITES: 0' -ForegroundColor Green
    Write-Host "OUTPUT: $FinalRoot" -ForegroundColor Cyan
    Write-Host ''
    exit 0
}
catch {
    try {
        foreach ($Writer in $Writers.Values) {$Writer.Dispose()}
        Write-RunLog $_.Exception.Message 'ERROR'
        Add-TestRecord 'Run_CP09A_PowerShell_LEP_Compositor_v1_0_2.ps1' 'LEP_COMPOSITING' 'FAIL' $_.Exception.Message 'BLOCK_AND_REVIEW'
        $_ | Out-String | Set-Content -LiteralPath (Join-Path $RecordRoot 'Fatal_Error.txt') -Encoding utf8
    } catch {}
    Write-Progress -Id 1 -Activity 'CP09A LEP composites' -Completed
    exit 1
}
