#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$RunRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\02_Composites_PowerShell\CP09A-PS-LEP-COMPOSITE-20261001-175040',
    [switch]$ContinueOnNonCriticalError
)
Clear-Host
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Classification = 'BHP Internal'
$Version = '1.0.7'
$Scales = @(1,3,6)
$Metrics = @('GT20_MEAN','GT30_MEAN','GT50_MEAN')
$Columns = @('DATA_TYPE','HOLE_ID','FROM','TO','CENTRE','SCALE_M','N_SAMPLES','EXPECTED_SAMPLES','SAMPLE_COVERAGE_RATIO','PARTIAL_COMPOSITE','OBSERVED_MIN_DEPTH','OBSERVED_MAX_DEPTH','GT20_MEAN','GT30_MEAN','GT50_MEAN','STATUS','SOURCE_FILE')
$Data = @{}
$Summary = @()
$MetricSummary = @()
$Warnings = @()
$FinalStatus = 'PASS'
function Test-SourcePatterns {
    param([string]$Path)
    $Text = Get-Content -LiteralPath $Path -Raw -Encoding utf8
    $OpenBracket = [regex]::Escape([string][char]91)
    $Quotes = '[' + [string][char]34 + [string][char]39 + ']'
    $Patterns = @(
        (('ret' + 'urn') + $OpenBracket),
        (('Write' + '-Host') + $Quotes),
        (('Write' + '-Warning') + $Quotes),
        (('Write' + '-Error') + $Quotes),
        (('thr' + 'ow') + $Quotes)
    )
    $Failures = @($Patterns | Where-Object { $Text -match $_ })
    if ($Failures.Count -gt 0) { throw ('SOURCE PATTERN VALIDATION FAILED: {0}' -f ($Failures -join ', ')) }
}
function Convert-Number {
    param([string]$Value,[string]$Field,[string]$Context)
    $Number = 0.0
    $Parsed = [double]::TryParse($Value.Trim().Replace(',','.'),[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$Number)
    if (-not $Parsed) { throw ('Invalid numeric value. Field={0}; Context={1}; Value={2}' -f $Field,$Context,$Value) }
    return $Number
}
function Get-Stats {
    param([object[]]$Rows,[string]$Column,[string]$Context,[switch]$ContinueAfterInvalid)
    $Values = [Collections.Generic.List[double]]::new()
    $Invalid = 0
    foreach ($Row in $Rows) {
        try {
            $Property = $Row.PSObject.Properties[$Column]
            if ($null -eq $Property -or [string]::IsNullOrWhiteSpace([string]$Property.Value)) { $Invalid++; continue }
            $Values.Add((Convert-Number -Value ([string]$Property.Value) -Field $Column -Context $Context))
        } catch {
            $Invalid++
            if (-not $ContinueAfterInvalid) { throw }
        }
    }
    if ($Values.Count -eq 0) { throw ('No valid numeric values. Column={0}; Context={1}' -f $Column,$Context) }
    $Measure = $Values | Measure-Object -Average -Minimum -Maximum
    return [pscustomobject]@{
        Valid = $Values.Count
        Invalid = $Invalid
        Mean = [math]::Round([double]$Measure.Average,6)
        Min = [math]::Round([double]$Measure.Minimum,6)
        Max = [math]::Round([double]$Measure.Maximum,6)
    }
}
try {
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { Test-SourcePatterns -Path $PSCommandPath }
    if (-not (Test-Path -LiteralPath $RunRoot -PathType Container)) { throw "RunRoot not found: $RunRoot" }
    $CompositeRoot = Join-Path $RunRoot 'Composites'
    if (-not (Test-Path -LiteralPath $CompositeRoot -PathType Container)) { throw "Composites folder not found: $CompositeRoot" }
    $ObservedFolders = @(Get-ChildItem -LiteralPath $CompositeRoot -Directory | Where-Object { $_.Name -match '^(1|3|6)m$' } | ForEach-Object { [int]($_.Name -replace 'm$','') } | Sort-Object)
    $MissingFolders = @($Scales | Where-Object { $_ -notin $ObservedFolders })
    if ($MissingFolders.Count -gt 0) { throw ('INVALID SCALE FOLDERS. Missing={0}' -f ($MissingFolders -join ',')) }
    foreach ($Scale in $Scales) {
        $Path = Join-Path $RunRoot "Composites\${Scale}m\LEP_Composite_${Scale}m.csv"
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "INVALID SCALE FILE. Expected ${Scale}m: $Path" }
        $Rows = @(Import-Csv -LiteralPath $Path)
        if ($Rows.Count -eq 0) { throw "INVALID SCALE FILE. Empty ${Scale}m: $Path" }
        $Headers = @($Rows[0].PSObject.Properties.Name)
        $Missing = @($Columns | Where-Object { $_ -notin $Headers })
        $Unexpected = @($Headers | Where-Object { $_ -notin $Columns })
        $OrderOk = (($Headers -join '|') -ceq ($Columns -join '|'))
        if ($Missing.Count -gt 0 -or $Unexpected.Count -gt 0 -or -not $OrderOk) { throw ('INVALID COLUMN SCHEMA at {0}m. Missing={1}; Unexpected={2}; OrderOk={3}' -f $Scale,($Missing -join ','),($Unexpected -join ','),$OrderOk) }
        $ObservedScales = @($Rows.SCALE_M | ForEach-Object { [int]$_ } | Sort-Object -Unique)
        if ($ObservedScales.Count -ne 1 -or $ObservedScales[0] -ne $Scale) { throw ('INVALID SCALE CONTENT. Expected={0}; Observed={1}' -f $Scale,($ObservedScales -join ',')) }
        $BadType = @($Rows | Where-Object { $_.DATA_TYPE -ne 'LEP' })
        $BadHole = @($Rows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.HOLE_ID) })
        $BadStatus = @($Rows | Where-Object { $_.STATUS -notin @('PASS','PARTIAL') })
        if ($BadType.Count -gt 0) { throw "INVALID DATA_TYPE at ${Scale}m. Rows=$($BadType.Count)" }
        if ($BadHole.Count -gt 0) { throw "MISSING HOLE_ID at ${Scale}m. Rows=$($BadHole.Count)" }
        if ($BadStatus.Count -gt 0) { throw "INVALID STATUS at ${Scale}m. Rows=$($BadStatus.Count)" }
        $Data[$Scale] = $Rows
    }
    foreach ($Scale in $Scales) {
        $Rows = @($Data[$Scale])
        $Holes = @($Rows.HOLE_ID | Sort-Object -Unique)
        $Pass = @($Rows | Where-Object { $_.STATUS -eq 'PASS' })
        $Partial = @($Rows | Where-Object { $_.STATUS -eq 'PARTIAL' })
        $Coverage = Get-Stats -Rows $Rows -Column 'SAMPLE_COVERAGE_RATIO' -Context "${Scale}m" -ContinueAfterInvalid:$ContinueOnNonCriticalError
        $Summary += [pscustomobject]@{ScaleM=$Scale;Rows=$Rows.Count;Holes=$Holes.Count;PassRows=$Pass.Count;PartialRows=$Partial.Count;PassPct=[math]::Round(100.0*$Pass.Count/$Rows.Count,6);PartialPct=[math]::Round(100.0*$Partial.Count/$Rows.Count,6);CoverageMean=$Coverage.Mean;CoverageMin=$Coverage.Min;CoverageMax=$Coverage.Max}
        foreach ($Metric in $Metrics) {
            try {
                $Stats = Get-Stats -Rows $Rows -Column $Metric -Context "${Scale}m" -ContinueAfterInvalid:$ContinueOnNonCriticalError
                $Status = if ($Stats.Invalid -eq 0) { 'PASS' } else { 'ANALYTICAL_WARN' }
                $MetricSummary += [pscustomobject]@{ScaleM=$Scale;Metric=$Metric;Valid=$Stats.Valid;Invalid=$Stats.Invalid;Mean=$Stats.Mean;Min=$Stats.Min;Max=$Stats.Max;Status=$Status}
                if ($Stats.Invalid -gt 0) { $Warnings += [pscustomobject]@{ScaleM=$Scale;Metric=$Metric;Count=$Stats.Invalid;Warning='Null or invalid values excluded.'} }
            } catch {
                if (-not $ContinueOnNonCriticalError) { throw }
                $Warnings += [pscustomobject]@{ScaleM=$Scale;Metric=$Metric;Count=$Rows.Count;Warning=$_.Exception.Message}
            }
        }
    }
    $H1 = @($Data[1].HOLE_ID | Sort-Object -Unique)
    $H3 = @($Data[3].HOLE_ID | Sort-Object -Unique)
    $H6 = @($Data[6].HOLE_ID | Sort-Object -Unique)
    $M3 = @($H1 | Where-Object { $_ -notin $H3 })
    $M6 = @($H1 | Where-Object { $_ -notin $H6 })
    $U3 = @($H3 | Where-Object { $_ -notin $H1 })
    $U6 = @($H6 | Where-Object { $_ -notin $H1 })
    if ($M3.Count -gt 0 -or $M6.Count -gt 0 -or $U3.Count -gt 0 -or $U6.Count -gt 0) { throw ('HOLE_ID inconsistency. M3={0}; M6={1}; U3={2}; U6={3}' -f $M3.Count,$M6.Count,$U3.Count,$U6.Count) }
    if (@($Warnings).Count -gt 0) { $FinalStatus = 'PASS_WITH_ANALYTICAL_WARNINGS' }
    Write-Host "$Classification | CP09A LEP ANALYTICAL QAQC v$Version" -ForegroundColor Cyan
    $Summary | Format-Table -AutoSize
    $MetricSummary | Format-Table -AutoSize
    Write-Host ('HOLE_ID CONSISTENCY | Missing3m={0} | Missing6m={1} | Unexpected3m={2} | Unexpected6m={3}' -f $M3.Count,$M6.Count,$U3.Count,$U6.Count)
    Write-Host "ANALYTICAL WARNINGS: $(@($Warnings).Count)"
    if (@($Warnings).Count -gt 0) { $Warnings | Format-Table -AutoSize }
    Write-Host "QAQC STATUS: $FinalStatus" -ForegroundColor $(if ($FinalStatus -eq 'PASS') { 'Green' } else { 'Yellow' })
    Write-Host 'NO FILES OR FOLDERS CREATED' -ForegroundColor Green
} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor Red
    exit 1
}
