# Classification: Internal
# Script: Initialize_CP09A_Composite_Run_v1_0_4.ps1
# Purpose: Deterministically resolve the four governed CP09A sources and create run-scoped records for 12 composite jobs.
# Safety: Audit mode is the default. Sources are read-only. No source delete, move, rename, overwrite, or modification is allowed.

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SourceRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\01_Processed_Exports\GEOAI-EXPORTS-20261002-225517',

    [Parameter()]
    [AllowEmptyString()]
    [string]$OutputRoot = '',

    [Parameter()]
    [ValidatePattern('^CP09A-COMPOSITE-\d{8}-\d{6}$')]
    [string]$RunId = ('CP09A-COMPOSITE-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),

    [Parameter()]
    [switch]$CommitRecords,

    [Parameter()]
    [ValidatePattern('^Exports_\d{4}-\d{2}-\d{2}$')]
    [string]$ActiveExportId = 'Exports_2026-09-28',

    [Parameter()]
    [AllowEmptyString()]
    [string]$AutoLithSourceFolder = '',

    [Parameter()]
    [ValidateRange(1, 100000)]
    [int]$AutoLithExpectedFiles = 150
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$ScriptName = 'Initialize_CP09A_Composite_Run_v1_0_4.ps1'
$ScriptVersion = '1.0.4'
$Classification = 'Internal'
$PipelineRunId = 'GEOAI-EXPORTS-20261002-225517'
$ProgressId = 1302
$Scales = @('1m', '3m', '6m')
$ExpectedSourceIds = @('SPENCE_MASTER', 'LAS_LEP', 'LEP_CONFIDENCE', 'AUTOLITH_CONFIDENCE')
$CreatedPaths = [System.Collections.Generic.List[string]]::new()

function Show-RunProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateRange(0, 100)][int]$Percent,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Status
    )
    Write-Progress -Id $ProgressId -Activity 'CP09A composite initialization v1.0.4' -Status $Status -PercentComplete $Percent
}

function Resolve-ExistingDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LiteralPath)
    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Container)) {
        throw "Required directory does not exist: $LiteralPath"
    }
    return (Get-Item -LiteralPath $LiteralPath -ErrorAction Stop).FullName
}

function Test-SourceFileExcluded {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    if ($File.Extension -ne '.csv') { return $true }
    if ($File.Name -match '__DUP_SHA256_EXACT|__COPY_\d+|\bCopiar\b') { return $true }
    if ($File.Name -match '(?i)(^|_)composite(_|\b)|_[136]m\.csv$') { return $true }
    return $false
}

function Get-EligibleFiles {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Folder)
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    return @(
        Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction Stop |
            Where-Object { ($_.Extension -eq '.csv') -and (-not (Test-SourceFileExcluded -File $_)) } |
            Sort-Object FullName -Unique
    )
}

function Get-FileSetHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.FileInfo[]]$Files)
    if (@($Files).Count -eq 0) { return '' }
    $hashMaterial = @(
        $Files | Sort-Object FullName | ForEach-Object {
            $fileHash = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
            '{0}|{1}' -f $_.FullName, $fileHash.Hash
        }
    ) -join "`n"
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($hashMaterial)
        return [System.Convert]::ToHexString($algorithm.ComputeHash($bytes))
    }
    finally {
        $algorithm.Dispose()
    }
}

function New-SourceRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('SPENCE_MASTER', 'LAS_LEP', 'LEP_CONFIDENCE', 'AUTOLITH_CONFIDENCE')][string]$SourceId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RelativeFolder,
        [Parameter(Mandatory)][ValidateSet('FILE', 'FOLDER_SET')][string]$SelectionMode,
        [Parameter(Mandatory)][ValidateRange(1, 2147483647)][int]$MinimumFiles,
        [Parameter()][ValidateRange(0, 2147483647)][int]$MaximumFiles = 0,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ExportRoot,
        [Parameter()][AllowEmptyString()][string]$SourceFolderOverride = '',
        [Parameter()][AllowEmptyString()][string]$ExportIdOverride = ''
    )

    $sourceFolder = if ([string]::IsNullOrWhiteSpace($SourceFolderOverride)) { Join-Path $ExportRoot $RelativeFolder } else { $SourceFolderOverride }
    $recordExportId = if ([string]::IsNullOrWhiteSpace($ExportIdOverride)) { $ActiveExportId } else { $ExportIdOverride }
    $folderExists = Test-Path -LiteralPath $sourceFolder -PathType Container
    $files = @(if ($folderExists) { Get-EligibleFiles -Folder $sourceFolder })
    $status = 'PASS'
    $reason = 'Deterministic governed source selected.'

    if (-not $folderExists) {
        $status = 'REQUIRES_VALIDATION'
        $reason = "Required governed folder does not exist: $sourceFolder"
    }
    elseif ($files.Count -lt $MinimumFiles) {
        $status = 'REQUIRES_VALIDATION'
        $reason = "Eligible file count $($files.Count) is below required minimum $MinimumFiles."
    }
    elseif ($MaximumFiles -gt 0 -and $files.Count -gt $MaximumFiles) {
        $status = 'REQUIRES_VALIDATION'
        $reason = "Eligible file count $($files.Count) exceeds allowed maximum $MaximumFiles."
    }
    elseif ($SelectionMode -eq 'FILE' -and $files.Count -ne 1) {
        $status = 'REQUIRES_VALIDATION'
        $reason = "FILE mode requires exactly one eligible file; found $($files.Count)."
    }

    $resolvedPath = ''
    $sha256 = ''
    $totalBytes = 0L
    if ($files.Count -gt 0) {
        $totalBytes = [long](($files | Measure-Object -Property Length -Sum).Sum)
        if ($SelectionMode -eq 'FILE' -and $files.Count -eq 1) {
            $resolvedPath = $files[0].FullName
            $sha256 = (Get-FileHash -LiteralPath $files[0].FullName -Algorithm SHA256).Hash
        }
        else {
            $resolvedPath = (Get-Item -LiteralPath $sourceFolder -ErrorAction Stop).FullName
            $sha256 = Get-FileSetHash -Files $files
        }
    }

    return [pscustomobject]@{
        SourceId = $SourceId
        Status = $status
        SelectionMode = $SelectionMode
        ExportId = $recordExportId
        RelativeFolder = $RelativeFolder
        SourcePath = $resolvedPath
        FileCount = $files.Count
        TotalBytes = $totalBytes
        SHA256 = $sha256
        Reason = $reason
    }
}

function Write-Utf8Text {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LiteralPath,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )
    $Content | Set-Content -LiteralPath $LiteralPath -Encoding utf8NoBOM
    $CreatedPaths.Add($LiteralPath)
}

function Export-Utf8Csv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$InputObject,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LiteralPath
    )
    $InputObject | Export-Csv -LiteralPath $LiteralPath -NoTypeInformation -Encoding utf8NoBOM
    $CreatedPaths.Add($LiteralPath)
}

try {
    Show-RunProgress -Percent 5 -Status 'Validating interface and context'

    $ResolvedSourceRoot = Resolve-ExistingDirectory -LiteralPath $SourceRoot
    if ([System.IO.Path]::GetFileName($ResolvedSourceRoot) -ne $PipelineRunId) {
        throw "CTX-001 failure. SourceRoot is not the expected pipeline RUN_ID: $ResolvedSourceRoot"
    }

    $AuditOnly = -not $CommitRecords.IsPresent
    if (-not $AuditOnly -and [string]::IsNullOrWhiteSpace($OutputRoot)) {
        throw 'OutputRoot is required when CommitRecords is used.'
    }

    $ExportRoot = Resolve-ExistingDirectory -LiteralPath (Join-Path $ResolvedSourceRoot $ActiveExportId)

    Show-RunProgress -Percent 20 -Status 'Resolving deterministic source contract'

    $SourceContract = @(
        [pscustomobject]@{ SourceId='SPENCE_MASTER'; RelativeFolder='06_Consolidated_Database\03_SPENCE_MASTER'; SelectionMode='FOLDER_SET'; MinimumFiles=1; MaximumFiles=0 },
        [pscustomobject]@{ SourceId='LAS_LEP'; RelativeFolder='05_Ready_For_Compositing\02_LEP_CONFIDENCE'; SelectionMode='FOLDER_SET'; MinimumFiles=1; MaximumFiles=0 },
        [pscustomobject]@{ SourceId='LEP_CONFIDENCE'; RelativeFolder='06_Consolidated_Database\02_LEP_CONFIDENCE'; SelectionMode='FILE'; MinimumFiles=1; MaximumFiles=1 },
        [pscustomobject]@{ SourceId='AUTOLITH_CONFIDENCE'; RelativeFolder='05_Ready_For_Compositing\01_AUTOLITH_CONFIDENCE'; SelectionMode='FOLDER_SET'; MinimumFiles=1; MaximumFiles=0 }
    )

    $AutoLithOverride = ''
    $AutoLithExportId = ''
    if (-not [string]::IsNullOrWhiteSpace($AutoLithSourceFolder)) {
        $AutoLithOverride = Resolve-ExistingDirectory -LiteralPath $AutoLithSourceFolder
        if (-not $AutoLithOverride.StartsWith($ResolvedSourceRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "CTX-001 failure. AutoLithSourceFolder must be inside SourceRoot: $AutoLithOverride"
        }
        if ($AutoLithOverride -notmatch '\\CP09A-AUTOLITH-CSV-\d{8}-\d{6}\\CSV$') {
            throw "AutoLithSourceFolder must be the CSV folder of a CP09A-AUTOLITH-CSV run: $AutoLithOverride"
        }
        $AutoLithRegister = Join-Path (Split-Path -Path $AutoLithOverride -Parent) 'Run_Register.json'
        if (-not (Test-Path -LiteralPath $AutoLithRegister -PathType Leaf)) { throw "AutoLith Run_Register.json not found: $AutoLithRegister" }
        $AutoLithRun = Get-Content -LiteralPath $AutoLithRegister -Raw | ConvertFrom-Json
        if ([string]$AutoLithRun.Decision -ne 'PASS') { throw ('AutoLith conversion decision is not PASS: {0}' -f $AutoLithRun.Decision) }
        $AutoLithExportId = 'Exports_13_07_2026'
    }

    $Sources = [System.Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $SourceContract.Count; $index++) {
        $contract = $SourceContract[$index]
        Show-RunProgress -Percent (@(30, 40, 50, 60)[$index]) -Status ("Resolving {0}" -f $contract.SourceId)
        $folderOverride = ''
        $exportOverride = ''
        $minimumFiles = $contract.MinimumFiles
        $maximumFiles = $contract.MaximumFiles
        if ($contract.SourceId -eq 'AUTOLITH_CONFIDENCE' -and -not [string]::IsNullOrEmpty($AutoLithOverride)) {
            $folderOverride = $AutoLithOverride
            $exportOverride = $AutoLithExportId
            $minimumFiles = $AutoLithExpectedFiles
            $maximumFiles = $AutoLithExpectedFiles
        }
        $record = New-SourceRecord -SourceId $contract.SourceId -RelativeFolder $contract.RelativeFolder -SelectionMode $contract.SelectionMode -MinimumFiles $minimumFiles -MaximumFiles $maximumFiles -ExportRoot $ExportRoot -SourceFolderOverride $folderOverride -ExportIdOverride $exportOverride
        if (-not [string]::IsNullOrEmpty($exportOverride) -and $record.Status -eq 'PASS') {
            $record.Reason = 'Governed AutoLith CSV conversion selected. Cross-export source: Exports_13_07_2026. Requires validation before COMPARE.'
        }
        $Sources.Add($record)
    }

    $Jobs = [System.Collections.Generic.List[object]]::new()
    foreach ($source in $Sources) {
        foreach ($scale in $Scales) {
            $Jobs.Add([pscustomobject]@{
                JobId = ('{0}-{1}' -f $source.SourceId, $scale)
                SourceId = $source.SourceId
                ExportId = $source.ExportId
                SourcePath = $source.SourcePath
                SourceSHA256 = $source.SHA256
                Scale = $scale
                Status = if ($source.Status -eq 'PASS') { 'READY' } else { 'BLOCKED' }
                BlockReason = if ($source.Status -eq 'PASS') { '' } else { $source.Reason }
            })
        }
    }

    $BlockingSources = @($Sources | Where-Object { $_.Status -ne 'PASS' })
    $OverallStatus = if ($Sources.Count -eq 4 -and $Jobs.Count -eq 12 -and $BlockingSources.Count -eq 0) { 'READY_FOR_COMPOSITING' } else { 'REQUIRES_VALIDATION' }

    $RunRecord = [ordered]@{
        Classification = $Classification
        ScriptName = $ScriptName
        ScriptVersion = $ScriptVersion
        Pipeline = 'CP09A'
        Step = '13_Composites'
        PipelineRunId = $PipelineRunId
        CompositeRunId = $RunId
        ActiveExportId = $ActiveExportId
        AuditOnly = $AuditOnly
        CommitRecords = $CommitRecords.IsPresent
        SourceRoot = $ResolvedSourceRoot
        ExportRoot = $ExportRoot
        OutputRoot = $OutputRoot
        AuthoritativeSources = $ExpectedSourceIds
        Scales = $Scales
        ExpectedSources = 4
        ResolvedSources = $Sources.Count
        ExpectedJobs = 12
        PlannedJobs = $Jobs.Count
        BlockingSources = $BlockingSources.Count
        Status = $OverallStatus
        SourceModified = $false
        CreatedAt = (Get-Date).ToString('o')
    }

    $QaRecords = @(
        [pscustomobject]@{ Control='CTX-001'; Status='PASS'; Evidence=$ResolvedSourceRoot },
        [pscustomobject]@{ Control='PATH-001'; Status='PASS'; Evidence=$ExportRoot },
        [pscustomobject]@{ Control='DeterministicContract'; Status='PASS'; Evidence='Exact governed subpaths; no score or fuzzy matching' },
        [pscustomobject]@{ Control='SourceCount'; Status=if ($Sources.Count -eq 4) { 'PASS' } else { 'FAIL' }; Evidence=("Sources={0}" -f $Sources.Count) },
        [pscustomobject]@{ Control='JobCount'; Status=if ($Jobs.Count -eq 12) { 'PASS' } else { 'FAIL' }; Evidence=("Jobs={0}" -f $Jobs.Count) },
        [pscustomobject]@{ Control='DuplicateExclusion'; Status='PASS'; Evidence='Duplicate, copy and prior composite filename patterns excluded' },
        [pscustomobject]@{ Control='SourceImmutability'; Status='PASS'; Evidence='Read-only discovery and SHA256 only' },
        [pscustomobject]@{ Control='ExecutionReadiness'; Status=$OverallStatus; Evidence=("BlockingSources={0}" -f $BlockingSources.Count) }
    )

    Write-Host ''
    Write-Host 'CP09A COMPOSITE SOURCE DISCOVERY v1.0.4' -ForegroundColor Cyan
    Write-Host "RUN_ID: $RunId"
    Write-Host "ACTIVE EXPORT: $ActiveExportId"
    Write-Host "MODE: $(if ($AuditOnly) { 'AUDIT_ONLY' } else { 'COMMIT_RECORDS' })"
    Write-Host "STATUS: $OverallStatus"
    Write-Host ''
    $Sources | Select-Object SourceId,Status,SelectionMode,FileCount,SourcePath,SHA256,Reason | Format-Table -AutoSize
    Write-Host ''
    $Jobs | Select-Object JobId,SourceId,Scale,Status,BlockReason | Format-Table -AutoSize

    if ($AuditOnly) {
        Write-Host ''
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED, MODIFIED, MOVED, RENAMED OR DELETED.' -ForegroundColor Green
        Write-Host "READINESS: $OverallStatus"
        return
    }

    $ResolvedOutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
    $RunOutput = Join-Path $ResolvedOutputRoot $RunId
    if (Test-Path -LiteralPath $RunOutput) {
        throw "Overwrite is prohibited. Run output already exists: $RunOutput"
    }

    if ($PSCmdlet.ShouldProcess($RunOutput, 'Create CP09A composite initialization records')) {
        Show-RunProgress -Percent 75 -Status 'Creating run-scoped records'
        New-Item -ItemType Directory -Path $RunOutput -ErrorAction Stop | Out-Null

        Write-Utf8Text -LiteralPath (Join-Path $RunOutput 'Composite_Run_Register.json') -Content ($RunRecord | ConvertTo-Json -Depth 12)
        Export-Utf8Csv -InputObject $Sources.ToArray() -LiteralPath (Join-Path $RunOutput 'Composite_Source_Register.csv')
        Export-Utf8Csv -InputObject $Jobs.ToArray() -LiteralPath (Join-Path $RunOutput 'Composite_Job_Plan.csv')
        Export-Utf8Csv -InputObject $QaRecords -LiteralPath (Join-Path $RunOutput 'Composite_QAQC_Register.csv')

        $LearningRecords = @(
            [pscustomobject]@{ LearningId='LL-COMP-010'; ObservedIssue='PowerShell parser failure in v1.0.0'; RootCause='Format expression was not parenthesized inside hashtable value'; Correction='Parenthesized format expression'; PreventiveControl='Real PowerShell ParseFile gate before execution'; Status='CLOSED' },
            [pscustomobject]@{ LearningId='LL-COMP-011'; ObservedIssue='Score-based selector mixed source families'; RootCause='Non-matching files retained a positive generic score'; Correction='Replaced scoring with exact governed subpaths'; PreventiveControl='Deterministic source contract only'; Status='CLOSED' },
            [pscustomobject]@{ LearningId='LL-COMP-012'; ObservedIssue='Consolidated AutoLith folder was empty'; RootCause='AutoLith governed inputs remain in Ready_For_Compositing'; Correction='AUTOLITH_CONFIDENCE maps to exact Ready_For_Compositing folder'; PreventiveControl='Validate expected layer and non-zero eligible files'; Status='CLOSED' },
            [pscustomobject]@{ LearningId='LL-COMP-013'; ObservedIssue='SPENCE_MASTER contains multiple valid consolidated files'; RootCause='SPENCE_MASTER is a governed file set, not a single file'; Correction='Use deterministic FOLDER_SET mode'; PreventiveControl='Source-specific cardinality contract'; Status='CLOSED' }
        )
        Export-Utf8Csv -InputObject $LearningRecords -LiteralPath (Join-Path $RunOutput 'Learning_Register.csv')

        $Readme = @"
# CP09A Composite Initialization v1.0.4

Classification: Internal
Pipeline RUN ID: $PipelineRunId
Composite RUN ID: $RunId
Active export: $ActiveExportId
Status: $OverallStatus

The engine resolves SPENCE_MASTER, LAS_LEP, LEP_CONFIDENCE and AUTOLITH_CONFIDENCE through exact governed subpaths. It creates a 12-job plan for 1m, 3m and 6m. It does not execute compositing and does not modify source data.
"@
        Write-Utf8Text -LiteralPath (Join-Path $RunOutput 'README.md') -Content $Readme

        $ChangeLog = @"
# Changelog

## v1.0.3
- Fixed StrictMode Count failure by forcing array collections.
- Removed wildcard filter from file discovery.

## v1.0.2
- Preserved the v1.0.1 interface and added the optional ActiveExportId parameter.
- Replaced score-based and fuzzy source selection with deterministic governed subpaths.
- Defined SPENCE_MASTER and AUTOLITH_CONFIDENCE as governed file sets.
- Defined LEP_CONFIDENCE as exactly one consolidated file.
- Preserved LAS_LEP as a governed file set.
- Added validation, AuditOnly default, CommitRecords, QA/QC, learning records, manifest and SHA256 verification.
"@
        Write-Utf8Text -LiteralPath (Join-Path $RunOutput 'CHANGELOG.md') -Content $ChangeLog

        Show-RunProgress -Percent 88 -Status 'Generating manifest and SHA256 evidence'

        $ManifestItems = @(
            Get-ChildItem -LiteralPath $RunOutput -File -ErrorAction Stop |
                Sort-Object Name |
                ForEach-Object {
                    $hash = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
                    [pscustomobject]@{ Name=$_.Name; Length=[long]$_.Length; SHA256=$hash.Hash; Classification=$Classification }
                }
        )
        $Manifest = [ordered]@{ Classification=$Classification; RunId=$RunId; CreatedAt=(Get-Date).ToString('o'); FileCount=$ManifestItems.Count; Files=$ManifestItems }
        Write-Utf8Text -LiteralPath (Join-Path $RunOutput 'Manifest.json') -Content ($Manifest | ConvertTo-Json -Depth 12)

        $HashLines = @(
            Get-ChildItem -LiteralPath $RunOutput -File -ErrorAction Stop |
                Where-Object { $_.Name -ne 'SHA256SUMS.txt' } |
                Sort-Object Name |
                ForEach-Object {
                    $hash = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
                    '{0}  {1}' -f $hash.Hash, $_.Name
                }
        )
        Write-Utf8Text -LiteralPath (Join-Path $RunOutput 'SHA256SUMS.txt') -Content ($HashLines -join "`n")

        Show-RunProgress -Percent 96 -Status 'Verifying generated evidence'

        foreach ($hashLine in $HashLines) {
            $parts = $hashLine -split '\s{2}', 2
            if ($parts.Count -ne 2) { throw "Invalid SHA256 register line: $hashLine" }
            $actualHash = (Get-FileHash -LiteralPath (Join-Path $RunOutput $parts[1]) -Algorithm SHA256).Hash
            if ($actualHash -ne $parts[0]) { throw "SHA256 verification failed: $($parts[1])" }
        }

        $RequiredArtifacts = @('Composite_Run_Register.json','Composite_Source_Register.csv','Composite_Job_Plan.csv','Composite_QAQC_Register.csv','Learning_Register.csv','README.md','CHANGELOG.md','Manifest.json','SHA256SUMS.txt')
        foreach ($artifact in $RequiredArtifacts) {
            if (-not (Test-Path -LiteralPath (Join-Path $RunOutput $artifact) -PathType Leaf)) { throw "Missing required artifact: $artifact" }
        }

        Write-Host ''
        Write-Host "RECORD GENERATION PASS: $RunOutput" -ForegroundColor Green
        Write-Host "STATUS: $OverallStatus"
    }
}
catch {
    Write-Host ''
    Write-Host ("FAIL: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
finally {
    Write-Progress -Id $ProgressId -Activity 'CP09A composite initialization v1.0.4' -Completed
}
