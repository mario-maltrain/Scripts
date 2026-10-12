#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$RepositoryRoot = 'C:\Python\GeoAi',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [switch]$ExecuteCleanup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.0.0'
$Classification = 'Internal'
$RunId = 'CP09A-DUPLICATE-CLEANUP-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$Mode = if ($ExecuteCleanup) { 'EXECUTE' } else { 'AUDIT_ONLY' }
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\Repository_Cleanup'
$InstalledPath = Join-Path $InstalledFolder 'CP09A_AutoLith_Duplicate_Audit_Cleanup_v1_0_0.ps1'
$AuditRoot = Join-Path $ApplicationRoot '_Audit\Repository_Cleanup'
$RunFolder = Join-Path $AuditRoot $RunId
$LogPath = Join-Path $RunFolder 'Execution.log'
$SourcePath = $MyInvocation.MyCommand.Path
$InventoryAuditRoot = Join-Path $ApplicationRoot '_Audit\Repository_Inventory'
$FamilyPattern = 'CP09A_AutoLith_Duplicate_Audit_Cleanup*.ps1'

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','PASS')][string]$Level='INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $line
    if (Test-Path -LiteralPath $RunFolder) { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 }
}

function Test-Syntax {
    param([string]$Path)
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if ($null -eq $parseErrors) { return $true }
    return ($parseErrors.Count -eq 0)
}

function Get-NormalizedPath {
    param([string]$Path)
    return [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
}

function Test-IsUnderRoot {
    param([string]$Path,[string]$Root)
    $normalizedPath = Get-NormalizedPath $Path
    $normalizedRoot = Get-NormalizedPath $Root
    return $normalizedPath.StartsWith($normalizedRoot + '\',[System.StringComparison]::OrdinalIgnoreCase)
}

function Install-CurrentScript {
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-Syntax $SourcePath)) { throw 'Downloaded script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    if (-not ((Get-NormalizedPath $SourcePath).Equals((Get-NormalizedPath $InstalledPath),[System.StringComparison]::OrdinalIgnoreCase))) {
        Copy-Item -LiteralPath $SourcePath -Destination $InstalledPath -Force
    }
    if (-not (Test-Syntax $InstalledPath)) { throw 'Installed script failed parser validation.' }
    foreach ($folder in @($InstalledFolder,(Join-Path $env:USERPROFILE 'Downloads'))) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($candidate in @(Get-ChildItem -LiteralPath $folder -Filter $FamilyPattern -File -ErrorAction SilentlyContinue)) {
            if ($candidate.FullName.Equals($SourcePath,[System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($candidate.FullName.Equals($InstalledPath,[System.StringComparison]::OrdinalIgnoreCase)) { continue }
            Remove-Item -LiteralPath $candidate.FullName -Force
            Write-Log ('Removed obsolete cleanup script: ' + $candidate.FullName) 'WARN'
        }
    }
}

function Get-LatestValidInventoryRun {
    if (-not (Test-Path -LiteralPath $InventoryAuditRoot -PathType Container)) { throw "Inventory audit root not found: $InventoryAuditRoot" }
    $validRuns = New-Object System.Collections.ArrayList
    foreach ($folder in @(Get-ChildItem -LiteralPath $InventoryAuditRoot -Directory -ErrorAction Stop)) {
        $summaryPath = Join-Path $folder.FullName 'Summary.json'
        $hashPath = Join-Path $folder.FullName 'Files_SHA256.csv'
        $filesPath = Join-Path $folder.FullName 'Files.csv'
        if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
        if (-not (Test-Path -LiteralPath $hashPath -PathType Leaf)) { continue }
        if (-not (Test-Path -LiteralPath $filesPath -PathType Leaf)) { continue }
        try {
            $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($summary.Status -eq 'PASS' -and [int]$summary.ScanErrorCount -eq 0 -and $summary.HashesGenerated -eq $true) {
                [void]$validRuns.Add([pscustomobject]@{Folder=$folder;Summary=$summary;HashPath=$hashPath;FilesPath=$filesPath})
            }
        } catch {
            Write-Log ('Ignored unreadable inventory run: ' + $folder.FullName) 'WARN'
        }
    }
    if ($validRuns.Count -eq 0) { throw 'No valid PASS inventory run with SHA256 data was found.' }
    return ($validRuns | Sort-Object { $_.Folder.LastWriteTime } -Descending | Select-Object -First 1)
}

function Get-FolderSignatureFromInventory {
    param([string]$FolderPath,[object[]]$HashRows)
    $normalizedFolder = (Get-NormalizedPath $FolderPath) + '\'
    $rows = @($HashRows | Where-Object { $_.Status -eq 'PASS' -and $_.FullPath.StartsWith($normalizedFolder,[System.StringComparison]::OrdinalIgnoreCase) })
    if ($rows.Count -eq 0) { return $null }
    $relativeHashes = foreach ($row in $rows) {
        $relative = $row.FullPath.Substring($normalizedFolder.Length).ToLowerInvariant()
        '{0}|{1}|{2}' -f $relative,$row.SizeBytes,$row.SHA256
    }
    $canonical = ($relativeHashes | Sort-Object) -join "`n"
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($canonical)
        $signature = ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','')
    } finally { $sha.Dispose() }
    return [pscustomobject]@{Signature=$signature;FileCount=$rows.Count;TotalBytes=[int64](($rows | Measure-Object SizeBytes -Sum).Sum)}
}

try {
    if (-not (Test-Path -LiteralPath $RepositoryRoot -PathType Container)) { throw "Repository root not found: $RepositoryRoot" }
    if (-not (Test-Path -LiteralPath $ApplicationRoot -PathType Container)) { throw "Application root not found: $ApplicationRoot" }
    New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null
    Write-Log "BHP INVENT | CP09A AutoLith Duplicate Audit and Cleanup | Version $Version | Classification $Classification"
    Write-Log "RUN_ID: $RunId | MODE: $Mode"
    Install-CurrentScript
    Write-Log "Installed script: $InstalledPath" 'PASS'

    $baseline = Get-LatestValidInventoryRun
    $baselineRunId = [string]$baseline.Summary.RunId
    Write-Log "Baseline selected: $baselineRunId" 'PASS'
    $hashRows = @(Import-Csv -LiteralPath $baseline.HashPath)
    $fileRows = @(Import-Csv -LiteralPath $baseline.FilesPath)
    if ($hashRows.Count -eq 0) { throw 'Files_SHA256.csv contains no rows.' }

    Write-Progress -Id 1 -Activity 'Duplicate analysis' -Status 'Grouping SHA256 records' -PercentComplete 15
    $duplicateGroups = @($hashRows | Where-Object { $_.Status -eq 'PASS' -and -not [string]::IsNullOrWhiteSpace($_.SHA256) } | Group-Object SHA256 | Where-Object { $_.Count -gt 1 })
    $duplicateFileRows = New-Object System.Collections.ArrayList
    $duplicateGroupRows = New-Object System.Collections.ArrayList
    $groupIndex = 0
    foreach ($group in $duplicateGroups) {
        $groupIndex++
        $items = @($group.Group)
        $size = [int64]$items[0].SizeBytes
        $recoverable = [int64]($size * ($items.Count - 1))
        [void]$duplicateGroupRows.Add([pscustomobject]@{GroupId=$groupIndex;SHA256=$group.Name;Copies=$items.Count;SizeEachBytes=$size;RecoverableBytes=$recoverable;RecoverableMB=[math]::Round($recoverable/1MB,3)})
        foreach ($item in $items) {
            [void]$duplicateFileRows.Add([pscustomobject]@{GroupId=$groupIndex;SHA256=$group.Name;SizeBytes=[int64]$item.SizeBytes;FullPath=$item.FullPath;RelativePath=$item.RelativePath})
        }
    }
    $duplicateGroupRows | Sort-Object RecoverableBytes -Descending | Export-Csv -LiteralPath (Join-Path $RunFolder 'Duplicate_Groups.csv') -NoTypeInformation -Encoding UTF8
    $duplicateFileRows | Sort-Object GroupId,FullPath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Duplicate_Files.csv') -NoTypeInformation -Encoding UTF8

    Write-Progress -Id 1 -Activity 'Duplicate analysis' -Status 'Validating duplicate folders' -PercentComplete 45
    $cleanupCandidates = New-Object System.Collections.ArrayList
    $copyFolders = @($fileRows | Select-Object -ExpandProperty RelativeFolder -Unique | Where-Object { $_ -match ' Copiar$' })
    foreach ($relativeCopy in $copyFolders) {
        $copyPath = Join-Path $RepositoryRoot $relativeCopy
        $relativeOriginal = $relativeCopy -replace ' Copiar$',''
        $originalPath = Join-Path $RepositoryRoot $relativeOriginal
        if (-not (Test-Path -LiteralPath $copyPath -PathType Container)) { continue }
        if (-not (Test-Path -LiteralPath $originalPath -PathType Container)) { continue }
        $copySignature = Get-FolderSignatureFromInventory $copyPath $hashRows
        $originalSignature = Get-FolderSignatureFromInventory $originalPath $hashRows
        $exactMatch = $null -ne $copySignature -and $null -ne $originalSignature -and $copySignature.Signature -eq $originalSignature.Signature
        [void]$cleanupCandidates.Add([pscustomobject]@{
            CandidateType='EXACT_DUPLICATE_COPY_FOLDER';DeletePath=$copyPath;KeepPath=$originalPath
            FileCount=if($null -ne $copySignature){$copySignature.FileCount}else{0}
            TotalBytes=if($null -ne $copySignature){$copySignature.TotalBytes}else{0}
            Evidence=if($exactMatch){'MATCHED_FOLDER_SIGNATURE'}else{'NOT_MATCHED'}
            ApprovedForDeletion=$exactMatch;ActionStatus='PENDING'
        })
    }

    Write-Progress -Id 1 -Activity 'Duplicate analysis' -Status 'Identifying obsolete inventory runs' -PercentComplete 65
    foreach ($oldRun in @(Get-ChildItem -LiteralPath $InventoryAuditRoot -Directory)) {
        if ($oldRun.Name -eq $baselineRunId) { continue }
        [void]$cleanupCandidates.Add([pscustomobject]@{
            CandidateType='OBSOLETE_REPOSITORY_INVENTORY_RUN';DeletePath=$oldRun.FullName;KeepPath=$baseline.Folder.FullName
            FileCount=@(Get-ChildItem -LiteralPath $oldRun.FullName -File -Recurse -ErrorAction SilentlyContinue).Count
            TotalBytes=[int64]((Get-ChildItem -LiteralPath $oldRun.FullName -File -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum)
            Evidence='SUPERSEDED_BY_LATEST_VALID_PASS_BASELINE';ApprovedForDeletion=$true;ActionStatus='PENDING'
        })
    }

    $approved = @($cleanupCandidates | Where-Object { $_.ApprovedForDeletion -eq $true })
    $rejected = @($cleanupCandidates | Where-Object { $_.ApprovedForDeletion -ne $true })
    $deletedCount = 0
    $deletedBytes = [int64]0

    Write-Progress -Id 1 -Activity 'Duplicate analysis' -Status 'Applying governed cleanup' -PercentComplete 80
    foreach ($candidate in $cleanupCandidates) {
        if (-not $candidate.ApprovedForDeletion) { $candidate.ActionStatus='RETAINED_REQUIRES_VALIDATION'; continue }
        if (-not $ExecuteCleanup) { $candidate.ActionStatus='AUDIT_ONLY_APPROVED'; continue }
        if (-not (Test-IsUnderRoot $candidate.DeletePath $RepositoryRoot)) { throw "Safety block: path outside repository root: $($candidate.DeletePath)" }
        if ($candidate.DeletePath.Equals($baseline.Folder.FullName,[System.StringComparison]::OrdinalIgnoreCase)) { throw 'Safety block: attempted deletion of baseline run.' }
        if (Test-Path -LiteralPath $candidate.DeletePath) {
            Remove-Item -LiteralPath $candidate.DeletePath -Recurse -Force
            if (Test-Path -LiteralPath $candidate.DeletePath) { throw "Deletion verification failed: $($candidate.DeletePath)" }
            $candidate.ActionStatus='DELETED_VERIFIED'
            $deletedCount++
            $deletedBytes += [int64]$candidate.TotalBytes
            Write-Log ('Deleted and verified: ' + $candidate.DeletePath) 'PASS'
        } else { $candidate.ActionStatus='ALREADY_ABSENT' }
    }
    $cleanupCandidates | Export-Csv -LiteralPath (Join-Path $RunFolder 'Cleanup_Register.csv') -NoTypeInformation -Encoding UTF8

    $recoverableBytes = [int64](($duplicateGroupRows | Measure-Object RecoverableBytes -Sum).Sum)
    $approvedBytes = [int64](($approved | Measure-Object TotalBytes -Sum).Sum)
    $architecture = @"
# CP09A AutoLith Logical Architecture

Classification: Internal
Baseline: $baselineRunId

## 1. Data layer
- 00_Base_Datos/Spence/00_DataBase: authoritative Spence source data, read-only.
- 00_Base_Datos/GeoAi/03_Original_Data_Exports/01_Raw_Data: authoritative GeoAI raw inputs, read-only.
- 00_Base_Datos/CP09A_Output: immutable historical CP09A outputs.

## 2. Application layer
- 01_Aplicacion/CP09A_Release_v2.1/00_Script: launchers, inventory, validation and cleanup tools.
- 01_Source: source adapters and input contracts.
- 02_Model: deterministic AutoLith compositing and argmax logic.
- 03_Output: current application outputs only.
- 04_Manifest: manifests, schemas and SHA256 records.
- 05_Logs: execution logs.
- 06_Backup: controlled rollback assets.
- 07_Archive: superseded releases retained only when required.

## 3. Processing contract
1. Discover and validate source paths.
2. Load Spence and GeoAI sources in read-only mode.
3. Validate HOLE_ID, intervals, probability columns, NULL rules and coverage threshold.
4. Composite each source independently.
5. Derive AutoLith class by argmax on aggregated raw probabilities without normalization.
6. Compare sources only in the comparison stage.
7. Export run-scoped outputs, QA/QC, manifest and SHA256.
8. Promote only validated PASS runs.

## 4. Governance
- Every execution requires a unique RUN_ID.
- Inputs remain immutable.
- Outputs are isolated by run.
- No deletion is allowed without hash or supersession evidence.
- Classification is Internal.
"@
    $architecture | Set-Content -LiteralPath (Join-Path $RunFolder 'CP09A_AutoLith_Logical_Architecture.md') -Encoding UTF8

    $plan = @"
# CP09A AutoLith Refactoring Plan

Classification: Internal
Baseline: $baselineRunId

## Phase 1 - Freeze and map
- Freeze the inventory baseline and preserve Summary.json, Files.csv and Files_SHA256.csv.
- Externalize all source, output, threshold and NULL settings.
- Replace hard-coded legacy paths with the new GeoAi root contract.

## Phase 2 - Separate responsibilities
- Isolate source discovery, schema validation, compositing, argmax, comparison, export and validation modules.
- Keep Spence and GeoAI processing independent until the comparison stage.
- Standardize run-scoped output folders and manifests.

## Phase 3 - Migrate safely
- Execute the current baseline and refactored workflow against the same read-only inputs.
- Compare row counts, interval coverage, derived classes, QA/QC and SHA256 outputs.
- Mark every difference as explained, accepted or blocking.

## Phase 4 - Promote
- Promote only after parser, smoke, regression and real-data validation PASS.
- Update launchers, README, manifest, changelog and deployment records.
- Archive the superseded implementation only after rollback validation.

## Decision gates
- Gate 1: path and schema contracts validated.
- Gate 2: numerical and classification equivalence validated.
- Gate 3: governance and rollback evidence complete.
- Gate 4: operational launcher promoted.
"@
    $plan | Set-Content -LiteralPath (Join-Path $RunFolder 'CP09A_AutoLith_Refactoring_Plan.md') -Encoding UTF8

    $summary = [ordered]@{
        Title='BHP INVENT | CP09A AutoLith Duplicate Audit and Cleanup';Classification=$Classification;Version=$Version
        RunId=$RunId;Mode=$Mode;BaselineRunId=$baselineRunId;DuplicateHashGroups=$duplicateGroups.Count
        DuplicateFileRecords=$duplicateFileRows.Count;PotentialRecoverableBytes=$recoverableBytes
        CleanupCandidates=$cleanupCandidates.Count;ApprovedCandidates=$approved.Count;RequiresValidation=$rejected.Count
        ApprovedBytes=$approvedBytes;DeletedCount=$deletedCount;DeletedBytes=$deletedBytes
        Status=if($ExecuteCleanup){'PASS_CLEANUP_EXECUTED'}else{'PASS_AUDIT_ONLY'}
    }
    $summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.txt') -Encoding UTF8

    $manifest = foreach ($output in @(Get-ChildItem -LiteralPath $RunFolder -File | Where-Object Name -ne 'SHA256SUMS.txt')) {
        $hash = Get-FileHash -LiteralPath $output.FullName -Algorithm SHA256
        '{0} *{1}' -f $hash.Hash,$output.Name
    }
    $manifest | Set-Content -LiteralPath (Join-Path $RunFolder 'SHA256SUMS.txt') -Encoding ASCII
    Write-Progress -Id 1 -Activity 'Duplicate analysis' -Completed
    Write-Log "STATUS: $($summary.Status)" 'PASS'
    Write-Log "OUTPUT: $RunFolder" 'PASS'

    if (-not ((Get-NormalizedPath $SourcePath).Equals((Get-NormalizedPath $InstalledPath),[System.StringComparison]::OrdinalIgnoreCase))) {
        try { Remove-Item -LiteralPath $SourcePath -Force; Write-Log "Removed downloaded source: $SourcePath" 'PASS' }
        catch { Write-Log "Downloaded source could not be removed: $SourcePath" 'WARN' }
    }
    Write-Host ('INSTALLED SCRIPT: ' + $InstalledPath) -ForegroundColor Cyan
    Write-Host ('OUTPUT: ' + $RunFolder) -ForegroundColor Cyan
    exit 0
}
catch {
    try {
        if (-not (Test-Path -LiteralPath $RunFolder)) { New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null }
        Write-Log $_.Exception.Message 'ERROR'
        @('MESSAGE: ' + $_.Exception.Message,'SCRIPT: ' + $_.InvocationInfo.ScriptName,'LINE: ' + $_.InvocationInfo.ScriptLineNumber,'POSITION: ' + $_.InvocationInfo.PositionMessage,'',($_ | Out-String)) | Set-Content -LiteralPath (Join-Path $RunFolder 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: ' + $_.Exception.Message) -ForegroundColor Red }
    exit 1
}
