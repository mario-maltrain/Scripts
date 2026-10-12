#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InventoryRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1\_Audit\Repository_Inventory\CP09A-REPOSITORY-INVENTORY-20261001-115947',
    [string]$ContractPath = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1\_Audit\Real_Test\source_contract_cp09a_autolith_v1_0_0.json',
    [string]$OutputRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1\_Audit\Duplicate_Audit'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$Version = '1.2.0'
$Classification = 'Internal'
$RunId = 'CP09A-DUPLICATE-AUDIT-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$RunFolder = Join-Path $OutputRoot $RunId
$LogPath = Join-Path $RunFolder 'Execution.log'
function Write-AuditLog {
    param([string]$Message,[ValidateSet('INFO','PASS','WARN','ERROR')][string]$Level='INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $line
    if (Test-Path -LiteralPath $RunFolder -PathType Container) { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 }
}
function Get-NormalizedWindowsPath {
    param([string]$Path)
    return [System.IO.Path]::GetFullPath($Path.Replace('/','\')).TrimEnd('\')
}
function Test-IsProtected {
    param([string]$CandidatePath,[string[]]$ProtectedPaths)
    $candidate = Get-NormalizedWindowsPath $CandidatePath
    foreach ($protectedPath in $ProtectedPaths) {
        $protected = Get-NormalizedWindowsPath $protectedPath
        if ($candidate.Equals($protected,[System.StringComparison]::OrdinalIgnoreCase)) { return $true }
        if ($candidate.StartsWith($protected + '\',[System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}
try {
    New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null
    Write-AuditLog "BHP INVENT | CP09A Duplicate Audit | Version $Version | Classification $Classification"
    Write-AuditLog "RUN_ID: $RunId | MODE: READ_ONLY"
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Validating inputs' -PercentComplete 5
    if (-not (Test-Path -LiteralPath $InventoryRoot -PathType Container)) { throw "Inventory root not found: $InventoryRoot" }
    if (-not (Test-Path -LiteralPath $ContractPath -PathType Leaf)) { throw "Source contract not found: $ContractPath" }
    $hashFile = Join-Path $InventoryRoot 'Files_SHA256.csv'
    $inventorySummaryFile = Join-Path $InventoryRoot 'Summary.json'
    if (-not (Test-Path -LiteralPath $hashFile -PathType Leaf)) { throw "Files_SHA256.csv not found: $hashFile" }
    if (-not (Test-Path -LiteralPath $inventorySummaryFile -PathType Leaf)) { throw "Summary.json not found: $inventorySummaryFile" }
    $contract = Get-Content -LiteralPath $ContractPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($contract.classification -ne 'Internal') { throw 'Unexpected contract classification.' }
    if ($contract.governance.audit_mode -ne 'READ_ONLY') { throw 'Contract is not configured for READ_ONLY mode.' }
    if ([bool]$contract.governance.deletion_allowed) { throw 'Safety block: contract permits deletion.' }
    [string[]]$protectedPaths = @($contract.protected_paths | ForEach-Object { Get-NormalizedWindowsPath ([string]$_.path) })
    if ($protectedPaths.Count -ne 5) { throw "Expected 5 protected paths; found $($protectedPaths.Count)." }
    $protectedRows = @($contract.protected_paths | ForEach-Object {
        $p = Get-NormalizedWindowsPath ([string]$_.path)
        [pscustomobject]@{SourceId=[string]$_.source_id;ProtectedPath=$p;Role=[string]$_.role;Authoritative=[bool]$_.authoritative;ReadOnly=[bool]$_.read_only;Protection=[string]$_.protection;ExistsAtAuditTime=(Test-Path -LiteralPath $p)}
    })
    $protectedRows | Export-Csv -LiteralPath (Join-Path $RunFolder 'Protected_Paths.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Loading SHA256 inventory' -PercentComplete 20
    [object[]]$hashRows = @(Import-Csv -LiteralPath $hashFile | Where-Object { $_.Status -eq 'PASS' -and -not [string]::IsNullOrWhiteSpace([string]$_.SHA256) -and -not [string]::IsNullOrWhiteSpace([string]$_.FullPath) })
    if ($hashRows.Count -eq 0) { throw 'No valid PASS hash records were found.' }
    $included = New-Object System.Collections.ArrayList
    $excluded = New-Object System.Collections.ArrayList
    for ($i=0; $i -lt $hashRows.Count; $i++) {
        $row = $hashRows[$i]
        $percent = 20 + [int](40 * (($i + 1) / [double]$hashRows.Count))
        Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Applying protected-path exclusions' -CurrentOperation ("Record {0} of {1}" -f ($i+1),$hashRows.Count) -PercentComplete $percent
        if (Test-IsProtected ([string]$row.FullPath) $protectedPaths) {
            [void]$excluded.Add([pscustomobject]@{FullPath=[string]$row.FullPath;RelativePath=[string]$row.RelativePath;SizeBytes=[int64]$row.SizeBytes;SHA256=[string]$row.SHA256;ExclusionReason='PROTECTED_PATH'})
        } else { [void]$included.Add($row) }
    }
    $excluded | Sort-Object FullPath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Excluded_Files_Register.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Grouping duplicates' -PercentComplete 65
    [object[]]$groups = @($included | Group-Object SHA256 | Where-Object { $_.Count -gt 1 })
    $groupRows = New-Object System.Collections.ArrayList
    $fileRows = New-Object System.Collections.ArrayList
    $groupId = 0
    foreach ($group in $groups) {
        $groupId++
        [object[]]$items = @($group.Group)
        $sizeEach = [int64]$items[0].SizeBytes
        $recoverable = [int64]($sizeEach * ($items.Count - 1))
        [void]$groupRows.Add([pscustomobject]@{GroupId=$groupId;SHA256=[string]$group.Name;Copies=$items.Count;SizeEachBytes=$sizeEach;RecoverableBytes=$recoverable;RecoverableMB=[math]::Round($recoverable/1MB,3);Decision='REVIEW_ONLY'})
        foreach ($item in $items) { [void]$fileRows.Add([pscustomobject]@{GroupId=$groupId;SHA256=[string]$group.Name;SizeBytes=[int64]$item.SizeBytes;FullPath=[string]$item.FullPath;RelativePath=[string]$item.RelativePath;Action='NO_DELETE'}) }
    }
    $groupRows | Sort-Object RecoverableBytes -Descending | Export-Csv -LiteralPath (Join-Path $RunFolder 'Duplicate_Groups_Outside_Protected.csv') -NoTypeInformation -Encoding UTF8
    $fileRows | Sort-Object GroupId,FullPath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Duplicate_Files_Outside_Protected.csv') -NoTypeInformation -Encoding UTF8
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Building folder summary' -PercentComplete 80
    [object[]]$folderSummary = @($fileRows | ForEach-Object { [pscustomobject]@{Folder=(Split-Path $_.FullPath -Parent);SizeBytes=[int64]$_.SizeBytes;GroupId=$_.GroupId} } | Group-Object Folder | ForEach-Object {
        [object[]]$rows = @($_.Group)
        $bytes = ($rows | Measure-Object SizeBytes -Sum).Sum
        if ($null -eq $bytes) { $bytes = 0 }
        [pscustomobject]@{Folder=$_.Name;DuplicateFileRecords=$rows.Count;DuplicateHashGroups=@($rows | Select-Object -ExpandProperty GroupId -Unique).Count;TotalDuplicateBytes=[int64]$bytes;TotalDuplicateMB=[math]::Round($bytes/1MB,3)}
    } | Sort-Object TotalDuplicateBytes -Descending)
    $folderSummary | Export-Csv -LiteralPath (Join-Path $RunFolder 'Duplicate_Summary_By_Folder.csv') -NoTypeInformation -Encoding UTF8
    $recoverableTotal = ($groupRows | Measure-Object RecoverableBytes -Sum).Sum
    if ($null -eq $recoverableTotal) { $recoverableTotal = 0 }
    $inventorySummary = Get-Content -LiteralPath $inventorySummaryFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $summary = [ordered]@{Title='BHP INVENT | CP09A Duplicate Audit';Classification=$Classification;Version=$Version;RunId=$RunId;Mode='READ_ONLY';DeleteOperations=0;BaselineRunId=[string]$inventorySummary.RunId;InventoryRowsLoaded=$hashRows.Count;ProtectedPathCount=$protectedPaths.Count;ExcludedFileRecords=$excluded.Count;IncludedFileRecords=$included.Count;DuplicateHashGroupsOutsideProtected=$groupRows.Count;DuplicateFileRecordsOutsideProtected=$fileRows.Count;PotentialRecoverableBytesOutsideProtected=[int64]$recoverableTotal;PotentialRecoverableMBOutsideProtected=[math]::Round($recoverableTotal/1MB,3);OutputFolder=$RunFolder;Status='PASS_READ_ONLY'}
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.txt') -Encoding UTF8
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Status 'Generating manifest' -PercentComplete 95
    $manifest = @(Get-ChildItem -LiteralPath $RunFolder -File | Where-Object Name -ne 'SHA256SUMS.txt' | ForEach-Object { $h=Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256; '{0} *{1}' -f $h.Hash,$_.Name })
    $manifest | Set-Content -LiteralPath (Join-Path $RunFolder 'SHA256SUMS.txt') -Encoding ASCII
    Write-Progress -Id 1 -Activity 'CP09A duplicate audit' -Completed
    Write-AuditLog 'AUDIT: PASS_READ_ONLY' 'PASS'
    Write-AuditLog 'DELETE OPERATIONS: 0' 'PASS'
    Write-AuditLog "OUTPUT: $RunFolder" 'PASS'
    Write-Host "OUTPUT: $RunFolder" -ForegroundColor Cyan
    exit 0
} catch {
    try {
        if (-not (Test-Path -LiteralPath $RunFolder)) { New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null }
        Write-AuditLog $_.Exception.Message 'ERROR'
        @('MESSAGE: '+$_.Exception.Message,'SCRIPT: '+$_.InvocationInfo.ScriptName,'LINE: '+$_.InvocationInfo.ScriptLineNumber,'POSITION: '+$_.InvocationInfo.PositionMessage,'',($_ | Out-String)) | Set-Content -LiteralPath (Join-Path $RunFolder 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: '+$_.Exception.Message) -ForegroundColor Red }
    exit 1
}
