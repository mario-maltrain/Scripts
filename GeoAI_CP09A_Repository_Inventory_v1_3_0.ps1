#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$RepositoryRoot = 'C:\Python\GeoAi',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [switch]$SkipHashes
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.3.0'
$Classification = 'Internal'
$RunId = 'CP09A-REPOSITORY-INVENTORY-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$FamilyPattern = 'GeoAI_CP09A_Repository_Inventory*.ps1'
$InstalledName = 'GeoAI_CP09A_Repository_Inventory_v1_3_0.ps1'
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\Repository_Inventory'
$InstalledPath = Join-Path $InstalledFolder $InstalledName
$AuditRoot = Join-Path $ApplicationRoot '_Audit\Repository_Inventory'
$RunFolder = Join-Path $AuditRoot $RunId
$LogPath = Join-Path $RunFolder 'Execution.log'
$SourcePath = $MyInvocation.MyCommand.Path

$ExcludedDirectoryNames = @(
    '_Discovery','.venv','venv','env','.env','virtualenv','__pycache__',
    '.git','.svn','.hg','node_modules','tmp','temp','.tmp','.temp','cache',
    '.cache','.pytest_cache','.mypy_cache','.ruff_cache','.tox','site-packages',
    '$RECYCLE.BIN','System Volume Information'
)
$ExcludedFilePatterns = @(
    '~$*','*.tmp','*.temp','*.bak','*.swp','*.swo','*.part','*.partial',
    '*.crdownload','*.download','*.lock','*.lck','*.dmp','*.pyc','*.pyo'
)

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

function Get-FullPath {
    param([string]$Path)
    return [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
}

function Get-RelativePath {
    param([string]$BasePath,[string]$TargetPath)
    $baseUri = New-Object System.Uri(($BasePath.TrimEnd('\') + '\'))
    $targetUri = New-Object System.Uri($TargetPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('/','\')
}

function Test-ExcludedDirectory {
    param([System.IO.DirectoryInfo]$Directory,[string]$RootPath,[string]$RunPath)
    $full = $Directory.FullName.TrimEnd('\')
    if ($full.Equals($RunPath,[System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($full.StartsWith($RunPath + '\',[System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($ExcludedDirectoryNames -contains $Directory.Name) { return $true }
    $relative = $full.Substring($RootPath.Length).TrimStart('\')
    foreach ($segment in @($relative -split '\\')) {
        if ($ExcludedDirectoryNames -contains $segment) { return $true }
    }
    return $false
}

function Test-ExcludedFile {
    param([System.IO.FileInfo]$File)
    foreach ($pattern in $ExcludedFilePatterns) {
        if ($File.Name -like $pattern) { return $true }
    }
    return $false
}

function Install-Script {
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-Syntax -Path $SourcePath)) { throw 'Current script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    if (-not ((Get-FullPath $SourcePath).Equals((Get-FullPath $InstalledPath),[System.StringComparison]::OrdinalIgnoreCase))) {
        Copy-Item -LiteralPath $SourcePath -Destination $InstalledPath -Force
    }
    if (-not (Test-Syntax -Path $InstalledPath)) { throw 'Installed script failed parser validation.' }
}

function Remove-ObsoleteScripts {
    foreach ($folder in @($InstalledFolder,(Join-Path $env:USERPROFILE 'Downloads'))) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($candidate in @(Get-ChildItem -LiteralPath $folder -Filter $FamilyPattern -File -ErrorAction SilentlyContinue)) {
            if ($candidate.FullName.Equals($SourcePath,[System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($candidate.FullName.Equals($InstalledPath,[System.StringComparison]::OrdinalIgnoreCase)) { continue }
            Remove-Item -LiteralPath $candidate.FullName -Force
            Write-Log ('Deleted obsolete script: ' + $candidate.FullName) 'WARN'
        }
    }
}

try {
    if (-not (Test-Path -LiteralPath $RepositoryRoot -PathType Container)) { throw "Repository root does not exist: $RepositoryRoot" }
    if (-not (Test-Path -LiteralPath $ApplicationRoot -PathType Container)) { throw "Application root does not exist: $ApplicationRoot" }

    New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null
    Write-Log "BHP INVENT | CP09A Repository Inventory | Version $Version | Classification $Classification"
    Write-Log "RUN_ID: $RunId"
    Install-Script
    Write-Log "Installed script: $InstalledPath" 'PASS'
    Remove-ObsoleteScripts

    $rootFull = Get-FullPath $RepositoryRoot
    $runFull = Get-FullPath $RunFolder
    $queue = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $directories = New-Object System.Collections.ArrayList
    $files = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $queue.Enqueue((Get-Item -LiteralPath $RepositoryRoot -Force))
    $visited = 0

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $visited++
        Write-Progress -Id 1 -Activity 'Scanning GeoAI repository' -Status $current.FullName -CurrentOperation ("Folders scanned: {0}" -f $visited)
        if (-not $current.FullName.Equals($rootFull,[System.StringComparison]::OrdinalIgnoreCase)) { [void]$directories.Add($current) }
        try {
            foreach ($child in @(Get-ChildItem -LiteralPath $current.FullName -Directory -Force -ErrorAction Stop)) {
                if (-not (Test-ExcludedDirectory $child $rootFull $runFull)) { $queue.Enqueue($child) }
            }
        } catch { [void]$errors.Add([pscustomobject]@{Path=$current.FullName;Stage='DirectoryEnumeration';Error=$_.Exception.Message}) }
        try {
            foreach ($file in @(Get-ChildItem -LiteralPath $current.FullName -File -Force -ErrorAction Stop)) {
                if (-not (Test-ExcludedFile $file)) { [void]$files.Add($file) }
            }
        } catch { [void]$errors.Add([pscustomobject]@{Path=$current.FullName;Stage='FileEnumeration';Error=$_.Exception.Message}) }
    }
    Write-Progress -Id 1 -Activity 'Scanning GeoAI repository' -Completed

    [object[]]$folderRows = @(
        foreach ($directory in $directories) {
            [pscustomobject]@{
                RelativePath = Get-RelativePath $RepositoryRoot $directory.FullName
                FullPath = $directory.FullName
                ParentPath = if ($null -ne $directory.Parent) { $directory.Parent.FullName } else { '' }
                Name = $directory.Name
                CreationTime = $directory.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
                LastWriteTime = $directory.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            }
        }
    )
    $folderRows | Sort-Object RelativePath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Folders.csv') -NoTypeInformation -Encoding UTF8

    $fileRows = New-Object System.Collections.ArrayList
    $hashRows = New-Object System.Collections.ArrayList
    $fileTotal = $files.Count
    for ($i=0; $i -lt $fileTotal; $i++) {
        $file = $files[$i]
        $percent = if ($fileTotal -gt 0) { [math]::Min(100,[int](($i + 1) * 100 / $fileTotal)) } else { 100 }
        Write-Progress -Id 2 -Activity 'Inventorying files and SHA256' -Status $file.FullName -PercentComplete $percent -CurrentOperation ("File {0} of {1}" -f ($i+1),$fileTotal)
        $relative = Get-RelativePath $RepositoryRoot $file.FullName
        $relativeFolder = Get-RelativePath $RepositoryRoot $file.DirectoryName
        $extension = if ([string]::IsNullOrWhiteSpace($file.Extension)) { '[NO_EXTENSION]' } else { $file.Extension.ToLowerInvariant() }
        [void]$fileRows.Add([pscustomobject]@{
            RelativePath=$relative;RelativeFolder=$relativeFolder;FullPath=$file.FullName;Name=$file.Name;Extension=$extension
            SizeBytes=[int64]$file.Length;SizeMB=[math]::Round($file.Length/1MB,6)
            CreationTime=$file.CreationTime.ToString('yyyy-MM-dd HH:mm:ss');LastWriteTime=$file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss');Attributes=$file.Attributes.ToString()
        })
        if (-not $SkipHashes) {
            try {
                $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256
                [void]$hashRows.Add([pscustomobject]@{RelativePath=$relative;FullPath=$file.FullName;SizeBytes=[int64]$file.Length;SHA256=$hash.Hash;Status='PASS'})
            } catch {
                [void]$hashRows.Add([pscustomobject]@{RelativePath=$relative;FullPath=$file.FullName;SizeBytes=[int64]$file.Length;SHA256='';Status='ERROR'})
                [void]$errors.Add([pscustomobject]@{Path=$file.FullName;Stage='SHA256';Error=$_.Exception.Message})
            }
        }
    }
    Write-Progress -Id 2 -Activity 'Inventorying files and SHA256' -Completed

    $fileRows | Sort-Object RelativePath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Files.csv') -NoTypeInformation -Encoding UTF8
    if (-not $SkipHashes) { $hashRows | Sort-Object RelativePath | Export-Csv -LiteralPath (Join-Path $RunFolder 'Files_SHA256.csv') -NoTypeInformation -Encoding UTF8 }

    Write-Progress -Id 3 -Activity 'Building grouped inventories' -Status 'Grouping by extension' -PercentComplete 20
    [object[]]$extensionRows = @(
        $fileRows | Group-Object Extension | ForEach-Object {
            $groupItems = @($_.Group)
            $bytes = ($groupItems | Measure-Object SizeBytes -Sum).Sum
            if ($null -eq $bytes) { $bytes = 0 }
            [pscustomobject]@{Extension=$_.Name;FileCount=$groupItems.Count;TotalBytes=[int64]$bytes;TotalMB=[math]::Round($bytes/1MB,3);TotalGB=[math]::Round($bytes/1GB,6)}
        } | Sort-Object -Property @{Expression='FileCount';Descending=$true},@{Expression='Extension';Descending=$false}
    )
    $extensionRows | Export-Csv -LiteralPath (Join-Path $RunFolder 'Inventory_By_Extension.csv') -NoTypeInformation -Encoding UTF8

    Write-Progress -Id 3 -Activity 'Building grouped inventories' -Status 'Grouping by folder' -PercentComplete 55
    [object[]]$folderGroupRows = @(
        $fileRows | Group-Object RelativeFolder | ForEach-Object {
            $groupItems = @($_.Group)
            $bytes = ($groupItems | Measure-Object SizeBytes -Sum).Sum
            if ($null -eq $bytes) { $bytes = 0 }
            $uniqueExtensions = @($groupItems | Select-Object -ExpandProperty Extension -Unique)
            [pscustomobject]@{RelativeFolder=$_.Name;FileCount=$groupItems.Count;TotalBytes=[int64]$bytes;TotalMB=[math]::Round($bytes/1MB,3);TotalGB=[math]::Round($bytes/1GB,6);ExtensionCount=$uniqueExtensions.Count}
        } | Sort-Object RelativeFolder
    )
    $folderGroupRows | Export-Csv -LiteralPath (Join-Path $RunFolder 'Inventory_By_Folder.csv') -NoTypeInformation -Encoding UTF8

    Write-Progress -Id 3 -Activity 'Building grouped inventories' -Status 'Building tree and summary' -PercentComplete 80
    $tree = New-Object System.Collections.ArrayList
    [void]$tree.Add($RepositoryRoot)
    foreach ($row in ($folderRows | Sort-Object RelativePath)) { [void]$tree.Add('[DIR]  ' + $row.RelativePath) }
    foreach ($row in ($fileRows | Sort-Object RelativePath)) { [void]$tree.Add('[FILE] ' + $row.RelativePath) }
    $tree | Set-Content -LiteralPath (Join-Path $RunFolder 'Repository_Tree.txt') -Encoding UTF8

    if ($errors.Count -gt 0) { $errors | Export-Csv -LiteralPath (Join-Path $RunFolder 'Scan_Errors.csv') -NoTypeInformation -Encoding UTF8 }
    else { [pscustomobject]@{Path='';Stage='';Error='No scan errors.'} | Export-Csv -LiteralPath (Join-Path $RunFolder 'Scan_Errors.csv') -NoTypeInformation -Encoding UTF8 }

    $totalBytes = ($fileRows | Measure-Object SizeBytes -Sum).Sum
    if ($null -eq $totalBytes) { $totalBytes = 0 }
    $summary = [ordered]@{
        Title='BHP INVENT | CP09A Repository Inventory';Classification=$Classification;Version=$Version;RunId=$RunId
        RepositoryRoot=$RepositoryRoot;ApplicationRoot=$ApplicationRoot;InstalledScript=$InstalledPath;OutputFolder=$RunFolder
        GeneratedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');FolderCount=$folderRows.Count;FileCount=$fileRows.Count
        TotalBytes=[int64]$totalBytes;TotalGB=[math]::Round($totalBytes/1GB,6);ExtensionGroups=$extensionRows.Count
        FolderGroups=$folderGroupRows.Count;HashesGenerated=(-not $SkipHashes);ScanErrorCount=$errors.Count
        Status=if ($errors.Count -eq 0) {'PASS'} else {'PASS_WITH_WARNINGS'}
    }
    $summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RunFolder 'Summary.txt') -Encoding UTF8

    $manifest = foreach ($outputFile in @(Get-ChildItem -LiteralPath $RunFolder -File | Where-Object Name -ne 'SHA256SUMS.txt')) {
        $outputHash = Get-FileHash -LiteralPath $outputFile.FullName -Algorithm SHA256
        '{0} *{1}' -f $outputHash.Hash,$outputFile.Name
    }
    $manifest | Set-Content -LiteralPath (Join-Path $RunFolder 'SHA256SUMS.txt') -Encoding ASCII
    Write-Progress -Id 3 -Activity 'Building grouped inventories' -Completed

    Write-Log "INVENTORY: $($summary.Status)" 'PASS'
    Write-Log "OUTPUT: $RunFolder" 'PASS'

    if (-not ((Get-FullPath $SourcePath).Equals((Get-FullPath $InstalledPath),[System.StringComparison]::OrdinalIgnoreCase))) {
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
        @(
            'MESSAGE: ' + $_.Exception.Message
            'SCRIPT: ' + $_.InvocationInfo.ScriptName
            'LINE: ' + $_.InvocationInfo.ScriptLineNumber
            'POSITION: ' + $_.InvocationInfo.PositionMessage
            ''
            ($_ | Out-String)
        ) | Set-Content -LiteralPath (Join-Path $RunFolder 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: ' + $_.Exception.Message) -ForegroundColor Red }
    exit 1
}
