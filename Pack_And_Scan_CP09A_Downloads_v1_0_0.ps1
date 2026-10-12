#Requires -Version 7.4
<#
BHP Internal
Pack_And_Scan_CP09A_Downloads v1.0.0
Part 1: compresses the recovery folder (including Recovery_Register.csv) into a new ZIP and verifies its entries.
Part 2: read-only scan of Downloads for CP09A-related files, grouped by type, with SHA256 and duplicate flag.
No source file is modified, moved, renamed or deleted. Existing ZIP or CSV outputs are never overwritten.
#>
[CmdletBinding()]
param(
    [Parameter()][string]$DownloadsRoot = (Join-Path $env:USERPROFILE 'Downloads'),
    [Parameter()][string]$RecoveryFolderName = 'CP09A-SESSION-RECOVERY-20261003-032432',
    [Parameter()][switch]$SkipPack,
    [Parameter()][switch]$ExportInventory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$Pattern = '(?i)(CP09A|GeoAI|Spence|AutoLith|LEP|Vendor|Composite|Cierre|Recovery|Programming_Delivery|PowerShell_.+Standard)'

function Get-TypeLabel {
    param([Parameter(Mandatory)][string]$Extension)
    switch ($Extension.ToLowerInvariant()) {
        '.ps1'  { return 'SCRIPT' }
        '.docx' { return 'DOCUMENT' }
        '.pdf'  { return 'DOCUMENT' }
        '.md'   { return 'DOCUMENT' }
        '.xlsx' { return 'SPREADSHEET' }
        '.zip'  { return 'ZIP' }
        '.csv'  { return 'DATA' }
        '.json' { return 'DATA' }
        default { return 'OTHER' }
    }
}

try {
    if (-not (Test-Path -LiteralPath $DownloadsRoot -PathType Container)) { throw ('Downloads not found: ' + $DownloadsRoot) }

    Write-Host ''
    Write-Host 'BHP Internal | Pack_And_Scan_CP09A_Downloads v1.0.0' -ForegroundColor Cyan
    Write-Host ('ROOT: ' + $DownloadsRoot)

    # ---------- PART 1: PACK RECOVERY FOLDER ----------
    if (-not $SkipPack) {
        Write-Host ''
        Write-Host '== PART 1: PACK RECOVERY FOLDER ==' -ForegroundColor Cyan
        $Run = Join-Path $DownloadsRoot $RecoveryFolderName
        $Register = Join-Path $Run 'Recovery_Register.csv'
        $Zip = Join-Path $DownloadsRoot ($RecoveryFolderName + '_' + $Stamp + '_Internal.zip')
        if (-not (Test-Path -LiteralPath $Run -PathType Container)) { throw ('Recovery folder not found: ' + $Run) }
        if (-not (Test-Path -LiteralPath $Register -PathType Leaf)) { throw ('Recovery_Register.csv not found: ' + $Register) }
        if (Test-Path -LiteralPath $Zip) { throw ('ZIP already exists: ' + $Zip) }
        $Files = @(Get-ChildItem -LiteralPath $Run -File | ForEach-Object { $_.FullName })
        if ($Files -notcontains $Register) { throw 'Recovery_Register.csv is not in the file list.' }
        Compress-Archive -LiteralPath $Files -DestinationPath $Zip
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $Archive = [System.IO.Compression.ZipFile]::OpenRead($Zip)
        try { $Entries = @($Archive.Entries | ForEach-Object { $_.Name }) } finally { $Archive.Dispose() }
        if ($Entries -notcontains 'Recovery_Register.csv') { throw 'Recovery_Register.csv missing inside ZIP.' }
        if ($Entries.Count -ne $Files.Count) { throw ('ZIP entries ' + $Entries.Count + ' differ from files ' + $Files.Count) }
        Import-Csv -LiteralPath $Register | Format-Table File, Status -AutoSize
        Write-Host ('ZIP ENTRIES: ' + $Entries.Count + ' (includes Recovery_Register.csv)') -ForegroundColor Green
        Write-Host ('ZIP: ' + $Zip) -ForegroundColor Green
        Write-Host ('ZIP SHA256: ' + (Get-FileHash -LiteralPath $Zip -Algorithm SHA256).Hash) -ForegroundColor Green
    }
    else {
        Write-Host 'PART 1 SKIPPED (SkipPack).' -ForegroundColor Yellow
    }

    # ---------- PART 2: READ-ONLY SCAN ----------
    Write-Host ''
    Write-Host '== PART 2: SCAN DOWNLOADS (READ-ONLY) ==' -ForegroundColor Cyan
    $Found = @(Get-ChildItem -LiteralPath $DownloadsRoot -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $Pattern })
    if ($Found.Count -eq 0) {
        Write-Host 'NO RELATED FILES FOUND' -ForegroundColor Yellow
        return
    }
    $Rows = [System.Collections.Generic.List[object]]::new()
    $Index = 0
    foreach ($File in $Found) {
        $Index++
        $Percent = [int](($Index / $Found.Count) / 0.01)
        if ($Percent -gt 100) { $Percent = 100 }
        Write-Progress -Activity 'Scanning Downloads' -Status $File.Name -PercentComplete $Percent
        $Rows.Add([pscustomobject]@{
            Type      = (Get-TypeLabel -Extension $File.Extension)
            Name      = $File.Name
            Folder    = $File.DirectoryName.Replace($DownloadsRoot, 'Downloads')
            SizeKB    = [int]($File.Length / 1024)
            LastWrite = $File.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
            SHA256    = (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256).Hash
            Duplicate = 'N'
        })
    }
    Write-Progress -Activity 'Scanning Downloads' -Completed

    $DupGroups = @($Rows | Group-Object SHA256 | Where-Object { $_.Count -gt 1 })
    $DupHashes = @($DupGroups | ForEach-Object { $_.Name })
    foreach ($Row in $Rows) { if ($DupHashes -contains $Row.SHA256) { $Row.Duplicate = 'Y' } }

    foreach ($Group in @($Rows | Group-Object Type | Sort-Object Name)) {
        Write-Host ''
        Write-Host ('-- ' + $Group.Name + ' (' + $Group.Count + ') --') -ForegroundColor Cyan
        $Group.Group | Sort-Object Name | Format-Table Name, Folder, SizeKB, LastWrite, Duplicate -AutoSize
    }
    Write-Host ('TOTAL FILES: ' + $Rows.Count + ' | DUPLICATE GROUPS: ' + $DupGroups.Count) -ForegroundColor Green

    if ($ExportInventory) {
        $Out = Join-Path $DownloadsRoot ('CP09A_Downloads_Inventory_' + $Stamp + '_Internal.csv')
        if (Test-Path -LiteralPath $Out) { throw ('Output exists: ' + $Out) }
        $Rows | Export-Csv -LiteralPath $Out -NoTypeInformation -Encoding utf8NoBOM
        Write-Host ('INVENTORY: ' + $Out) -ForegroundColor Green
    }
    else {
        Write-Host 'SCAN READ-ONLY: NO INVENTORY FILE CREATED. Use -ExportInventory to save it.' -ForegroundColor Green
    }
    Write-Host 'DECISION: PASS' -ForegroundColor Green
}
catch {
    Write-Progress -Activity 'Scanning Downloads' -Completed
    Write-Host ('FAIL: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
