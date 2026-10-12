#Requires -Version 7.4
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ToolRelativePath = '00_Script\Application_Governance',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AuditRelativePath = '_Audit\Application_Governance',

    [Parameter()]
    [switch]$AuditOnly,

    [Parameter()]
    [switch]$InstallOnly,

    [Parameter()]
    [switch]$IncludeFileHashes,

    [Parameter()]
    [switch]$ContinueOnNonCriticalError
)

Clear-Host
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$EngineName = 'Application_Governance_Engine'
$EngineVersion = '1.1.0'
$Classification = 'BHP Internal'
$RunId = 'APP-GOV-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$StartedAt = Get-Date
$NonCriticalErrors = [System.Collections.Generic.List[object]]::new()

function Assert-TextValue {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value,
        [Parameter(Mandatory)]
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw ('{0} cannot be empty.' -f $Name)
    }
}

function Get-NormalizedFullPath {
    param(
        [Parameter(Mandatory)]
        [string]$PathValue
    )
    Assert-TextValue -Value $PathValue -Name 'PathValue'
    return [System.IO.Path]::GetFullPath($PathValue)
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory)]
        [string]$BasePath,
        [Parameter(Mandatory)]
        [string]$FullPath
    )
    return [System.IO.Path]::GetRelativePath(
        (Get-NormalizedFullPath -PathValue $BasePath),
        (Get-NormalizedFullPath -PathValue $FullPath)
    )
}

function Test-IsBelowPath {
    param(
        [Parameter(Mandatory)]
        [string]$CandidatePath,
        [Parameter(Mandatory)]
        [string]$ParentPath
    )
    $candidate = (Get-NormalizedFullPath -PathValue $CandidatePath).TrimEnd('\')
    $parent = (Get-NormalizedFullPath -PathValue $ParentPath).TrimEnd('\')
    return $candidate.StartsWith($parent + '\',[System.StringComparison]::OrdinalIgnoreCase) -or ($candidate -ieq $parent)
}

function Get-TextSafely {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )
    try {
        return [System.IO.File]::ReadAllText($File.FullName)
    }
    catch {
        if ($ContinueOnNonCriticalError) {
            $NonCriticalErrors.Add([pscustomobject]@{
                Stage = 'READ_TEXT'
                Path = $File.FullName
                Message = $_.Exception.Message
            })
            return $null
        }
        throw
    }
}

function Get-HashSafely {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )
    try {
        return (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
    }
    catch {
        if ($ContinueOnNonCriticalError) {
            $NonCriticalErrors.Add([pscustomobject]@{
                Stage = 'HASH'
                Path = $File.FullName
                Message = $_.Exception.Message
            })
            return $null
        }
        throw
    }
}

function Test-ScriptSyntax {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath
    )
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath,
        [ref]$tokens,
        [ref]$errors
    )
    $parserErrors = @($errors)
    if ($parserErrors.Count -gt 0) {
        $parserErrors | Select-Object Message,ErrorId,Extent | Format-List
        throw 'SYNTAX FAIL'
    }
    $content = [System.IO.File]::ReadAllText($ScriptPath)
    $forbiddenPatterns = @(
        ('return' + '['),
        ('Write-Host' + [char]34),
        ('Write-Warning' + [char]34),
        ('Write-Error' + [char]34),
        ('throw' + [char]34)
    )
    foreach ($pattern in $forbiddenPatterns) {
        if ($content.Contains($pattern,[System.StringComparison]::Ordinal)) {
            throw ('FORBIDDEN SYNTAX PATTERN: {0}' -f $pattern)
        }
    }
    return $true
}

function Get-ApplicationFiles {
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,
        [Parameter(Mandatory)]
        [string]$ToolPath,
        [Parameter(Mandatory)]
        [string]$AuditPath
    )
    return @(
        Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction Stop |
        Where-Object {
            -not (Test-IsBelowPath -CandidatePath $_.FullName -ParentPath $ToolPath) -and
            -not (Test-IsBelowPath -CandidatePath $_.FullName -ParentPath $AuditPath)
        } |
        Sort-Object -Property FullName
    )
}

function Get-SourceState {
    param(
        [Parameter(Mandatory)]
        [array]$Files
    )
    return @(
        foreach ($file in @($Files)) {
            [pscustomobject]@{
                FullName = $file.FullName
                Length = $file.Length
                LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
            }
        }
    )
}

function Get-Disposition {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,
        [Parameter(Mandatory)]
        [bool]$Referenced,
        [Parameter(Mandatory)]
        [bool]$ManifestListed,
        [Parameter(Mandatory)]
        [bool]$HashListed,
        [Parameter(Mandatory)]
        [bool]$ExactDuplicate,
        [Parameter(Mandatory)]
        [bool]$ExecutableOrCode
    )

    $pathText = $File.FullName.ToLowerInvariant()
    $nameText = $File.Name.ToLowerInvariant()
    $extension = $File.Extension.ToLowerInvariant()

    $archiveTokens = @('backup','archive','legacy','deprecated','previous','superseded','old_version')
    $reviewTokens = @('copy','duplicate','draft','sandbox','test','temp','tmp','bak')
    $removeExtensions = @('.tmp','.temp','.cache')

    if ($Referenced -or $ManifestListed -or $HashListed -or $ExecutableOrCode) {
        return 'KEEP'
    }
    if ($ExactDuplicate) {
        return 'REVIEW'
    }
    foreach ($token in $archiveTokens) {
        if ($pathText.Contains($token,[System.StringComparison]::OrdinalIgnoreCase)) {
            return 'ARCHIVE'
        }
    }
    foreach ($token in $reviewTokens) {
        if ($nameText.Contains($token,[System.StringComparison]::OrdinalIgnoreCase)) {
            return 'REVIEW'
        }
    }
    if ($removeExtensions -contains $extension) {
        return 'REMOVE_CANDIDATE'
    }
    return 'REVIEW'
}

function Write-ReportSet {
    param(
        [Parameter(Mandatory)]
        [string]$ReportRoot,
        [Parameter(Mandatory)]
        [array]$FileInventory,
        [Parameter(Mandatory)]
        [array]$FolderInventory,
        [Parameter(Mandatory)]
        [array]$ExecutableInventory,
        [Parameter(Mandatory)]
        [array]$DependencyInventory,
        [Parameter(Mandatory)]
        [array]$ReferencedInventory,
        [Parameter(Mandatory)]
        [array]$ManifestCoverage,
        [Parameter(Mandatory)]
        [array]$HashCoverage,
        [Parameter(Mandatory)]
        [array]$DuplicateInventory,
        [Parameter(Mandatory)]
        [array]$GovernanceDisposition,
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Summary
    )

    $reportsRoot = Join-Path -Path $ReportRoot -ChildPath 'reports'
    $logsRoot = Join-Path -Path $ReportRoot -ChildPath 'logs'
    $signatureRoot = Join-Path -Path $ReportRoot -ChildPath 'signature'

    foreach ($folder in @($ReportRoot,$reportsRoot,$logsRoot,$signatureRoot)) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            New-Item -ItemType Directory -Path $folder -ErrorAction Stop | Out-Null
        }
    }

    $FileInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'File_Inventory.csv') -NoTypeInformation -Encoding utf8
    $FolderInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Folder_Inventory.csv') -NoTypeInformation -Encoding utf8
    $ExecutableInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Executable_Inventory.csv') -NoTypeInformation -Encoding utf8
    $DependencyInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Dependency_Inventory.csv') -NoTypeInformation -Encoding utf8
    $ReferencedInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Referenced_Files.csv') -NoTypeInformation -Encoding utf8
    $ManifestCoverage | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Manifest_Coverage.csv') -NoTypeInformation -Encoding utf8
    $HashCoverage | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Hash_Coverage.csv') -NoTypeInformation -Encoding utf8
    $DuplicateInventory | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Exact_Duplicates.csv') -NoTypeInformation -Encoding utf8
    $GovernanceDisposition | Export-Csv -LiteralPath (Join-Path $reportsRoot 'Governance_Disposition.csv') -NoTypeInformation -Encoding utf8
    @($NonCriticalErrors) | Export-Csv -LiteralPath (Join-Path $logsRoot 'NonCritical_Errors.csv') -NoTypeInformation -Encoding utf8
    $Summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $ReportRoot 'Summary.json') -Encoding utf8

    $readme = @(
        '# Application Governance Engine'
        ''
        ('Classification: {0}' -f $Classification)
        ('Version: {0}' -f $EngineVersion)
        ('RUN_ID: {0}' -f $RunId)
        ('Application root: {0}' -f $ApplicationRoot)
        ''
        'This run inventories application files, code, dependencies, references, manifests, SHA256 coverage and exact duplicates.'
        'No application file is deleted, moved, renamed or overwritten.'
        'REMOVE_CANDIDATE is a review status only and never authorizes deletion.'
    )
    $readme | Set-Content -LiteralPath (Join-Path $ReportRoot 'README.md') -Encoding utf8

    $manifestEntries = @(
        Get-ChildItem -LiteralPath $ReportRoot -File -Recurse -Force -ErrorAction Stop |
        Where-Object { $_.Name -ne 'Manifest.json' -and $_.Name -ne 'SHA256SUMS.txt' } |
        Sort-Object -Property FullName |
        ForEach-Object {
            [pscustomobject]@{
                RelativePath = Get-RelativePathSafe -BasePath $ReportRoot -FullPath $_.FullName
                LengthBytes = $_.Length
                SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
            }
        }
    )
    [ordered]@{
        Application = $EngineName
        Version = $EngineVersion
        Classification = $Classification
        RunId = $RunId
        Files = $manifestEntries
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $ReportRoot 'Manifest.json') -Encoding utf8

    $filesToSign = @(
        Get-ChildItem -LiteralPath $ReportRoot -File -Recurse -Force -ErrorAction Stop |
        Where-Object { $_.Name -ne 'SHA256SUMS.txt' } |
        Sort-Object -Property FullName
    )
    $checksumLines = @(
        foreach ($file in $filesToSign) {
            $hashValue = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
            $relative = Get-RelativePathSafe -BasePath $ReportRoot -FullPath $file.FullName
            '{0}  {1}' -f $hashValue,$relative
        }
    )
    $checksumLines | Set-Content -LiteralPath (Join-Path $signatureRoot 'SHA256SUMS.txt') -Encoding ascii
}

try {
    Assert-TextValue -Value $ApplicationRoot -Name 'ApplicationRoot'
    Assert-TextValue -Value $ToolRelativePath -Name 'ToolRelativePath'
    Assert-TextValue -Value $AuditRelativePath -Name 'AuditRelativePath'

    $ApplicationRoot = Get-NormalizedFullPath -PathValue $ApplicationRoot
    if (-not (Test-Path -LiteralPath $ApplicationRoot -PathType Container)) {
        throw ('Application root not found: {0}' -f $ApplicationRoot)
    }

    $ToolRoot = Join-Path -Path $ApplicationRoot -ChildPath $ToolRelativePath
    $AuditBaseRoot = Join-Path -Path $ApplicationRoot -ChildPath $AuditRelativePath
    $RunRoot = Join-Path -Path $AuditBaseRoot -ChildPath $RunId
    $InstalledScriptPath = Join-Path -Path $ToolRoot -ChildPath 'Application_Governance_Engine_v1_1_0.ps1'
    $CurrentScriptPath = $MyInvocation.MyCommand.Path

    if ([string]::IsNullOrWhiteSpace($CurrentScriptPath)) {
        throw 'Current script path is unavailable.'
    }

    [void](Test-ScriptSyntax -ScriptPath $CurrentScriptPath)

    Write-Host ''
    Write-Host ('{0} | {1} v{2}' -f $Classification,$EngineName,$EngineVersion) -ForegroundColor Cyan
    Write-Host ('APPLICATION ROOT: {0}' -f $ApplicationRoot)
    Write-Host ('PROPOSED TOOL ROOT: {0}' -f $ToolRoot)
    Write-Host ('PROPOSED AUDIT ROOT: {0}' -f $RunRoot)

    $Files = @(Get-ApplicationFiles -RootPath $ApplicationRoot -ToolPath $ToolRoot -AuditPath $AuditBaseRoot)
    $Folders = @(
        Get-ChildItem -LiteralPath $ApplicationRoot -Directory -Recurse -Force -ErrorAction Stop |
        Where-Object {
            -not (Test-IsBelowPath -CandidatePath $_.FullName -ParentPath $ToolRoot) -and
            -not (Test-IsBelowPath -CandidatePath $_.FullName -ParentPath $AuditBaseRoot)
        } |
        Sort-Object -Property FullName
    )
    $SourceStateBefore = @(Get-SourceState -Files $Files)

    $CodeExtensions = @('.ps1','.psm1','.psd1','.py','.r','.bat','.cmd')
    $ConfigExtensions = @('.json','.yaml','.yml','.toml','.config','.xml')
    $ExecutableExtensions = @('.exe','.com','.bat','.cmd','.ps1','.psm1','.py','.r','.jar','.dll')
    $AnalysisExtensions = @($CodeExtensions + $ConfigExtensions | Sort-Object -Unique)

    $CodeAndConfigFiles = @($Files | Where-Object { $AnalysisExtensions -contains $_.Extension.ToLowerInvariant() })
    $ExecutableFiles = @($Files | Where-Object { $ExecutableExtensions -contains $_.Extension.ToLowerInvariant() })
    $ManifestFiles = @($Files | Where-Object {
        $lower = $_.Name.ToLowerInvariant()
        $lower.Contains('manifest') -or $lower.Contains('release_register') -or $lower -eq 'release_version.json'
    })
    $HashRegisterFiles = @($Files | Where-Object {
        $lower = $_.Name.ToLowerInvariant()
        $lower.Contains('sha256') -or $lower.Contains('checksum')
    })

    Write-Progress -Id 1 -Activity 'Application governance' -Status 'Hashing application files' -PercentComplete 15
    $HashRecords = [System.Collections.Generic.List[object]]::new()
    $fileNumber = 0
    foreach ($file in $Files) {
        $fileNumber++
        $percent = 15 + [int](25.0 * $fileNumber / [math]::Max(1,$Files.Count))
        Write-Progress -Id 1 -Activity 'Application governance' -Status ('Hashing {0} of {1}' -f $fileNumber,$Files.Count) -PercentComplete $percent
        $HashRecords.Add([pscustomobject]@{
            File = $file
            RelativePath = Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $file.FullName
            SHA256 = Get-HashSafely -File $file
        })
    }

    Write-Progress -Id 1 -Activity 'Application governance' -Status 'Analyzing dependencies and references' -PercentComplete 45
    $DependencyInventory = [System.Collections.Generic.List[object]]::new()
    $ReferencedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ReferencedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $referencePatterns = @(
        [pscustomobject]@{ Type='PowerShellModule'; Pattern='(?im)^\s{0,}Import-Module\s+[''"]?([^''"\r\n]+)' },
        [pscustomobject]@{ Type='PowerShellUsingModule'; Pattern='(?im)^\s{0,}using\s+module\s+[''"]?([^''"\r\n]+)' },
        [pscustomobject]@{ Type='PowerShellDotSource'; Pattern='(?im)^\s{0,}\.\s+[''"]([^''"]+\.ps1)[''"]' },
        [pscustomobject]@{ Type='PythonImport'; Pattern='(?im)^\s{0,}(?:from|import)\s+([A-Za-z0-9_\.]+)' },
        [pscustomobject]@{ Type='RDependency'; Pattern='(?im)^\s{0,}(?:library|require|source)\s{0,}\(\s{0,}[''"]?([^''"\)]+)' },
        [pscustomobject]@{ Type='QuotedFileReference'; Pattern='(?i)[''"]([^''"]+\.(?:ps1|psm1|psd1|py|r|json|yaml|yml|toml|config|xml|csv|parquet|feather|las|exe|dll|bat|cmd))[''"]' }
    )

    foreach ($sourceFile in $CodeAndConfigFiles) {
        $content = Get-TextSafely -File $sourceFile
        if ($null -eq $content) {
            continue
        }
        foreach ($definition in $referencePatterns) {
            foreach ($match in [regex]::Matches($content,$definition.Pattern)) {
                $referenceValue = $match.Groups[1].Value.Trim()
                if ([string]::IsNullOrWhiteSpace($referenceValue)) {
                    continue
                }
                $normalizedReference = $referenceValue.Replace('/','\')
                $referenceName = [System.IO.Path]::GetFileName($normalizedReference)
                if (-not [string]::IsNullOrWhiteSpace($referenceName)) {
                    [void]$ReferencedNames.Add($referenceName)
                }
                if (-not [System.IO.Path]::IsPathRooted($normalizedReference)) {
                    [void]$ReferencedPaths.Add($normalizedReference.TrimStart('.','\'))
                }
                $matchingFiles = @($Files | Where-Object {
                    $_.Name -ieq $referenceName -or
                    (Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $_.FullName) -ieq $normalizedReference
                })
                $DependencyInventory.Add([pscustomobject]@{
                    SourceFile = Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $sourceFile.FullName
                    DependencyType = $definition.Type
                    Reference = $referenceValue
                    ReferenceName = $referenceName
                    MatchCount = $matchingFiles.Count
                    ExistsInApplication = ($matchingFiles.Count -gt 0)
                    MatchedPaths = (@($matchingFiles | ForEach-Object { Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $_.FullName }) -join '; ')
                })
            }
        }
    }

    $ManifestListedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ManifestListedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($manifest in $ManifestFiles) {
        $content = Get-TextSafely -File $manifest
        if ($null -eq $content) {
            continue
        }
        foreach ($file in $Files) {
            $relative = Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $file.FullName
            if ($content.IndexOf($file.Name,[System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                [void]$ManifestListedNames.Add($file.Name)
            }
            if ($content.IndexOf($relative,[System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                [void]$ManifestListedPaths.Add($relative)
            }
        }
    }

    $RegisteredHashNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $RegisteredHashPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $RegisteredHashes = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($hashFile in $HashRegisterFiles) {
        $content = Get-TextSafely -File $hashFile
        if ($null -eq $content) {
            continue
        }
        foreach ($line in ($content -split "`r?`n")) {
            if ($line -match '^\s{0,}([A-Fa-f0-9]{64})\s+[^A-Za-z0-9]?(.+?)\s{0,}$') {
                $hashValue = $Matches[1].ToUpperInvariant()
                $registeredPath = $Matches[2].Trim().Replace('/','\')
                $registeredName = [System.IO.Path]::GetFileName($registeredPath)
                if (-not [string]::IsNullOrWhiteSpace($registeredName)) {
                    [void]$RegisteredHashNames.Add($registeredName)
                }
                [void]$RegisteredHashPaths.Add($registeredPath)
                if (-not $RegisteredHashes.ContainsKey($registeredPath)) {
                    $RegisteredHashes.Add($registeredPath,$hashValue)
                }
            }
        }
    }

    $DuplicateGroups = @(
        $HashRecords |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.SHA256) } |
        Group-Object -Property SHA256 |
        Where-Object { @($_.Group).Count -gt 1 }
    )
    $DuplicatePaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $DuplicateInventory = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $DuplicateGroups) {
        foreach ($item in @($group.Group)) {
            [void]$DuplicatePaths.Add($item.RelativePath)
            $DuplicateInventory.Add([pscustomobject]@{
                SHA256 = $group.Name
                GroupSize = @($group.Group).Count
                RelativePath = $item.RelativePath
                LengthBytes = $item.File.Length
            })
        }
    }

    $ManifestCoverage = [System.Collections.Generic.List[object]]::new()
    $HashCoverage = [System.Collections.Generic.List[object]]::new()
    $GovernanceDisposition = [System.Collections.Generic.List[object]]::new()
    $ReferencedInventory = [System.Collections.Generic.List[object]]::new()
    $FileInventory = [System.Collections.Generic.List[object]]::new()

    foreach ($record in $HashRecords) {
        $file = $record.File
        $relative = $record.RelativePath
        $referenced = $ReferencedNames.Contains($file.Name) -or $ReferencedPaths.Contains($relative)
        $manifestListed = $ManifestListedNames.Contains($file.Name) -or $ManifestListedPaths.Contains($relative)
        $hashListed = $RegisteredHashNames.Contains($file.Name) -or $RegisteredHashPaths.Contains($relative)
        $exactDuplicate = $DuplicatePaths.Contains($relative)
        $executableOrCode = $ExecutableExtensions -contains $file.Extension.ToLowerInvariant()
        $disposition = Get-Disposition -File $file -Referenced $referenced -ManifestListed $manifestListed -HashListed $hashListed -ExactDuplicate $exactDuplicate -ExecutableOrCode $executableOrCode

        $FileInventory.Add([pscustomobject]@{
            RelativePath = $relative
            Name = $file.Name
            Extension = $file.Extension
            LengthBytes = $file.Length
            LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
            ExecutableOrCode = $executableOrCode
            Referenced = $referenced
            ManifestListed = $manifestListed
            HashRegisterListed = $hashListed
            ExactDuplicate = $exactDuplicate
            SHA256 = if ($IncludeFileHashes) { $record.SHA256 } else { $null }
        })
        $ManifestCoverage.Add([pscustomobject]@{
            RelativePath = $relative
            ExistsOnDisk = $true
            ListedInManifest = $manifestListed
            Status = if ($manifestListed) { 'COVERED' } else { 'NOT_LISTED' }
        })
        $HashCoverage.Add([pscustomobject]@{
            RelativePath = $relative
            ExistsOnDisk = $true
            ListedInHashRegister = $hashListed
            CalculatedSHA256 = $record.SHA256
            Status = if ($hashListed) { 'REGISTERED' } else { 'NOT_REGISTERED' }
        })
        if ($referenced) {
            $ReferencedInventory.Add([pscustomobject]@{
                RelativePath = $relative
                Name = $file.Name
                ExecutableOrCode = $executableOrCode
                ManifestListed = $manifestListed
                HashRegisterListed = $hashListed
            })
        }
        $GovernanceDisposition.Add([pscustomobject]@{
            Disposition = $disposition
            RelativePath = $relative
            Reason = if ($referenced) { 'Referenced by code or configuration' }
                     elseif ($manifestListed) { 'Listed in manifest or release record' }
                     elseif ($hashListed) { 'Listed in SHA256 register' }
                     elseif ($executableOrCode) { 'Executable or source code' }
                     elseif ($exactDuplicate) { 'Exact SHA256 duplicate requires review' }
                     else { 'No verified dependency evidence; manual review required' }
            Referenced = $referenced
            ManifestListed = $manifestListed
            HashRegisterListed = $hashListed
            ExactDuplicate = $exactDuplicate
        })
    }

    $FolderInventory = @(
        foreach ($folder in $Folders) {
            [pscustomobject]@{
                RelativePath = Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $folder.FullName
                LastWriteTimeUtc = $folder.LastWriteTimeUtc.ToString('o')
            }
        }
    )
    $ExecutableInventory = @(
        foreach ($file in $ExecutableFiles) {
            $relative = Get-RelativePathSafe -BasePath $ApplicationRoot -FullPath $file.FullName
            [pscustomobject]@{
                RelativePath = $relative
                Extension = $file.Extension
                Referenced = $ReferencedNames.Contains($file.Name) -or $ReferencedPaths.Contains($relative)
                ManifestListed = $ManifestListedNames.Contains($file.Name) -or $ManifestListedPaths.Contains($relative)
                HashRegisterListed = $RegisteredHashNames.Contains($file.Name) -or $RegisteredHashPaths.Contains($relative)
            }
        }
    )

    $DispositionSummary = @(
        $GovernanceDisposition |
        Group-Object -Property Disposition |
        Sort-Object -Property Name |
        ForEach-Object {
            [pscustomobject]@{
                Disposition = $_.Name
                Count = @($_.Group).Count
            }
        }
    )

    $Summary = [ordered]@{
        Application = $EngineName
        Version = $EngineVersion
        Classification = $Classification
        RunId = $RunId
        ApplicationRoot = $ApplicationRoot
        Files = $Files.Count
        Folders = $Folders.Count
        CodeAndConfigFiles = $CodeAndConfigFiles.Count
        ExecutablesOrCode = $ExecutableFiles.Count
        DependencyReferences = $DependencyInventory.Count
        ReferencedFiles = $ReferencedInventory.Count
        ManifestFiles = $ManifestFiles.Count
        HashRegisters = $HashRegisterFiles.Count
        ExactDuplicateGroups = $DuplicateGroups.Count
        NonCriticalErrors = $NonCriticalErrors.Count
        DispositionSummary = $DispositionSummary
        AuditOnly = [bool]$AuditOnly
        InstalledToolPath = $InstalledScriptPath
        ProposedReportRoot = $RunRoot
        StartedAt = $StartedAt.ToString('o')
        CompletedAt = (Get-Date).ToString('o')
    }

    if ($AuditOnly) {
        Write-Host ('FILES: {0} | FOLDERS: {1}' -f $Files.Count,$Folders.Count)
        Write-Host ('CODE AND CONFIG: {0} | EXECUTABLES OR CODE: {1}' -f $CodeAndConfigFiles.Count,$ExecutableFiles.Count)
        Write-Host ('DEPENDENCY REFERENCES: {0} | REFERENCED FILES: {1}' -f $DependencyInventory.Count,$ReferencedInventory.Count)
        Write-Host ('MANIFESTS: {0} | HASH REGISTERS: {1}' -f $ManifestFiles.Count,$HashRegisterFiles.Count)
        Write-Host ('EXACT DUPLICATE GROUPS: {0}' -f $DuplicateGroups.Count)
        $DispositionSummary | Format-Table -AutoSize
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED.' -ForegroundColor Green
        Write-Progress -Id 1 -Activity 'Application governance' -Completed
        exit 0
    }

    if (-not (Test-Path -LiteralPath $ToolRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $ToolRoot -ErrorAction Stop | Out-Null
    }
    if ((Get-NormalizedFullPath -PathValue $CurrentScriptPath) -ine (Get-NormalizedFullPath -PathValue $InstalledScriptPath)) {
        if (Test-Path -LiteralPath $InstalledScriptPath -PathType Leaf) {
            $sourceHash = (Get-FileHash -LiteralPath $CurrentScriptPath -Algorithm SHA256).Hash.ToUpperInvariant()
            $installedHash = (Get-FileHash -LiteralPath $InstalledScriptPath -Algorithm SHA256).Hash.ToUpperInvariant()
            if ($sourceHash -ne $installedHash) {
                throw ('Installed tool already exists with different content: {0}' -f $InstalledScriptPath)
            }
        }
        else {
            Copy-Item -LiteralPath $CurrentScriptPath -Destination $InstalledScriptPath -ErrorAction Stop
        }
    }
    [void](Test-ScriptSyntax -ScriptPath $InstalledScriptPath)

    if (-not $InstallOnly) {
        Write-Progress -Id 1 -Activity 'Application governance' -Status 'Writing governed reports' -PercentComplete 90
        Write-ReportSet -ReportRoot $RunRoot -FileInventory @($FileInventory) -FolderInventory @($FolderInventory) -ExecutableInventory @($ExecutableInventory) -DependencyInventory @($DependencyInventory) -ReferencedInventory @($ReferencedInventory) -ManifestCoverage @($ManifestCoverage) -HashCoverage @($HashCoverage) -DuplicateInventory @($DuplicateInventory) -GovernanceDisposition @($GovernanceDisposition) -Summary $Summary
    }

    $FilesAfter = @(Get-ApplicationFiles -RootPath $ApplicationRoot -ToolPath $ToolRoot -AuditPath $AuditBaseRoot)
    $SourceStateAfter = @(Get-SourceState -Files $FilesAfter)
    $SourceUnchanged = (($SourceStateBefore | ConvertTo-Json -Depth 4 -Compress) -ceq ($SourceStateAfter | ConvertTo-Json -Depth 4 -Compress))
    if (-not $SourceUnchanged) {
        throw 'SOURCE VALIDATION FAILED: application files outside governance paths changed.'
    }

    Write-Progress -Id 1 -Activity 'Application governance' -Completed
    Write-Host ('INSTALLED TOOL: {0}' -f $InstalledScriptPath) -ForegroundColor Green
    if (-not $InstallOnly) {
        Write-Host ('REPORT ROOT: {0}' -f $RunRoot) -ForegroundColor Green
    }
    Write-Host ('SOURCE UNCHANGED OUTSIDE GOVERNANCE PATHS: {0}' -f $SourceUnchanged) -ForegroundColor Green
    Write-Host 'APPLICATION GOVERNANCE PASS' -ForegroundColor Green
}
catch {
    Write-Progress -Id 1 -Activity 'Application governance' -Completed
    Write-Host ('CRITICAL ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor Red
    exit 1
}
