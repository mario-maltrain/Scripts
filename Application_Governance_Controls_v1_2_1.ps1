#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [string]$ToolRelativePath = '00_Script\Application_Governance',
    [string]$AuditRelativePath = '_Audit\Application_Governance',
    [switch]$AuditOnly,
    [switch]$ContinueOnNonCriticalError
)

Clear-Host
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.2.1'
$Classification = 'BHP Internal'
$RunId = 'APP-GOV-CONTROLS-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$Warnings = [System.Collections.Generic.List[object]]::new()

function Assert-Value {
    param([AllowEmptyString()][string]$Value,[string]$Name)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw ('{0} cannot be empty.' -f $Name) }
}

function Get-FullPathSafe {
    param([string]$PathValue)
    Assert-Value -Value $PathValue -Name 'PathValue'
    return [System.IO.Path]::GetFullPath($PathValue)
}

function Get-RelativeSafe {
    param([string]$BasePath,[string]$FullPath)
    return [System.IO.Path]::GetRelativePath((Get-FullPathSafe $BasePath),(Get-FullPathSafe $FullPath))
}

function Test-BelowPath {
    param([string]$Candidate,[string]$Parent)
    $candidateFull = (Get-FullPathSafe $Candidate).TrimEnd([char]92)
    $parentFull = (Get-FullPathSafe $Parent).TrimEnd([char]92)
    return ($candidateFull -ieq $parentFull) -or $candidateFull.StartsWith($parentFull + [char]92,[System.StringComparison]::OrdinalIgnoreCase)
}

function Read-TextSafe {
    param([System.IO.FileInfo]$File)
    try { return [System.IO.File]::ReadAllText($File.FullName) }
    catch {
        if ($ContinueOnNonCriticalError) {
            $Warnings.Add([pscustomobject]@{ Stage='READ'; Path=$File.FullName; Message=$_.Exception.Message })
            return $null
        }
        throw
    }
}

function Get-SHA256Safe {
    param([System.IO.FileInfo]$File)
    try { return (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256).Hash.ToUpperInvariant() }
    catch {
        if ($ContinueOnNonCriticalError) {
            $Warnings.Add([pscustomobject]@{ Stage='HASH'; Path=$File.FullName; Message=$_.Exception.Message })
            return $null
        }
        throw
    }
}

function Test-PowerShellFile {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if (@($errors).Count -gt 0) {
        $errors | Select-Object Message,ErrorId,Extent | Format-List
        throw 'SYNTAX FAIL'
    }
}

function Write-CsvSafe {
    param([array]$Data,[string]$Path)
    @($Data) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8
}

try {
    Assert-Value -Value $ApplicationRoot -Name 'ApplicationRoot'
    Assert-Value -Value $ToolRelativePath -Name 'ToolRelativePath'
    Assert-Value -Value $AuditRelativePath -Name 'AuditRelativePath'

    $ApplicationRoot = Get-FullPathSafe $ApplicationRoot
    if (-not (Test-Path -LiteralPath $ApplicationRoot -PathType Container)) { throw ('Application root not found: {0}' -f $ApplicationRoot) }

    Test-PowerShellFile -Path $MyInvocation.MyCommand.Path

    $ToolRoot = Join-Path $ApplicationRoot $ToolRelativePath
    $AuditBase = Join-Path $ApplicationRoot $AuditRelativePath
    $RunRoot = Join-Path $AuditBase $RunId
    $ReportsRoot = Join-Path $RunRoot 'reports'
    $SignatureRoot = Join-Path $RunRoot 'signature'
    $InstalledPath = Join-Path $ToolRoot 'Application_Governance_Controls_v1_2_1.ps1'

    Write-Host ('{0} | Application Governance Controls v{1}' -f $Classification,$Version) -ForegroundColor Cyan
    Write-Host ('APPLICATION ROOT: {0}' -f $ApplicationRoot)
    Write-Host ('PROPOSED TOOL PATH: {0}' -f $InstalledPath)
    Write-Host ('PROPOSED REPORT ROOT: {0}' -f $RunRoot)

    $Files = @(Get-ChildItem -LiteralPath $ApplicationRoot -File -Recurse -Force | Where-Object {
        -not (Test-BelowPath -Candidate $_.FullName -Parent $ToolRoot) -and
        -not (Test-BelowPath -Candidate $_.FullName -Parent $AuditBase)
    } | Sort-Object FullName)
    $Folders = @(Get-ChildItem -LiteralPath $ApplicationRoot -Directory -Recurse -Force | Where-Object {
        -not (Test-BelowPath -Candidate $_.FullName -Parent $ToolRoot) -and
        -not (Test-BelowPath -Candidate $_.FullName -Parent $AuditBase)
    } | Sort-Object FullName)

    $StateBefore = @($Files | ForEach-Object { '{0}|{1}|{2}' -f $_.FullName,$_.Length,$_.LastWriteTimeUtc.Ticks })

    $CodeExtensions = @('.ps1','.psm1','.psd1','.py','.r','.bat','.cmd')
    $ConfigExtensions = @('.json','.yaml','.yml','.toml','.config','.xml')
    $ExecutableExtensions = @('.exe','.com','.bat','.cmd','.ps1','.psm1','.py','.r','.jar','.dll')
    $AnalysisExtensions = @($CodeExtensions + $ConfigExtensions | Sort-Object -Unique)

    $CodeFiles = @($Files | Where-Object { $AnalysisExtensions -contains $_.Extension.ToLowerInvariant() })
    $ExecutableFiles = @($Files | Where-Object { $ExecutableExtensions -contains $_.Extension.ToLowerInvariant() })

    $manifestTokens = @('manifest','release_register','release-version','release_version','inventory','catalog','deployment','signature','source_contract')
    $ManifestFiles = @($Files | Where-Object {
        $name = $_.Name.ToLowerInvariant()
        $matched = $false
        foreach ($token in $manifestTokens) { if ($name.Contains($token)) { $matched = $true; break } }
        $matched -and @('.json','.csv','.md','.txt','.yaml','.yml','.xml') -contains $_.Extension.ToLowerInvariant()
    })
    $GovernanceArtifactFiles = @($Files | Where-Object {
        $name = $_.Name.ToLowerInvariant()
        $name.Contains('readme') -or $name.Contains('changelog') -or $name.Contains('version')
    })

    $HashRegisterFiles = @($Files | Where-Object {
        $name = $_.Name.ToLowerInvariant()
        $name.Contains('sha256') -or $name.Contains('checksum') -or $name.Contains('hash_register')
    })

    Write-Progress -Id 1 -Activity 'Governance controls' -Status 'Hashing files' -PercentComplete 15
    $FileRecords = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($file in $Files) {
        $index++
        Write-Progress -Id 1 -Activity 'Governance controls' -Status ('Hashing {0} of {1}' -f $index,$Files.Count) -PercentComplete (15 + [int](20.0 * $index / [math]::Max(1,$Files.Count)))
        $FileRecords.Add([pscustomobject]@{
            File = $file
            RelativePath = Get-RelativeSafe $ApplicationRoot $file.FullName
            SHA256 = Get-SHA256Safe $file
        })
    }

    Write-Progress -Id 1 -Activity 'Governance controls' -Status 'Analyzing dependencies' -PercentComplete 40
    $DependencyGraph = [System.Collections.Generic.List[object]]::new()
    $RuntimeDependencies = [System.Collections.Generic.List[object]]::new()
    $ReferencedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ReferencedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $patterns = @(
        [pscustomobject]@{ Type='PowerShellModule'; Pattern='(?im)^\s{0,}Import-Module\s+[''"]?([^''"\r\n]+)' },
        [pscustomobject]@{ Type='PowerShellDotSource'; Pattern='(?im)^\s{0,}\.\s+[''"]([^''"]+\.ps1)[''"]' },
        [pscustomobject]@{ Type='PowerShellPathCommand'; Pattern='(?im)(?:Import-Csv|Export-Csv|Get-Content|Set-Content|Add-Content|Test-Path|Get-ChildItem|Copy-Item|Start-Process)\s+(?:-LiteralPath\s+|-Path\s+)?[''"]([^''"]+)[''"]' },
        [pscustomobject]@{ Type='PythonImport'; Pattern='(?im)^\s{0,}(?:from|import)\s+([A-Za-z0-9_\.]+)' },
        [pscustomobject]@{ Type='RDependency'; Pattern='(?im)^\s{0,}(?:library|require|source)\s{0,}\(\s{0,}[''"]?([^''"\)]+)' },
        [pscustomobject]@{ Type='QuotedFileReference'; Pattern='(?i)[''"]([^''"]+\.(?:ps1|psm1|psd1|py|r|json|yaml|yml|toml|config|xml|csv|parquet|feather|las|exe|dll|bat|cmd))[''"]' }
    )
    $runtimePattern = '(?im)(?:Start-Process\s+|&\s+|\b)(python(?:\.exe)?|pwsh(?:\.exe)?|powershell(?:\.exe)?|rscript(?:\.exe)?|java(?:\.exe)?|cmd(?:\.exe)?)\b'

    foreach ($source in $CodeFiles) {
        $content = Read-TextSafe $source
        if ($null -eq $content) { continue }
        $sourceRelative = Get-RelativeSafe $ApplicationRoot $source.FullName
        foreach ($definition in $patterns) {
            foreach ($match in [regex]::Matches($content,$definition.Pattern)) {
                $value = $match.Groups[1].Value.Trim()
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                $normalized = $value.Replace('/','\')
                $leaf = [System.IO.Path]::GetFileName($normalized)
                if (-not [string]::IsNullOrWhiteSpace($leaf)) { [void]$ReferencedNames.Add($leaf) }
                if (-not [System.IO.Path]::IsPathRooted($normalized)) { [void]$ReferencedPaths.Add($normalized.TrimStart([char]46,[char]92)) }
                $matches = @($Files | Where-Object {
                    $_.Name -ieq $leaf -or (Get-RelativeSafe $ApplicationRoot $_.FullName) -ieq $normalized
                })
                $DependencyGraph.Add([pscustomobject]@{
                    SourceFile=$sourceRelative
                    TargetReference=$value
                    DependencyType=$definition.Type
                    MatchCount=$matches.Count
                    ExistsInApplication=($matches.Count -gt 0)
                    MatchedPaths=(@($matches | ForEach-Object { Get-RelativeSafe $ApplicationRoot $_.FullName }) -join '; ')
                })
            }
        }
        foreach ($match in [regex]::Matches($content,$runtimePattern)) {
            $runtimeRaw = $match.Groups[1].Value.ToLowerInvariant()
            $runtimeNormalized = $runtimeRaw.Replace('.exe','')
            $RuntimeDependencies.Add([pscustomobject]@{
                SourceFile=$sourceRelative
                RuntimeRaw=$runtimeRaw
                RuntimeNormalized=$runtimeNormalized
                DetectedCommand=$match.Value.Trim()
            })
        }
    }

    $DependencyGraph = [System.Collections.Generic.List[object]]::new(@(
        $DependencyGraph |
        Sort-Object SourceFile,DependencyType,TargetReference,MatchedPaths -Unique
    ))
    $RuntimeDependencies = [System.Collections.Generic.List[object]]::new(@(
        $RuntimeDependencies |
        Sort-Object SourceFile,RuntimeNormalized,DetectedCommand -Unique
    ))

    Write-Progress -Id 1 -Activity 'Governance controls' -Status 'Analyzing manifests and hashes' -PercentComplete 60
    $ManifestEvidence = [System.Collections.Generic.List[object]]::new()
    $ManifestEntries = [System.Collections.Generic.List[object]]::new()
    $ManifestListedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ManifestListedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $manifestPathPattern = '(?i)([A-Za-z0-9_\.\- ]+\.(?:ps1|psm1|psd1|py|r|json|yaml|yml|toml|config|xml|csv|parquet|feather|las|exe|dll|bat|cmd|md|txt))'
    foreach ($manifest in $ManifestFiles) {
        $content = Read-TextSafe $manifest
        if ($null -eq $content) { continue }
        $entryCount = 0
        $resolvedCount = 0
        foreach ($match in [regex]::Matches($content,$manifestPathPattern)) {
            $entry = $match.Groups[1].Value.Trim().Replace('/','\')
            if ([string]::IsNullOrWhiteSpace($entry)) { continue }
            $entryCount++
            $leaf = [System.IO.Path]::GetFileName($entry)
            $manifestDir = $manifest.Directory.FullName
            $candidatePaths = [System.Collections.Generic.List[string]]::new()
            if ([System.IO.Path]::IsPathRooted($entry)) {
                $candidatePaths.Add((Get-FullPathSafe $entry))
            }
            else {
                $candidatePaths.Add((Get-FullPathSafe (Join-Path $manifestDir $entry)))
                $candidatePaths.Add((Get-FullPathSafe (Join-Path $ApplicationRoot $entry)))
            }
            $exact = @($FileRecords | Where-Object {
                $full = Get-FullPathSafe $_.File.FullName
                @($candidatePaths | Where-Object { $_ -ieq $full }).Count -gt 0
            })
            $nameMatches = @($FileRecords | Where-Object { $_.File.Name -ieq $leaf })
            $resolution = 'UNRESOLVED'
            $resolved = @()
            if ($exact.Count -eq 1) { $resolution='EXACT_PATH'; $resolved=@($exact) }
            elseif ($exact.Count -gt 1) { $resolution='AMBIGUOUS_EXACT_PATH'; $resolved=@($exact) }
            elseif ($nameMatches.Count -eq 1) { $resolution='UNIQUE_NAME_FALLBACK'; $resolved=@($nameMatches) }
            elseif ($nameMatches.Count -gt 1) { $resolution='AMBIGUOUS_NAME'; $resolved=@($nameMatches) }
            if ($resolution -eq 'EXACT_PATH' -or $resolution -eq 'UNIQUE_NAME_FALLBACK') {
                $resolvedCount++
                [void]$ManifestListedNames.Add($leaf)
                foreach ($item in $resolved) { [void]$ManifestListedPaths.Add($item.RelativePath) }
            }
            $ManifestEntries.Add([pscustomobject]@{
                ManifestPath=Get-RelativeSafe $ApplicationRoot $manifest.FullName
                Entry=$entry
                Resolution=$resolution
                MatchCount=$resolved.Count
                MatchedPaths=(@($resolved | ForEach-Object { $_.RelativePath }) -join '; ')
            })
        }
        $ManifestEvidence.Add([pscustomobject]@{
            ManifestPath=Get-RelativeSafe $ApplicationRoot $manifest.FullName
            Extension=$manifest.Extension
            EntriesParsed=$entryCount
            EntriesResolved=$resolvedCount
            LengthBytes=$manifest.Length
        })
    }

    $HashEntries = [System.Collections.Generic.List[object]]::new()
    $HashListedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $HashListedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($register in $HashRegisterFiles) {
        $content = Read-TextSafe $register
        if ($null -eq $content) { continue }
        $registerDirectory = $register.Directory.FullName
        foreach ($line in ($content -split "`r?`n")) {
            if ($line -match '^\s{0,}([A-Fa-f0-9]{64})\s+[^A-Za-z0-9]?(.+?)\s{0,}$') {
                $registeredHash = $Matches[1].ToUpperInvariant()
                $registeredPathRaw = $Matches[2].Trim().TrimStart([char]46,[char]92,[char]47)
                $registeredPath = $registeredPathRaw.Replace('/','\')
                $registeredName = [System.IO.Path]::GetFileName($registeredPath)

                $candidatePaths = [System.Collections.Generic.List[string]]::new()
                if ([System.IO.Path]::IsPathRooted($Matches[2].Trim())) {
                    $candidatePaths.Add((Get-FullPathSafe $Matches[2].Trim()))
                }
                else {
                    $candidatePaths.Add((Get-FullPathSafe (Join-Path $registerDirectory $registeredPath)))
                    $candidatePaths.Add((Get-FullPathSafe (Join-Path $ApplicationRoot $registeredPath)))
                }

                $exactMatches = @($FileRecords | Where-Object {
                    $discoveredFull = Get-FullPathSafe $_.File.FullName
                    @($candidatePaths | Where-Object { $_ -ieq $discoveredFull }).Count -gt 0
                })
                $nameMatches = @($FileRecords | Where-Object { $_.File.Name -ieq $registeredName })

                $resolution = 'UNRESOLVED'
                $diskMatches = @()
                if ($exactMatches.Count -eq 1) {
                    $resolution = 'EXACT_PATH'
                    $diskMatches = @($exactMatches)
                }
                elseif ($exactMatches.Count -gt 1) {
                    $resolution = 'AMBIGUOUS_EXACT_PATH'
                    $diskMatches = @($exactMatches)
                }
                elseif ($nameMatches.Count -eq 1) {
                    $resolution = 'UNIQUE_NAME_FALLBACK'
                    $diskMatches = @($nameMatches)
                }
                elseif ($nameMatches.Count -gt 1) {
                    $resolution = 'AMBIGUOUS_NAME'
                    $diskMatches = @($nameMatches)
                }

                if ($resolution -eq 'EXACT_PATH' -or $resolution -eq 'UNIQUE_NAME_FALLBACK') {
                    [void]$HashListedNames.Add($registeredName)
                    foreach ($diskMatch in $diskMatches) { [void]$HashListedPaths.Add($diskMatch.RelativePath) }
                }

                $hashMatch = @($diskMatches | Where-Object { $_.SHA256 -eq $registeredHash }).Count -gt 0
                $HashEntries.Add([pscustomobject]@{
                    RegisterPath=Get-RelativeSafe $ApplicationRoot $register.FullName
                    RegisteredPath=$registeredPath
                    RegisteredSHA256=$registeredHash
                    Resolution=$resolution
                    MatchCount=$diskMatches.Count
                    MatchedPaths=(@($diskMatches | ForEach-Object { $_.RelativePath }) -join '; ')
                    HashMatch=$hashMatch
                })
            }
        }
    }

    $DuplicateGroups = @($FileRecords | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.SHA256) } | Group-Object SHA256 | Where-Object { @($_.Group).Count -gt 1 })
    $DuplicateAnalysis = [System.Collections.Generic.List[object]]::new()
    $DuplicatePaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($group in $DuplicateGroups) {
        $members = @($group.Group)
        $hasGovernedPath = @($members | Where-Object { $_.RelativePath -match '(?i)(backup|archive|snapshot|release)' }).Count -gt 0
        $hasCode = @($members | Where-Object { $CodeExtensions -contains $_.File.Extension.ToLowerInvariant() }).Count -gt 0
        $classification = if ($hasCode -and -not $hasGovernedPath) { 'ACTIVE_CODE_DUPLICATE' } elseif ($hasGovernedPath) { 'GOVERNED_COPY' } else { 'EXACT_DUPLICATE_REVIEW' }
        foreach ($member in $members) {
            [void]$DuplicatePaths.Add($member.RelativePath)
            $DuplicateAnalysis.Add([pscustomobject]@{
                SHA256=$group.Name
                GroupSize=$members.Count
                Classification=$classification
                RelativePath=$member.RelativePath
                LengthBytes=$member.File.Length
            })
        }
    }

    Write-Progress -Id 1 -Activity 'Governance controls' -Status 'Classifying orphans and risks' -PercentComplete 80
    $OrphanAssets = [System.Collections.Generic.List[object]]::new()
    $RiskRegister = [System.Collections.Generic.List[object]]::new()
    $GovernanceMatrix = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $FileRecords) {
        $relative = $record.RelativePath
        $referenced = $ReferencedNames.Contains($record.File.Name) -or $ReferencedPaths.Contains($relative)
        $manifestListed = $ManifestListedNames.Contains($record.File.Name) -or $ManifestListedPaths.Contains($relative)
        $hashListed = $HashListedNames.Contains($record.File.Name) -or $HashListedPaths.Contains($relative)
        $duplicate = $DuplicatePaths.Contains($relative)
        $isCode = $ExecutableExtensions -contains $record.File.Extension.ToLowerInvariant()
        $isRootAsset = $record.File.Name -match '(?i)^(start|run|launch|launcher|main|install|bootstrap)'
        $orphan = -not $referenced -and -not $manifestListed -and -not $hashListed -and -not $isRootAsset
        if ($orphan) {
            $OrphanAssets.Add([pscustomobject]@{
                RelativePath=$relative
                Extension=$record.File.Extension
                ExecutableOrCode=$isCode
                RootAsset=$isRootAsset
                SHA256=$record.SHA256
            })
        }
        $disposition = if ($referenced -or $manifestListed -or $hashListed -or $isRootAsset -or $isCode) { 'KEEP' }
                       elseif ($duplicate) { 'REVIEW' }
                       elseif ($relative -match '(?i)(backup|archive|legacy|deprecated|superseded)') { 'ARCHIVE' }
                       elseif ($record.File.Extension.ToLowerInvariant() -in @('.tmp','.temp','.cache')) { 'REMOVE_CANDIDATE' }
                       else { 'REVIEW' }
        $GovernanceMatrix.Add([pscustomobject]@{
            Disposition=$disposition
            RelativePath=$relative
            Referenced=$referenced
            ManifestListed=$manifestListed
            HashListed=$hashListed
            RootAsset=$isRootAsset
            ExecutableOrCode=$isCode
            ExactDuplicate=$duplicate
            Orphan=$orphan
        })
        if ($isCode -and $orphan) { $RiskRegister.Add([pscustomobject]@{ Severity='P1'; Control='ORPHAN_CODE'; Path=$relative; Message='Code or executable has no verified reference, manifest, hash registration or root-entry evidence.' }) }
        if ($duplicate -and $isCode) { $RiskRegister.Add([pscustomobject]@{ Severity='P1'; Control='DUPLICATE_CODE'; Path=$relative; Message='Code is part of an exact SHA256 duplicate group.' }) }
        if (-not $hashListed -and ($isCode -or $manifestListed)) { $RiskRegister.Add([pscustomobject]@{ Severity='P2'; Control='HASH_COVERAGE'; Path=$relative; Message='Governed or executable file is not listed in a discovered hash register.' }) }
    }
    foreach ($edge in $DependencyGraph) {
        if (-not $edge.ExistsInApplication -and $edge.DependencyType -ne 'PythonImport' -and $edge.DependencyType -ne 'RDependency' -and $edge.DependencyType -ne 'PowerShellPathCommand') {
            $RiskRegister.Add([pscustomobject]@{ Severity='P1'; Control='MISSING_DEPENDENCY'; Path=$edge.SourceFile; Message=('Reference not resolved: {0}' -f $edge.TargetReference) })
        }
    }
    foreach ($hashEntry in $HashEntries) {
        if (($hashEntry.Resolution -eq 'EXACT_PATH' -or $hashEntry.Resolution -eq 'UNIQUE_NAME_FALLBACK') -and -not $hashEntry.HashMatch) {
            $RiskRegister.Add([pscustomobject]@{ Severity='P0'; Control='HASH_MISMATCH'; Path=$hashEntry.RegisteredPath; Message='Resolved file does not match the registered SHA256.' })
        }
        elseif ($hashEntry.Resolution -eq 'AMBIGUOUS_NAME' -or $hashEntry.Resolution -eq 'AMBIGUOUS_EXACT_PATH') {
            $RiskRegister.Add([pscustomobject]@{ Severity='P1'; Control='HASH_REFERENCE_AMBIGUOUS'; Path=$hashEntry.RegisteredPath; Message='Hash register entry resolves to more than one file and cannot support a P0 mismatch conclusion.' })
        }
        elseif ($hashEntry.Resolution -eq 'UNRESOLVED') {
            $RiskRegister.Add([pscustomobject]@{ Severity='P1'; Control='HASH_REFERENCE_UNRESOLVED'; Path=$hashEntry.RegisteredPath; Message='Hash register entry does not resolve to a discovered file.' })
        }
    }

    $referencedCount = @($GovernanceMatrix | Where-Object Referenced).Count
    $manifestCovered = @($GovernanceMatrix | Where-Object ManifestListed).Count
    $hashCovered = @($GovernanceMatrix | Where-Object HashListed).Count
    $orphanCount = @($GovernanceMatrix | Where-Object Orphan).Count
    $eligibleGovernedFiles = @($GovernanceMatrix | Where-Object { $_.ExecutableOrCode -or $_.ManifestListed -or $_.HashListed -or $_.RootAsset }).Count
    $eligibleDenominator = [math]::Max(1,$eligibleGovernedFiles)
    $fileDenominator = [math]::Max(1,$Files.Count)
    $manifestParsed = @($ManifestEntries).Count
    $manifestResolved = @($ManifestEntries | Where-Object { $_.Resolution -eq 'EXACT_PATH' -or $_.Resolution -eq 'UNIQUE_NAME_FALLBACK' }).Count
    $hashParsed = @($HashEntries).Count
    $hashResolved = @($HashEntries | Where-Object { $_.Resolution -eq 'EXACT_PATH' -or $_.Resolution -eq 'UNIQUE_NAME_FALLBACK' }).Count
    $dependencyResolved = @($DependencyGraph | Where-Object ExistsInApplication).Count
    $dependencyResolvable = @($DependencyGraph | Where-Object { $_.DependencyType -ne 'PythonImport' -and $_.DependencyType -ne 'RDependency' }).Count
    $confirmedP0 = @($RiskRegister | Where-Object { $_.Severity -eq 'P0' -and $_.Control -eq 'HASH_MISMATCH' }).Count
    $p1Count = @($RiskRegister | Where-Object Severity -eq 'P1').Count
    $p2Count = @($RiskRegister | Where-Object Severity -eq 'P2').Count
    $decision = if ($confirmedP0 -gt 0) { 'NO_GO' } elseif ($p1Count -gt 0) { 'REVIEW_REQUIRED' } elseif ($p2Count -gt 0) { 'PASS_WITH_WARNINGS' } else { 'PASS' }

    $Kpis = [pscustomobject]@{
        Files=$Files.Count
        Folders=$Folders.Count
        CodeAndConfig=$CodeFiles.Count
        DependencyEdges=$DependencyGraph.Count
        DependencyResolutionPct=if ($dependencyResolvable -gt 0) { [math]::Round(100.0 * $dependencyResolved / $dependencyResolvable,2) } else { 0 }
        RuntimeReferenceCount=@($RuntimeDependencies).Count
        UniqueRuntimeCount=@($RuntimeDependencies.RuntimeNormalized | Sort-Object -Unique).Count
        UniqueRuntimeNames=(@($RuntimeDependencies.RuntimeNormalized | Sort-Object -Unique) -join '; ')
        ManifestArtifactsFound=$ManifestFiles.Count
        GovernanceArtifactsFound=$GovernanceArtifactFiles.Count
        ManifestEntriesParsed=$manifestParsed
        ManifestEntriesResolved=$manifestResolved
        UniqueFilesListedInManifest=$manifestCovered
        ManifestCoveragePct=[math]::Round(100.0 * $manifestCovered / $eligibleDenominator,2)
        ManifestResolutionPct=if ($manifestParsed -gt 0) { [math]::Round(100.0 * $manifestResolved / $manifestParsed,2) } else { 0 }
        HashRegistersFound=$HashRegisterFiles.Count
        HashEntriesParsed=$hashParsed
        HashEntriesResolved=$hashResolved
        HashEntriesAmbiguous=@($HashEntries | Where-Object { $_.Resolution -like 'AMBIGUOUS?' }).Count
        HashEntriesUnresolved=@($HashEntries | Where-Object Resolution -eq 'UNRESOLVED').Count
        HashEntriesVerified=@($HashEntries | Where-Object HashMatch).Count
        UniqueFilesHashCovered=$hashCovered
        HashCoveragePct=[math]::Round(100.0 * $hashCovered / $eligibleDenominator,2)
        HashResolutionPct=if ($hashParsed -gt 0) { [math]::Round(100.0 * $hashResolved / $hashParsed,2) } else { 0 }
        ExactDuplicateGroups=$DuplicateGroups.Count
        OrphanAssets=$orphanCount
        ReferencedCoveragePct=[math]::Round(100.0 * $referencedCount / $fileDenominator,2)
        OrphanRatePct=[math]::Round(100.0 * $orphanCount / $fileDenominator,2)
        ConfirmedP0Count=$confirmedP0
        P1Count=$p1Count
        P2Count=$p2Count
        Decision=$decision
    }

    if ($AuditOnly) {
        $Kpis | Format-List
        @($GovernanceMatrix | Group-Object Disposition | ForEach-Object { [pscustomobject]@{ Disposition=$_.Name; Count=@($_.Group).Count } }) | Format-Table -AutoSize
        Write-Host 'AUDITONLY PASS: NO FILES OR FOLDERS CREATED.' -ForegroundColor Green
        Write-Progress -Id 1 -Activity 'Governance controls' -Completed
        exit 0
    }

    foreach ($folder in @($ToolRoot,$RunRoot,$ReportsRoot,$SignatureRoot)) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { New-Item -ItemType Directory -Path $folder | Out-Null }
    }
    if ((Get-FullPathSafe $MyInvocation.MyCommand.Path) -ine (Get-FullPathSafe $InstalledPath)) {
        if (Test-Path -LiteralPath $InstalledPath -PathType Leaf) {
            $sourceHash = (Get-FileHash -LiteralPath $MyInvocation.MyCommand.Path -Algorithm SHA256).Hash
            $installedHash = (Get-FileHash -LiteralPath $InstalledPath -Algorithm SHA256).Hash
            if ($sourceHash -ne $installedHash) { throw ('Installed file exists with different content: {0}' -f $InstalledPath) }
        } else { Copy-Item -LiteralPath $MyInvocation.MyCommand.Path -Destination $InstalledPath }
    }

    Write-CsvSafe @($DependencyGraph) (Join-Path $ReportsRoot 'Dependency_Graph.csv')
    Write-CsvSafe @($RuntimeDependencies) (Join-Path $ReportsRoot 'Runtime_Dependencies.csv')
    Write-CsvSafe @($ManifestEvidence) (Join-Path $ReportsRoot 'Manifest_Discovery.csv')
    Write-CsvSafe @($ManifestEntries) (Join-Path $ReportsRoot 'Manifest_Entry_Resolution.csv')
    Write-CsvSafe @($HashEntries) (Join-Path $ReportsRoot 'Hash_Register_Validation.csv')
    Write-CsvSafe @($DuplicateAnalysis) (Join-Path $ReportsRoot 'Duplicate_Analysis.csv')
    Write-CsvSafe @($OrphanAssets) (Join-Path $ReportsRoot 'Orphan_Assets.csv')
    Write-CsvSafe @($RiskRegister) (Join-Path $ReportsRoot 'Governance_Risk_Register.csv')
    Write-CsvSafe @($GovernanceMatrix) (Join-Path $ReportsRoot 'Governance_Disposition.csv')
    Write-CsvSafe @($Warnings) (Join-Path $ReportsRoot 'NonCritical_Warnings.csv')
    @($Kpis.PSObject.Properties | ForEach-Object { [pscustomobject]@{ KPI=$_.Name; Value=[string]$_.Value } }) | Export-Csv -LiteralPath (Join-Path $ReportsRoot 'Governance_Executive_Summary.csv') -NoTypeInformation -Encoding utf8
    $Kpis | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $RunRoot 'Governance_KPIs.json') -Encoding utf8
    @(
        '# Application Governance Controls',
        '',
        ('Classification: {0}' -f $Classification),
        ('Version: {0}' -f $Version),
        ('RUN_ID: {0}' -f $RunId),
        '',
        'Read-only analysis outside the dedicated tool and audit paths.',
        'No delete, move, rename or source overwrite operation is implemented.',
        'All dispositions are recommendations requiring evidence review.'
    ) | Set-Content -LiteralPath (Join-Path $RunRoot 'README.md') -Encoding utf8
    @(
        '# Changelog',
        '',
        'Version 1.2.1',
        '- Deduplicated dependency and runtime edges.',
        '- Separated manifest discovery from manifest coverage.',
        '- Resolved SHA256 entries by exact path, register-relative path, root-relative path or unique-name fallback.',
        '- Downgraded ambiguous hash references from false P0 mismatch conclusions to P1 review.',
        '- Added hash resolution KPIs and complete evidence packaging.',
        '- Added manifest entry resolution and eligible-file denominators.',
        '- Added normalized unique runtime KPIs.',
        '- Added evidence-based decision gate: NO_GO, REVIEW_REQUIRED, PASS_WITH_WARNINGS or PASS.'
    ) | Set-Content -LiteralPath (Join-Path $RunRoot 'CHANGELOG.md') -Encoding utf8

    $manifestItems = @(Get-ChildItem -LiteralPath $RunRoot -File -Recurse | Sort-Object FullName | ForEach-Object {
        [pscustomobject]@{ RelativePath=Get-RelativeSafe $RunRoot $_.FullName; LengthBytes=$_.Length; SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToUpperInvariant() }
    })
    [ordered]@{ Application='Application_Governance_Controls'; Version=$Version; Classification=$Classification; RunId=$RunId; Files=$manifestItems } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $RunRoot 'Manifest.json') -Encoding utf8
    $signLines = @(Get-ChildItem -LiteralPath $RunRoot -File -Recurse | Where-Object Name -ne 'SHA256SUMS.txt' | Sort-Object FullName | ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToUpperInvariant(),(Get-RelativeSafe $RunRoot $_.FullName) })
    $signLines | Set-Content -LiteralPath (Join-Path $SignatureRoot 'SHA256SUMS.txt') -Encoding ascii

    $FilesAfter = @(Get-ChildItem -LiteralPath $ApplicationRoot -File -Recurse -Force | Where-Object { -not (Test-BelowPath $_.FullName $ToolRoot) -and -not (Test-BelowPath $_.FullName $AuditBase) } | Sort-Object FullName)
    $StateAfter = @($FilesAfter | ForEach-Object { '{0}|{1}|{2}' -f $_.FullName,$_.Length,$_.LastWriteTimeUtc.Ticks })
    if (($StateBefore | ConvertTo-Json -Compress) -cne ($StateAfter | ConvertTo-Json -Compress)) { throw 'SOURCE UNCHANGED VALIDATION FAILED' }

    Write-Progress -Id 1 -Activity 'Governance controls' -Completed
    Write-Host ('INSTALLED CONTROL SCRIPT: {0}' -f $InstalledPath) -ForegroundColor Green
    Write-Host ('REPORT ROOT: {0}' -f $RunRoot) -ForegroundColor Green
    Write-Host 'APPLICATION GOVERNANCE CONTROLS PASS' -ForegroundColor Green
}
catch {
    Write-Progress -Id 1 -Activity 'Governance controls' -Completed
    Write-Host ('CRITICAL ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor Red
    exit 1
}
