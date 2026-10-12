#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$RegisterPath = "C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1\INTEGRATION_REGISTER.json",
    [string]$RepositoryRoot = "C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1",
    [string]$OutputRoot = "C:\Python\GeoAi\00_Base_Datos\GeoAi\14_Register_Verification",
    [string]$ApplicationRoot = "C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1",
    [switch]$AuditOnly
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ScriptName = "Verify_CP09A_Integration_Register"
$Version = "1.0.0"
$Classification = "BHP Internal"
$RunId = "REGISTER-VERIFY-{0}" -f (Get-Date -Format "yyyyMMdd-HHmmss")
$StagingRoot = Join-Path $OutputRoot (".staging-" + $RunId)
$FinalRoot = Join-Path $OutputRoot $RunId
$InstalledFolder = Join-Path $ApplicationRoot "00_Script\Register_Verification"
$InstalledPath = Join-Path $InstalledFolder "Verify_CP09A_Integration_Register_v1_0_0.ps1"
$SourceScriptPath = $MyInvocation.MyCommand.Path
function Assert-PathValue { param([string]$Value,[string]$Name) if([string]::IsNullOrWhiteSpace($Value)){throw ("{0} is empty." -f $Name)} if(-not [IO.Path]::IsPathFullyQualified($Value)){throw ("{0} must be absolute." -f $Name)} }
function Test-ScriptSyntax { param([string]$Path) $tokens=$null;$errors=$null;[void][System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors);return ($null -eq $errors -or $errors.Count -eq 0) }
# Extracts the _vMAJOR_MINOR_PATCH token from a file name and returns MAJOR.MINOR.PATCH, or $null if absent.
function Get-VersionFromName { param([string]$Name) $m=[regex]::Match($Name,'_v(\d+)_(\d+)_(\d+)');if($m.Success){return ("{0}.{1}.{2}" -f $m.Groups[1].Value,$m.Groups[2].Value,$m.Groups[3].Value)};return $null }
# Extracts the in-file declared version from $Version or $ScriptVersion, or $null if the file defines neither.
function Get-VersionFromContent { param([string]$Path) try{$t=Get-Content -LiteralPath $Path -Raw -Encoding utf8 -ErrorAction Stop}catch{return $null};$m=[regex]::Match($t,'(?im)^\s*\$(?:Script)?Version\s*=\s*["'']([0-9]+\.[0-9]+\.[0-9]+)["'']');if($m.Success){return $m.Groups[1].Value};return $null }
# Resolves a registered component to a real file: exact registered path first, then by leaf name under RepositoryRoot.
function Resolve-Component { param([string]$RegisteredPath,[string]$Root,[object[]]$RepoFiles) $r=[ordered]@{Status="FILE_NOT_FOUND";Target=$null;Method="NONE"};if(Test-Path -LiteralPath $RegisteredPath -PathType Leaf){$r.Status="RESOLVED";$r.Target=[IO.Path]::GetFullPath($RegisteredPath);$r.Method="EXACT_PATH";return [pscustomobject]$r};$leaf=Split-Path $RegisteredPath -Leaf;$m=@($RepoFiles|Where-Object{$_.Name -ieq $leaf});if($m.Count -eq 1){$r.Status="RESOLVED";$r.Target=$m[0].FullName;$r.Method="LEAF_NAME"}elseif($m.Count -gt 1){$r.Status="AMBIGUOUS";$r.Method="MULTIPLE_LEAF_NAMES"};return [pscustomobject]$r }
Assert-PathValue $RegisterPath "RegisterPath"
Assert-PathValue $RepositoryRoot "RepositoryRoot"
if(-not(Test-Path -LiteralPath $RegisterPath -PathType Leaf)){throw ("Register file not found: {0}" -f $RegisterPath)}
if(-not(Test-Path -LiteralPath $RepositoryRoot -PathType Container)){throw ("RepositoryRoot not found: {0}" -f $RepositoryRoot)}
$RepositoryRoot=[IO.Path]::GetFullPath($RepositoryRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
$register=Get-Content -LiteralPath $RegisterPath -Raw -Encoding utf8|ConvertFrom-Json
if($null -eq $register.Components){throw "Register JSON contains no Components array."}
$repoFiles=@(Get-ChildItem -LiteralPath $RepositoryRoot -File -Recurse -Force)
$results=@()
foreach($component in $register.Components){
    $registeredPath=[string]$component.Path
    $registeredHash=([string]$component.SHA256).ToUpperInvariant()
    $leaf=Split-Path $registeredPath -Leaf
    $resolution=Resolve-Component $registeredPath $RepositoryRoot $repoFiles
    $actualHash=$null;$hashStatus="FILE_NOT_FOUND";$fileVersion=$null;$nameVersion=Get-VersionFromName $leaf;$versionStatus="VERSION_NOT_CHECKED"
    if($resolution.Status -eq "RESOLVED"){
        $actualHash=(Get-FileHash -LiteralPath $resolution.Target -Algorithm SHA256).Hash.ToUpperInvariant()
        $hashStatus=if($actualHash -eq $registeredHash){"HASH_MATCH"}else{"HASH_MISMATCH"}
        $fileVersion=Get-VersionFromContent $resolution.Target
        if($null -eq $nameVersion){$versionStatus="NAME_VERSION_ABSENT"}
        elseif($null -eq $fileVersion){$versionStatus="FILE_VERSION_ABSENT"}
        elseif($fileVersion -eq $nameVersion){$versionStatus="VERSION_MATCH"}
        else{$versionStatus="VERSION_MISMATCH"}
    }elseif($resolution.Status -eq "AMBIGUOUS"){$hashStatus="FILE_NOT_FOUND";$versionStatus="AMBIGUOUS_RESOLUTION"}
    $results+=[pscustomobject]@{
        Component=$leaf;RegisteredPath=$registeredPath;ResolutionStatus=$resolution.Status;ResolutionMethod=$resolution.Method;ResolvedTarget=$resolution.Target
        RegisteredSHA256=$registeredHash;ActualSHA256=$actualHash;HashStatus=$hashStatus
        NameVersion=$nameVersion;FileVersion=$fileVersion;VersionStatus=$versionStatus
    }
}
# Decision idiom aligned with the governance/compare engines:
#   NO_GO            -> at least one resolved file has a confirmed hash mismatch
#   REVIEW_REQUIRED  -> any missing/ambiguous file or any version drift
#   PASS             -> every component resolved, hashes match, versions consistent
$hashMismatch=@($results|Where-Object{$_.HashStatus -eq "HASH_MISMATCH"})
$notResolved=@($results|Where-Object{$_.ResolutionStatus -ne "RESOLVED"})
$versionMismatch=@($results|Where-Object{$_.VersionStatus -eq "VERSION_MISMATCH"})
$decision=if($hashMismatch.Count -gt 0){"NO_GO"}elseif($notResolved.Count -gt 0 -or $versionMismatch.Count -gt 0){"REVIEW_REQUIRED"}else{"PASS"}
$summary=[ordered]@{Classification=$Classification;Version=$Version;RunId=$RunId;RegisterPath=$RegisterPath;RepositoryRoot=$RepositoryRoot;IntegrationVersion=$register.IntegrationVersion;ComponentCount=$results.Count;Resolved=@($results|Where-Object{$_.ResolutionStatus -eq "RESOLVED"}).Count;HashMatches=@($results|Where-Object{$_.HashStatus -eq "HASH_MATCH"}).Count;HashMismatches=$hashMismatch.Count;FilesNotFound=$notResolved.Count;VersionMismatches=$versionMismatch.Count;Decision=$decision;CompletedAt=(Get-Date).ToString("o")}
Write-Host "";Write-Host ("{0} | {1} v{2}" -f $Classification,$ScriptName,$Version) -ForegroundColor Cyan;Write-Host ("REGISTER: {0}" -f $RegisterPath);Write-Host ("REPOSITORY ROOT: {0}" -f $RepositoryRoot);Write-Host ("PROPOSED OUTPUT: {0}" -f $FinalRoot);$results|Format-Table Component,HashStatus,VersionStatus,ResolutionMethod -AutoSize;Write-Host ("DECISION: {0}" -f $decision)
if($AuditOnly){Write-Host "AUDITONLY PASS: NO FILES OR FOLDERS CREATED." -ForegroundColor Green;exit 0}
if(Test-Path -LiteralPath $FinalRoot){throw ("Final output already exists: {0}" -f $FinalRoot)}
if(Test-Path -LiteralPath $StagingRoot){throw ("Staging output already exists: {0}" -f $StagingRoot)}
if(-not(Test-Path -LiteralPath $SourceScriptPath -PathType Leaf)){throw "Current script path could not be resolved."}
if(-not(Test-ScriptSyntax $SourceScriptPath)){throw "Current script failed parser validation."}
New-Item -ItemType Directory -Path $InstalledFolder -Force|Out-Null;Copy-Item -LiteralPath $SourceScriptPath -Destination $InstalledPath -Force;if(-not(Test-ScriptSyntax $InstalledPath)){throw "Installed script failed parser validation."}
New-Item -ItemType Directory -Path $StagingRoot -ErrorAction Stop|Out-Null
$results|Export-Csv (Join-Path $StagingRoot "Register_Verification.csv") -NoTypeInformation -Encoding utf8
$summary|ConvertTo-Json -Depth 10|Set-Content (Join-Path $StagingRoot "Summary.json") -Encoding utf8
$summary.GetEnumerator()|ForEach-Object{"{0}: {1}" -f $_.Key,$_.Value}|Set-Content (Join-Path $StagingRoot "Summary.txt") -Encoding utf8
$lines=@();foreach($f in @(Get-ChildItem -LiteralPath $StagingRoot -File|Sort-Object Name)){$lines+=("{0}  {1}" -f (Get-FileHash $f.FullName -Algorithm SHA256).Hash,$f.Name)};$lines|Set-Content (Join-Path $StagingRoot "SHA256SUMS.txt") -Encoding ascii
Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot;if(-not(Test-Path -LiteralPath $FinalRoot -PathType Container)){throw "Final promotion failed."}
Write-Host ("INSTALLED SCRIPT: {0}" -f $InstalledPath) -ForegroundColor Green;Write-Host ("REPORT ROOT: {0}" -f $FinalRoot) -ForegroundColor Green;Write-Host ("DECISION: {0}" -f $decision);Write-Host "INTEGRATION REGISTER VERIFICATION COMPLETE" -ForegroundColor Green
