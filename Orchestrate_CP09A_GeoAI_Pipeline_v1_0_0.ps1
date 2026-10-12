#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$ScriptRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1\00_Script',
    [string]$OutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\15_Pipeline_Orchestration',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [string]$ZipPath = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\00_Export\Exports_2026-09-28.zip',
    [string]$ConsolidationRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\00_Consolidacion_BD_GeoAI',
    [string]$LasInputRoot = '',
    [string]$LasCsvRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\01_LAS_to_CSV',
    [string]$CompositeOutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\02_Composites',
    [string]$ProcessedRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\03_Original_Data_Exports\01_Processed_Exports',
    [string]$ComparePipelineRunId = 'GEOAI-EXPORTS-20261002-225517',
    [string]$BaselineExportId = 'Exports_2026-09-26',
    [string]$CandidateExportId = 'Exports_2026-09-28',
    [string]$MasterRunId = '',
    [switch]$AuditOnly
)

# ---------------------------------------------------------------------------
# #requires -Version 7.4 is intentional: this orchestrator must be able to host
# the Cross_Export_Comparison_Engine Compare stage, which itself requires 7.4.
# Running the whole chain in a single 7.4 host avoids a version split-brain.
# The upstream 5.1 stages (Consolidation, LAS-to-CSV, Composite) are forward
# compatible with a 7.4 host, so one shared runtime is safe.
#
# EXECUTION MODEL: each stage is invoked in a CHILD pwsh process via
# 'pwsh -NoProfile -File <stage> @tokens', NOT with the in-process call operator.
# Every stage script ends with 'exit 0'/'exit 1'; in PowerShell, 'exit' inside a
# script invoked with '&' in the current runspace terminates the HOST process, so
# an in-process '& $stage.Script' would kill this orchestrator at stage 1's exit
# before fail-fast, the stage register, or staging->promote could run. A child
# process confines the terminal 'exit' to the child; $LASTEXITCODE after the
# native-exe call then carries the child's exit code for the fail-fast check.
# ---------------------------------------------------------------------------

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Version = '1.0.0'
$Classification = 'Internal'
if ([string]::IsNullOrWhiteSpace($MasterRunId)) { $MasterRunId = 'CP09A-GEOAI-PIPELINE-' + (Get-Date -Format 'yyyyMMdd-HHmmss') }
$RunId = $MasterRunId
$StagingRoot = Join-Path $OutputRoot ('.staging-' + $RunId)
$FinalRoot = Join-Path $OutputRoot $RunId
$RecordRoot = Join-Path $StagingRoot '_Pipeline_Record'
$LogPath = Join-Path $RecordRoot 'Execution.log'
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\Pipeline_Orchestration'
$InstalledPath = Join-Path $InstalledFolder 'Orchestrate_CP09A_GeoAI_Pipeline_v1_0_0.ps1'
$SourceScriptPath = $MyInvocation.MyCommand.Path

function Write-RunLog {
    param([string]$Message,[ValidateSet('INFO','PASS','WARN','ERROR')][string]$Level='INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $line
    if (Test-Path -LiteralPath $RecordRoot -PathType Container) { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 }
}

function Test-ScriptSyntax {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    return ($null -eq $errors -or $errors.Count -eq 0)
}

# Locates a stage script by file-name glob under ScriptRoot (recursive), newest first.
function Resolve-StageScript {
    param([string]$Pattern)
    if (-not (Test-Path -LiteralPath $ScriptRoot -PathType Container)) { throw ('ScriptRoot not found: {0}' -f $ScriptRoot) }
    $match = Get-ChildItem -LiteralPath $ScriptRoot -Recurse -File -Filter $Pattern -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1
    if ($null -eq $match) { throw ('Stage script not found under ScriptRoot: {0}' -f $Pattern) }
    return $match.FullName
}

# Converts a stage argument hashtable into an ordered -Name Value token array suitable
# for 'pwsh -File <script> @tokens'. Switch params ($true) collapse to a bare -Name.
function ConvertTo-StageArgTokens {
    param([hashtable]$StageArgs)
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($key in ($StageArgs.Keys | Sort-Object)) {
        $value = $StageArgs[$key]
        if ($value -is [bool]) {
            if ($value) { [void]$tokens.Add('-' + $key) }
            continue
        }
        [void]$tokens.Add('-' + $key)
        [void]$tokens.Add([string]$value)
    }
    return ,$tokens.ToArray()
}

function Install-CurrentScript {
    if (-not (Test-Path -LiteralPath $SourceScriptPath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-ScriptSyntax $SourceScriptPath)) { throw 'Current script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    Copy-Item -LiteralPath $SourceScriptPath -Destination $InstalledPath -Force
    if (-not (Test-ScriptSyntax $InstalledPath)) { throw 'Installed script failed parser validation.' }
}

try {
    Write-Host ''
    Write-Host ('{0} | CP09A GEOAI PIPELINE ORCHESTRATOR v{1}' -f $Classification,$Version) -ForegroundColor Cyan
    Write-Host ('MASTER RUN_ID: {0}' -f $RunId)
    Write-Host ('SCRIPT ROOT: {0}' -f $ScriptRoot)
    Write-Host ('MODE: {0}' -f $(if ($AuditOnly) { 'AUDIT_ONLY (dry run, no stage executed)' } else { 'EXECUTE' }))

    # Resolve each stage script path. The Compare stage prefers v1_0_2; falls back to any version.
    $consolidationScript = Resolve-StageScript 'Run_CP09A_GeoAI_Consolidation_v*.ps1'
    $lasToCsvScript      = Resolve-StageScript 'Convert_CP09A_LAS_to_CSV_v*.ps1'
    $compositeScript     = Resolve-StageScript 'Composite_CP09A_GeoAI_1m_3m_6m_v*.ps1'
    try { $compareScript = Resolve-StageScript 'Cross_Export_Comparison_Engine_v1_0_2.ps1' }
    catch { $compareScript = Resolve-StageScript 'Cross_Export_Comparison_Engine_v*.ps1' }

    # Define the ordered chain. Each stage carries the argument map used for invocation.
    # NOTE ON RUN_ID COUPLING: the shared MasterRunId tags this orchestrator's own
    # evidence. The downstream stages do not accept a RunId parameter; instead the
    # pipeline couples through folder discovery. Composite auto-discovers the latest
    # LAS-to-CSV run via its Resolve-InputRoot (newest CP09A-LAS-TO-CSV-* folder under
    # LasCsvRoot), so running LAS-to-CSV then Composite in sequence links them without
    # a direct parameter. Compare reads a pre-existing processed-exports pipeline run.
    $stages = @(
        [pscustomobject]@{ Order=1; Name='Consolidation'; Script=$consolidationScript; Args=@{ ZipPath=$ZipPath; ConsolidationRoot=$ConsolidationRoot; ApplicationRoot=$ApplicationRoot } }
        [pscustomobject]@{ Order=2; Name='LAS_to_CSV';    Script=$lasToCsvScript;      Args=@{ InputRoot=$LasInputRoot; OutputRoot=$LasCsvRoot; ApplicationRoot=$ApplicationRoot } }
        [pscustomobject]@{ Order=3; Name='Composite';     Script=$compositeScript;     Args=@{ LasCsvRoot=$LasCsvRoot; OutputRoot=$CompositeOutputRoot; ApplicationRoot=$ApplicationRoot } }
        [pscustomobject]@{ Order=4; Name='Compare';       Script=$compareScript;       Args=@{ ProcessedRoot=$ProcessedRoot; PipelineRunId=$ComparePipelineRunId; BaselineExportId=$BaselineExportId; CandidateExportId=$CandidateExportId } }
    )

    # LAS_to_CSV InputRoot is optional; drop it from the invocation when blank so the
    # stage script keeps its own default.
    foreach ($stage in $stages) {
        if ($stage.Name -eq 'LAS_to_CSV' -and [string]::IsNullOrWhiteSpace($LasInputRoot)) { $stage.Args.Remove('InputRoot') }
    }

    function Format-InvocationChain {
        param([object]$Stage)
        $parts = foreach ($key in ($Stage.Args.Keys | Sort-Object)) { "-{0} '{1}'" -f $key,$Stage.Args[$key] }
        return ('& pwsh -NoProfile -File "{0}" {1}' -f $Stage.Script,($parts -join ' '))
    }

    Write-Host ''
    Write-Host 'RESOLVED INVOCATION CHAIN:' -ForegroundColor Cyan
    foreach ($stage in $stages) {
        Write-Host ('  [{0}] {1}' -f $stage.Order,$stage.Name) -ForegroundColor Cyan
        Write-Host ('      {0}' -f (Format-InvocationChain $stage))
    }

    if ($AuditOnly) {
        Write-Host ''
        Write-Host 'AUDITONLY PASS: NO STAGE EXECUTED, NO FILES OR FOLDERS CREATED.' -ForegroundColor Green
        exit 0
    }

    if (Test-Path -LiteralPath $FinalRoot) { throw ('Final output already exists: {0}' -f $FinalRoot) }
    if (Test-Path -LiteralPath $StagingRoot) { throw ('Staging output already exists: {0}' -f $StagingRoot) }
    New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null
    Write-RunLog ('BHP INVENT | CP09A GeoAI Pipeline Orchestrator | Version {0} | Classification {1}' -f $Version,$Classification)
    Write-RunLog ('MASTER RUN_ID: {0}' -f $RunId)

    Install-CurrentScript
    Write-RunLog ('Installed script: {0}' -f $InstalledPath) 'PASS'

    $stageRecords = New-Object System.Collections.ArrayList
    $chainFailed = $false
    $failedStage = $null
    foreach ($stage in $stages) {
        $start = Get-Date
        Write-RunLog ('STAGE {0} START: {1} -> {2}' -f $stage.Order,$stage.Name,$stage.Script)
        $status = 'PASS'
        $exitCode = 0
        $errorMessage = ''
        # Each stage runs in a CHILD pwsh process (not the orchestrator's own runspace).
        # The stage scripts terminate with 'exit 0'/'exit 1'; invoking them with the call
        # operator in-process would run that 'exit' against THIS orchestrator and kill the
        # chain at stage 1. Hosting the stage in 'pwsh -NoProfile -File' confines the
        # terminal 'exit' to the child, and $LASTEXITCODE after the native-exe call then
        # reflects the child's exit code so fail-fast, the stage register, and the
        # staging->promote all remain reachable.
        $stageTokens = ConvertTo-StageArgTokens -StageArgs $stage.Args
        try {
            & pwsh -NoProfile -File $stage.Script @stageTokens *>> $LogPath
            $exitCode = $LASTEXITCODE
            if ($null -eq $exitCode) { $exitCode = 0 }
            if ($exitCode -ne 0) { $status = 'FAIL' }
        }
        catch {
            $status = 'FAIL'
            $exitCode = if ($null -ne $LASTEXITCODE -and $LASTEXITCODE -ne 0) { $LASTEXITCODE } else { 1 }
            $errorMessage = $_.Exception.Message
        }
        $end = Get-Date
        [void]$stageRecords.Add([pscustomobject]@{
            Order=$stage.Order;Stage=$stage.Name;Script=$stage.Script;ExitCode=$exitCode;Status=$status
            StartedAt=$start.ToString('o');CompletedAt=$end.ToString('o');DurationSeconds=[math]::Round(($end-$start).TotalSeconds,2)
            Error=$errorMessage;MasterRunId=$RunId
        })
        if ($status -ne 'PASS') {
            Write-RunLog ('STAGE {0} FAIL (exit {1}): {2}' -f $stage.Order,$exitCode,$stage.Name) 'ERROR'
            $chainFailed = $true
            $failedStage = $stage.Name
            break
        }
        Write-RunLog ('STAGE {0} PASS: {1}' -f $stage.Order,$stage.Name) 'PASS'
    }

    $decision = if ($chainFailed) { 'NO_GO' } else { 'PASS' }
    $stageRecords | Export-Csv -LiteralPath (Join-Path $RecordRoot 'Pipeline_Stage_Register.csv') -NoTypeInformation -Encoding UTF8
    $summary = [ordered]@{
        Title='BHP INVENT | CP09A GeoAI Pipeline Orchestration';Classification=$Classification;Version=$Version;MasterRunId=$RunId
        ScriptRoot=$ScriptRoot;StagesPlanned=$stages.Count;StagesExecuted=$stageRecords.Count
        StagesPassed=@($stageRecords | Where-Object Status -eq 'PASS').Count;FailedStage=$failedStage;Decision=$decision;CompletedAt=(Get-Date).ToString('o')
    }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.json') -Encoding UTF8
    $summary.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key,$_.Value } | Set-Content -LiteralPath (Join-Path $RecordRoot 'Summary.txt') -Encoding UTF8
    $manifestFiles = @(Get-ChildItem -LiteralPath $RecordRoot -File | Where-Object Name -ne 'SHA256SUMS.txt')
    $manifest = foreach ($file in $manifestFiles) { $hash=Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256; '{0} *{1}' -f $hash.Hash,$file.Name }
    $manifest | Set-Content -LiteralPath (Join-Path $RecordRoot 'SHA256SUMS.txt') -Encoding ASCII

    Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot
    if (-not (Test-Path -LiteralPath $FinalRoot -PathType Container)) { throw 'Final promotion failed.' }

    Write-Host ''
    Write-Host ('PIPELINE DECISION: {0}' -f $decision) -ForegroundColor $(if ($chainFailed) { 'Red' } else { 'Green' })
    Write-Host ('STAGES PASSED: {0}/{1}' -f @($stageRecords | Where-Object Status -eq 'PASS').Count,$stages.Count)
    Write-Host ('OUTPUT: {0}' -f $FinalRoot) -ForegroundColor Cyan
    Write-Host ''
    if ($chainFailed) { exit 1 }
    exit 0
}
catch {
    try {
        if (-not (Test-Path -LiteralPath $RecordRoot -PathType Container)) { New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null }
        Write-RunLog $_.Exception.Message 'ERROR'
        @('MESSAGE: '+$_.Exception.Message,'SCRIPT: '+$_.InvocationInfo.ScriptName,'LINE: '+$_.InvocationInfo.ScriptLineNumber,'POSITION: '+$_.InvocationInfo.PositionMessage,'',($_ | Out-String)) | Set-Content -LiteralPath (Join-Path $RecordRoot 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: '+$_.Exception.Message) -ForegroundColor Red }
    exit 1
}
