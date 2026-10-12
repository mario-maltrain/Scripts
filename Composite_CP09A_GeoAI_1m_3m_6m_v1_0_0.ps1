#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputRoot = '',
    [string]$LasCsvRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\01_LAS_to_CSV',
    [string]$OutputRoot = 'C:\Python\GeoAi\00_Base_Datos\GeoAi\02_Composites',
    [string]$ApplicationRoot = 'C:\Python\GeoAi\01_Aplicacion\CP09A_Release_v2.1',
    [double]$MinimumCoverage = 0.80,
    [double[]]$Scales = @(1.0,3.0,6.0)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

$Version = '1.0.0'
$Classification = 'Internal'
$RunId = 'CP09A-GEOAI-COMPOSITE-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$FinalRoot = Join-Path $OutputRoot $RunId
$StagingRoot = Join-Path $OutputRoot ('.staging-' + $RunId)
$DataRoot = Join-Path $StagingRoot 'Composites'
$RecordRoot = Join-Path $StagingRoot '_Composite_Record'
$FailureRegisterPath = Join-Path $RecordRoot 'Consolidated_Failure_Real_Test_Register.csv'
$InstalledFolder = Join-Path $ApplicationRoot '00_Script\GeoAI_Composites'
$InstalledPath = Join-Path $InstalledFolder 'Composite_CP09A_GeoAI_1m_3m_6m_v1_0_0.ps1'
$SourceScriptPath = $MyInvocation.MyCommand.Path
$PythonPath = Join-Path $RecordRoot 'composite_engine.py'
$LogPath = Join-Path $RecordRoot 'Execution.log'

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

function Add-FailureRecord {
    param([string]$Artifact,[string]$VersionValue,[string]$Stage,[string]$Status,[string]$Evidence,[string]$Disposition)
    [pscustomobject]@{
        Timestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Artifact=$Artifact;Version=$VersionValue;Stage=$Stage
        Status=$Status;Evidence=$Evidence;Disposition=$Disposition;Classification=$Classification;RunId=$RunId
    } | Export-Csv -LiteralPath $FailureRegisterPath -NoTypeInformation -Append -Encoding UTF8
}

function Resolve-InputRoot {
    if (-not [string]::IsNullOrWhiteSpace($InputRoot)) { return $InputRoot }
    if (-not (Test-Path -LiteralPath $LasCsvRoot -PathType Container)) { throw ('LAS-to-CSV root not found: {0}' -f $LasCsvRoot) }
    $candidate = Get-ChildItem -LiteralPath $LasCsvRoot -Directory | Where-Object { $_.Name -like 'CP09A-LAS-TO-CSV-*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $candidate) { throw 'No completed LAS-to-CSV run folder was found.' }
    $summaryPath = Join-Path $candidate.FullName '_Conversion_Record\Summary.json'
    $csvRoot = Join-Path $candidate.FullName 'CSV'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { throw ('LAS-to-CSV Summary.json not found: {0}' -f $summaryPath) }
    $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($summary.Status -ne 'PASS_REAL_TEST') { throw ('LAS-to-CSV run is not PASS_REAL_TEST: {0}' -f $candidate.FullName) }
    if (-not (Test-Path -LiteralPath $csvRoot -PathType Container)) { throw ('LAS-to-CSV data folder not found: {0}' -f $csvRoot) }
    return $csvRoot
}

try {
    if ($MinimumCoverage -le 0 -or $MinimumCoverage -gt 1) { throw 'MinimumCoverage must be greater than 0 and less than or equal to 1.' }
    if (@($Scales).Count -ne 3 -or @($Scales | Sort-Object -Unique).Count -ne 3) { throw 'Exactly three unique scales are required.' }
    foreach ($requiredScale in @(1.0,3.0,6.0)) { if ($requiredScale -notin $Scales) { throw 'Scales must be exactly 1, 3 and 6 metres.' } }
    if (Test-Path -LiteralPath $FinalRoot) { throw ('Final output already exists: {0}' -f $FinalRoot) }
    if (Test-Path -LiteralPath $StagingRoot) { throw ('Staging output already exists: {0}' -f $StagingRoot) }

    New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null
    Write-RunLog ('BHP INVENT | CP09A GeoAI Composite | Version {0} | Classification {1}' -f $Version,$Classification)
    Write-RunLog ('RUN_ID: {0}' -f $RunId)

    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Status 'Validating and installing script' -PercentComplete 3
    if (-not (Test-Path -LiteralPath $SourceScriptPath -PathType Leaf)) { throw 'Current script path could not be resolved.' }
    if (-not (Test-ScriptSyntax $SourceScriptPath)) { throw 'Current script failed parser validation.' }
    New-Item -ItemType Directory -Path $InstalledFolder -Force | Out-Null
    Copy-Item -LiteralPath $SourceScriptPath -Destination $InstalledPath -Force
    if (-not (Test-ScriptSyntax $InstalledPath)) { throw 'Installed script failed parser validation.' }

    Add-FailureRecord 'Run_CP09A_GeoAI_Consolidation_v1_0_0.ps1' '1.0.0' 'Consolidation' 'FAIL' 'ParserError InvalidVariableReferenceWithDrive' 'DO_NOT_USE_REPLACED_BY_v1_0_1'
    Add-FailureRecord 'Run_CP09A_GeoAI_Consolidation_v1_0_1.ps1' '1.0.1' 'Consolidation' 'PASS_REAL_TEST' 'CP09A-GEOAI-CONSOLIDATION-20261001-134313' 'ACTIVE_VALIDATED_BASELINE'
    Add-FailureRecord 'Convert_CP09A_LAS_to_CSV_v1_0_0.ps1' '1.0.0' 'LAS_TO_CSV' 'IN_PROGRESS' 'User confirmed running on real data' 'AWAIT_FINAL_SUMMARY'

    $ResolvedInputRoot = Resolve-InputRoot
    Write-RunLog ('INPUT: {0}' -f $ResolvedInputRoot)
    Write-RunLog ('OUTPUT: {0}' -f $FinalRoot)
    $csvFiles = @(Get-ChildItem -LiteralPath $ResolvedInputRoot -Recurse -File -Filter '*.csv' | Sort-Object FullName)
    if ($csvFiles.Count -eq 0) { throw 'No input CSV files found.' }

    $python = Get-Command python -ErrorAction Stop
    & $python.Source -c "import sys,pandas,numpy; print(sys.version); print('DEPENDENCIES_OK')" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Python dependency validation failed.' }

    $engine = @'
import argparse, csv, hashlib, json, math, os, re, sys, traceback
from pathlib import Path
import numpy as np
import pandas as pd

NULL_SENTINELS = {-999.25, -999.0, -9999.0}
DEPTH_NAMES = {'DEPT','DEPTH','MD','DIST','FROM','Z'}
NON_PROBABILITY_NAMES = {'H-CONFIDENCE','H_CONFIDENCE','CONFIDENCE','ENTROPY'}

def sha256(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''): h.update(block)
    return h.hexdigest().upper()

def safe_write_csv(frame,path):
    if path.exists(): raise RuntimeError(f'No-overwrite safety block: {path}')
    path.parent.mkdir(parents=True,exist_ok=True)
    frame.to_csv(path,index=False,encoding='utf-8-sig',float_format='%.10g')

def read_input(path):
    frame=pd.read_csv(path,low_memory=False)
    if frame.empty: raise RuntimeError('Input CSV is empty')
    frame.columns=[str(c).strip() for c in frame.columns]
    upper={c:c.upper() for c in frame.columns}
    depth_candidates=[c for c in frame.columns if upper[c] in DEPTH_NAMES]
    if not depth_candidates: raise RuntimeError('No recognised depth column')
    depth=depth_candidates[0]
    frame[depth]=pd.to_numeric(frame[depth],errors='coerce')
    numeric=[]
    for c in frame.columns:
        if c==depth: continue
        converted=pd.to_numeric(frame[c],errors='coerce')
        if converted.notna().any():
            frame[c]=converted
            numeric.append(c)
    if not numeric: raise RuntimeError('No numeric value columns')
    frame=frame.dropna(subset=[depth]).sort_values(depth).drop_duplicates(subset=[depth],keep='last')
    if len(frame)<2: raise RuntimeError('Fewer than two unique depth rows')
    if not np.all(np.diff(frame[depth].to_numpy(dtype=float))>0): raise RuntimeError('Depth is not strictly increasing')
    for c in numeric:
        frame.loc[frame[c].isin(NULL_SENTINELS),c]=np.nan
    return frame,depth,numeric

def composite(frame,depth,numeric,scale,min_coverage,hole_id,source_file):
    d=frame[depth].to_numpy(dtype=float)
    diffs=np.diff(d)
    step=float(np.median(diffs[diffs>0]))
    if not np.isfinite(step) or step<=0: raise RuntimeError('Invalid depth step')
    start=math.floor(float(d.min())/scale)*scale
    stop=math.ceil(float(d.max()+step)/scale)*scale
    probability=[c for c in numeric if c.upper().endswith('_PROB') or 'PROBABILITY' in c.upper()]
    if not probability:
        probability=[c for c in numeric if c.upper() not in NON_PROBABILITY_NAMES]
    rows=[]
    for frm in np.arange(start,stop,scale):
        to=frm+scale
        mask=(d>=frm-1e-10)&(d<to-1e-10)
        sub=frame.loc[mask,numeric]
        expected=max(1,int(round(scale/step)))
        observed=int(mask.sum())
        coverage=min(1.0,observed/expected)
        out={'HOLE_ID':hole_id,'FROM':round(float(frm),6),'TO':round(float(to),6),'SCALE_M':scale,'SOURCE_ROWS':observed,'EXPECTED_ROWS':expected,'COVERAGE':coverage,'SOURCE_FILE':source_file}
        for c in numeric:
            values=pd.to_numeric(sub[c],errors='coerce')
            out[c]=float(values.mean()) if values.notna().any() else np.nan
        if coverage>=min_coverage and probability:
            vals={c:out[c] for c in probability if pd.notna(out[c])}
            out['ARGMAX_CLASS']=max(vals,key=vals.get).removesuffix('_PROB') if vals else 'NO_PROBABILITY_INFORMATION'
            out['ARGMAX_VALUE']=max(vals.values()) if vals else np.nan
            out['STATUS']='PASS' if vals else 'NO_PROBABILITY_INFORMATION'
        else:
            out['ARGMAX_CLASS']='INCOMPLETE'
            out['ARGMAX_VALUE']=np.nan
            out['STATUS']='INCOMPLETE'
        rows.append(out)
    return pd.DataFrame(rows),step

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--input',required=True); ap.add_argument('--output',required=True); ap.add_argument('--record',required=True)
    ap.add_argument('--coverage',type=float,required=True); ap.add_argument('--scales',nargs='+',type=float,required=True)
    args=ap.parse_args()
    root=Path(args.input); out=Path(args.output); record=Path(args.record)
    files=sorted(root.rglob('*.csv'))
    register=[]; failures=[]; scale_frames={s:[] for s in args.scales}
    for i,path in enumerate(files,1):
        print(f'PROGRESS|{i}|{len(files)}|{path.name}',flush=True)
        try:
            frame,depth,numeric=read_input(path)
            hole=re.sub(r'_SPENCE.*$','',path.stem,flags=re.I)
            for scale in args.scales:
                comp,step=composite(frame,depth,numeric,scale,args.coverage,hole,str(path))
                scale_frames[scale].append(comp)
                register.append({'SourceFile':str(path),'HoleId':hole,'ScaleM':scale,'InputRows':len(frame),'OutputRows':len(comp),'DepthColumn':depth,'NumericColumns':'|'.join(numeric),'SourceStepM':step,'Status':'PASS'})
        except Exception as exc:
            failures.append({'SourceFile':str(path),'Error':str(exc),'Status':'FAIL'})
            break
    pd.DataFrame(register).to_csv(record/'Composite_Register.csv',index=False,encoding='utf-8-sig')
    if failures:
        pd.DataFrame(failures).to_csv(record/'Composite_Failures.csv',index=False,encoding='utf-8-sig')
        raise RuntimeError(failures[0]['Error'])
    summaries=[]
    for scale,parts in scale_frames.items():
        combined=pd.concat(parts,ignore_index=True)
        scale_name=f'{int(scale)}m'
        target=out/scale_name/f'GeoAI_Composite_{scale_name}.csv'
        safe_write_csv(combined,target)
        summaries.append({'ScaleM':scale,'Rows':len(combined),'Holes':combined.HOLE_ID.nunique(),'PassRows':int((combined.STATUS=='PASS').sum()),'IncompleteRows':int((combined.STATUS=='INCOMPLETE').sum()),'NoProbabilityRows':int((combined.STATUS=='NO_PROBABILITY_INFORMATION').sum()),'OutputFile':str(target),'SHA256':sha256(target)})
    pd.DataFrame(summaries).to_csv(record/'Scale_Summary.csv',index=False,encoding='utf-8-sig')
    summary={'Classification':'Internal','InputRoot':str(root),'InputFiles':len(files),'ScalesM':args.scales,'MinimumCoverage':args.coverage,'FailedFiles':0,'OverwriteCount':0,'Status':'PASS_REAL_TEST','ScaleResults':summaries}
    (record/'Summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8')
    (record/'Summary.txt').write_text('\n'.join(f'{k}: {v}' for k,v in summary.items() if k!='ScaleResults'),encoding='utf-8')
    manifest=[]
    for p in sorted(list(out.rglob('*.csv'))+list(record.glob('*.csv'))+list(record.glob('*.json'))+list(record.glob('*.txt'))):
        if p.name!='SHA256SUMS.txt': manifest.append(f'{sha256(p)} *{p.relative_to(record.parent)}')
    (record/'SHA256SUMS.txt').write_text('\n'.join(manifest)+'\n',encoding='ascii')
    print('RESULT|PASS_REAL_TEST',flush=True)

if __name__=='__main__':
    try: main()
    except Exception:
        traceback.print_exc(); sys.exit(1)
'@
    Set-Content -LiteralPath $PythonPath -Value $engine -Encoding UTF8

    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Status 'Compositing 1 m, 3 m and 6 m' -PercentComplete 10
    $arguments = @($PythonPath,'--input',$ResolvedInputRoot,'--output',$DataRoot,'--record',$RecordRoot,'--coverage',$MinimumCoverage.ToString([System.Globalization.CultureInfo]::InvariantCulture),'--scales','1','3','6')
    & $python.Source @arguments 2>&1 | ForEach-Object {
        $text = $_.ToString()
        if ($text -match '^PROGRESS\|(\d+)\|(\d+)\|(.+)$') {
            $done=[int]$Matches[1]; $total=[int]$Matches[2]; $name=$Matches[3]
            $percent=10+[int](78*($done/[double][math]::Max(1,$total)))
            Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Status 'Compositing source CSV files' -CurrentOperation ('File {0} of {1}: {2}' -f $done,$total,$name) -PercentComplete $percent
        } else { Write-RunLog $text }
    }
    if ($LASTEXITCODE -ne 0) { throw ('Composite engine failed with exit code {0}' -f $LASTEXITCODE) }

    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Status 'Validating outputs' -PercentComplete 92
    $summaryPath = Join-Path $RecordRoot 'Summary.json'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { throw 'Summary.json was not generated.' }
    $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($summary.Status -ne 'PASS_REAL_TEST') { throw ('Unexpected status: {0}' -f $summary.Status) }
    foreach ($scaleName in @('1m','3m','6m')) {
        $target = Join-Path $DataRoot ('{0}\GeoAI_Composite_{0}.csv' -f $scaleName)
        if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw ('Missing composite output: {0}' -f $target) }
        if ((Get-Item -LiteralPath $target).Length -le 0) { throw ('Empty composite output: {0}' -f $target) }
    }
    Add-FailureRecord 'Composite_CP09A_GeoAI_1m_3m_6m_v1_0_0.ps1' '1.0.0' 'COMPOSITING' 'PASS_REAL_TEST' $RunId 'ACTIVE_VALIDATED_BASELINE'
    Remove-Item -LiteralPath $PythonPath -Force

    Write-RunLog 'RESULT REVIEW: PASS_REAL_TEST' 'PASS'
    Write-RunLog 'SCALES: 1 m, 3 m, 6 m' 'PASS'
    Write-RunLog 'OVERWRITES: 0' 'PASS'
    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Status 'Promoting completed run' -PercentComplete 98
    Move-Item -LiteralPath $StagingRoot -Destination $FinalRoot
    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Completed
    Write-Host ''
    Write-Host 'COMPOSITES: PASS_REAL_TEST' -ForegroundColor Green
    Write-Host 'SCALES: 1 m, 3 m, 6 m' -ForegroundColor Green
    Write-Host 'OVERWRITES: 0' -ForegroundColor Green
    Write-Host ('OUTPUT: {0}' -f $FinalRoot) -ForegroundColor Cyan
    Write-Host ''
    exit 0
}
catch {
    try {
        if (-not (Test-Path -LiteralPath $RecordRoot -PathType Container)) { New-Item -ItemType Directory -Path $RecordRoot -Force | Out-Null }
        Write-RunLog $_.Exception.Message 'ERROR'
        Add-FailureRecord 'Composite_CP09A_GeoAI_1m_3m_6m_v1_0_0.ps1' '1.0.0' 'COMPOSITING' 'FAIL' $_.Exception.Message 'BLOCK_AND_REVIEW'
        @('MESSAGE: '+$_.Exception.Message,'SCRIPT: '+$_.InvocationInfo.ScriptName,'LINE: '+$_.InvocationInfo.ScriptLineNumber,'POSITION: '+$_.InvocationInfo.PositionMessage,'',($_ | Out-String)) | Set-Content -LiteralPath (Join-Path $RecordRoot 'Fatal_Error.txt') -Encoding UTF8
    } catch { Write-Host ('FATAL: '+$_.Exception.Message) -ForegroundColor Red }
    Write-Progress -Id 1 -Activity 'CP09A GeoAI composites' -Completed
    exit 1
}
