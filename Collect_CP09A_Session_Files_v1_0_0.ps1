#Requires -Version 7.4
# BHP Internal. Read-only search; copies matching files into a new run folder and zips it. Sources are never modified.
[CmdletBinding()]
param(
    [Parameter()][string[]]$SearchRoots = @("$env:USERPROFILE\Downloads", 'C:\Python\GeoAi'),
    [Parameter()][string]$OutputRoot = "$env:USERPROFILE\Downloads"
)
Clear-Host
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Expected = @(
    @{ Name = 'Convert_CP09A_AutoLith_LAS_To_CSV_v1_0_0.ps1'; Hash = 'F780FBED7E09BA2F4731362351382352527969446F4636CBB4D37C0E988DCE5C' },
    @{ Name = 'Initialize_CP09A_Composite_Run_v1_0_4.ps1'; Hash = 'CD5E739751B77ADA0562FDA6AB87B1A47E97C99DEA8583E4AF63C241B555029E' },
    @{ Name = 'Initialize_CP09A_Composite_Run_v1_0_3.ps1'; Hash = '675D9870C755A532C77B02ADE42E702F89ADD9E691FF8DDA945B011AAF4AE82E' },
    @{ Name = 'Initialize_CP09A_Composite_Run_v1_0_2.ps1'; Hash = 'F4A02F7C1C98E233D973D67312794137971FD9D1B10407021937AB518858B58A' },
    @{ Name = 'Initialize_CP09A_Composite_Run_v1_0_1.ps1'; Hash = 'A7E894769D58BC9AE8F5FED67C8958C12265FFE5B1928B21C86375B32C5264C3' },
    @{ Name = 'Initialize_CP09A_Composite_Run_v1_0_0.ps1'; Hash = '46A5A2A2B6FF75B5357BD72BEDF94C68F37967EF97409203E231AB94CE3FEA97' },
    @{ Name = 'Run_CP09A_Composite_Jobs_v1_0_0.ps1'; Hash = '6F4A29E6D691734E1E905F8B1A7385E587552CC2DDE688329C5353E573FCCAA0' },
    @{ Name = 'Inventory_CP09A_Spence_ToComposite_v1_0_0.ps1'; Hash = '46090810082F58261E79ACA12470B972D4F06C9E8B68F680D95E29BCEE98F986' },
    @{ Name = 'Vendor_Hole_List.csv'; Hash = '84E6F3CFA7D3248AC7162992942273DE47A41B28305712470EFC38981F4747FC' }
)
if ([string]::IsNullOrWhiteSpace($OutputRoot) -or -not (Test-Path -LiteralPath $OutputRoot -PathType Container)) { throw "OutputRoot not found: $OutputRoot" }
$RunId = 'CP09A-SESSION-RECOVERY-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
$RunFolder = Join-Path $OutputRoot $RunId
if (Test-Path -LiteralPath $RunFolder) { throw "Output already exists: $RunFolder" }
New-Item -ItemType Directory -Path $RunFolder | Out-Null
$Results = [System.Collections.Generic.List[object]]::new()
$Index = 0
foreach ($Item in $Expected) {
    $Index++
    Write-Progress -Activity 'Collecting CP09A session files' -Status $Item.Name -PercentComplete ([int]([Math]::Floor(($Index / $Expected.Count) / 0.01)))
    $Found = @()
    foreach ($Root in $SearchRoots) {
        if (Test-Path -LiteralPath $Root -PathType Container) {
            $Found += @(Get-ChildItem -LiteralPath $Root -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Item.Name })
        }
    }
    $Match = @($Found | Where-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash -eq $Item.Hash } | Select-Object -First 1)
    if ($Match.Count -eq 1) {
        Copy-Item -LiteralPath $Match[0].FullName -Destination (Join-Path $RunFolder $Item.Name)
        $Results.Add([pscustomobject]@{ File = $Item.Name; Status = 'COLLECTED'; Source = $Match[0].FullName; SHA256 = $Item.Hash })
    }
    elseif (@($Found).Count -gt 0) { $Results.Add([pscustomobject]@{ File = $Item.Name; Status = 'FOUND_HASH_MISMATCH'; Source = (@($Found)[0].FullName); SHA256 = $Item.Hash }) }
    else { $Results.Add([pscustomobject]@{ File = $Item.Name; Status = 'NOT_FOUND'; Source = ''; SHA256 = $Item.Hash }) }
}
Write-Progress -Activity 'Collecting CP09A session files' -Completed
$Results | Export-Csv -LiteralPath (Join-Path $RunFolder 'Recovery_Register.csv') -NoTypeInformation -Encoding utf8NoBOM
$Zip = "$RunFolder.zip"
Compress-Archive -Path (Join-Path $RunFolder 'Recovery_Register.csv'), (Get-ChildItem -LiteralPath $RunFolder -File | Where-Object { $_.Name -ne 'Recovery_Register.csv' } | ForEach-Object { $_.FullName }) -DestinationPath $Zip
$Results | Format-Table File, Status -AutoSize
Write-Host ('ZIP: ' + $Zip) -ForegroundColor Green
Write-Host ('ZIP SHA256: ' + (Get-FileHash -LiteralPath $Zip -Algorithm SHA256).Hash) -ForegroundColor Green
