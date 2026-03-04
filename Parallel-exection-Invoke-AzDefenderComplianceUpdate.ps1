# ================================
# Script: Invoke-DefenderUpdate-AndStatus-Parallel.ps1
# Purpose: Install Defender update and fetch signature status (Parallel with Single Login)
# ================================

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Compute  -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# -------------------------------
# CONFIGURATION
# -------------------------------

$ExcelPath   = "C:\Scripts\VMInput.xlsx"
$OutputExcel = "C:\Scripts\DefenderUpdateReport.xlsx"
$MpamSasUrl     = "PASTE_MPAM_SAS_URL"
$PlatformSasUrl = "PASTE_PLATFORM_SAS_URL"

$ThrottleLimit  = 4   # Number of parallel VMs

# -------------------------------
# LOGIN (ONLY ONCE)
# -------------------------------

Write-Host "Authenticating to Azure..." -ForegroundColor Cyan
Connect-AzAccount -ErrorAction Stop

$ContextFile = Join-Path $env:TEMP "AzContext.json"
Save-AzContext -Path $ContextFile -Force

# -------------------------------
# LOAD INPUT
# -------------------------------

$VMList = Import-Excel -Path $ExcelPath |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.VMName) -and
        -not [string]::IsNullOrWhiteSpace($_.ResourceGroup) -and
        -not [string]::IsNullOrWhiteSpace($_.Subscription)
    }

$Jobs = @()

# -------------------------------
# REMOTE SCRIPT TEMPLATE
# -------------------------------

$RemoteScript = @'
$ErrorActionPreference = "Stop"

$DownloadFolder = "C:\Users\swetha\Downloads\DefenderUpdate"

if (-not (Test-Path $DownloadFolder)) {
    New-Item -ItemType Directory -Path $DownloadFolder -Force | Out-Null
}

$MpamFile = Join-Path $DownloadFolder "mpam-fe.exe"
$PlatformFile = Join-Path $DownloadFolder "updateplatform.exe"

$mpamExit = "N/A"
$platformExit = "N/A"

try {

    Invoke-WebRequest -Uri "__PLATFORM_URL__" -OutFile $PlatformFile -UseBasicParsing
    Invoke-WebRequest -Uri "__MPAM_URL__" -OutFile $MpamFile -UseBasicParsing

    # Install Platform FIRST
    $Platform = Start-Process -FilePath $PlatformFile -ArgumentList "/quiet /norestart" -Wait -PassThru
    $platformExit = $Platform.ExitCode

    # Install Signature SECOND
    $Mpam = Start-Process -FilePath $MpamFile -ArgumentList "/quiet /norestart" -Wait -PassThru
    $mpamExit = $Mpam.ExitCode

    # Allow Defender service to refresh
    Start-Sleep -Seconds 10
    Update-MpSignature -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
}
catch {
    $mpamExit = "ERROR"
    $platformExit = "ERROR"
}

# Cleanup files
foreach ($File in @($MpamFile, $PlatformFile)) {
    if (Test-Path $File) {
        Remove-Item $File -Force -ErrorAction SilentlyContinue
    }
}

try {
    $status = Get-MpComputerStatus

    $output = [PSCustomObject]@{
        ComputerName                  = $env:COMPUTERNAME
        MpamExitCode                  = $mpamExit
        PlatformExitCode              = $platformExit
        AntivirusSignatureVersion     = $status.AntivirusSignatureVersion
        AntivirusSignatureLastUpdated = $status.AntivirusSignatureLastUpdated
        NISSignatureVersion           = $status.NISSignatureVersion
        NISSignatureLastUpdated       = $status.NISSignatureLastUpdated
        RealTimeProtectionEnabled     = $status.RealTimeProtectionEnabled
        AntivirusEnabled              = $status.AntivirusEnabled
    }

    $output | ConvertTo-Json -Compress
}
catch {
    [PSCustomObject]@{
        ComputerName                  = $env:COMPUTERNAME
        MpamExitCode                  = $mpamExit
        PlatformExitCode              = $platformExit
        AntivirusSignatureVersion     = "ERROR"
        AntivirusSignatureLastUpdated = "ERROR"
        NISSignatureVersion           = "ERROR"
        NISSignatureLastUpdated       = "ERROR"
        RealTimeProtectionEnabled     = "ERROR"
        AntivirusEnabled              = "ERROR"
    } | ConvertTo-Json -Compress
}
'@

$RemoteScript = $RemoteScript.Replace("__MPAM_URL__", $MpamSasUrl)
$RemoteScript = $RemoteScript.Replace("__PLATFORM_URL__", $PlatformSasUrl)

# -------------------------------
# START PARALLEL JOBS
# -------------------------------

foreach ($Row in $VMList) {

    while (@($Jobs | Where-Object { $_.State -eq "Running" }).Count -ge $ThrottleLimit) {
        Start-Sleep -Seconds 2
    }

    Write-Host "Queueing $($Row.VMName)..." -ForegroundColor Yellow

    $Jobs += Start-Job -ArgumentList $Row, $RemoteScript, $ContextFile -ScriptBlock {

        param($Row, $RemoteScript, $ContextFile)

        Import-Module Az.Accounts -ErrorAction Stop
        Import-Module Az.Compute  -ErrorAction Stop

        Import-AzContext -Path $ContextFile -ErrorAction Stop | Out-Null

        try {
            $sub = Get-AzSubscription -SubscriptionName $Row.Subscription -ErrorAction Stop
            Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null

            $run = Invoke-AzVMRunCommand `
                -ResourceGroupName $Row.ResourceGroup `
                -VMName $Row.VMName `
                -CommandId "RunPowerShellScript" `
                -ScriptString $RemoteScript `
                -ErrorAction Stop

            $message = $run.Value[0].Message

            # Extract JSON safely (handles RunCommand wrapping text)
            if ($message -match '(?s)\{.*\}') {
                $json = $matches[0]
                $result = $json | ConvertFrom-Json
            }
            else {
                throw "JSON output not found in RunCommand response"
            }

            [PSCustomObject]@{
                VMName                        = $Row.VMName
                ResourceGroup                 = $Row.ResourceGroup
                Subscription                  = $Row.Subscription
                MpamExitCode                  = $result.MpamExitCode
                PlatformExitCode              = $result.PlatformExitCode
                AntivirusSignatureVersion     = $result.AntivirusSignatureVersion
                AntivirusSignatureLastUpdated = $result.AntivirusSignatureLastUpdated
                AntivirusEnabled              = $result.AntivirusEnabled
                NISSignatureVersion           = $result.NISSignatureVersion
                NISSignatureLastUpdated       = $result.NISSignatureLastUpdated
                RealTimeProtectionEnabled     = $result.RealTimeProtectionEnabled
            }
        }
        catch {
            [PSCustomObject]@{
                VMName                        = $Row.VMName
                ResourceGroup                 = $Row.ResourceGroup
                Subscription                  = $Row.Subscription
                MpamExitCode                  = "ERROR"
                PlatformExitCode              = "ERROR"
                AntivirusSignatureVersion     = "ERROR"
                AntivirusSignatureLastUpdated = "ERROR"
                AntivirusEnabled              = "ERROR"
                NISSignatureVersion           = "ERROR"
                NISSignatureLastUpdated       = "ERROR"
                RealTimeProtectionEnabled     = "ERROR"
            }
        }
    }
}

# -------------------------------
# WAIT FOR COMPLETION
# -------------------------------

Write-Host "Waiting for all jobs to complete..." -ForegroundColor Cyan
Wait-Job -Job $Jobs

# -------------------------------
# COLLECT RESULTS
# -------------------------------

$finalResults = foreach ($job in $Jobs) {
    Receive-Job -Job $job
}

Remove-Job -Job $Jobs

if (Test-Path $ContextFile) {
    Remove-Item $ContextFile -Force
}

# -------------------------------
# EXPORT REPORT
# -------------------------------

$finalResults | Export-Excel `
    -Path $OutputExcel `
    -WorksheetName "DefenderUpdateStatus" `
    -AutoSize `
    -BoldTopRow `
    -FreezeTopRow `
    -AutoFilter `
    -ClearSheet

Write-Host "Report saved to $OutputExcel" -ForegroundColor Green
#I have tested the parallel execution on 3 vms which means the script runs in 3vms in a single point of time 
#12:53 march 1st initiated. till 13:06 march 1st Completed. -- for 3 vms
<#13:29 march 2nd initiated.
13:42 march 2nd Completed. -- for 8 vms
QA
----------------------------
14:41 march 2nd initiated.
15:04 march 2nd Completed. -- for 11 vms
PROD
WARNING: INITIALIZATION: Fallback context save mode to process because of error during checking token cache persistence: 
Persistence check fails due to unknown error.
#>