# ================================
# Script: Invoke-DefenderUpdate-AndStatus.ps1
# Purpose: Install Defender update and fetch signature status
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

# -------------------------------
# LOGIN
# -------------------------------

Connect-AzAccount

# -------------------------------
# LOAD INPUT
# -------------------------------

$VMList = Import-Excel -Path $ExcelPath
$finalResults = @()

# -------------------------------
# REMOTE SCRIPT (UPDATE + STATUS)
# -------------------------------

$RemoteScript = @'
$DownloadFolder = "C:\Temp\DefenderUpdate"

if (-not (Test-Path $DownloadFolder)) {
    New-Item -ItemType Directory -Path $DownloadFolder -Force | Out-Null
}

$MpamFile = Join-Path $DownloadFolder "mpam-fe.exe"
$PlatformFile = Join-Path $DownloadFolder "updateplatform.exe"

$mpamExit = "N/A"
$platformExit = "N/A"

try {
    Invoke-WebRequest -Uri "__MPAM_URL__" -OutFile $MpamFile -UseBasicParsing -ErrorAction Stop
    Invoke-WebRequest -Uri "__PLATFORM_URL__" -OutFile $PlatformFile -UseBasicParsing -ErrorAction Stop

    $Mpam = Start-Process -FilePath $MpamFile -ArgumentList "/quiet /norestart" -Wait -PassThru
    $Platform = Start-Process -FilePath $PlatformFile -ArgumentList "/quiet /norestart" -Wait -PassThru

    $mpamExit = $Mpam.ExitCode
    $platformExit = $Platform.ExitCode
}
catch {
    $mpamExit = "ERROR"
    $platformExit = "ERROR"
}

# Cleanup only files
foreach ($File in @($MpamFile, $PlatformFile)) {
    if (Test-Path $File) {
        Remove-Item $File -Force -ErrorAction SilentlyContinue
    }
}

# Get Defender status
try {
    $status = Get-MpComputerStatus

    [PSCustomObject]@{
        ComputerName                  = $env:COMPUTERNAME
        MpamExitCode                  = $mpamExit
        PlatformExitCode              = $platformExit
        AntivirusSignatureVersion     = $status.AntivirusSignatureVersion
        AntivirusSignatureLastUpdated = $status.AntivirusSignatureLastUpdated
        NISSignatureVersion           = $status.NISSignatureVersion
        NISSignatureLastUpdated       = $status.NISSignatureLastUpdated
        RealTimeProtectionEnabled     = $status.RealTimeProtectionEnabled
        AntivirusEnabled              = $status.AntivirusEnabled
    } | ConvertTo-Json -Compress
}
catch {
    [PSCustomObject]@{
        ComputerName                  = $env:COMPUTERNAME
        MpamExitCode                  = $mpamExit
        PlatformExitCode              = $platformExit
        AntivirusSignatureVersion     = "ERROR"
        AntivirusSignatureLastUpdated = "ERROR"
        RealTimeProtectionEnabled     = "ERROR"
        NISSignatureVersion           = "ERROR"
        NISSignatureLastUpdated       = "ERROR"
        AntivirusEnabled              = "ERROR"
    } | ConvertTo-Json -Compress
}
'@

$RemoteScript = $RemoteScript.Replace("__MPAM_URL__", $MpamSasUrl)
$RemoteScript = $RemoteScript.Replace("__PLATFORM_URL__", $PlatformSasUrl)

# -------------------------------
# PROCESS VMs
# -------------------------------

foreach ($Row in $VMList) {

    Write-Host "Processing $($Row.VMName)..." -ForegroundColor Cyan

    try {
        $sub = Get-AzSubscription -SubscriptionName $Row.Subscription -ErrorAction Stop
        Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null

        $run = Invoke-AzVMRunCommand `
            -ResourceGroupName $Row.ResourceGroup `
            -VMName $Row.VMName `
            -CommandId "RunPowerShellScript" `
            -ScriptString $RemoteScript `
            -ErrorAction Stop

        $json = $run.Value[0].Message.Trim()

        $result = $json | ConvertFrom-Json

        $finalResults += [PSCustomObject]@{
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
        Write-Warning "Failed on $($Row.VMName)"

        $finalResults += [PSCustomObject]@{
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
#===================================================================
#This is the final working version of the script.