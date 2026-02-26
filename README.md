Azure Defender Update Automation

Automated PowerShell solution to deploy Microsoft Defender signature and platform updates across multiple Azure Windows Virtual Machines using Azure Run Command.

This solution enables centralized update orchestration, signature validation, and reporting via Excel — without requiring WinRM, RDP, or inbound firewall changes.

🚀 Overview

This repository provides an automation framework that:

Reads VM details from an Excel file

Connects to Azure using PIM-enabled access

Downloads Defender update binaries from Azure Storage (SAS URLs)

Installs updates silently on target VMs

Retrieves Defender signature status post-installation

Generates a structured Excel report for auditing

The solution leverages:

Azure PowerShell (Az Modules)

Azure VM Run Command

ImportExcel PowerShell module

🏗 Architecture
Local Machine
    │
    ├── Reads VM list from Excel
    ├── Connects to Azure (PIM)
    ├── Invokes Azure Run Command
    │
Azure VM
    ├── Downloads mpam-fe.exe
    ├── Downloads updateplatform.exe
    ├── Installs updates silently
    ├── Retrieves Defender status
    └── Returns JSON response
    │
Local Machine
    └── Exports consolidated Excel report

No direct connectivity to VMs is required.

📦 Repository Structure
azure-defender-update-automation/
│
├── Invoke-DefenderUpdate-AndStatus.ps1
├── VMInput.xlsx (Sample only)
├── README.md
📋 Requirements
PowerShell Version

Windows PowerShell 5.1 or PowerShell 7+

Required Modules
Install-Module Az -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
Azure Permissions

You must have:

Contributor or higher on target VMs

PIM activation enabled (if required by your organization)

📄 Excel Input Format

The Excel file must contain the following columns:

VMName	ResourceGroup	Subscription
VM01	RG-Prod	Prod-Sub
VM02	RG-Test	Test-Sub

Column names must match exactly.

🔐 Storage Requirements

The Defender update binaries must be stored in Azure Storage and accessible via SAS URL:

mpam-fe.exe

updateplatform.exe

SAS URLs must have read (sp=r) permission and valid expiry time.

▶ Usage

Update configuration values inside the script:

$ExcelPath   = "C:\Scripts\VMInput.xlsx"
$OutputExcel = "C:\Scripts\DefenderUpdateReport.xlsx"

$MpamSasUrl     = "PASTE_MPAM_SAS_URL"
$PlatformSasUrl = "PASTE_PLATFORM_SAS_URL"

Run the script:

.\Invoke-DefenderUpdate-AndStatus.ps1

Authenticate via Azure login (PIM activation if required)

Review generated Excel report

📊 Output Report

The generated Excel file contains:

VM Name

Resource Group

Subscription

Mpam Exit Code

Platform Exit Code

Antivirus Signature Version

Antivirus Signature Last Updated

Antivirus Enabled

Real-Time Protection Enabled

Exit Code Reference:

Code	Meaning
0	Success
3010	Success, reboot required
Other	Installation issue
✅ Features

Multi-subscription support

Excel-driven VM targeting

Safe file cleanup (no folder deletion)

JSON-based structured VM response

Consolidated Excel reporting

Error handling per VM

No WinRM or RDP required

🛡 Security Considerations

Do not commit SAS URLs to public repositories

Use short-lived SAS tokens

Rotate storage access regularly

Ensure least-privilege Azure RBAC access

Store sensitive configuration securely (Key Vault recommended for production use)

🔄 Future Enhancements

Potential improvements:

Parallel execution for large VM sets

Version comparison (before vs after update)

Reboot-required detection

Tag-based targeting

Centralized logging to Log Analytics

Azure Automation integration

CI/CD integration

📘 Example Use Cases

Enterprise Defender signature compliance checks

MSP-managed Azure environments

Security patch validation workflows

Scheduled signature enforcement automation
👤 Author

Azure Automation Framework
PowerShell-based Azure VM security automation
