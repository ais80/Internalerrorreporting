# ==============================================================================
# Provision-SharePointList.ps1
# UHBFT Pharmacy Near Miss Reporting - Automated SharePoint List Provisioner
# ==============================================================================
# This script provisions the complete "Pharmacy Near Miss Log" list structure
# inside your chosen NHS Microsoft 365 SharePoint Team or Communication Site.
#
# USAGE:
#   .\Provision-SharePointList.ps1 -SiteUrl "https://uhb.sharepoint.com/teams/PharmacySafety"
#
# PREREQUISITES:
#   Install-Module PnP.PowerShell -Scope CurrentUser (if not already installed)
# ==============================================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "The URL of the target SharePoint Site (e.g. https://uhb.sharepoint.com/teams/PharmacySafety)")]
    [string]$SiteUrl,

    [string]$ListName = "Pharmacy Near Miss Log"
)

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " UHBFT Pharmacy - SharePoint List Provisioning Engine (M365)" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Target Site : $SiteUrl" -ForegroundColor Green
Write-Host " List Name   : $ListName" -ForegroundColor Green

# 1. Verify PnP.PowerShell
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Warning "PnP.PowerShell module is not installed."
    Write-Host "Installing PnP.PowerShell for CurrentUser..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

# 2. Connect to SharePoint Online (Interactive NHS M365 Authentication)
Write-Host "`nAuthenticating to SharePoint Online via Microsoft 365 Single Sign-On..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# 3. Check or Create List
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Creating list '$ListName'..." -ForegroundColor Green
    $list = New-PnPList -Title $ListName -Template GenericList -EnableVersioning
    Write-Host "List created successfully." -ForegroundColor Green
} else {
    Write-Host "List '$ListName' already exists. Adding/updating schema columns..." -ForegroundColor Yellow
}

# 4. Define Columns from Schema
$columns = @(
    @{ Name = "SubmittedAt"; Title = "Submitted At"; Type = "DateTime" },
    @{ Name = "Site"; Title = "Hospital Site"; Type = "Choice"; Choices = @("Queen Elizabeth", "Heartlands", "Good Hope", "Solihull", "Other") },
    @{ Name = "SiteOther"; Title = "Other Site Details"; Type = "Text" },
    @{ Name = "NearMissType"; Title = "Near Miss Type"; Type = "Choice"; Choices = @("Dispensing", "Clinical") },
    @{ Name = "IsSatelliteUnit"; Title = "Satellite Unit"; Type = "Choice"; Choices = @("Yes", "No") },
    @{ Name = "SatelliteUnitName"; Title = "Satellite Name"; Type = "Text" },
    @{ Name = "CompletedBy"; Title = "Completed By (Initials)"; Type = "Text" },
    @{ Name = "CompletedOn"; Title = "Completed On"; Type = "Text" },
    @{ Name = "TimeOfDay"; Title = "Time of Day"; Type = "Choice"; Choices = @("Before 9am", "9-10am", "10-11am", "11-12am", "12-1pm", "1-2pm", "2-3pm", "3-4pm", "4-5pm", "After 5pm") },
    @{ Name = "ErrorCategories"; Title = "Error Categories"; Type = "MultiChoice"; Choices = @(
        "Wrong Medication (Drug)", "Wrong Strength", "Wrong Form", "Wrong Brand (NTI)", "Wrong Quantity", "Out of Date",
        "Wrong Dose", "Wrong Frequency", "Wrong Duration", "Wrong Labelling Directions (Other)", "Wrong Patient Details", "Wrong Location / Ward",
        "Dose Error (Renal / Hepatic / Weight / Age)", "Wrong Dose / Strength / Frequency Prescribed", "Drug-Drug Interaction",
        "Allergy / Cross-Sensitivity", "Omission of Critical / Regular Medication", "Therapeutic Duplication", "Contraindication",
        "Wrong Formulation / Route Prescribed", "High-Risk / VTE / Anticoagulant / Insulin Error", "TDM / Monitoring / Other Clinical"
    )},
    @{ Name = "ErrorDrugName"; Title = "Error Medication Name"; Type = "Text" },
    @{ Name = "ErrorDrugCMM"; Title = "Error CMM Code"; Type = "Text" },
    @{ Name = "ErrorDrugVTMID"; Title = "Error VTMID"; Type = "Text" },
    @{ Name = "ErrorDrugVTMName"; Title = "Error Generic Molecule"; Type = "Text" },
    @{ Name = "ErrorDrugVMPID"; Title = "Error VMPID"; Type = "Text" },
    @{ Name = "ErrorDrugVMPName"; Title = "Error VMP Name"; Type = "Text" },
    @{ Name = "ErrorDrugAMPID"; Title = "Error AMPID"; Type = "Text" },
    @{ Name = "ErrorDrugAMPPID"; Title = "Error AMPPID"; Type = "Text" },
    @{ Name = "ErrorDrugGTIN14"; Title = "Error GTIN14 Barcode"; Type = "Text" },
    @{ Name = "ErrorDrugBatch"; Title = "Error Batch Number"; Type = "Text" },
    @{ Name = "ErrorDrugExpiry"; Title = "Error Expiry Date"; Type = "Text" },
    @{ Name = "ErrorDrugIsExpired"; Title = "Error Was Expired"; Type = "Choice"; Choices = @("Yes", "No") },
    @{ Name = "ErrorDrugIsNTI"; Title = "Error Is NTI Drug"; Type = "Choice"; Choices = @("Yes", "No") },
    @{ Name = "IntendedDrugName"; Title = "Intended Medication Name"; Type = "Text" },
    @{ Name = "IntendedDrugCMM"; Title = "Intended CMM Code"; Type = "Text" },
    @{ Name = "IntendedDrugVTMID"; Title = "Intended VTMID"; Type = "Text" },
    @{ Name = "IntendedDrugVMPID"; Title = "Intended VMPID"; Type = "Text" },
    @{ Name = "IntendedDrugAMPID"; Title = "Intended AMPID"; Type = "Text" },
    @{ Name = "IntendedDrugIsNTI"; Title = "Intended Is NTI Drug"; Type = "Choice"; Choices = @("Yes", "No") },
    @{ Name = "QuantityDispensed"; Title = "Quantity Dispensed"; Type = "Text" },
    @{ Name = "QuantityPrescribed"; Title = "Quantity Prescribed"; Type = "Text" },
    @{ Name = "ContributingFactors"; Title = "Contributing Factors"; Type = "MultiChoice"; Choices = @(
        "Similar Packaging / Look-Alike", "Sound-Alike Name", "Robot Pick Error", "Split Pack / Owings",
        "High Workload / Interruption", "Handwritten / Unclear Rx", "PICS Selection Error"
    )},
    @{ Name = "LearningNote"; Title = "Learning Note (Zero PID)"; Type = "Note" }
)

# 5. Add Each Field to the List
Write-Host "`nProvisioning schema fields..." -ForegroundColor Cyan
foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host " + Adding column: $($col.Title) ($($col.Type))" -ForegroundColor Green
        switch ($col.Type) {
            "Text" {
                Add-PnPField -List $ListName -InternalName $col.Name -DisplayName $col.Title -Type Text -AddToDefaultView | Out-Null
            }
            "DateTime" {
                Add-PnPField -List $ListName -InternalName $col.Name -DisplayName $col.Title -Type DateTime -AddToDefaultView | Out-Null
            }
            "Choice" {
                Add-PnPField -List $ListName -InternalName $col.Name -DisplayName $col.Title -Type Choice -Choices $col.Choices -AddToDefaultView | Out-Null
            }
            "MultiChoice" {
                Add-PnPField -List $ListName -InternalName $col.Name -DisplayName $col.Title -Type MultiChoice -Choices $col.Choices -AddToDefaultView | Out-Null
            }
            "Note" {
                Add-PnPField -List $ListName -InternalName $col.Name -DisplayName $col.Title -Type Note -AddToDefaultView | Out-Null
            }
        }
    } else {
        Write-Host " . Column exists: $($col.Title)" -ForegroundColor DarkGray
    }
}

Write-Host "`n================================================================================" -ForegroundColor Cyan
Write-Host " SharePoint List Provisioning Complete!                                         " -ForegroundColor Green
Write-Host " List URL: $SiteUrl/Lists/$($ListName -replace ' ', '%20')" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

