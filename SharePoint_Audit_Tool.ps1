<#
.SYNOPSIS
    Generates various SharePoint Online reports using PnP PowerShell and SPO Management Shell.
    Reports include site storage, inactive sites, sharing settings, duplicate files, large files, and access reviews.

.DESCRIPTION
    Offers a menu to choose which report to generate. Each report establishes its own connection via 
    the required modules and outputs the results to CSV.

    For the Inactive Sites Report, Duplicate Files Report, Large Files Report, and User Access Review,
    if a site returns an error stating "Attempted to perform an unauthorized operation", the script will 
    prompt whether you want to automatically add an admin account as Owner. If chosen, the specified admin 
    account is added to the site’s Owners and tracked. Once the report completes, the admin account is removed.

.NOTES
    Requires: PnP.PowerShell, Microsoft.Online.SharePoint.PowerShell, PowerShell v7.5+, Tenant Admin permissions
#>

#############################
### Module & Helper Functions
#############################

function Check-PnPModule {
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Host "🔍 PnP.PowerShell module is not installed." -ForegroundColor Yellow
        $install = Read-Host "Would you like to install it now? (Y/N): "
        if ($install.ToUpper() -eq "Y") {
            try {
                Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
                Write-Host "✅ PnP.PowerShell installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "❌ Failed to install PnP.PowerShell: $_" -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "🚫 Cannot continue without PnP.PowerShell. Exiting." -ForegroundColor Red
            exit
        }
    }
}

function Check-SPOModule {
    if (-not (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")) {
        Write-Host "🔍 SPO Management Shell module is not installed." -ForegroundColor Yellow
        $installSPO = Read-Host "Would you like to install it now? (Y/N): "
        if ($installSPO.ToUpper() -eq "Y") {
            try {
                Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
                Write-Host "✅ SPO module installed." -ForegroundColor Green
            } catch {
                Write-Host "❌ Failed to install SPO module: $_" -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "🚫 Cannot continue without SPO module. Exiting." -ForegroundColor Red
            exit
        }
    }
}

function Read-HostGreen {
    param([string]$Prompt)
    Write-Host $Prompt -ForegroundColor Green -NoNewline
    return Read-Host
}

function Get-SavePath {
    param(
        [string]$Title,
        [string]$DefaultFileName
    )
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $Title
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.FileName = $DefaultFileName

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        Write-Host "❌ Save cancelled. No file created." -ForegroundColor Yellow
        return $null
    }
}

function Check-For-Cancel {
    if ([System.Console]::KeyAvailable) {
        $key = [System.Console]::ReadKey($true)
        if ($key.Key -ieq 'q') {
            throw "UserCancelled"
        }
    }
}

#############################
### Global Variables, Getters & Tracking
#############################

$Global:ClientID = $null
$Global:SharePointAdminURL = $null
$Global:AdminPrivilegesAdded = @()

function Get-SharePointAdminURL {
    if (-not $Global:SharePointAdminURL) {
        $Global:SharePointAdminURL = Read-HostGreen "Enter the SharePoint Admin URL (e.g. https://contoso-admin.sharepoint.com): "
    }
    return $Global:SharePointAdminURL
}

function Get-ClientID {
    if (-not $Global:ClientID) {
        $Global:ClientID = Read-HostGreen "Enter your Client ID: "
    }
    return $Global:ClientID
}

#############################
### Function to Add Admin as Owner and Track the Site
#############################

function Add-OwnerAndRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteURL,
        [Parameter(Mandatory = $true)]
        [string]$ClientID,
        [Parameter(Mandatory = $true)]
        [string]$AdminAccount
    )
    Write-Host "🔐 Attempting to add $AdminAccount as a temporary Owner to $SiteURL..." -ForegroundColor Yellow
    try {
        Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
        $siteInfo = Get-PnPTenantSite -Identity $SiteURL
        $CurrentOwners = @()
        if ($siteInfo.Owners) { $CurrentOwners = $siteInfo.Owners }
        if ($CurrentOwners -notcontains $AdminAccount) {
            $NewOwners = $CurrentOwners + $AdminAccount
            Set-PnPTenantSite -Identity $SiteURL -Owners $NewOwners
        }
        $Global:AdminPrivilegesAdded += $SiteURL
        Write-Host "✅ $AdminAccount added as Owner to $SiteURL. Retrying site access..." -ForegroundColor Green
        Connect-PnPOnline -Url $SiteURL -ClientID $ClientID -Interactive
        $web = Get-PnPWeb
        return $web
    } catch {
        Write-Host "⚠️  Retried operation failed for ${SiteURL}: $_" -ForegroundColor Red
        return $null
    }
}

#############################
### Function to Remove Elevated Privileges from Tracked Sites
#############################

function Remove-ElevatedAdminPrivileges {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientID,
        [Parameter(Mandatory = $true)]
        [string]$AdminAccount
    )
    if ($Global:AdminPrivilegesAdded.Count -gt 0) {
        foreach ($site in $Global:AdminPrivilegesAdded) {
            try {
                Connect-PnPOnline -Url $site -ClientID $ClientID -Interactive
                $siteInfo = Get-PnPTenantSite -Identity $site
                $CurrentOwners = @()
                if ($siteInfo.Owners) { $CurrentOwners = $siteInfo.Owners }
                if ($CurrentOwners -contains $AdminAccount) {
                    $NewOwners = $CurrentOwners | Where-Object { $_ -ne $AdminAccount }
                    Set-PnPTenantSite -Identity $site -Owners $NewOwners
                    Write-Host "Removed $AdminAccount from Owners at $site" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "Could not remove elevated privileges from ${site}: $_" -ForegroundColor Red
            }
        }
        $Global:AdminPrivilegesAdded = @()
    }
}

#############################
### Encapsulated Function for File Scan Reports (Used in Duplicate and Large Files Reports)
#############################

function Process-Library {
    param (
        $SiteUrl,
        $LibraryTitle,
        $ClientID,
        $SizeThreshold,
        $OnlyLargeFiles,
        $OutputDuplicatesOnly
    )

    Check-For-Cancel
    Write-Host "Checking Library: $LibraryTitle" -ForegroundColor Yellow

    Check-For-Cancel
    $Files = Get-PnPListItem -List $LibraryTitle -PageSize 500 -Fields "FileLeafRef", "FileRef", "Modified", "Editor", "Author", "FSObjType", "SMTotalFileStreamSize"

    $FileHashTable = @{}

    foreach ($file in $Files) {
        Check-For-Cancel

        if ($file["FSObjType"] -eq 0) {
            $fileName = $file["FileLeafRef"]
            $fileSize = $file.FieldValues["SMTotalFileStreamSize"]

            if (-not $fileSize) { continue }

            $fileSize = [int64]$fileSize

            if ($OnlyLargeFiles -eq "Y" -and $fileSize -lt $SizeThreshold) {
                continue
            }

            $filePath = $file["FileRef"]

            $fileInfo = [PSCustomObject]@{
                SiteURL         = $SiteUrl
                Library         = $LibraryTitle
                FileName        = $fileName
                "File Size (MB)"= [math]::Round($fileSize / 1MB, 2)
                LastModified    = $file["Modified"]
                CreatedBy       = $file["Author"].LookupValue
                ModifiedBy      = $file["Editor"].LookupValue
                FolderLocation  = $filePath
            }

            if ($OutputDuplicatesOnly -eq "Y") {
                $hashKey = "$fileName|$fileSize"
                if (-not $FileHashTable.ContainsKey($hashKey)) {
                    $FileHashTable[$hashKey] = @()
                }
                $FileHashTable[$hashKey] += $fileInfo
            } else {
                $Global:FileResults.Add($fileInfo)
            }
        }
    }

    if ($OutputDuplicatesOnly -eq "Y") {
        $duplicates = $FileHashTable.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
        foreach ($duplicate in $duplicates) {
            Check-For-Cancel
            foreach ($fileInfo in $duplicate.Value) {
                $Global:FileResults.Add([PSCustomObject]@{
                    "Duplicate File"    = $fileInfo.FileName
                    "Occurrences"       = $duplicate.Value.Count
                    "File Size (MB)"    = $fileInfo."File Size (MB)"
                    "Site URL"          = $fileInfo.SiteURL
                    "Library"           = $fileInfo.Library
                    "Last Modified"     = $fileInfo.LastModified
                    "Modified By"       = $fileInfo.ModifiedBy
                    "Created By"        = $fileInfo.CreatedBy
                    "Folder Location"   = $fileInfo.FolderLocation
                })
            }
        }
    }
}

function Process-FileScanReport {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Duplicates", "AllFiles")]
        [string]$ReportMode,

        [Parameter(Mandatory = $true)]
        [string]$ClientID,

        [Parameter(Mandatory = $false)]
        [string]$AdminAccount,

        [Parameter(Mandatory = $true)]
        [string]$OutputCSV,

        [Parameter(Mandatory = $true)]
        [string]$ErrorSitesCSV
    )

    if ($ReportMode -eq "Duplicates") {
        $OutputDuplicatesOnly = "Y"
        $OnlyLargeFiles       = (Read-HostGreen "`nOnly include files over a certain size? (Y/N): ").ToUpper()
        $MinFileSizeMB        = 0
        if ($OnlyLargeFiles -eq "Y") {
            $MinFileSizeInput = Read-HostGreen "Enter minimum file size in MB: "
            if ([int]::TryParse($MinFileSizeInput, [ref]$null)) {
                $MinFileSizeMB = [int]$MinFileSizeInput
            } else {
                Write-Host "Invalid number entered. Defaulting to 100MB." -ForegroundColor Yellow
                $MinFileSizeMB = 100
            }
        }
    }
    else {
        $OutputDuplicatesOnly = "N"
        $OnlyLargeFiles       = "Y"
        $MinFileSizeInput     = Read-HostGreen "Enter minimum file size in MB: "
        if ([int]::TryParse($MinFileSizeInput, [ref]$null)) {
            $MinFileSizeMB = [int]$MinFileSizeInput
        } else {
            Write-Host "Invalid number entered. Defaulting to 100MB." -ForegroundColor Yellow
            $MinFileSizeMB = 100
        }
    }

    $SearchAllLibraries = (Read-HostGreen "`nSearch ALL document libraries (Y) or only Documents (N)? ").ToUpper()
    $SizeThreshold      = $MinFileSizeMB * 1MB

    $Global:FileResults = [System.Collections.Generic.List[object]]::new()
    $ErrorSites         = [System.Collections.Generic.List[object]]::new()

    $AllSites = Get-PnPTenantSite | Where-Object {
        $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*"
    }

    $totalSites = $AllSites.Count
    $i = 0

    foreach ($Site in $AllSites) {
        $i++
        Write-Host "`n--- Processing Site ($i of $totalSites): $($Site.Url) ---" -ForegroundColor Cyan

        try {
            Connect-PnPOnline -Url $Site.Url -ClientID $ClientID -Interactive

            $DocumentLibraries = Get-PnPList | Where-Object {
                $_.BaseType -eq "DocumentLibrary" -and
                $_.Hidden   -eq $false -and
                ($SearchAllLibraries -eq "Y" -or $_.Title -eq "Documents")
            }

            foreach ($Library in $DocumentLibraries) {
                Process-Library `
                    -SiteUrl         $Site.Url `
                    -LibraryTitle    $Library.Title `
                    -ClientID        $ClientID `
                    -SizeThreshold   $SizeThreshold `
                    -OnlyLargeFiles  $OnlyLargeFiles `
                    -OutputDuplicatesOnly $OutputDuplicatesOnly
            }
        }
        catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    throw $_
                }

                $errorMessage = $_.Exception.Message

            if ( ($errorMessage -match "unauthorized operation|Access is denied") -and $AdminAccount ) {
                $web = Add-OwnerAndRetry `
                    -SiteURL      $Site.Url `
                    -ClientID     $ClientID `
                    -AdminAccount $AdminAccount

                if ($web) {
                    Write-Host "🔁 Re-running file scan on $($Site.Url) after elevation…" -ForegroundColor Cyan

                    Connect-PnPOnline -Url $Site.Url -ClientID $ClientID -Interactive
                    $DocumentLibraries = Get-PnPList | Where-Object {
                        $_.BaseType -eq "DocumentLibrary" -and
                        $_.Hidden   -eq $false -and
                        ($SearchAllLibraries -eq "Y" -or $_.Title -eq "Documents")
                    }
                    foreach ($Library in $DocumentLibraries) {
                        Process-Library `
                            -SiteUrl         $Site.Url `
                            -LibraryTitle    $Library.Title `
                            -ClientID        $ClientID `
                            -SizeThreshold   $SizeThreshold `
                            -OnlyLargeFiles  $OnlyLargeFiles `
                            -OutputDuplicatesOnly $OutputDuplicatesOnly
                    }
                }
                else {
                    Write-Host "⚠️  Skipping $($Site.Url): unable to elevate and access." -ForegroundColor Red
                    $ErrorSites.Add([PSCustomObject]@{
                        SiteURL      = $Site.Url
                        ErrorMessage = "Elevation failed"
                    })
                }
            }
            else {
                Write-Host "⚠️  Skipping $($Site.Url): $errorMessage" -ForegroundColor Red
                $ErrorSites.Add([PSCustomObject]@{
                    SiteURL      = $Site.Url
                    ErrorMessage = $errorMessage
                })
            }
        }
    }

    return @{ 
        "Results" = $Global:FileResults
        "Errors"  = $ErrorSites 
    }
}

#############################
### Main Menu
#############################

function Show-MainMenu {
    Write-Host "`nChoose a report to generate:" -ForegroundColor Cyan
    Write-Host "1. Site Storage Report"
    Write-Host "2. Inactive Sites Report"
    Write-Host "3. Sharing Settings Audit"
    Write-Host "4. Duplicate Files Report"
    Write-Host "5. Large Files Report"
    Write-Host "6. User Access Review"
    Write-Host "7. External Users Access Report"
    Write-Host "8. Update Connection Settings"
    Write-Host "0. Exit"
    Write-Host "`n💡 Tip: Press Q at any time to cancel and return to the main menu." -ForegroundColor DarkCyan
}

$continue = $true
while ($continue) {
    Show-MainMenu
    $choice = Read-Host "`nEnter choice number (0-8)"
    switch ($choice) {
        '1' {
            try {
                Check-SPOModule
                Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
                Connect-SPOService -Url (Get-SharePointAdminURL)

                $outPath = Get-SavePath -Title "Save Site Storage Report" -DefaultFileName "Site_Storage_Report.csv"
                if (-not $outPath) { return }

                $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "
                if ($includeOneDrive.ToUpper() -eq 'Y') {
                    $allSites = Get-SPOSite -IncludePersonalSite $true -Limit All
                } else {
                    $allSites = Get-SPOSite -Limit All
                }

                $report = @()
                $i = 0
                foreach ($site in $allSites) {
                    Check-For-Cancel
                    $i++
                    Write-Host "Processing site $i of $($allSites.Count): $($site.Url)" -ForegroundColor Cyan
                    $report += [PSCustomObject]@{
                        SiteUrl        = $site.Url
                        StorageUsedMB  = [math]::Round($site.StorageUsageCurrent, 2)
                        StorageQuotaMB = [math]::Round($site.StorageQuota, 2)
                        PercentUsed    = if ($site.StorageQuota -gt 0) { [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 2) } else { 'N/A' }
                    }
                }

                $total = ($allSites | Measure-Object -Property StorageUsageCurrent -Sum).Sum
                $report += [PSCustomObject]@{
                    SiteUrl        = '*** Total ***'
                    StorageUsedMB  = [math]::Round($total,2)
                    StorageQuotaMB = ''
                    PercentUsed    = ''
                }

                $report | Export-Csv -Path $outPath -NoTypeInformation
                $msgSuffix = if ($includeOneDrive.ToUpper() -eq 'Y') { ' (including OneDrive)' } else { '' }
                Write-Host "`n✅ Site storage report$msgSuffix with totals saved to: $outPath" -ForegroundColor Green
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan
                    continue
                } else {
                    throw $_
                }
            }
        }
        '2' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $autoAddOwner = Read-HostGreen "`nWould you like to temporarily add an admin account as Owner on unauthorized sites? (Y/N): "
                $AdminAccount = if ($autoAddOwner.ToUpper() -eq "Y") {
                    Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
                } else {
                    $null
                }

                $AllSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
                $outPath = Get-SavePath -Title "Save Inactive Sites Report" -DefaultFileName "Inactive_Sites.csv"
                if (-not $outPath) { return }

                $InactiveSites = @()
                $i = 0
                foreach ($site in $AllSites) {
                    Check-For-Cancel
                    $i++
                    Write-Host "Scanning site ($i of $($AllSites.Count)): $($site.Url)" -ForegroundColor Cyan

                    try {
                        Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                        $web = Get-PnPWeb -Includes LastItemUserModifiedDate
                        $InactiveSites += [PSCustomObject]@{
                            SiteUrl      = $site.Url
                            LastModified = $web.LastItemUserModifiedDate
                        }
                    } catch {
                        if (($_ | Out-String) -match "unauthorized operation" -and $AdminAccount) {
                            $web = Add-OwnerAndRetry -SiteURL $site.Url -ClientID $ClientID -AdminAccount $AdminAccount
                            if ($web) {
                                $InactiveSites += [PSCustomObject]@{
                                    SiteUrl      = $site.Url
                                    LastModified = $web.LastItemUserModifiedDate
                                }
                            }
                        } else {
                            Write-Host "⚠️  Skipping site $($site.Url) due to error: $_" -ForegroundColor Red
                        }
                    }
                }

                $InactiveSites | Export-Csv -Path $outPath -NoTypeInformation
                Write-Host "`n✅ Inactive sites report saved to: $outPath" -ForegroundColor Green

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                    Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                }
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan

                    if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                        Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                    }

                    continue
                } else {
                    throw $_
                }
            }
        }
        '3' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $AllSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
                $outPath = Get-SavePath -Title "Save Sharing Settings Report" -DefaultFileName "Sharing_Settings.csv"
                if (-not $outPath) { return }

                $SharingReport = @()
                $i = 0
                foreach ($site in $AllSites) {
                    Check-For-Cancel
                    $i++
                    Write-Host "Checking site ($i of $($AllSites.Count)): $($site.Url)" -ForegroundColor Cyan
                    $SharingReport += [PSCustomObject]@{
                        Url              = $site.Url
                        SharingCapability = $site.SharingCapability
                    }
                }

                $SharingReport | Export-Csv -Path $outPath -NoTypeInformation
                Write-Host "`n✅ Sharing settings report saved to: $outPath" -ForegroundColor Green
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan
                    continue
                } else {
                    throw $_
                }
            }
        }
        '4' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $OutputCSV = Get-SavePath -Title "Save Duplicate Files Report" -DefaultFileName "Duplicate_Files_Report.csv"
                if (-not $OutputCSV) { return }

                $ErrorSitesCSV = Get-SavePath -Title "Save Sites That Errored" -DefaultFileName "Errored_Sites.csv"
                if (-not $ErrorSitesCSV) { return }

                $autoAddOwner = Read-HostGreen "`nWould you like to temporarily add an admin account as Owner on unauthorized sites? (Y/N): "
                $AdminAccount = if ($autoAddOwner.ToUpper() -eq "Y") {
                    Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
                } else {
                    $null
                }

                $fileScanResult = Process-FileScanReport `
                    -ReportMode "Duplicates" `
                    -ClientID $ClientID `
                    -AdminAccount $AdminAccount `
                    -OutputCSV $OutputCSV `
                    -ErrorSitesCSV $ErrorSitesCSV

                # Only export if the scan was completed (not interrupted)
                if ($fileScanResult -and $fileScanResult.Results) {
                    $fileScanResult.Results | Export-Csv -Path $OutputCSV -NoTypeInformation
                    $fileScanResult.Errors | Export-Csv -Path $ErrorSitesCSV -NoTypeInformation
                    Write-Host "`n✅ Duplicate Files Report saved to: $OutputCSV" -ForegroundColor Green
                }

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                    Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                }
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                    Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                }

                    continue
                } else {
                    throw $_
                }
            }
        }
        '5' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $OutputCSV = Get-SavePath -Title "Save Large Files Report" -DefaultFileName "Large_Files_Report.csv"
                if (-not $OutputCSV) { return }

                $ErrorSitesCSV = Get-SavePath -Title "Save Sites That Errored" -DefaultFileName "Errored_Sites.csv"
                if (-not $ErrorSitesCSV) { return }

                $autoAddOwner = Read-HostGreen "`nWould you like to temporarily add an admin account as Owner on unauthorized sites? (Y/N): "
                $AdminAccount = if ($autoAddOwner.ToUpper() -eq "Y") {
                    Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
                } else {
                    $null
                }

                $fileScanResult = Process-FileScanReport `
                    -ReportMode "AllFiles" `
                    -ClientID $ClientID `
                    -AdminAccount $AdminAccount `
                    -OutputCSV $OutputCSV `
                    -ErrorSitesCSV $ErrorSitesCSV

                # Only export if scan wasn’t interrupted
                if ($fileScanResult -and $fileScanResult.Results) {
                    $fileScanResult.Results | Export-Csv -Path $OutputCSV -NoTypeInformation
                    $fileScanResult.Errors | Export-Csv -Path $ErrorSitesCSV -NoTypeInformation
                    Write-Host "`n✅ Large Files Report saved to: $OutputCSV" -ForegroundColor Green
                }

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                    Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                }
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                    Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                }

                    continue
                } else {
                    throw $_
                }
            }
        }
        '6' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $autoAddOwner = Read-HostGreen "`nWould you like to temporarily add an admin account as Owner on unauthorized sites? (Y/N): "
                $AdminAccount = if ($autoAddOwner.ToUpper() -eq "Y") {
                    Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
                } else {
                    $null
                }

                $outPath = Get-SavePath -Title "Save User Access Review Report" -DefaultFileName "Access_Review_Report.csv"
                if (-not $outPath) { return }

                $AllSites     = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
                $totalSites   = $AllSites.Count
                $i            = 0
                $AccessReport = @()

                foreach ($site in $AllSites) {
                    Check-For-Cancel
                    $i++
                    Write-Host "`n--- Processing Site ($i of $totalSites): $($site.Url) ---" -ForegroundColor Cyan

                    try {
                        Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                        $groups = Get-PnPGroup
                        foreach ($group in $groups) {
                            $members = Get-PnPGroupMember -Group $group
                            foreach ($member in $members) {
                                $AccessReport += [PSCustomObject]@{
                                    SiteUrl   = $site.Url
                                    GroupName = $group.Title
                                    UserName  = $member.Title
                                    UserLogin = $member.LoginName
                                    UserEmail = $member.Email
                                }
                            }
                        }
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        if ( ($errorMessage -match "unauthorized operation|Access is denied") -and $AdminAccount ) {
                            $web = Add-OwnerAndRetry -SiteURL $site.Url -ClientID $ClientID -AdminAccount $AdminAccount
                            if ($web) {
                                Write-Host "🔁 Re-running access review on $($site.Url) after elevation…" -ForegroundColor Cyan
                                Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                                $groups = Get-PnPGroup
                                foreach ($group in $groups) {
                                    $members = Get-PnPGroupMember -Group $group
                                    foreach ($member in $members) {
                                        $AccessReport += [PSCustomObject]@{
                                            SiteUrl   = $site.Url
                                            GroupName = $group.Title
                                            UserName  = $member.Title
                                            UserLogin = $member.LoginName
                                            UserEmail = $member.Email
                                        }
                                    }
                                }
                            }
                            else {
                                Write-Host "⚠️  Skipping $($site.Url): unable to elevate and access." -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "⚠️  Skipping $($site.Url): $errorMessage" -ForegroundColor Red
                        }
                    }
                }

                # Export and cleanup
                if ($AccessReport.Count -gt 0) {
                    $AccessReport | Export-Csv -Path $outPath -NoTypeInformation
                    Write-Host "`n✅ User access review report saved to: $outPath" -ForegroundColor Green
                }

                if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
            }
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan

                    if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                        Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
                    }

                    continue
                } else {
                    throw $_
                }
            }
        }
        '7' {
            try {
                Check-PnPModule
                $ClientID = Get-ClientID
                Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

                $outPath = Get-SavePath -Title "Save External Users Report" -DefaultFileName "External_Users_Report.csv"
                if (-not $outPath) { return }

                $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "

                $AllSites = if ($includeOneDrive.ToUpper() -eq 'Y') {
                    Get-PnPTenantSite -IncludeOneDriveSites
                } else {
                    Get-PnPTenantSite
                }

                $totalSites     = $AllSites.Count
                $i              = 0
                $ExternalReport = @()

                foreach ($site in $AllSites) {
                    Check-For-Cancel
                    $i++
                    Write-Host "`n--- Processing Site ($i of $totalSites): $($site.Url) ---" -ForegroundColor Cyan

                    try {
                        Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                        $extUsers = Get-PnPExternalUser -PageSize 50 -SiteUrl $site.Url
                        foreach ($u in $extUsers) {
                            $ExternalReport += [PSCustomObject]@{
                                SiteUrl     = $site.Url
                                DisplayName = $u.DisplayName
                                Email       = $u.Email
                                AcceptedAs  = $u.AcceptedAs
                                WhenCreated = $u.WhenCreated
                                InvitedBy   = $u.InvitedBy
                                InvitedAs   = $u.InvitedAs
                            }
                        }
                    }
                    catch {
                        Write-Warning "⚠️  Failed on $($site.Url): $_"
                    }
                }

                if ($ExternalReport.Count -gt 0) {
                    $ExternalReport | Export-Csv -Path $outPath -NoTypeInformation
                    $suffix = if ($includeOneDrive.ToUpper() -eq 'Y') { ' (including OneDrive)' } else { '' }
                    Write-Host "`n✅ External users‑by‑site report$suffix saved to: $outPath" -ForegroundColor Green
                }
            } catch {
                if ($_.Exception.Message -eq 'UserCancelled') {
                    Write-Host "`n↩️  Cancelled — returning to main menu..." -ForegroundColor Cyan
                    continue
                } else {
                    throw $_
                }
}
        }
        '8' {
            Write-Host "Updating connection settings..." -ForegroundColor Cyan
            $newURL = Read-HostGreen "Enter the new SharePoint Admin URL (e.g. https://newtenant-admin.sharepoint.com): "
            if ($newURL -ne $Global:SharePointAdminURL) {
                $Global:SharePointAdminURL = $newURL
                $Global:ClientID = $null  # Clear stored Client ID to force re-entry
                Write-Host "Connection settings updated. Client ID has been reset and will be requested on next use." -ForegroundColor Green
            } else {
                Write-Host "The URL entered is the same as the current connection. No changes made." -ForegroundColor Yellow
            }
            continue
        }
        '0' {
            Write-Host "Exiting..." -ForegroundColor Yellow
            $continue = $false
        }
        default {
            Write-Host "`n❌ Invalid choice. Please choose an option between 0 and 8." -ForegroundColor Red
        }
    }
    if (($choice -ne '0') -and ($choice -ne '8')) {
        $return = Read-Host "`nWould you like to return to the main menu? (Y/N): "
        if ($return.ToUpper() -ne 'Y') {
            $continue = $false
        }
    }
}
