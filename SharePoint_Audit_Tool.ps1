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

#############################
### Global Variables and Getters & Tracking
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
    Write-Host "🔐 Attempting to add $AdminAccount as an Owner to $SiteURL..." -ForegroundColor Yellow
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
}

$continue = $true
while ($continue) {
    Show-MainMenu
    $choice = Read-Host "`nEnter choice number (0-8)"
    switch ($choice) {
        '1' {
            Check-SPOModule
            Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
            try {
                Connect-SPOService -Url (Get-SharePointAdminURL)
            } catch {
                Write-Host "❌ Failed to connect to SPO: $_" -ForegroundColor Red
                continue
            }
            $outPath = Read-Host "`nEnter the path to save the storage report (e.g. C:\SiteStorageReport.csv)"
            $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "
            if ($includeOneDrive.ToUpper() -eq 'Y') {
                $allSites = Get-SPOSite -IncludePersonalSite $true -Limit All
            } else {
                $allSites = Get-SPOSite -Limit All
            }
            $report = $allSites | Select-Object `
                @{Name='SiteUrl';        Expression={$_.Url}}, `
                @{Name='StorageUsedMB';  Expression={[math]::Round($_.StorageUsageCurrent,2)}}, `
                @{Name='StorageQuotaMB'; Expression={[math]::Round($_.StorageQuota,2)}}, `
                @{Name='PercentUsed';    Expression={ if ($_.StorageQuota -gt 0) {[math]::Round(($_.StorageUsageCurrent / $_.StorageQuota)*100,2)} else {'N/A'} }}
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
}
        '2' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
            $autoAddOwner = Read-HostGreen "`nWould you like to automatically add an admin account as Owner on unauthorized sites? (Y/N): "
            if ($autoAddOwner.ToUpper() -eq "Y") {
                $adminAccountForAuto = Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
            } else {
                $adminAccountForAuto = $null
            }
            $AllSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
            $outPath = Read-Host "`nEnter the path to save the inactive sites report (e.g. C:\InactiveSites.csv)"
            $InactiveSites = @()
            foreach ($site in $AllSites) {
                Write-Host "Scanning site: $($site.Url)" -ForegroundColor Cyan
                try {
                    Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                    $web = Get-PnPWeb -Includes LastItemUserModifiedDate
                    $InactiveSites += [PSCustomObject]@{
                        SiteUrl      = $site.Url
                        LastModified = $web.LastItemUserModifiedDate
                    }
                } catch {
                    if (($_ | Out-String) -match "unauthorized operation" -and $adminAccountForAuto) {
                        $web = Add-OwnerAndRetry -SiteURL $site.Url -ClientID $ClientID -AdminAccount $adminAccountForAuto
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

            if ($adminAccountForAuto -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $adminAccountForAuto
            }
        }
        '3' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
            $AllSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
            $outPath = Read-Host "`nEnter the path to save the sharing settings report (e.g. C:\SharingSettings.csv)"
            $AllSites | Select-Object Url, SharingCapability | Export-Csv -Path $outPath -NoTypeInformation
            Write-Host "`n✅ Sharing settings report saved to: $outPath" -ForegroundColor Green
        }
        '4' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
            $OutputCSV = Read-HostGreen "`nEnter the path to export the Duplicate Files Report CSV (e.g. C:\DuplicateFilesReport.csv): "
            $ErrorSitesCSV = Read-HostGreen "`nEnter the path to export the errored sites CSV (e.g. C:\SitesThatErrored.csv): "
            $autoAddOwner = Read-HostGreen "`nWould you like to automatically add an admin account as Owner on unauthorized sites? (Y/N): "
            if ($autoAddOwner.ToUpper() -eq "Y") {
                $AdminAccount = Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
            } else {
                $AdminAccount = $null
            }
            $fileScanResult = Process-FileScanReport -ReportMode "Duplicates" -ClientID $ClientID -AdminAccount $AdminAccount -OutputCSV $OutputCSV -ErrorSitesCSV $ErrorSitesCSV
            $fileScanResult.Results | Export-Csv -Path $OutputCSV -NoTypeInformation
            $fileScanResult.Errors | Export-Csv -Path $ErrorSitesCSV -NoTypeInformation
            Write-Host "`n✅ Duplicate Files Report saved to: $OutputCSV" -ForegroundColor Green

            if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
            }
        }
        '5' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
            $OutputCSV = Read-HostGreen "`nEnter the path to export the Large Files Report CSV (e.g. C:\LargeFilesReport.csv): "
            $ErrorSitesCSV = Read-HostGreen "`nEnter the path to export the errored sites CSV (e.g. C:\SitesThatErrored.csv): "
            $autoAddOwner = Read-HostGreen "`nWould you like to automatically add an admin account as Owner on unauthorized sites? (Y/N): "
            if ($autoAddOwner.ToUpper() -eq "Y") {
                $AdminAccount = Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
            } else {
                $AdminAccount = $null
            }
            $fileScanResult = Process-FileScanReport -ReportMode "AllFiles" -ClientID $ClientID -AdminAccount $AdminAccount -OutputCSV $OutputCSV -ErrorSitesCSV $ErrorSitesCSV
            $fileScanResult.Results | Export-Csv -Path $OutputCSV -NoTypeInformation
            $fileScanResult.Errors | Export-Csv -Path $ErrorSitesCSV -NoTypeInformation
            Write-Host "`n✅ Large Files Report saved to: $OutputCSV" -ForegroundColor Green

            if ($AdminAccount -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount
            }
        }
        '6' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

            $autoAddOwner = Read-HostGreen "`nWould you like to automatically add an admin account as Owner on unauthorized sites? (Y/N): "
            if ($autoAddOwner.ToUpper() -eq "Y") {
                $adminAccountForAuto = Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
            } else {
                $adminAccountForAuto = $null
            }

            $outPath      = Read-Host "`nEnter the path to save the user access review report (e.g. C:\AccessReview.csv)"
            $AllSites     = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
            $totalSites   = $AllSites.Count
            $i            = 0
            $AccessReport = @()

            foreach ($site in $AllSites) {
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
                    if ( ($errorMessage -match "unauthorized operation|Access is denied") -and $adminAccountForAuto ) {
                        $web = Add-OwnerAndRetry -SiteURL $site.Url -ClientID $ClientID -AdminAccount $adminAccountForAuto
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
            $AccessReport | Export-Csv -Path $outPath -NoTypeInformation
            Write-Host "`n✅ User access review report saved to: $outPath" -ForegroundColor Green

            if ($adminAccountForAuto -and $Global:AdminPrivilegesAdded.Count -gt 0) {
                Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $adminAccountForAuto
            }
        }
        '7' {
            Check-PnPModule
            $ClientID = Get-ClientID
            Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive

            $outPath        = Read-Host "`nEnter the path to save the external users‑by‑site report (e.g. C:\ExternalUsers.csv)"
            $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "

            if ($includeOneDrive.ToUpper() -eq 'Y') {
                $AllSites = Get-PnPTenantSite -IncludeOneDriveSites
            } else {
                $AllSites = Get-PnPTenantSite
            }

            $totalSites     = $AllSites.Count
            $i              = 0
            $ExternalReport = @()

            foreach ($site in $AllSites) {
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
                    Write-Warning "Failed on $($site.Url): $_"
                }
            }

            $ExternalReport | Export-Csv -Path $outPath -NoTypeInformation
            $suffix = if ($includeOneDrive.ToUpper() -eq 'Y') { ' (including OneDrive)' } else { '' }
            Write-Host "`n✅ External users‑by‑site report$suffix saved to: $outPath" -ForegroundColor Green
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
