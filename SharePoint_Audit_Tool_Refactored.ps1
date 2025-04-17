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
### Global Variables & Getters
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
### Elevation Helpers
#############################

function Add-OwnerAndRetry {
    param(
        [string]$SiteURL,
        [string]$ClientID,
        [string]$AdminAccount
    )
    Write-Host "🔐 Adding $AdminAccount as Owner to $SiteURL..." -ForegroundColor Yellow
    try {
        Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $ClientID -Interactive
        $siteInfo = Get-PnPTenantSite -Identity $SiteURL
        $owners = if ($siteInfo.Owners) { $siteInfo.Owners } else { @() }
        if ($owners -notcontains $AdminAccount) {
            Set-PnPTenantSite -Identity $SiteURL -Owners ($owners + $AdminAccount)
        }
        $Global:AdminPrivilegesAdded += $SiteURL
        Write-Host "✅ $AdminAccount added as Owner." -ForegroundColor Green
        Connect-PnPOnline -Url $SiteURL -ClientID $ClientID -Interactive
        return Get-PnPWeb
    } catch {
        Write-Host "⚠️ Retry failed for $SiteURL: $_" -ForegroundColor Red
        return $null
    }
}

function Remove-ElevatedAdminPrivileges {
    param(
        [string]$ClientID,
        [string]$AdminAccount
    )
    if ($Global:AdminPrivilegesAdded.Count -gt 0) {
        foreach ($site in $Global:AdminPrivilegesAdded) {
            try {
                Connect-PnPOnline -Url $site -ClientID $ClientID -Interactive
                $siteInfo = Get-PnPTenantSite -Identity $site
                $owners = if ($siteInfo.Owners) { $siteInfo.Owners } else { @() }
                if ($owners -contains $AdminAccount) {
                    Set-PnPTenantSite -Identity $site -Owners ($owners | Where-Object { $_ -ne $AdminAccount })
                    Write-Host "Removed $AdminAccount from Owners at $site" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "Could not remove elevated privileges from $site: $_" -ForegroundColor Red
            }
        }
        $Global:AdminPrivilegesAdded = @()
    }
}

#############################
### New Reusable Helpers
#############################

function Initialize-PnPContext {
    param([switch]$EnsureModule = $true)
    if ($EnsureModule) { Check-PnPModule }
    $client = Get-ClientID
    Connect-PnPOnline -Url (Get-SharePointAdminURL) -ClientID $client -Interactive
    return $client
}

function Prompt-Elevation {
    param([string]$Prompt = "Automatically add an admin account as Owner on unauthorized sites?")
    $resp = Read-HostGreen "`n$Prompt (Y/N): "
    if ($resp.ToUpper() -eq 'Y') {
        return Read-HostGreen "Enter the admin UPN (e.g. admin@contoso.com): "
    }
    return $null
}

function Invoke-TenantSitesAction {
    param(
        [ScriptBlock]$Action,
        [string]$ClientID,
        [string]$AdminAccount,
        [switch]$IncludeOneDrive
    )
    $exclude = @("SPSPERS*","APPCATALOG*")
    $sites = if ($IncludeOneDrive) {
        Get-PnPTenantSite -IncludeOneDriveSites
    } else {
        Get-PnPTenantSite
    } | Where-Object { $exclude -notcontains $_.Template }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($site in $sites) {
        Write-Host "`n--- Processing Site: $($site.Url) ---" -ForegroundColor Cyan
        try {
            Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
            $results.Add(& $Action $site $ClientID)
        } catch {
            $msg = $_.Exception.Message
            if (($msg -match "unauthorized operation|Access is denied") -and $AdminAccount) {
                $web = Add-OwnerAndRetry -SiteURL $site.Url -ClientID $ClientID -AdminAccount $AdminAccount
                if ($web) {
                    Write-Host "🔁 Re-running action on $($site.Url) after elevation…" -ForegroundColor Cyan
                    Connect-PnPOnline -Url $site.Url -ClientID $ClientID -Interactive
                    $results.Add(& $Action $site $ClientID)
                } else {
                    Write-Host "⚠️ Skipping $($site.Url): elevation failed" -ForegroundColor Red
                }
            } else {
                Write-Host "⚠️ Skipping $($site.Url): $msg" -ForegroundColor Red
            }
        }
    }
    return $results
}

function Write-CsvAndNotify {
    param(
        $Data,
        [string]$Path,
        [string]$Label
    )
    $Data | Export-Csv -Path $Path -NoTypeInformation
    Write-Host "`n✅ $Label saved to: $Path" -ForegroundColor Green
}

#############################
### Encapsulated File Scan Function
#############################

function Process-FileScanReport {
    param(
        [ValidateSet("Duplicates","AllFiles")] [string]$ReportMode,
        [string]$ClientID,
        [string]$AdminAccount,
        [string]$OutputCSV,
        [string]$ErrorSitesCSV
    )

    if ($ReportMode -eq "Duplicates") {
        $OutputDuplicatesOnly = "Y"
        $OnlyLargeFiles       = (Read-HostGreen "`nOnly include files over a certain size? (Y/N): ").ToUpper()
        $MinFileSizeMB        = 0
        if ($OnlyLargeFiles -eq "Y") {
            $input = Read-HostGreen "Enter minimum file size in MB: "
            if ([int]::TryParse($input,[ref]$null)) { $MinFileSizeMB = [int]$input } else { Write-Host "Invalid, defaulting to 100MB." -ForegroundColor Yellow; $MinFileSizeMB = 100 }
        }
    } else {
        $OutputDuplicatesOnly = "N"
        $OnlyLargeFiles       = "Y"
        $input                = Read-HostGreen "`nEnter minimum file size in MB: "
        if ([int]::TryParse($input,[ref]$null)) { $MinFileSizeMB = [int]$input } else { Write-Host "Invalid, defaulting to 100MB." -ForegroundColor Yellow; $MinFileSizeMB = 100 }
    }
    $SizeThreshold = $MinFileSizeMB * 1MB
    $SearchAllLibraries = (Read-HostGreen "`nSearch ALL document libraries (Y) or only Documents (N)? ").ToUpper()

    $Global:FileResults = [System.Collections.Generic.List[object]]::new()
    $ErrorSites         = [System.Collections.Generic.List[object]]::new()

    $AllSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
    foreach ($Site in $AllSites) {
        Write-Host "`n--- Processing Site: $($Site.Url) ---" -ForegroundColor Cyan
        try {
            Connect-PnPOnline -Url $Site.Url -ClientID $ClientID -Interactive
            $libs = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and ($SearchAllLibraries -eq "Y" -or $_.Title -eq "Documents") }
            foreach ($lib in $libs) {
                Process-Library -SiteUrl $Site.Url -LibraryTitle $lib.Title -ClientID $ClientID -SizeThreshold $SizeThreshold -OnlyLargeFiles $OnlyLargeFiles -OutputDuplicatesOnly $OutputDuplicatesOnly
            }
        } catch {
            $err = $_.Exception.Message
            if (($err -match "unauthorized operation|Access is denied") -and $AdminAccount) {
                $web = Add-OwnerAndRetry -SiteURL $Site.Url -ClientID $ClientID -AdminAccount $AdminAccount
                if ($web) {
                    Write-Host "🔁 Re-running file scan on $($Site.Url) after elevation…" -ForegroundColor Cyan
                    Connect-PnPOnline -Url $Site.Url -ClientID $ClientID -Interactive
                    $libs = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and ($SearchAllLibraries -eq "Y" -or $_.Title -eq "Documents") }
                    foreach ($lib in $libs) {
                        Process-Library -SiteUrl $Site.Url -LibraryTitle $lib.Title -ClientID $ClientID -SizeThreshold $SizeThreshold -OnlyLargeFiles $OnlyLargeFiles -OutputDuplicatesOnly $OutputDuplicatesOnly
                    }
                } else {
                    $ErrorSites.Add([PSCustomObject]@{SiteURL=$Site.Url;ErrorMessage="Elevation failed"})
                }
            } else {
                $ErrorSites.Add([PSCustomObject]@{SiteURL=$Site.Url;ErrorMessage=$err})
            }
        }
    }
    return @{Results=$Global:FileResults;Errors=$ErrorSites}
}

#############################
### Main Menu & Loop
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
            try { Connect-SPOService -Url (Get-SharePointAdminURL) } catch { Write-Host "❌ Failed to connect to SPO: $_" -ForegroundColor Red; continue }
            $out = Read-Host "`nEnter path for storage report (e.g. C:\SiteStorageReport.csv)"
            $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "
            $sites = if ($includeOneDrive.ToUpper() -eq 'Y') { Get-SPOSite -IncludePersonalSite $true -Limit All } else { Get-SPOSite -Limit All }
            $report = $sites | Select-Object @{N='SiteUrl';E={$_.Url}},@{N='StorageUsedMB';E={[math]::Round($_.StorageUsageCurrent,2)}},@{N='StorageQuotaMB';E={[math]::Round($_.StorageQuota,2)}},@{N='PercentUsed';E={if ($_.StorageQuota -gt 0){[math]::Round(($_.StorageUsageCurrent / $_.StorageQuota)*100,2)}else{'N/A'}}}
            $total = ($sites | Measure-Object -Property StorageUsageCurrent -Sum).Sum
            $report += [PSCustomObject]@{SiteUrl='*** Total ***';StorageUsedMB=[math]::Round($total,2);StorageQuotaMB='';PercentUsed=''}
            Write-CsvAndNotify -Data $report -Path $out -Label "Site storage report"
        }
        '2' {
            $ClientID = Initialize-PnPContext
            $AdminAccount = Prompt-Elevation
            $out = Read-Host "`nEnter path for inactive sites report (e.g. C:\InactiveSites.csv)"
            $action = {
                param($site,$client)
                Connect-PnPOnline -Url $site.Url -ClientID $client -Interactive | Out-Null
                $web = Get-PnPWeb -Includes LastItemUserModifiedDate
                [PSCustomObject]@{SiteUrl=$site.Url;LastModified=$web.LastItemUserModifiedDate}
            }
            $report = Invoke-TenantSitesAction -Action $action -ClientID $ClientID -AdminAccount $AdminAccount
            Write-CsvAndNotify -Data $report -Path $out -Label "Inactive sites report"
            if ($AdminAccount) { Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount }
        }
        '3' {
            $ClientID = Initialize-PnPContext
            $out = Read-Host "`nEnter path for sharing settings report (e.g. C:\SharingSettings.csv)"
            $sites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*" }
            $report = $sites | Select-Object Url,SharingCapability
            Write-CsvAndNotify -Data $report -Path $out -Label "Sharing settings report"
        }
        '4' {
            $ClientID = Initialize-PnPContext
            $out1 = Read-HostGreen "`nEnter path for Duplicate Files Report CSV: "
            $out2 = Read-HostGreen "`nEnter path for errored sites CSV: "
            $AdminAccount = Prompt-Elevation
            $res = Process-FileScanReport -ReportMode "Duplicates" -ClientID $ClientID -AdminAccount $AdminAccount -OutputCSV $out1 -ErrorSitesCSV $out2
            Write-CsvAndNotify -Data $res.Results -Path $out1 -Label "Duplicate Files Report"
            Write-CsvAndNotify -Data $res.Errors -Path $out2 -Label "Errored sites"
            if ($AdminAccount) { Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount }
        }
        '5' {
            $ClientID = Initialize-PnPContext
            $out1 = Read-HostGreen "`nEnter path for Large Files Report CSV: "
            $out2 = Read-HostGreen "`nEnter path for errored sites CSV: "
            $AdminAccount = Prompt-Elevation
            $res = Process-FileScanReport -ReportMode "AllFiles" -ClientID $ClientID -AdminAccount $AdminAccount -OutputCSV $out1 -ErrorSitesCSV $out2
            Write-CsvAndNotify -Data $res.Results -Path $out1 -Label "Large Files Report"
            Write-CsvAndNotify -Data $res.Errors -Path $out2 -Label "Errored sites"
            if ($AdminAccount) { Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount }
        }
        '6' {
            $ClientID = Initialize-PnPContext
            $AdminAccount = Prompt-Elevation
            $out = Read-Host "`nEnter path for user access review report (e.g. C:\AccessReview.csv)"
            $action = {
                param($site,$client)
                Connect-PnPOnline -Url $site.Url -ClientID $client -Interactive | Out-Null
                $groups = Get-PnPGroup
                $data = @()
                foreach ($g in $groups) {
                    $ms = Get-PnPGroupMember -Group $g
                    foreach ($m in $ms) {
                        $data += [PSCustomObject]@{SiteUrl=$site.Url;GroupName=$g.Title;UserName=$m.Title;UserLogin=$m.LoginName;UserEmail=$m.Email}
                    }
                }
                return $data
            }
            $report = Invoke-TenantSitesAction -Action $action -ClientID $ClientID -AdminAccount $AdminAccount
            Write-CsvAndNotify -Data $report -Path $out -Label "User access review report"
            if ($AdminAccount) { Remove-ElevatedAdminPrivileges -ClientID $ClientID -AdminAccount $AdminAccount }
        }
        '7' {
            $ClientID = Initialize-PnPContext
            $includeOneDrive = Read-HostGreen "`nInclude OneDrive (personal) sites? (Y/N): "
            $out = Read-Host "`nEnter path for External Users report (e.g. C:\ExternalUsers.csv)"
            $action = {
                param($site,$client)
                Connect-PnPOnline -Url $site.Url -ClientID $client -Interactive | Out-Null
                $users = Get-PnPExternalUser -PageSize 50 -SiteUrl $site.Url
                $data = @()
                foreach ($u in $users) {
                    $data += [PSCustomObject]@{SiteUrl=$site.Url;DisplayName=$u.DisplayName;Email=$u.Email;AcceptedAs=$u.AcceptedAs;WhenCreated=$u.WhenCreated;InvitedBy=$u.InvitedBy;InvitedAs=$u.InvitedAs}
                }
                return $data
            }
            $report = Invoke-TenantSitesAction -Action $action -ClientID $ClientID -AdminAccount $null -IncludeOneDrive:($includeOneDrive.ToUpper() -eq 'Y')
            Write-CsvAndNotify -Data $report -Path $out -Label "External users-by-site report"
        }
        '8' {
            Write-Host "Updating connection settings..." -ForegroundColor Cyan
            $newURL = Read-HostGreen "Enter the new SharePoint Admin URL (e.g. https://newtenant-admin.sharepoint.com): "
            if ($newURL -ne $Global:SharePointAdminURL) { $Global:SharePointAdminURL = $newURL; $Global:ClientID = $null; Write-Host "Connection settings updated. Client ID will be requested next time." -ForegroundColor Green } else { Write-Host "No changes made." -ForegroundColor Yellow }
            continue
        }
        '0' {
            Write-Host "Exiting..." -ForegroundColor Yellow
            $continue = $false
        }
        default {
            Write-Host "❌ Invalid choice. Please choose 0-8." -ForegroundColor Red
        }
    }
    if (($choice -ne '0') -and ($choice -ne '8')) {
        $ret = Read-Host "`nReturn to menu? (Y/N): "
        if ($ret.ToUpper() -ne 'Y') { $continue = $false }
    }
}