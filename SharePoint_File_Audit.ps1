<#
.SYNOPSIS
    Scans SharePoint sites for all or duplicate files, with 100MB filter.

.DESCRIPTION
    Connects to SharePoint using PnP PowerShell and scans either all or just duplicate files across libraries.
    Optionally filters files >100MB and uses a memory-efficient list structure.

.NOTES
    Requires: PnP PowerShell, PowerShell v7.5+, Admin permissions
#>

function Read-HostGreen {
    param([string]$Prompt)
    Write-Host $Prompt -ForegroundColor Green -NoNewline
    return Read-Host
}

# --- Get User Inputs ---
$ClientID           = Read-HostGreen "Enter your Client ID: "
$SharePointAdminURL = Read-HostGreen "`nEnter the SharePoint Admin URL (e.g. https://contoso-admin.sharepoint.com): "
$AdminAccount       = Read-HostGreen "`nEnter the admin UPN (e.g. admin@contoso.com): "
$OutputCSV          = Read-HostGreen "`nEnter the path to export the results CSV (e.g. C:\AllFilesOrDuplicates.csv): "
$ErrorSitesCSV      = Read-HostGreen "`nEnter the path to export the errored sites CSV (e.g. C:\SitesThatErrored.csv): "

# --- Options ---
$OutputDuplicatesOnly = (Read-HostGreen "`nOutput Duplicates only (Y) or All files (N)?: ").ToUpper()
$OnlyLargeFiles       = (Read-HostGreen "`nOnly include files over 100MB? (Y/N): ").ToUpper()
$SearchAllLibraries   = (Read-HostGreen "`nSearch ALL document libraries (Y) or only Documents (N)? ").ToUpper()

# --- Threshold for large files (100MB) ---
$SizeThreshold = 104857600

# --- Connect to SharePoint Admin Center ---
Connect-PnPOnline -Url $SharePointAdminURL -ClientID $ClientID -Interactive

# --- Get All Sites ---
$AllSites = Get-PnPTenantSite | Where-Object {
    $_.Template -notlike "SPSPERS*" -and $_.Template -notlike "APPCATALOG*"
}

# --- Efficient Lists ---
$Global:FileResults = [System.Collections.Generic.List[object]]::new()
$ErrorSites = [System.Collections.Generic.List[object]]::new()

# --- Function to Process Library ---
function Process-Library {
    param (
        $SiteUrl,
        $LibraryTitle,
        $ClientID,
        $SizeThreshold,
        $OnlyLargeFiles,
        $OutputDuplicatesOnly
    )

    Write-Host "Checking Library: $LibraryTitle" -ForegroundColor Yellow

    $Files = Get-PnPListItem -List $LibraryTitle -PageSize 500 -Fields "FileLeafRef", "FileRef", "Modified", "Editor", "Author", "FSObjType", "SMTotalFileStreamSize"

    $FileHashTable = @{}

    foreach ($file in $Files) {
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

# --- Loop through Sites ---
foreach ($Site in $AllSites) {
    Write-Host "`n--- Processing Site: $($Site.Url) ---" -ForegroundColor Cyan

    try {
        $ErrorActionPreference = "Stop"
        Connect-PnPOnline -Url $Site.Url -ClientID $ClientID -Interactive

        $DocumentLibraries = Get-PnPList | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and
            ($SearchAllLibraries -eq "Y" -or $_.Title -eq "Documents")
        }

        foreach ($Library in $DocumentLibraries) {
            Process-Library -SiteUrl $Site.Url -LibraryTitle $Library.Title -ClientID $ClientID -SizeThreshold $SizeThreshold -OnlyLargeFiles $OnlyLargeFiles -OutputDuplicatesOnly $OutputDuplicatesOnly
        }
    }
    catch {
        Write-Host "`n⚠️  Error accessing $($Site.Url): $_" -ForegroundColor Red
        $ErrorSites.Add([PSCustomObject]@{
            SiteURL      = $Site.Url
            ErrorMessage = $_.Exception.Message
        })
    }
}

# --- Export results ---
$Global:FileResults | Export-Csv -Path $OutputCSV -NoTypeInformation
Write-Host "`n✅ File report exported to: $OutputCSV" -ForegroundColor Green

$ErrorSites | Export-Csv -Path $ErrorSitesCSV -NoTypeInformation
Write-Host "`n✅ Sites with errors exported to: $ErrorSitesCSV" -ForegroundColor Red