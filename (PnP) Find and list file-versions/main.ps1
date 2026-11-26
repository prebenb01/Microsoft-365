Clear-Host

# Ensure required modules are installed
function Ensure-Module {
    param([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module $ModuleName not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -AllowClobber
    }
}

function Check-PowerShellVersion {
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "PowerShell 7+ detected." -ForegroundColor Green
    } else {
        Write-Host "PowerShell 7 or higher is required." -ForegroundColor Red
        Exit
    }
}

Check-PowerShellVersion
Ensure-Module -ModuleName "PnP.PowerShell"
Import-Module PnP.PowerShell

# Variables
$VersionLimit = 10
$ClientID     = "XXXXXXXXXXXXXXXX"
$Tenant       = "XXXXXXXXXXXXXXXX"
$ThumbPrint   = "XXXXXXXXXXXXXXXX"
$AdminUrl     = "https://<domain>-admin.sharepoint.com/"

# ── Target Sites ─────────────────────────────────────────────────────────────
$TargetSites = @(
    "https://<domain>.sharepoint.com/sites/<site>"
)

# ── Validate ──────────────────────────────────────────────────────────────────
if ($TargetSites.Count -eq 0) {
    Write-Host "ERROR: No sites specified in `$TargetSites. Please add at least one URL." -ForegroundColor Red
    Exit
}

# ── Connect to SharePoint Admin ───────────────────────────────────────────────
try {
    Connect-PnPOnline -Url $AdminUrl `
        -ClientId $ClientID `
        -Thumbprint $ThumbPrint `
        -Tenant $Tenant
    Write-Host "Connected to tenant admin." -ForegroundColor Green
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    Exit
}

# ── System libraries to skip ──────────────────────────────────────────────────
$SkipLibraries = @(
    "appfiles",
    "Bibliotek for vedlikeholdslogg",
    "Galleri for hoveddokumenter",
    "Konverterte skjemaer",
    "Listemalgalleri",
    "Løsningsgalleri",
    "Nettstedsobjekter",
    "Områdesider",
    "Skjemamaler",
    "Stilbibliotek",
    "Temagalleri",
    "Webdel-galleri",
    "Teams Wiki Data",
    "Style Library",
    "Form Templates",
    "Site Assets",
    "Site Pages"
)

# ── Helper: get all user-facing document libraries ────────────────────────────
function Get-DocLibraries {
    $lists = Get-PnPList -Includes RootFolder
    return $lists | Where-Object {
        $_.BaseType -eq "DocumentLibrary" -and
        $_.Hidden -eq $false -and
        $_.Title -notin $SkipLibraries
    }
}

# ── Helper: recursively get all files in a folder ─────────────────────────────
function Get-AllFiles {
    param(
        [string]$LibraryTitle
    )

    Write-Host "    Fetching items from list: $LibraryTitle" -ForegroundColor DarkCyan

    $items = Get-PnPListItem -List $LibraryTitle -PageSize 500 -Fields FileLeafRef, FileRef, File_x0020_Size, File_x0020_Type, Modified |
        Where-Object { $_.FileSystemObjectType -eq "File" }

    return $items
}

# ── Functions ─────────────────────────────────────────────────────────────────
function Get-StorageUsage {
    $web = Get-PnPWeb -Includes Title
    Write-Host "  Site: $($web.Title)" -ForegroundColor White
}

function Get-VersionHistorySettings {
    $lists = Get-PnPList
    foreach ($list in $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" }) {
        Write-Host "  Library: $($list.Title) | Versioning: $($list.EnableVersioning) | Max Versions: $($list.MajorVersionLimit)"
    }
}

function Generate-StorageReport {
    param(
        [string]$SiteUrl,
        [string]$LibraryTitle,
        [string]$ReportFile = "SharePointStorageReport_All.csv",
        [int]$MaxVersions = 5
    )

    $Items = Get-AllFiles -LibraryTitle $LibraryTitle

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "    No files found in library: $LibraryTitle" -ForegroundColor Yellow
        return
    }

    Write-Host "    Found $($Items.Count) file(s)" -ForegroundColor DarkCyan
    $FileData = @()

    foreach ($item in $Items) {
        $fileUrl = $item["FileRef"]

        try {
            $versions     = Get-PnPFileVersion -Url $fileUrl -ErrorAction SilentlyContinue
            $versionCount = if ($versions) { $versions.Count } else { 0 }

            # Skip files that are within the version limit
            if ($versionCount -le $MaxVersions) { continue }

            Write-Host "    Processing: $fileUrl ($versionCount versions)" -ForegroundColor Cyan

            $currentSizeMB  = $item["File_x0020_Size"] / 1MB
            $totalSizeInMB  = if ($versions) { ($versions | Measure-Object -Property Size -Sum).Sum / 1MB } else { 0 }
            $totalSizeInMB += $currentSizeMB

            $excessSizeInMB = 0
            for ($i = 0; $i -lt ($versionCount - $MaxVersions); $i++) {
                $excessSizeInMB += $versions[$i].Size / 1MB
            }

            $FileData += [PSCustomObject]@{
                Site            = $SiteUrl
                Library         = $LibraryTitle
                Name            = $item["FileLeafRef"]
                CurrentSizeInMB = [math]::Round($currentSizeMB, 3)
                TotalSizeInMB   = [math]::Round($totalSizeInMB, 3)
                URL             = $fileUrl
                Modified        = $item["Modified"]
                Versions        = $versionCount
                ExcessSizeInMB  = [math]::Round($excessSizeInMB, 3)
            }
        } catch {
            Write-Host "    ERROR on file $fileUrl : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($FileData.Count -gt 0) {
        Write-Host "    $($FileData.Count) file(s) exceed version limit — appending to report." -ForegroundColor Green
        $FileData | Sort-Object TotalSizeInMB -Descending |
            Export-Csv -Path $ReportFile -NoTypeInformation -Append
    } else {
        Write-Host "    No files exceed the version limit of $MaxVersions." -ForegroundColor Yellow
    }
}

function Delete-OldVersions {
    $libs = Get-DocLibraries
    foreach ($lib in $libs) {
        $Items = Get-PnPListItem -List $lib.Title -Fields FileLeafRef, FileRef
        foreach ($item in $Items) {
            $fileUrl  = $item["FileRef"]
            $Versions = Get-PnPFileVersion -Url $fileUrl -ErrorAction SilentlyContinue
            if ($Versions -and $Versions.Count -gt $VersionLimit) {
                for ($i = $VersionLimit; $i -lt $Versions.Count; $i++) {
                    Remove-PnPFileVersion -Url $fileUrl -Identity $Versions[$i].Id -Force
                    Write-Host "  Deleted version $($Versions[$i].Id) of $fileUrl"
                }
            }
        }
    }
}


Write-Host "Processing $($TargetSites.Count) site(s)." -ForegroundColor Cyan

foreach ($siteUrl in $TargetSites) {
    Write-Host "`n── Processing site: $siteUrl" -ForegroundColor Magenta

    try {
        Connect-PnPOnline -Url $siteUrl `
            -ClientId $ClientID `
            -Thumbprint $ThumbPrint `
            -Tenant $Tenant

        Get-StorageUsage
        Get-VersionHistorySettings

        $libs = Get-DocLibraries
        if ($libs) {
            foreach ($lib in $libs) {
                Write-Host "`n  ── Library: $($lib.Title)" -ForegroundColor White
                Generate-StorageReport `
                    -SiteUrl      $siteUrl `
                    -LibraryTitle $lib.Title `
                    -ReportFile   "SharePointStorageReport_All.csv" `
                    -MaxVersions  $VersionLimit
            }
        } else {
            Write-Host "  No user libraries found — skipping." -ForegroundColor Yellow
        }

        # Uncomment to enable deletion:
        # Delete-OldVersions

    } catch {
        Write-Host "  ERROR on $($siteUrl): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nDone. Report saved to SharePointStorageReport_All.csv" -ForegroundColor Green