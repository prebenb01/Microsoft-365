#Requires -Version 7.0
#Requires -Modules PnP.PowerShell
<#
    SharePoint storage inventory (PnP.PowerShell)

    Walks every document library in the listed site(s), reads every file and folder,
    and exports storage use to CSV:

      SP-Files-<stamp>.csv      One row per file
      SP-Folders-<stamp>.csv    One row per folder, recursive totals
      SP-Libraries-<stamp>.csv  One row per library
      SP-FileTypes-<stamp>.csv  Totals per file extension

    Size columns (megabytes, 2 decimals, numeric in the local decimal format so Excel can sort/sum):
      SizeMB      Current version of the file(s)
      VersionMB   Version history (SMTotalSize minus current version)
      TotalMB     SizeMB + VersionMB = what counts against quota
#>

# ============================ CONFIG ============================
$SiteUrls = @(
    "https://<domain>.sharepoint.com/sites/<site>"
)
$ClientId      = "######"           # Entra app reg used for PnP
$Tenant        = "######"           # domain name
$Thumbprint    = "######"           # certificate thumbprint
$OutputFolder  = "C:\Temp\SPStorage"
$IncludeHidden = $false      # $true = also scan hidden/system libraries
$ExcludeLibs   = @()         # e.g. @("Style Library", "Form Templates")
# ================================================================

$Fields = "FileLeafRef", "FileRef", "FileDirRef", "FSObjType", "File_x0020_Size",
          "SMTotalSize", "SMTotalFileStreamSize", "Modified", "Editor", "_UIVersionString"

$Stamp      = Get-Date -Format "yyyyMMdd-HHmm"
$FilesCsv   = Join-Path $OutputFolder "SP-Files-$Stamp.csv"
$FoldersCsv = Join-Path $OutputFolder "SP-Folders-$Stamp.csv"
$LibsCsv    = Join-Path $OutputFolder "SP-Libraries-$Stamp.csv"
$TypesCsv   = Join-Path $OutputFolder "SP-FileTypes-$Stamp.csv"

# UTF-8 with BOM so Excel shows æøå, locale delimiter so it opens in columns
$CsvParams  = @{ NoTypeInformation = $true; Encoding = "utf8BOM"; UseCulture = $true }

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

# SMTotalSize comes back as a lookup value, File_x0020_Size as a string -> normalise to Int64
function Get-Bytes($Value) {
    if ($null -eq $Value) { return [int64]0 }
    $n = [int64]0
    if ($Value.PSObject.Properties["LookupId"]) {
        if ([int64]::TryParse([string]$Value.LookupValue, [ref]$n) -and $n -gt 0) { return $n }
        return [int64]$Value.LookupId
    }
    [void][int64]::TryParse([string]$Value, [ref]$n)
    return $n
}

function ConvertTo-MB([int64]$Bytes) {
    [math]::Round($Bytes / 1MB, 2)
}

function Get-FolderStat($Map, [string]$Path) {
    $s = $Map[$Path]
    if (-not $s) {
        $s = [pscustomobject]@{
            Path           = $Path
            DirectFiles    = 0
            DirectBytes    = [int64]0
            TotalFiles     = 0
            SizeBytes      = [int64]0
            VersionBytes   = [int64]0
            LatestModified = $null
        }
        $Map[$Path] = $s
    }
    return $s
}

$folderRows = [System.Collections.Generic.List[object]]::new()
$libRows    = [System.Collections.Generic.List[object]]::new()
$extMap     = @{}
$grandTotal = [int64]0

foreach ($SiteUrl in $SiteUrls) {
    Write-Host "`n=== $SiteUrl ===" -ForegroundColor Cyan
    try {
        if ($Thumbprint) {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Could not connect to $SiteUrl : $($_.Exception.Message)"
        continue
    }

    $tenantRoot = ([Uri]$SiteUrl).GetLeftPart([UriPartial]::Authority)

    $libraries = Get-PnPList | Where-Object {
        $_.BaseType -eq "DocumentLibrary" -and
        ($IncludeHidden -or -not $_.Hidden) -and
        $_.Title -notin $ExcludeLibs
    }

    foreach ($lib in $libraries) {
        $null = Get-PnPProperty -ClientObject $lib -Property RootFolder
        $root = $lib.RootFolder.ServerRelativeUrl.TrimEnd('/')
        Write-Host ("  {0} ({1} items)" -f $lib.Title, $lib.ItemCount)

        try {
            $items = Get-PnPListItem -List $lib -PageSize 2000 -Fields $Fields -ErrorAction Stop
        }
        catch {
            Write-Warning "  Skipped $($lib.Title): $($_.Exception.Message)"
            continue
        }

        $folderMap   = @{}
        $fileRows    = [System.Collections.Generic.List[object]]::new()
        $folderCount = 0
        $null = Get-FolderStat $folderMap $root

        foreach ($item in $items) {
            $fv = $item.FieldValues

            # Folder: register it so empty folders also show up
            if ($fv["FSObjType"] -eq 1) {
                $null = Get-FolderStat $folderMap $fv["FileRef"]
                $folderCount++
                continue
            }

            $size = Get-Bytes $fv["File_x0020_Size"]
            if ($size -eq 0) { $size = Get-Bytes $fv["SMTotalFileStreamSize"] }
            $total = Get-Bytes $fv["SMTotalSize"]
            if ($total -lt $size) { $total = $size }
            $versions = $total - $size

            $modified = [datetime]$fv["Modified"]
            $name     = $fv["FileLeafRef"]
            $dir      = $fv["FileDirRef"]
            $ext      = [IO.Path]::GetExtension($name).ToLower()
            if (-not $ext) { $ext = "(none)" }

            $relDir = $dir.Substring([math]::Min($root.Length, $dir.Length))
            if (-not $relDir) { $relDir = "/" }

            $fileRows.Add([pscustomobject]@{
                Site         = $SiteUrl
                Library      = $lib.Title
                Folder       = $relDir
                FileName     = $name
                Extension    = $ext
                SizeMB       = ConvertTo-MB $size
                VersionMB    = ConvertTo-MB $versions
                TotalMB      = ConvertTo-MB $total
                Version      = $fv["_UIVersionString"]
                Modified     = $modified.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
                ModifiedBy   = $fv["Editor"].LookupValue
                Url          = $tenantRoot + $fv["FileRef"]
            })

            # Per extension
            $e = $extMap[$ext]
            if (-not $e) {
                $e = [pscustomobject]@{ Files = 0; SizeBytes = [int64]0; VersionBytes = [int64]0 }
                $extMap[$ext] = $e
            }
            $e.Files++
            $e.SizeBytes    += $size
            $e.VersionBytes += $versions
            $grandTotal     += $total

            # Files sitting directly in this folder
            $d = Get-FolderStat $folderMap $dir
            $d.DirectFiles++
            $d.DirectBytes += $size

            # Roll up through every parent folder to the library root
            $path = $dir
            while ($true) {
                $f = Get-FolderStat $folderMap $path
                $f.TotalFiles++
                $f.SizeBytes    += $size
                $f.VersionBytes += $versions
                if (-not $f.LatestModified -or $modified -gt $f.LatestModified) { $f.LatestModified = $modified }
                if ($path.Length -le $root.Length) { break }
                $path = $path.Substring(0, $path.LastIndexOf('/'))
            }
        }

        # Files are written per library to keep memory down on big sites
        if ($fileRows.Count) { $fileRows | Export-Csv -Path $FilesCsv -Append @CsvParams }

        foreach ($f in $folderMap.Values) {
            $rel = $f.Path.Substring([math]::Min($root.Length, $f.Path.Length))
            if (-not $rel) { $rel = "/" }
            $folderRows.Add([pscustomobject]@{
                Site           = $SiteUrl
                Library        = $lib.Title
                Folder         = $rel
                Depth          = $rel.Split('/', [StringSplitOptions]::RemoveEmptyEntries).Count
                DirectFiles    = $f.DirectFiles
                DirectMB       = ConvertTo-MB $f.DirectBytes
                TotalFiles     = $f.TotalFiles
                SizeMB         = ConvertTo-MB $f.SizeBytes
                VersionMB      = ConvertTo-MB $f.VersionBytes
                TotalMB        = ConvertTo-MB ($f.SizeBytes + $f.VersionBytes)
                LatestModified = if ($f.LatestModified) { $f.LatestModified.ToLocalTime().ToString("yyyy-MM-dd") } else { "" }
                Url            = $tenantRoot + $f.Path
            })
        }

        $r = $folderMap[$root]
        $libRows.Add([pscustomobject]@{
            Site         = $SiteUrl
            Library      = $lib.Title
            Files        = $r.TotalFiles
            Folders      = $folderCount
            SizeMB       = ConvertTo-MB $r.SizeBytes
            VersionMB    = ConvertTo-MB $r.VersionBytes
            TotalMB      = ConvertTo-MB ($r.SizeBytes + $r.VersionBytes)
            Url          = $tenantRoot + $root
        })
    }
}

# ---------- Export aggregates ----------
$folderRows = $folderRows | Sort-Object TotalMB -Descending
$libRows    = $libRows    | Sort-Object TotalMB -Descending

$typeRows = foreach ($k in $extMap.Keys) {
    $e = $extMap[$k]
    $t = $e.SizeBytes + $e.VersionBytes
    [pscustomobject]@{
        Extension      = $k
        Files          = $e.Files
        SizeMB         = ConvertTo-MB $e.SizeBytes
        VersionMB      = ConvertTo-MB $e.VersionBytes
        TotalMB        = ConvertTo-MB $t
        PercentOfTotal = if ($grandTotal) { "{0:N1} %" -f (100 * $t / $grandTotal) } else { "0 %" }
        AvgFileMB      = ConvertTo-MB ([int64]($e.SizeBytes / $e.Files))
    }
}
$typeRows = $typeRows | Sort-Object TotalMB -Descending

$folderRows | Export-Csv -Path $FoldersCsv @CsvParams
$libRows    | Export-Csv -Path $LibsCsv    @CsvParams
$typeRows   | Export-Csv -Path $TypesCsv   @CsvParams

# ---------- Console summary ----------
Write-Host "`nLibraries" -ForegroundColor Cyan
$libRows | Select-Object Library, Files, Folders, TotalMB, VersionMB |
    Format-Table -AutoSize | Out-Host

Write-Host "Top 10 file types" -ForegroundColor Cyan
$typeRows | Select-Object -First 10 Extension, Files, TotalMB, PercentOfTotal, AvgFileMB |
    Format-Table -AutoSize | Out-Host

Write-Host "Top 15 folders (recursive)" -ForegroundColor Cyan
$folderRows | Where-Object Depth -gt 0 | Select-Object -First 15 Library, Folder, TotalFiles, TotalMB |
    Format-Table -AutoSize | Out-Host

Write-Host "CSV files written to $OutputFolder" -ForegroundColor Green