<#
.SYNOPSIS
    Download all files from SharePoint "Documents" library to Z:\Sharepoint
.DESCRIPTION
    - Uses PnP.PowerShell with app registration (Client ID).
    - Recursively enumerates all files in the library.
    - Downloads while preserving folder structure.
    - Handles very large libraries using paging.
    - Skips files that already exist unless source is newer.
    - Logs all output to a transcript file.
    - Use -Info switch to show detailed output including skipped files.
.NOTES
    Requires: PnP.PowerShell (Install-Module PnP.PowerShell -Force)
#>
param(
    [string]$SiteUrl       = "https://allenbutlerconstruction.sharepoint.com/sites/ProjectManagement",
    [string]$LibraryName   = "Documents",
    [string]$TargetPath    = "Z:\Sharepoint",
    [string]$ClientId      = "4e43c4bb-1a03-4fa2-ad63-c1f358d27996",
    [string]$Tenant        = "allenbutlerconstruction.onmicrosoft.com",   # <-- update with your tenant domain
    [string]$LogPath       = "",  # Optional: specify custom log path, otherwise uses default
    [switch]$Info          # Show detailed output including skipped files
)

# ---------------------------
# SETUP LOGGING
# ---------------------------
if ([string]::IsNullOrEmpty($LogPath)) {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $LogPath = Join-Path $TargetPath "SharePointSync_$timestamp.log"
}

# Ensure log directory exists
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

Write-Host "Starting transcript log: $LogPath" -ForegroundColor Cyan
Start-Transcript -Path $LogPath -Force

Write-Host "========================================" -ForegroundColor White
Write-Host "SharePoint Document Sync Started" -ForegroundColor White
Write-Host "Start Time: $(Get-Date)" -ForegroundColor White
Write-Host "Site URL: $SiteUrl" -ForegroundColor White
Write-Host "Library: $LibraryName" -ForegroundColor White
Write-Host "Target Path: $TargetPath" -ForegroundColor White
Write-Host "Log File: $LogPath" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

# ---------------------------
# CONNECT
# ---------------------------
$startTime = Get-Date
Write-Host "Connecting to $SiteUrl using client ID $ClientId ..."
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -Interactive

# Verify library
$library = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue
if (-not $library) {
    Write-Error "Library '$LibraryName' not found at $SiteUrl"
    Stop-Transcript
    exit
}
Write-Host "Connected. Library found: $($library.Title)"

# ---------------------------
# ENUMERATE ITEMS
# ---------------------------
Write-Host "Enumerating all files in $LibraryName ..."
$pageSize = 5000
$allItems = @()
$page = 0

$listItems = Get-PnPListItem -List $LibraryName -PageSize $pageSize -Fields FileLeafRef,FileRef,FSObjType,Modified -ScriptBlock {
    param($items)
    $script:allItems += $items
    $script:page++
    Write-Progress -Activity "Fetching items" -Status "Page $script:page ($($script:allItems.Count) items)" `
        -PercentComplete (($script:allItems.Count / $library.ItemCount) * 100)
}

# Filter only files (FSObjType=0)
$files = $allItems | Where-Object { $_.FieldValues.FSObjType -eq 0 }
Write-Host "📁 Found $($files.Count) files to process."

# ---------------------------
# DOWNLOAD FILES
# ---------------------------
$downloadedCount = 0
$skippedCount = 0
$errorCount = 0

foreach ($file in $files) {
    $serverRelativePath = $file.FieldValues.FileRef
    $fileName = $file.FieldValues.FileLeafRef
    $sourceModified = $file.FieldValues.Modified
    
    # Build local path
    $relativePath = $serverRelativePath.Replace("/sites/ProjectManagement/Shared Documents", "")
    $localPath = Join-Path $TargetPath $relativePath
    
    # Check if file already exists and compare dates
    $shouldDownload = $true
    if (Test-Path $localPath) {
        try {
            $localFile = Get-Item $localPath -ErrorAction Stop
            $localModified = $localFile.LastWriteTime
            
            # Convert SharePoint time to local time for comparison
            $sourceModifiedLocal = $sourceModified.ToLocalTime()
            
            if ($sourceModifiedLocal -le $localModified) {
                if ($Info) {
                    Write-Host "⏭️  Skipping $fileName (local file is up to date)" -ForegroundColor Yellow
                }
                $skippedCount++
                $shouldDownload = $false
            } else {
                Write-Host "🔄 Updating $fileName (source is newer: $sourceModifiedLocal vs $localModified)" -ForegroundColor Cyan
            }
        }
        catch {
            Write-Host "⚠️  Cannot read existing file $localPath, will re-download" -ForegroundColor Yellow
        }
    } else {
        Write-Host "⬇️  Downloading $fileName (new file)" -ForegroundColor Green
    }
    
    if ($shouldDownload) {
        # Ensure folder exists
        $localDir = Split-Path $localPath -Parent
        if (-not (Test-Path $localDir)) {
            try {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }
            catch {
                Write-Warning "❌ Failed to create directory $localDir : $_"
                $errorCount++
                continue
            }
        }
        
        # Download file
        try {
            Get-PnPFile -Url $serverRelativePath -Path $localDir -FileName $fileName -AsFile -Force
            
            # Verify the file was downloaded and set timestamp
            if (Test-Path $localPath) {
                try {
                    $downloadedFile = Get-Item $localPath -ErrorAction Stop
                    if ($downloadedFile -and $downloadedFile.PSObject.Properties['LastWriteTime']) {
                        $downloadedFile.LastWriteTime = $sourceModified.ToLocalTime()
                        $downloadedCount++
                    } else {
                        Write-Warning "⚠️  Downloaded $fileName but cannot set timestamp"
                        $downloadedCount++
                    }
                }
                catch {
                    Write-Warning "⚠️  Downloaded $fileName but failed to set timestamp: $_"
                    $downloadedCount++
                }
            } else {
                Write-Warning "❌ File download appeared successful but $localPath does not exist"
                $errorCount++
            }
        }
        catch {
            Write-Warning "❌ Failed to download $serverRelativePath : $_"
            $errorCount++
        }
    }
}

# ---------------------------
# SUMMARY
# ---------------------------
$endTime = Get-Date
$duration = $endTime - $startTime
$totalProcessed = $downloadedCount + $skippedCount + $errorCount

Write-Host "`n========================================" -ForegroundColor White
Write-Host "📊 Detailed Sync Summary:" -ForegroundColor White
Write-Host "   Total files in SharePoint: $($files.Count)" -ForegroundColor White
Write-Host "   Total files processed: $totalProcessed" -ForegroundColor White
Write-Host "" -ForegroundColor White
Write-Host "   ✅ Files downloaded/updated: $downloadedCount" -ForegroundColor Green
Write-Host "   ⏭️  Files skipped (up to date): $skippedCount" -ForegroundColor Yellow
Write-Host "" -ForegroundColor White
Write-Host "   ❌ Total errors: $errorCount" -ForegroundColor Red
Write-Host "" -ForegroundColor White
Write-Host "   ⏱️  Duration: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "   🕐 Start Time: $startTime" -ForegroundColor White
Write-Host "   🕐 End Time: $endTime" -ForegroundColor White
Write-Host "" -ForegroundColor White
if ($errorCount -eq 0) {
    Write-Host "✅ Sync completed successfully! All files processed without errors." -ForegroundColor Green
} elseif ($errorCount -lt ($files.Count * 0.05)) {
    Write-Host "⚠️  Sync completed with minor issues. Less than 5% error rate." -ForegroundColor Yellow
} else {
    Write-Host "❌ Sync completed with significant errors. Please review the log." -ForegroundColor Red
}
Write-Host "📁 Files saved to: $TargetPath" -ForegroundColor Green
Write-Host "📄 Complete log saved to: $LogPath" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor White

# ---------------------------
# ERROR REPORT
# ---------------------------
if ($errorCount -gt 0) {
    Write-Host "`n🚨 ERROR SUMMARY:" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "   ❌ Total errors encountered: $errorCount" -ForegroundColor Red
    Write-Host "   📄 Check the transcript log for detailed error information: $LogPath" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Red
}

Write-Host "`nStopping transcript..." -ForegroundColor Cyan
Stop-Transcript
