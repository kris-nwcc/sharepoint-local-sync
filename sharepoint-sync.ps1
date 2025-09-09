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
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -LiteralPath $logDir -Force | Out-Null
}

Write-Host "Starting transcript log: $LogPath" -ForegroundColor Cyan

# Initialize transcript with error handling
$transcriptStarted = $false
try {
    Start-Transcript -Path $LogPath -Force -Append
    $transcriptStarted = $true
    Write-Host "✅ Transcript successfully started" -ForegroundColor Green
    
    # Verify transcript is working by writing a test message
    Write-Host "🔍 Testing transcript capture..." -ForegroundColor Cyan
    [System.Console]::Out.Flush()
    
} catch {
    Write-Warning "⚠️  Failed to start transcript: $($_.Exception.Message)"
    Write-Warning "Script will continue but logging may be limited"
}

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
    if ($transcriptStarted) {
        try {
            Stop-Transcript
            Write-Host "Transcript stopped due to library not found error" -ForegroundColor Yellow
        } catch {
            Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
        }
    }
    exit 1
}
Write-Host "Connected. Library found: $($library.Title)"

# ---------------------------
# MAIN EXECUTION WITH ERROR HANDLING
# ---------------------------
try {
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

# Specific error counters
$directoryErrorCount = 0
$downloadErrorCount = 0
$missingAfterDownloadCount = 0
$timestampErrorCount = 0
$otherErrorCount = 0

# Error tracking arrays
$directoryErrors = @()
$downloadErrors = @()
$missingAfterDownloadErrors = @()
$timestampErrors = @()
$otherErrors = @()

foreach ($file in $files) {
    $serverRelativePath = $file.FieldValues.FileRef
    $fileName = $file.FieldValues.FileLeafRef
    $sourceModified = $file.FieldValues.Modified
    
    # Decode URL-encoded characters in the server relative path and filename
    # This fixes issues with special characters like semicolons (%3b) in filenames
    try {
        $serverRelativePath = [System.Web.HttpUtility]::UrlDecode($serverRelativePath)
        $fileName = [System.Web.HttpUtility]::UrlDecode($fileName)
    }
    catch {
        Write-Warning "⚠️  Failed to decode URL-encoded characters for file: $fileName"
        Write-Warning "   Using original paths. Error: $($_.Exception.Message)"
        # Continue with original paths if decoding fails
    }
    
    # Periodic transcript verification (every 100 files)
    if (($downloadedCount + $skippedCount + $errorCount) % 100 -eq 0 -and ($downloadedCount + $skippedCount + $errorCount) -gt 0) {
        Write-Host "📊 Progress: $($downloadedCount + $skippedCount + $errorCount) files processed..." -ForegroundColor Cyan
        # Force output buffer flush to ensure transcript captures progress
        [System.Console]::Out.Flush()
    }
    
    # Build local path
    $relativePath = $serverRelativePath.Replace("/sites/ProjectManagement/Shared Documents", "")
    $localPath = Join-Path $TargetPath $relativePath
    
    # Check if file already exists and compare dates
    $shouldDownload = $true
    if (Test-Path -LiteralPath $localPath) {
        try {
            $localFile = Get-Item -LiteralPath $localPath -ErrorAction Stop
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
        if (-not (Test-Path -LiteralPath $localDir)) {
            try {
                New-Item -ItemType Directory -LiteralPath $localDir -Force | Out-Null
            }
            catch {
                Write-Warning "❌ Failed to create directory $localDir : $_"
                $directoryErrorCount++
                $errorCount++
                $directoryErrors += [PSCustomObject]@{
                    FileName = $fileName
                    Path = $localDir
                    Error = $_.Exception.Message
                }
                continue
            }
        }
        
        # Download file
        try {
            # Log the actual paths being used for debugging
            if ($Info) {
                Write-Host "🔍 Downloading from SharePoint path: $serverRelativePath" -ForegroundColor Gray
                Write-Host "🔍 Saving to local path: $localPath" -ForegroundColor Gray
            }
            Get-PnPFile -Url $serverRelativePath -Path $localDir -FileName $fileName -AsFile -Force
            
            # Verify the file was downloaded and set timestamp
            if (Test-Path -LiteralPath $localPath) {
                try {
                    $downloadedFile = Get-Item -LiteralPath $localPath -ErrorAction Stop
                    if ($downloadedFile -and $downloadedFile.PSObject.Properties['LastWriteTime']) {
                        $downloadedFile.LastWriteTime = $sourceModified.ToLocalTime()
                        $downloadedCount++
                    } else {
                        Write-Warning "⚠️  Downloaded $fileName but cannot set timestamp"
                        $timestampErrorCount++
                        $errorCount++
                        $timestampErrors += [PSCustomObject]@{
                            FileName = $fileName
                            Path = $localPath
                            Error = "Cannot set timestamp - file properties not accessible"
                        }
                        $downloadedCount++
                    }
                }
                catch {
                    Write-Warning "⚠️  Downloaded $fileName but failed to set timestamp: $_"
                    $timestampErrorCount++
                    $errorCount++
                    $timestampErrors += [PSCustomObject]@{
                        FileName = $fileName
                        Path = $localPath
                        Error = $_.Exception.Message
                    }
                    $downloadedCount++
                }
            } else {
                Write-Warning "❌ File download appeared successful but $localPath does not exist"
                $missingAfterDownloadCount++
                $errorCount++
                $missingAfterDownloadErrors += [PSCustomObject]@{
                    FileName = $fileName
                    SharePointPath = $serverRelativePath
                    Path = $localPath
                    Error = "File missing after download"
                }
            }
        }
        catch {
            Write-Warning "❌ Failed to download $serverRelativePath : $_"
            $downloadErrorCount++
            $errorCount++
            $downloadErrors += [PSCustomObject]@{
                FileName = $fileName
                SharePointPath = $serverRelativePath
                Path = $localPath
                Error = $_.Exception.Message
            }
        }
    }
}

# ---------------------------
# SUMMARY
# ---------------------------
$endTime = Get-Date
$duration = $endTime - $startTime

# Calculate catch-all error count (errors not categorized above)
$categorizedErrors = $directoryErrorCount + $downloadErrorCount + $missingAfterDownloadCount + $timestampErrorCount
$otherErrorCount = $errorCount - $categorizedErrors

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
Write-Host "      • Directory creation errors: $directoryErrorCount" -ForegroundColor Red
Write-Host "      • Download failures: $downloadErrorCount" -ForegroundColor Red
Write-Host "      • Files missing after download: $missingAfterDownloadCount" -ForegroundColor Red
Write-Host "      • Timestamp setting errors: $timestampErrorCount" -ForegroundColor Yellow
Write-Host "      • Other errors: $otherErrorCount" -ForegroundColor Red
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
# DETAILED ERROR REPORT
# ---------------------------
if ($errorCount -gt 0) {
    Write-Host "`n🚨 DETAILED ERROR REPORT:" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    
    if ($directoryErrors.Count -gt 0) {
        Write-Host "`n📁 DIRECTORY CREATION ERRORS ($($directoryErrors.Count)):" -ForegroundColor Red
        $directoryErrors | ForEach-Object {
            Write-Host "   • $($_.FileName)" -ForegroundColor White
            Write-Host "     Path: $($_.Path)" -ForegroundColor Gray
            Write-Host "     Error: $($_.Error)" -ForegroundColor Yellow
            Write-Host ""
        }
    }
    
    if ($downloadErrors.Count -gt 0) {
        Write-Host "`n⬇️  DOWNLOAD FAILURES ($($downloadErrors.Count)):" -ForegroundColor Red
        $downloadErrors | ForEach-Object {
            Write-Host "   • $($_.FileName)" -ForegroundColor White
            Write-Host "     SharePoint: $($_.SharePointPath)" -ForegroundColor Gray
            Write-Host "     Local: $($_.Path)" -ForegroundColor Gray
            Write-Host "     Error: $($_.Error)" -ForegroundColor Yellow
            Write-Host ""
        }
    }
    
    if ($missingAfterDownloadErrors.Count -gt 0) {
        Write-Host "`n👻 FILES MISSING AFTER DOWNLOAD ($($missingAfterDownloadErrors.Count)):" -ForegroundColor Red
        $missingAfterDownloadErrors | ForEach-Object {
            Write-Host "   • $($_.FileName)" -ForegroundColor White
            Write-Host "     SharePoint: $($_.SharePointPath)" -ForegroundColor Gray
            Write-Host "     Expected at: $($_.Path)" -ForegroundColor Gray
            Write-Host "     Error: $($_.Error)" -ForegroundColor Yellow
            Write-Host ""
        }
    }
    
    if ($timestampErrors.Count -gt 0) {
        Write-Host "`n🕐 TIMESTAMP SETTING ERRORS ($($timestampErrors.Count)):" -ForegroundColor Yellow
        $timestampErrors | ForEach-Object {
            Write-Host "   • $($_.FileName)" -ForegroundColor White
            Write-Host "     Path: $($_.Path)" -ForegroundColor Gray
            Write-Host "     Error: $($_.Error)" -ForegroundColor Yellow
            Write-Host ""
        }
    }
    
    if ($otherErrorCount -gt 0) {
        Write-Host "`n❓ OTHER ERRORS ($otherErrorCount):" -ForegroundColor Red
        Write-Host "   • These errors were not categorized into specific types" -ForegroundColor White
        Write-Host "   • Check the transcript log for detailed information: $LogPath" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Save detailed error report to file
    $errorReportPath = $LogPath.Replace(".log", "_ErrorReport.txt")
    $errorReport = @()
    $errorReport += "SharePoint Sync Error Report"
    $errorReport += "Generated: $(Get-Date)"
    $errorReport += "="*50
    $errorReport += ""
    
    if ($directoryErrors.Count -gt 0) {
        $errorReport += "DIRECTORY CREATION ERRORS ($($directoryErrors.Count)):"
        $errorReport += "-"*40
        $directoryErrors | ForEach-Object {
            $errorReport += "File: $($_.FileName)"
            $errorReport += "Path: $($_.Path)"
            $errorReport += "Error: $($_.Error)"
            $errorReport += ""
        }
    }
    
    if ($downloadErrors.Count -gt 0) {
        $errorReport += "DOWNLOAD FAILURES ($($downloadErrors.Count)):"
        $errorReport += "-"*40
        $downloadErrors | ForEach-Object {
            $errorReport += "File: $($_.FileName)"
            $errorReport += "SharePoint Path: $($_.SharePointPath)"
            $errorReport += "Local Path: $($_.Path)"
            $errorReport += "Error: $($_.Error)"
            $errorReport += ""
        }
    }
    
    if ($missingAfterDownloadErrors.Count -gt 0) {
        $errorReport += "FILES MISSING AFTER DOWNLOAD ($($missingAfterDownloadErrors.Count)):"
        $errorReport += "-"*40
        $missingAfterDownloadErrors | ForEach-Object {
            $errorReport += "File: $($_.FileName)"
            $errorReport += "SharePoint Path: $($_.SharePointPath)"
            $errorReport += "Expected Local Path: $($_.Path)"
            $errorReport += "Error: $($_.Error)"
            $errorReport += ""
        }
    }
    
    if ($timestampErrors.Count -gt 0) {
        $errorReport += "TIMESTAMP SETTING ERRORS ($($timestampErrors.Count)):"
        $errorReport += "-"*40
        $timestampErrors | ForEach-Object {
            $errorReport += "File: $($_.FileName)"
            $errorReport += "Path: $($_.Path)"
            $errorReport += "Error: $($_.Error)"
            $errorReport += ""
        }
    }
    
    if ($otherErrorCount -gt 0) {
        $errorReport += "OTHER ERRORS ($otherErrorCount):"
        $errorReport += "-"*40
        $errorReport += "These errors were not categorized into specific types."
        $errorReport += "Check the transcript log for detailed information."
        $errorReport += ""
    }
    
    $errorReport | Out-File -FilePath $errorReportPath -Encoding UTF8
    Write-Host "`n📄 Detailed error report saved to: $errorReportPath" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Red
}

} catch {
    Write-Error "❌ Fatal error during script execution: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
    $errorCount++
}

# ---------------------------
# STOP TRANSCRIPT
# ---------------------------
Write-Host "`nStopping transcript..." -ForegroundColor Cyan

# Final transcript verification
if ($transcriptStarted) {
    Write-Host "🔍 Final transcript verification..." -ForegroundColor Cyan
    [System.Console]::Out.Flush()
    
    try {
        Stop-Transcript
        Write-Host "✅ Transcript successfully stopped" -ForegroundColor Green
        
        # Verify transcript file exists and has content
        if (Test-Path -LiteralPath $LogPath) {
            $logSize = (Get-Item -LiteralPath $LogPath).Length
            Write-Host "📄 Transcript file size: $logSize bytes" -ForegroundColor Green
            if ($logSize -lt 1000) {
                Write-Warning "⚠️  Transcript file seems unusually small - may indicate incomplete capture"
            }
        } else {
            Write-Warning "⚠️  Transcript file not found at expected location: $LogPath"
        }
        
    } catch {
        Write-Warning "⚠️  Failed to stop transcript: $($_.Exception.Message)"
    }
} else {
    Write-Host "ℹ️  No transcript was started, so no need to stop it" -ForegroundColor Yellow
}
