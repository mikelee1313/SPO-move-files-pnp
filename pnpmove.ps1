# Requires: PnP.PowerShell

# Install-Module PnP.PowerShell -Scope CurrentUser
 
# =========================

# CONFIGURATION

# =========================

# --- App Registration ---
$AppId = "abc64618-283f-47ba-a185-50d935d51d57"

$TenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

$CertThumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"

# --- Sites ---
$SourceSiteUrl = "https://m365cpi13246019.sharepoint.com/sites/BrandCenter"

$TargetSiteUrl = "https://m365cpi13246019.sharepoint.com/sites/SalesandMarketing"
 
# --- Libraries ---
$SourceLibrary = "Documents"

$TargetLibrary = "Documents"
 
# --- Folder paths inside each library (leave blank "" for the entire library) ---
$SourceRootFolder = "Folder1"

$TargetRootFolder = "DestFolder"
 
# --- Behavior ---
# Skip files that already exist at the destination
$SkipIfExists = $true

# When a file is skipped because it already exists at the destination,
# delete the source copy to complete the "move" semantic.
# Files are sent to the source Recycle Bin (not permanently deleted).
# WARNING: Only enable this AFTER confirming files are visible at the destination.
$DeleteSourceIfTargetExists = $false
 
# Overwrite existing destination files (takes precedence over SkipIfExists)
$OverwriteExisting = $false

# --- Logging ---
$LogPath = "C:\Temp\PnP_FB_MoveFiles_Log.csv"
$CheckpointPath = "C:\Temp\PnP_FB_MoveFiles_Checkpoint.json"
$ResumeFromCheckpoint = $true
$CheckpointEvery = 100

# --- Throttling (per MS guidance: https://aka.ms/SPOThrottling) ---
# Retry-After header is always honoured first; exponential backoff is the fallback.
$MaxRetries = 10   # Max retry attempts per call before giving up
$BaseDelaySec = 5    # Backoff base (seconds) when no Retry-After header is present
$MaxDelaySec = 300  # Hard cap on any single wait period (5 minutes)
# Small pause between every API request to avoid burst spikes.
# Increase to 200-500 ms if you see frequent 429s despite the retry logic.
$RequestDelayMs = 50
 
# =========================

# CONNECT

# =========================
 
Write-Host "Connecting to source site..." -ForegroundColor Cyan

$SourceConn = Connect-PnPOnline -Url $SourceSiteUrl -ClientId $AppId -Tenant $TenantId -Thumbprint $CertThumbprint -ReturnConnection
 
Write-Host "Connecting to target site..." -ForegroundColor Cyan

$TargetConn = Connect-PnPOnline -Url $TargetSiteUrl -ClientId $AppId -Tenant $TenantId -Thumbprint $CertThumbprint -ReturnConnection
 
$SourceWeb = Get-PnPWeb -Connection $SourceConn -Includes ServerRelativeUrl

$TargetWeb = Get-PnPWeb -Connection $TargetConn -Includes ServerRelativeUrl
 
$SourceWebRelUrl = $SourceWeb.ServerRelativeUrl.TrimEnd("/")

$TargetWebRelUrl = $TargetWeb.ServerRelativeUrl.TrimEnd("/")
 
if ([string]::IsNullOrWhiteSpace($SourceWebRelUrl)) { $SourceWebRelUrl = "" }

if ([string]::IsNullOrWhiteSpace($TargetWebRelUrl)) { $TargetWebRelUrl = "" }
 
# =========================

# HELPER FUNCTIONS

# =========================
 
function Join-Url {

    param(

        [string]$Part1,

        [string]$Part2

    )
 
    if ([string]::IsNullOrWhiteSpace($Part1)) {

        return "/" + $Part2.TrimStart("/")

    }
 
    return $Part1.TrimEnd("/") + "/" + $Part2.TrimStart("/")

}

# ---------------------------------------------------------------------------
# Invoke-WithRetry
# Wraps a script block and retries it on HTTP 429 / 503 throttle responses,
# honouring the Retry-After header exactly as required by Microsoft guidance:
# https://learn.microsoft.com/sharepoint/dev/general-development/
#         how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
# ---------------------------------------------------------------------------
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName = 'SPO call'
    )

    $attempt = 0
    while ($true) {
        try {
            # Honour the inter-request delay to prevent burst spikes
            if ($RequestDelayMs -gt 0) { Start-Sleep -Milliseconds $RequestDelayMs }
            return (& $ScriptBlock)
        }
        catch {
            $attempt++
            $ex = $_.Exception
            $isThrottled = $false
            $retryAfterSec = 0

            # Walk the full exception chain looking for throttle signals
            $cur = $ex
            while ($null -ne $cur) {
                $msg = $cur.Message + ''
                if ($msg -match '429|Too\s+Many\s+Requests|503|Server\s+Too\s+Busy|throttl') {
                    $isThrottled = $true
                    # Extract Retry-After value if SharePoint embedded it in the message
                    if ($msg -match 'Retry-After[:\s]+(\d+)') {
                        $retryAfterSec = [int]$Matches[1]
                    }
                    break
                }
                $cur = $cur.InnerException
            }

            if (-not $isThrottled -or $attempt -ge $MaxRetries) {
                # Not throttling, or retries exhausted — propagate the error
                throw
            }

            # Honour Retry-After header (MS #1 requirement); fall back to
            # exponential backoff capped at $MaxDelaySec
            if ($retryAfterSec -le 0) {
                $retryAfterSec = [Math]::Min(
                    [Math]::Pow(2, $attempt - 1) * $BaseDelaySec,
                    $MaxDelaySec
                )
            }
            else {
                $retryAfterSec = [Math]::Min($retryAfterSec, $MaxDelaySec)
            }

            Write-Host "  [Throttled] $OperationName — waiting $retryAfterSec s (attempt $attempt / $MaxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $retryAfterSec
        }
    }
}
 
function Ensure-FolderPath {

    param(

        # Full server-relative URL of the folder to create, e.g.
        # /sites/SalesandMarketing/Documents/DestFolder/Sub
        [string]$FolderServerRelUrl,

        # Verified server-relative URL of the library root, e.g.
        # /sites/SalesandMarketing/Documents
        # Obtained via Get-PnPList -Includes RootFolder so it is always correct.
        [string]$LibRootServerRelUrl,

        $Connection

    )

    if ([string]::IsNullOrWhiteSpace($FolderServerRelUrl)) { return }

    $libRoot = $LibRootServerRelUrl.TrimEnd('/')
    $target = $FolderServerRelUrl.TrimEnd('/')

    # Strip the library root prefix to get just the sub-path to create
    $subPath = $target.Substring($libRoot.Length).TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($subPath)) { return }  # Already at library root

    $parts = $subPath -split '/'
    for ($i = 0; $i -lt $parts.Length; $i++) {

        # Capture loop variables into plain locals — used directly in the
        # Add-PnPFolder call below (no script block, no closure, no scope risk).
        $fn = $parts[$i]
        $parent = if ($i -eq 0) { $libRoot } else { $libRoot + '/' + ($parts[0..($i - 1)] -join '/') }
        $fullPath = "$parent/$fn"

        Write-Host "  Ensuring folder: $fullPath" -ForegroundColor DarkCyan

        # Inline retry loop — avoids passing variables through a script-block
        # closure into Invoke-WithRetry, which can silently lose scope in
        # PowerShell's dynamic variable resolution model.
        $attempt = 0
        $done = $false
        while (-not $done) {
            $attempt++
            try {
                if ($RequestDelayMs -gt 0) { Start-Sleep -Milliseconds $RequestDelayMs }
                Add-PnPFolder -Name $fn -Folder $parent -Connection $Connection -ErrorAction Stop | Out-Null
                Write-Host "  Created folder: $fullPath" -ForegroundColor Green
                $done = $true
            }
            catch {
                $errMsg = $_.Exception.Message + ''
                if ($errMsg -like '*already exist*' -or $errMsg -like '*same name*') {
                    # Folder is already there — that is fine
                    Write-Host "  Folder already exists: $fullPath" -ForegroundColor DarkGray
                    $done = $true
                }
                elseif ($errMsg -match '429|Too\s+Many\s+Requests|503|Server\s+Too\s+Busy|throttl') {
                    if ($attempt -ge $MaxRetries) { throw }
                    $delay = [Math]::Min([Math]::Pow(2, $attempt - 1) * $BaseDelaySec, $MaxDelaySec)
                    Write-Host "  [Throttled] Ensure-FolderPath — waiting $delay s (attempt $attempt / $MaxRetries)..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $delay
                }
                else {
                    throw "Failed creating folder '$fullPath': $errMsg"
                }
            }
        }
    }

}

function Initialize-RunStorage {
    $logDir = Split-Path -Parent $LogPath
    if (-not [string]::IsNullOrWhiteSpace($logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $checkpointDir = Split-Path -Parent $CheckpointPath
    if (-not [string]::IsNullOrWhiteSpace($checkpointDir)) {
        New-Item -ItemType Directory -Path $checkpointDir -Force | Out-Null
    }

    if (-not $ResumeFromCheckpoint) {
        if (Test-Path -LiteralPath $LogPath) {
            Remove-Item -LiteralPath $LogPath -Force
        }

        if (Test-Path -LiteralPath $CheckpointPath) {
            Remove-Item -LiteralPath $CheckpointPath -Force
        }
    }
}

function Write-MoveLogEntry {
    param(
        [string]$SourceUrl,
        [string]$TargetUrl,
        [string]$Status,
        [string]$Message
    )

    $entry = [PSCustomObject]@{
        SourceUrl = $SourceUrl
        TargetUrl = $TargetUrl
        Status    = $Status
        Message   = $Message
    }

    if (Test-Path -LiteralPath $LogPath) {
        $entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    }
}

function Save-Checkpoint {
    param(
        [string]$LastProcessedSourceUrl,
        [int]$ProcessedCount,
        [string]$LastStatus
    )

    $state = [ordered]@{
        LastProcessedSourceUrl = $LastProcessedSourceUrl
        ProcessedCount         = $ProcessedCount
        LastStatus             = $LastStatus
        UpdatedAt              = (Get-Date).ToString('o')
    }

    $state | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $CheckpointPath -Encoding UTF8
}

function Load-Checkpoint {
    if (-not $ResumeFromCheckpoint -or -not (Test-Path -LiteralPath $CheckpointPath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $CheckpointPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Host "Warning: could not read checkpoint file '$CheckpointPath': $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

# Recursive function: walks one folder at a time to avoid List View Threshold.
# Each Get-PnPListItem call is scoped to a single folder (stays well under 5000).
# Get-PnPFolderItem enumerates sub-folders without touching the list index at all.
function Process-Folder {

    param(

        [string]$SourceFolderServerRelUrl,

        [string]$SourceFolderSiteRelUrl,

        [string]$TargetFolderSiteRelUrl,

        [string]$TargetFolderServerRelUrl

    )

    Write-Host "Scanning folder: $SourceFolderServerRelUrl" -ForegroundColor Cyan

    # Ensure the matching target folder exists before moving anything into it
    Ensure-FolderPath -FolderServerRelUrl $TargetFolderServerRelUrl -LibRootServerRelUrl $TargetLibRootServerRelUrl -Connection $TargetConn

    # ---- FILES in this folder only ----
    # Scoped to a single folder so the query never hits the List View Threshold,
    # even when the library as a whole contains millions of items.
    $gliArgs = @{
        List                    = $SourceLibrary
        FolderServerRelativeUrl = $SourceFolderServerRelUrl
        PageSize                = 500
        Fields                  = @('FileRef', 'FileLeafRef', 'FSObjType')
        Connection              = $SourceConn
        ErrorAction             = 'Stop'
    }
    $FolderFiles = Invoke-WithRetry -OperationName "Get-PnPListItem ($SourceFolderServerRelUrl)" -ScriptBlock {
        Get-PnPListItem @gliArgs |
            Where-Object { $_.FileSystemObjectType -eq 'File' } |
            Sort-Object { $_['FileRef'] }
    }

    foreach ($Item in $FolderFiles) {
        $SourceFileUrl = $Item["FileRef"]

        if ($script:ResumeAfterSourceUrl -and -not $script:ResumeStarted) {
            if ([string]::Compare($SourceFileUrl, $script:ResumeAfterSourceUrl, [System.StringComparison]::OrdinalIgnoreCase) -le 0) {
                Write-Host "Skipping already processed file during resume: $SourceFileUrl" -ForegroundColor DarkGray
                continue
            }

            $script:ResumeStarted = $true
        }

        $script:Counter++

        $FileName = $Item["FileLeafRef"]

        # TargetUrl for Move-PnPFile should be just the folder, not the full file path.
        # The filename is preserved automatically during the move.
        $TargetFileUrl = Join-Url $TargetFolderServerRelUrl $FileName

        Write-Host "[$($script:Counter)] Processing: $SourceFileUrl" -ForegroundColor Cyan

        try {
            # Same tenant, so Move-PnPFile can handle cross-site moves with full server-relative URLs.
            # TargetUrl is the destination folder path (filename is preserved).
            $moveArgs = @{
                SourceUrl                             = $SourceFileUrl
                TargetUrl                             = $TargetFolderServerRelUrl
                AllowSchemaMismatch                   = $true
                AllowSmallerVersionLimitOnDestination = $true
                Force                                 = $true
                Connection                            = $SourceConn
                ErrorAction                           = 'Stop'
            }
            if ($OverwriteExisting) { $moveArgs['Overwrite'] = $true }
            
            Invoke-WithRetry -OperationName "Move-PnPFile ($SourceFileUrl)" -ScriptBlock {
                Move-PnPFile @moveArgs
            }

            Write-Host "Moved to: $TargetFileUrl" -ForegroundColor Green
            Write-MoveLogEntry -SourceUrl $SourceFileUrl -TargetUrl $TargetFileUrl -Status "Moved" -Message ""
            Save-Checkpoint -LastProcessedSourceUrl $SourceFileUrl -ProcessedCount $script:Counter -LastStatus "Moved"
        }
        catch {
            # Check if the error is because the destination file already exists
            $errMsg = $_.Exception.Message + ''
            if ($errMsg -match 'already exist|duplicate|conflict') {
                if ($DeleteSourceIfTargetExists) {
                    # Target has the file — delete the source copy
                    try {
                        Invoke-WithRetry -OperationName "Remove-PnPFile ($SourceFileUrl)" -ScriptBlock {
                            Remove-PnPFile -ServerRelativeUrl $SourceFileUrl -Force -Recycle -Connection $SourceConn
                        }
                        Write-Host "Source deleted (target already had file): $SourceFileUrl" -ForegroundColor Green
                                                Write-MoveLogEntry -SourceUrl $SourceFileUrl -TargetUrl $TargetFileUrl -Status "Source Deleted - Target Already Existed" -Message ""
                                                Save-Checkpoint -LastProcessedSourceUrl $SourceFileUrl -ProcessedCount $script:Counter -LastStatus "Source Deleted - Target Already Existed"
                    }
                    catch {
                        Write-Host "Failed to delete source: $SourceFileUrl" -ForegroundColor Red
                                                Write-MoveLogEntry -SourceUrl $SourceFileUrl -TargetUrl $TargetFileUrl -Status "Failed - Could Not Delete Source" -Message $_.Exception.Message
                    }
                }
                else {
                    Write-Host "Skipped, file already exists at destination: $TargetFileUrl" -ForegroundColor Yellow
                                            Write-MoveLogEntry -SourceUrl $SourceFileUrl -TargetUrl $TargetFileUrl -Status "Skipped - Already Exists at Destination" -Message ""
                                            Save-Checkpoint -LastProcessedSourceUrl $SourceFileUrl -ProcessedCount $script:Counter -LastStatus "Skipped - Already Exists at Destination"
                }
            }
            else {
                # Some other error — log it
                Write-Host "Failed: $SourceFileUrl" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor Red
                                        Write-MoveLogEntry -SourceUrl $SourceFileUrl -TargetUrl $TargetFileUrl -Status "Failed" -Message $_.Exception.Message
            }
        }

                                if ($CheckpointEvery -gt 0 -and $script:Counter % $CheckpointEvery -eq 0) {
                                    Write-Host "  [Checkpoint] Saved at $($script:Counter) items." -ForegroundColor DarkGray
                                }

    }

    # ---- SUB-FOLDERS ----
    # Get-PnPFolderItem with -ItemType Folder does NOT query the list index,
    # so it is completely immune to the List View Threshold.
    # System folders (Forms, _t, _w, _vti_*) are skipped - they are SharePoint
    # infrastructure and should not be migrated.
    $systemFolders = @('Forms', '_t', '_w', '_vti_bin', '_vti_pvt', '_vti_cnf', '_vti_script', '_catalogs')

    try {
        $gfiArgs = @{
            FolderSiteRelativeUrl = $SourceFolderSiteRelUrl
            ItemType              = 'Folder'
            Connection            = $SourceConn
            ErrorAction           = 'Stop'
        }
        $SubFolders = Invoke-WithRetry -OperationName "Get-PnPFolderItem ($SourceFolderSiteRelUrl)" -ScriptBlock {
            Get-PnPFolderItem @gfiArgs | Sort-Object Name
        }
    }
    catch {
        Write-Host "Warning: could not enumerate sub-folders of '$SourceFolderSiteRelUrl': $($_.Exception.Message)" -ForegroundColor Yellow
        $SubFolders = @()
    }

    foreach ($SubFolder in $SubFolders) {
        $Name = $SubFolder.Name
        if ($systemFolders -contains $Name -or $Name -like '_vti_*') {
            Write-Host "  Skipping system folder: $Name" -ForegroundColor DarkGray
            continue
        }
        Process-Folder `
            -SourceFolderServerRelUrl (Join-Url $SourceFolderServerRelUrl $Name) `
            -SourceFolderSiteRelUrl   (Join-Url $SourceFolderSiteRelUrl   $Name) `
            -TargetFolderSiteRelUrl   (Join-Url $TargetFolderSiteRelUrl   $Name) `
            -TargetFolderServerRelUrl (Join-Url $TargetFolderServerRelUrl $Name)
    }

}

# =========================

# BUILD SOURCE/TARGET PATHS

# =========================

# Resolve actual library root folder server-relative URLs.
# The display name ($SourceLibrary / $TargetLibrary) may differ from the URL slug,
# so we look it up directly from the list's RootFolder property.
Write-Host "Resolving library root URLs..." -ForegroundColor Cyan

$SourceLibList = Get-PnPList -Identity $SourceLibrary -Connection $SourceConn -Includes RootFolder
$SourceLibRootServerRelUrl = $SourceLibList.RootFolder.ServerRelativeUrl.TrimEnd('/')

$TargetLibList = Get-PnPList -Identity $TargetLibrary -Connection $TargetConn -Includes RootFolder
$TargetLibRootServerRelUrl = $TargetLibList.RootFolder.ServerRelativeUrl.TrimEnd('/')

Write-Host "Source library root: $SourceLibRootServerRelUrl" -ForegroundColor Yellow
Write-Host "Target library root: $TargetLibRootServerRelUrl" -ForegroundColor Yellow

# Build server-relative root paths using the verified library URLs
$SourceRootServerRelative = if ([string]::IsNullOrWhiteSpace($SourceRootFolder)) {
    $SourceLibRootServerRelUrl
}
else {
    "$SourceLibRootServerRelUrl/$($SourceRootFolder.Trim('/'))"
}

$TargetRootServerRelative = if ([string]::IsNullOrWhiteSpace($TargetRootFolder)) {
    $TargetLibRootServerRelUrl
}
else {
    "$TargetLibRootServerRelUrl/$($TargetRootFolder.Trim('/'))"
}

# Site-relative versions (used by Get-PnPFolderItem -FolderSiteRelativeUrl)
$SourceRootSiteRelative = $SourceRootServerRelative.Substring($SourceWebRelUrl.Length).TrimStart('/')
$TargetRootSiteRelative = $TargetRootServerRelative.Substring($TargetWebRelUrl.Length).TrimStart('/')
 
Write-Host "Source root: $SourceRootServerRelative" -ForegroundColor Yellow

Write-Host "Target root: $TargetRootServerRelative" -ForegroundColor Yellow
 
# Ensure target root exists

Ensure-FolderPath -FolderServerRelUrl $TargetRootServerRelative -LibRootServerRelUrl $TargetLibRootServerRelUrl -Connection $TargetConn
 
# =========================

# MOVE FILES (recursive, folder-by-folder)

# =========================

# Processing is done folder-by-folder inside Process-Folder so that no single
# API call ever spans more than one folder's contents, keeping every query
# well below SharePoint's 5000-item List View Threshold.

Initialize-RunStorage

$checkpointState = Load-Checkpoint
$script:ResumeAfterSourceUrl = $null
$script:ResumeStarted = $false
$script:Counter = 0

if ($null -ne $checkpointState -and $checkpointState.LastProcessedSourceUrl) {
    $script:ResumeAfterSourceUrl = [string]$checkpointState.LastProcessedSourceUrl
    if ($null -ne $checkpointState.ProcessedCount) {
        $script:Counter = [int]$checkpointState.ProcessedCount
    }

    Write-Host "Resuming after checkpoint: $($script:ResumeAfterSourceUrl)" -ForegroundColor Yellow
}
else {
    Write-Host "Starting fresh run." -ForegroundColor Yellow
}

Write-Host "Starting recursive move from: $SourceRootServerRelative" -ForegroundColor Cyan

$pfArgs = @{
    SourceFolderServerRelUrl = $SourceRootServerRelative
    SourceFolderSiteRelUrl   = $SourceRootSiteRelative
    TargetFolderSiteRelUrl   = $TargetRootSiteRelative
    TargetFolderServerRelUrl = $TargetRootServerRelative
}
Process-Folder @pfArgs

if (Test-Path -LiteralPath $CheckpointPath) {
    Remove-Item -LiteralPath $CheckpointPath -Force
}
 
Write-Host "Move completed." -ForegroundColor Green

Write-Host "Log file: $LogPath" -ForegroundColor Green
 