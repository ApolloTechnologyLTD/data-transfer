<#
.SYNOPSIS
    Apollo Technology Data Migration Utility (Smart Restore & Backup) v2.0
.DESCRIPTION
    Menu-driven utility to Backup data or Restore data.
    - AUTO-INSTALLS Google Chrome if missing during Restore.
    - BACKUPS Edge, Chrome, Firefox, Opera, Brave, and full AppData.
    - GENERATES HTML/PDF Reports (Same style as Health Check).
#>

# --- 0. CONFIGURATION ---
# CHANGE THIS TO $FALSE WHEN READY FOR REAL USE
$DemoMode = $false
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"

# --- 1. AUTO-ELEVATE TO ADMINISTRATOR ---
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    try {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
    catch {
        Write-Error "Failed to elevate. Please run as Administrator manually."
        Pause
        Exit
    }
}

# --- 2. HELPER FUNCTIONS ---
function Show-Header {
    Clear-Host
    $Banner = @'
    ___    ____  ____  __    __    ____     ____________________  ___   ______  __    ____  ________  __
   /   |  / __ \/ __ \/ /   / /   / __ \   /_  __/ ____/ ____/ / / / | / / __ \/ /   / __ \/ ____/\ \/ /
  / /| | / /_/ / / / / /   / /   / / / /    / / / __/ / /   / /_/ /  |/ / / / / /   / / / / / __   \  / 
 / ___ |/ ____/ /_/ / /___/ /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
'@
    Write-Host $Banner -ForegroundColor Cyan
    Write-Host "`n   DATA MIGRATION & BACKUP TOOL" -ForegroundColor White
    Write-Host "=================================================================================" -ForegroundColor DarkGray
    if ($DemoMode) {
        Write-Host "`n   *** DEMO MODE ACTIVE - NO REAL DATA WILL BE COPIED ***" -ForegroundColor Magenta
    }
}

function Install-GoogleChrome {
    Write-Host "`n[ CHROME INSTALLER ]" -ForegroundColor Yellow
    Write-Host "   Chrome not found. Downloading Enterprise Installer..." -ForegroundColor Cyan
    
    $InstallerUrl = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
    $InstallerPath = "$env:TEMP\ChromeEnterprise.msi"

    try {
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing
        Write-Host "   Installing Chrome (Silent)..." -ForegroundColor Cyan
        
        $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /quiet /norestart" -Wait -PassThru
        
        if ($Process.ExitCode -eq 0) {
            Write-Host "   SUCCESS: Google Chrome installed." -ForegroundColor Green
        } else {
            Write-Host "   WARNING: Install exited with code $($Process.ExitCode)." -ForegroundColor Red
        }
    }
    catch {
        Write-Error "   Failed to install Chrome. Please install manually."
    }
}

function Run-Robocopy {
    param (
        [string]$Source,
        [string]$Destination,
        [string]$LogFile
    )

    # Standard Exclusions for AppData to avoid loops/bloat
    $Excludes = @("Temp", "Temporary Internet Files", "Application Data", "History", "Cookies")

    if ($DemoMode) {
        # Create Dummy Dest for visualization
        if (!(Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
        
        Write-Host "   [DEMO] Processing: $Source" -ForegroundColor Magenta
        Add-Content -Path $LogFile -Value "DEMO COPY: $Source -> $Destination"
        Start-Sleep -Milliseconds 100
        return "Success (Demo)"
    }
    else {
        if (Test-Path $Source) {
            Write-Host "   Copying: $Source" -ForegroundColor Green
            # Robocopy Args: /E (Recursive) /XO (Exclude Older) /R:1 /W:1 /NP (No progress) /XD (Exclude Dirs)
            robocopy $Source $Destination /E /XO /R:1 /W:1 /NP /LOG+:"$LogFile" /TEE /XD $Excludes | Out-Null
            return "Completed"
        } else {
            Write-Host "   Skipping: Source not found ($Source)" -ForegroundColor DarkGray
            Add-Content -Path $LogFile -Value "SKIPPED: Source Missing - $Source"
            return "Skipped (Not Found)"
        }
    }
}

# --- 3. MAIN MENU ---
Show-Header
Write-Host "`n[ SELECT OPERATION MODE ]" -ForegroundColor Yellow
Write-Host "   1. BACKUP  (Computer -> External Drive)"
Write-Host "   2. RESTORE (External Drive -> New Computer)"
Write-Host "   3. EXIT"
Write-Host "---------------------------------------------------------------------------------"

$MenuSelection = Read-Host "   > Enter Option (1, 2, or 3)"

switch ($MenuSelection) {
    '1' { $Mode = "BACKUP" }
    '2' { $Mode = "RESTORE" }
    '3' { Exit }
    Default { Write-Host "Invalid selection."; Pause; Exit }
}

# --- 4. INPUT COLLECTION ---
Show-Header
Write-Host "`n[ $Mode CONFIGURATION ]" -ForegroundColor Yellow

# Shared Inputs
$EngineerName = Read-Host "   > Enter Engineer Name"
$TicketNumber = Read-Host "   > Enter Ticket Number"

# Drive Selection
Write-Host "`n[ DRIVE SELECTION ]" -ForegroundColor Yellow
Write-Host "Detecting drives..." -ForegroundColor DarkGray
Get-PSDrive -PSProvider FileSystem | Select-Object Name, Used, Free, Root | Format-Table -AutoSize

$DriveLetterInput = Read-Host "   > Enter External Drive Letter (e.g. D or E)"
$DriveLetter = $DriveLetterInput -replace ":", ""

if (!(Test-Path "$($DriveLetter):")) {
    Write-Error "Drive $($DriveLetter): not found."
    Pause
    Exit
}

# --- 5. PATH LOGIC & SEARCH ---
$UserProfile = $env:USERPROFILE

if ($Mode -eq "BACKUP") {
    $CustomerName = Read-Host "   > Enter Customer Full Name"
    $ExternalStorePath = "$($DriveLetter):\${TicketNumber}-$($CustomerName)"
    
} elseif ($Mode -eq "RESTORE") {
    # SEARCH for the ticket number
    Write-Host "`n[ SEARCHING FOR BACKUP ]" -ForegroundColor Yellow
    Write-Host "   Searching drive $DriveLetter for Ticket $TicketNumber..." -ForegroundColor DarkGray
    
    $SearchPath = "$($DriveLetter):\${TicketNumber}-*"
    $FoundFolders = Get-ChildItem -Path $SearchPath -Directory -ErrorAction SilentlyContinue
    
    if ($null -eq $FoundFolders -or $FoundFolders.Count -eq 0) {
        Write-Host "`n[ ERROR ] NO BACKUP FOUND" -ForegroundColor Red
        Pause; Exit
    } elseif ($FoundFolders.Count -gt 1) {
        Write-Host "`n[ ERROR ] MULTIPLE BACKUPS FOUND" -ForegroundColor Red
        $FoundFolders | ForEach-Object { Write-Host "   - $($_.Name)" }
        Pause; Exit
    }

    $ExternalStorePath = $FoundFolders.FullName
    $CustomerName = $FoundFolders.Name 
    Write-Host "   FOUND: $CustomerName" -ForegroundColor Green

    # --- CHROME CHECK (RESTORE ONLY) ---
    $ChromePath64 = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    $ChromePath32 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    
    if ((-not (Test-Path $ChromePath64)) -and (-not (Test-Path $ChromePath32))) {
        Install-GoogleChrome
    } else {
        Write-Host "   Chrome is already installed." -ForegroundColor Gray
    }
}

$LogPath = "$ExternalStorePath\_Logs"

# --- 6. CONFIRMATION ---
Show-Header
Write-Host "`n[ CONFIRMATION ]" -ForegroundColor Yellow
Write-Host "   Operation:   $Mode" -ForegroundColor $(If ($Mode -eq 'BACKUP') {'Green'} Else {'Cyan'})
Write-Host "   Ticket:      $TicketNumber"
Write-Host "   Customer:    $CustomerName"
if ($Mode -eq "BACKUP") {
    Write-Host "   Source:      THIS COMPUTER"
    Write-Host "   Destination: $ExternalStorePath"
} else {
    Write-Host "   Source:      $ExternalStorePath"
    Write-Host "   Destination: THIS COMPUTER"
}

Write-Host "---------------------------------------------------------------------------------"
$Confirm = Read-Host "Type 'Y' to proceed"
if ($Confirm -ne 'Y') { Exit }

# Setup Logs
if (!(Test-Path $ExternalStorePath)) { New-Item -ItemType Directory -Path $ExternalStorePath -Force | Out-Null }
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
$MainLog = "$LogPath\${Mode}_Log.txt"
Add-Content -Path $MainLog -Value "Operation: $Mode | Date: $(Get-Date)"

# Stop Processes
Write-Host "`n[ SYSTEM PREP ]" -ForegroundColor Yellow
if (-not $DemoMode) {
    Write-Host "Stopping Browsers and Outlook..." -ForegroundColor Gray
    Stop-Process -Name "chrome", "msedge", "outlook", "firefox" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# --- 7. MAPPING & EXECUTION ---
Write-Host "`n[ STARTING TRANSFER ]" -ForegroundColor Yellow

# Initialize Map
$FoldersMap = @Ordered{
    "Desktop"   = "Desktop"
    "Documents" = "Documents"
    "Downloads" = "Downloads"
    "Pictures"  = "Pictures"
    "Music"     = "Music"
    "Videos"    = "Videos"
    "AppData\Local\Google\Chrome\User Data" = "Chrome_UserData"
    "AppData\Local\Microsoft\Outlook"       = "Outlook_AppData"
    "Documents\Outlook Files"               = "Outlook_Documents"
}

# --- DYNAMIC BROWSER & APPDATA LOGIC ---
if ($Mode -eq "BACKUP") {
    # 1. Edge User Data (Requested)
    if (Test-Path "$UserProfile\AppData\Local\Microsoft\Edge\User Data") {
        $FoldersMap["AppData\Local\Microsoft\Edge\User Data"] = "Edge_UserData"
    }

    # 2. Check for Other Browsers (Requested: "All Browsers")
    $BrowserChecks = @{
        "Firefox" = "AppData\Roaming\Mozilla\Firefox"
        "Opera"   = "AppData\Roaming\Opera Software\Opera Stable"
        "Brave"   = "AppData\Local\BraveSoftware\Brave-Browser\User Data"
        "Vivaldi" = "AppData\Local\Vivaldi\User Data"
    }
    
    foreach ($Browser in $BrowserChecks.Keys) {
        $Path = $BrowserChecks[$Browser]
        if (Test-Path "$UserProfile\$Path") {
            Write-Host "   Detected Browser: $Browser" -ForegroundColor Cyan
            $FoldersMap[$Path] = "${Browser}_Data"
        }
    }

    # 3. Full AppData (Requested: "copy app data directory")
    # Note: Robocopy handles overlapping files fine, but we list these last.
    $FoldersMap["AppData\Roaming"] = "AppData_Roaming"
    $FoldersMap["AppData\Local"]   = "AppData_Local"
}

# Store results for report
$ReportItems = @()

foreach ($LocalSubPath in $FoldersMap.Keys) {
    $ExternalSubName = $FoldersMap[$LocalSubPath]
    
    $LocalFull  = "$UserProfile\$LocalSubPath"
    $ExternalFull = "$ExternalStorePath\$ExternalSubName"

    if ($Mode -eq "BACKUP") {
        $Result = Run-Robocopy -Source $LocalFull -Destination $ExternalFull -LogFile $MainLog
        $ReportItems += [PSCustomObject]@{ Item = $LocalSubPath; Status = $Result }
    }
    elseif ($Mode -eq "RESTORE") {
        # For restore, we check if the external folder exists before trying to copy back
        if (Test-Path $ExternalFull) {
            $Result = Run-Robocopy -Source $ExternalFull -Destination $LocalFull -LogFile $MainLog
            $ReportItems += [PSCustomObject]@{ Item = $ExternalSubName; Status = $Result }
        }
    }
}

# --- 8. ADVANCED REPORT GENERATION (Health Check Style) ---
Write-Host "`n[ REPORT GENERATION ]" -ForegroundColor Yellow

$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$ComputerInfo = Get-CimInstance Win32_ComputerSystem
$TransferTableRows = ""

foreach ($Row in $ReportItems) {
    $StatusColor = if ($Row.Status -match "Completed|Success") { "green" } else { "red" }
    $TransferTableRows += "<tr><td>$($Row.Item)</td><td><span style='color:$StatusColor'>$($Row.Status)</span></td></tr>"
}

# Build HTML
$HtmlFile = "$ExternalStorePath\${Mode}_Report.html"
$PdfFile  = "$ExternalStorePath\${Mode}_Report.pdf"

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', sans-serif; color: #333; padding: 20px; }
    .header { text-align: center; margin-bottom: 20px; }
    .header img { max-height: 100px; }
    h1 { color: #0056b3; margin-bottom: 5px; }
    .meta { font-size: 0.9em; color: #666; text-align: center; margin-bottom: 30px; }
    .section { background: #f9f9f9; padding: 15px; border-left: 6px solid #0056b3; margin-bottom: 20px; }
    .item { margin-bottom: 12px; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px; }
    .label { font-weight: bold; color: #444; display: block; margin-bottom: 2px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
    th { text-align: left; background: #eee; padding: 8px; border-bottom: 1px solid #ddd; }
    td { padding: 8px; border-bottom: 1px solid #ddd; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" onerror="this.style.display='none'">
    <h1>Data $Mode Report</h1>
    <p>Report generated by <strong>$EngineerName</strong> for ticket (<strong>$TicketNumber</strong>)</p>
    <div class="meta">
        <strong>Date:</strong> $CurrentDate | <strong>Customer:</strong> $CustomerName
    </div>
</div>

<h2>Job Information</h2>
<div class="section">
    <div class="item"><span class="label">Operation Mode:</span> $Mode</div>
    <div class="item"><span class="label">Workstation:</span> $($ComputerInfo.Name)</div>
    <div class="item"><span class="label">Engineer:</span> $EngineerName</div>
    <div class="item"><span class="label">Storage Path:</span> $ExternalStorePath</div>
</div>

<h2>Transfer Details</h2>
<div class="section">
    <table>
        <thead>
            <tr><th>Data Folder</th><th>Status</th></tr>
        </thead>
        <tbody>
            $TransferTableRows
        </tbody>
    </table>
</div>

<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $(Get-Date -Format yyyy) by Apollo Technology.</p>
</body>
</html>
"@

$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8

# Convert to PDF using Edge (Headless)
$EdgeLoc1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$EdgeLoc2 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$EdgeExe = if (Test-Path $EdgeLoc1) { $EdgeLoc1 } elseif (Test-Path $EdgeLoc2) { $EdgeLoc2 } else { $null }

if ($EdgeExe) {
    Write-Host "   Generating PDF Report..." -ForegroundColor Cyan
    $EdgeUserData = "$ExternalStorePath\EdgeTemp"
    if (-not (Test-Path $EdgeUserData)) { New-Item -Path $EdgeUserData -ItemType Directory -Force | Out-Null }
    try {
        $Process = Start-Process -FilePath $EdgeExe -ArgumentList "--headless", "--disable-gpu", "--print-to-pdf=`"$PdfFile`"", "--no-pdf-header-footer", "--user-data-dir=`"$EdgeUserData`"", "`"$HtmlFile`"" -PassThru -Wait
        Start-Sleep -Seconds 2 
        if (Test-Path $PdfFile) {
            Write-Host "   Report saved to: $PdfFile" -ForegroundColor Green
            # Clean up HTML and Temp
            Remove-Item $HtmlFile -ErrorAction SilentlyContinue
            Remove-Item $EdgeUserData -Recurse -Force -ErrorAction SilentlyContinue
            
            # Open PDF
            Start-Process $PdfFile
        } else { throw "PDF creation failed" }
    } catch {
        Write-Warning "PDF Conversion failed. Report saved as HTML."
        Start-Process $HtmlFile
    }
} else {
    Write-Warning "Edge not found. Report saved as HTML."
    Start-Process $HtmlFile
}

Write-Host "`n[ COMPLETE ]" -ForegroundColor Green
Write-Host "Operation finished."
Pause