# Edge Removal Script for Windows 11 ISO (PowerShell 7+)
# Disclaimer: Use at your own risk. Always back up your data before making system changes. Please be advised if you use Edge Browser and WebView components,
# the script is not intended for such use cases as removing them does not make sense

param(
    [Parameter(Mandatory = $true)]
    [string]$IsoOrWimPath
)

# Import only required modules
Import-Module -Name Dism -ErrorAction Stop
Import-Module -Name Storage -ErrorAction Stop
Import-Module -Name CimCmdlets -ErrorAction Stop

# Start timer
$scriptStartTime = Get-Date

# Todo (Done): Add Optimized Export Image to Rebuild WIM after edits to reduce size (credits: abbodi1406 from MDL Forums)
# Todo: ESD to WIM conversion option for ESD inputs
# Todo: Improve slow processing time with powershell dism modules

# Function to optimize/rebuild WIM image (credits: abbodi1406)
function Optimize-WimImage {
    param([string]$WimPath)
    Write-Host "Starting WIM optimization/export... (credits: abbodi1406)" -ForegroundColor Cyan
    # Ensure $WimPath points to sources\install.wim under the root folder
    if (!(Test-Path $WimPath) -or ($WimPath -notmatch "sources\\install\.wim$")) {
        $possibleWim = Join-Path ([IO.Path]::GetDirectoryName($WimPath)) 'sources\install.wim'
        if (Test-Path $possibleWim) {
            $WimPath = $possibleWim
        }
        else {
            Write-Error "Could not locate install.wim under sources\ in the root folder."
            return
        }
    }
    $WimTemp = [IO.Path]::GetDirectoryName($WimPath) + '\temp.wim'
    $count = (Get-WindowsImage -ImagePath $WimPath).Count
    try {
        1..$count | foreach {
            Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $_ -CheckIntegrity -DestinationImagePath $WimTemp
        }
        if (Test-Path $WimTemp) {
            Move-Item -Path $WimTemp -Destination $WimPath -Force
            Write-Host "Optimized WIM has replaced original install.wim" -ForegroundColor Green
        }
        else {
            Write-Host "WIM optimization failed: temp.wim not found." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Export failed or was interrupted. Cleaning up temp.wim..." -ForegroundColor Red
        if (Test-Path $WimTemp) { Remove-Item -Path $WimTemp -Force }
    }
}

# Cleanup Old Mounts and Temp Folders
function Cleanup-WimMounts {
    $oldWimMounts = Get-ChildItem -Path $env:TEMP -Directory -Filter 'WimMount_*' -ErrorAction SilentlyContinue
    foreach ($wm in $oldWimMounts) {
        try {
            Remove-Item -Path $wm.FullName -Recurse -Force
            Write-Host "Removed old WimMount folder: $($wm.FullName)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "Could not remove old WimMount folder: $($wm.FullName): ${_}" -ForegroundColor Red
        }
    }
}

# Cleanup ISO Extract Folders if old ones exist and use fresh ones every session
function Cleanup-ISOExtracts {
    $oldIsoExtracts = Get-ChildItem -Path $env:TEMP -Directory -Filter 'ISOExtract_*' -ErrorAction SilentlyContinue
    foreach ($old in $oldIsoExtracts) {
        try {
            Remove-Item -Path $old.FullName -Recurse -Force
            Write-Host "Removed old extracted ISO folder: $($old.FullName)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "Could not remove old extracted ISO folder: $($old.FullName): ${_}" -ForegroundColor Red
        }
    }
}

function Cleanup-Mountpoints {
    Write-Host "Running: dism /Cleanup-Mountpoints" -ForegroundColor Cyan
    dism /Cleanup-Mountpoints
}

Cleanup-WimMounts
Cleanup-ISOExtracts
Cleanup-Mountpoints

# Remove extra quotes if present
$cleanPath = $IsoOrWimPath.Trim('"').Trim()
$isoExtracted = $false
$tempExtractPath = ""
if (Test-Path $cleanPath) {
    $item = Get-Item $cleanPath
    if ($item.PSIsContainer) {
        $WimPath = Join-Path $cleanPath 'sources\install.wim'
    }
    elseif ($cleanPath -match "\.iso$") {
        # Extract ISO contents to temp folder
        $tempExtractPath = Join-Path $env:TEMP ("ISOExtract_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempExtractPath | Out-Null
        Write-Host "Mounting ISO: $cleanPath" -ForegroundColor Cyan
        $mountResult = Mount-DiskImage -ImagePath $cleanPath -PassThru
        $driveLetter = ($mountResult | Get-Volume).DriveLetter
        if ($driveLetter) {
            $isoDrive = "${driveLetter}:\"
            Write-Host "Copying ISO contents to $tempExtractPath..." -ForegroundColor Cyan
            Copy-Item -Path $isoDrive\* -Destination $tempExtractPath -Recurse
            Dismount-DiskImage -ImagePath $cleanPath
            # Remove read-only attribute from all files in extracted folder
            Get-ChildItem -Path $tempExtractPath -Recurse -File | ForEach-Object { Set-ItemProperty -Path $_.FullName -Name Attributes -Value ((Get-ItemProperty -Path $_.FullName -Name Attributes).Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)) }
            $WimPath = Join-Path $tempExtractPath 'sources\install.wim'
            $isoExtracted = $true
        }
        else {
            Write-Error "Failed to mount ISO."
            exit 1
        }
    }
    else {
        $WimPath = $cleanPath
    }
}
else {
    Write-Error "Path '$cleanPath' does not exist."
    exit 1
}

# Check for DISM powershell module
if (-not (Get-Module -ListAvailable -Name DISM)) {
    Write-Error "DISM PowerShell module is not available. Please install it before running this script."
    exit 1
}

# Validate file extension
if ($WimPath -notmatch "\.wim$") {
    if ($WimPath -match "\.esd$") {
        Write-Error "ESD files are unsupported. Please provide a WIM file."
        exit 1
    }
    else {
        Write-Error "File '$WimPath' is not a WIM file."
        exit 1
    }
}

# Check for Admin Privileges before continuing
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }
}

function Get-WimEditions {
    param([string]$WimPath)
    $images = Get-WindowsImage -ImagePath $WimPath
    $editions = @()
    foreach ($img in $images) {
        $editions += [PSCustomObject]@{
            Index   = $img.ImageIndex
            Edition = $img.ImageName
        }
    }
    return $editions
}

# Mount offline WIM image for servicing
function Mount-Wim {
    param([string]$WimPath, [int]$Index, [string]$MountPath)
    Mount-WindowsImage -ImagePath $WimPath -Index $Index -Path $MountPath -CheckIntegrity -Optimize
}

# Commit changes and unmount WIM image with proper integrity checks
function Commit-Wim {
    param([string]$MountPath)
    Dismount-WindowsImage -Path $MountPath -Save -CheckIntegrity
}

function Run-DismRemove {
    param([string]$MountPath, [string]$Option)
    Write-Host "Running: dism /image:'$MountPath' $Option"
    dism /image:"$MountPath" $Option
}

Ensure-Admin


# Ensure $WimPath exists before continuing
if (!(Test-Path $WimPath)) {
    Write-Error "WIM file not found: $WimPath"
    exit 1
}

$editions = Get-WimEditions -WimPath $WimPath
Write-Host "Available Editions in ${WimPath}:" -ForegroundColor Cyan
foreach ($e in $editions) {
    Write-Host "$($e.Index): $($e.Edition)"
}
Write-Host "*: All editions" -ForegroundColor Yellow



$indexInput = Read-Host "Enter the index number(s) of the edition(s) to modify (e.g. 1,3,5 or * for all editions)"

Write-Host "Select operation:" -ForegroundColor Cyan
Write-Host "0: Cancel operation"
Write-Host "1: Remove All Edge Components"
Write-Host "2: Remove Edge Browser"
Write-Host "3: Remove Edge WebView"
Write-Host "4: Optimize WIM image for export (credits: abbodi1406)"

$choice = Read-Host "Enter your choice (0/1/2/3/4)"
if ($choice -eq '4') {
    Optimize-WimImage -WimPath $WimPath
    # Cleanup and exit after optimization
    Cleanup-WimMounts
    Cleanup-ISOExtracts
    Cleanup-Mountpoints
    $scriptEndTime = Get-Date
    $elapsed = $scriptEndTime - $scriptStartTime
    if ($elapsed.TotalMinutes -ge 1) {
        $elapsedMsg = "Time elapsed for WIM optimization: {0:N2} minutes" -f $elapsed.TotalMinutes
    }
    else {
        $elapsedMsg = "Time elapsed for WIM optimization: {0:N2} seconds" -f $elapsed.TotalSeconds
    }
    Write-Host $elapsedMsg -ForegroundColor Cyan
    exit 0
}
if ($choice -eq '0') {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    # Discard any mounted image
    $mountedImages = Get-WindowsImage -Mounted | Where-Object { $_.Mounted } | Select-Object -ExpandProperty Path
    foreach ($mp in $mountedImages) {
        try {
            Dismount-WindowsImage -Path $mp -Discard
            Write-Host "Discarded mounted image at $mp" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "Could not discard mounted image at ${mp}: ${_}" -ForegroundColor Red
        }
    }
    Cleanup-ISOExtracts
    Cleanup-WimMounts
    Cleanup-Mountpoints
    exit 0
}

function Process-Edition {
    param([int]$idx)
    $mountPath = "$env:TEMP\WimMount_${idx}"
    if (!(Test-Path $mountPath)) { New-Item -ItemType Directory -Path $mountPath | Out-Null }
    $errorsFound = $false
    try {
        Mount-Wim -WimPath $WimPath -Index $idx -MountPath $mountPath
    }
    catch {
        Write-Host "Failed to mount edition index ${idx}: ${_}" -ForegroundColor Red
        $errorsFound = $true
        return
    }
    try {
        switch ($choice) {
            '0' { Write-Host "Operation cancelled by user." -ForegroundColor Yellow; exit 0 }
            '1' { Run-DismRemove -MountPath $mountPath -Option "/Remove-Edge" }
            '2' { Run-DismRemove -MountPath $mountPath -Option "/Remove-EdgeBrowser" }
            '3' { Run-DismRemove -MountPath $mountPath -Option "/Remove-EdgeWebView" }
            default { Write-Host "Invalid choice."; exit 1 }
        }
        Commit-Wim -MountPath $mountPath
        Write-Host "Changes committed to ${WimPath} for edition index ${idx}." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to commit edition index ${idx}: ${_}" -ForegroundColor Red
    }
}



# Validate selected indexes
$mountPaths = @()
if ($indexInput -eq '*') {
    $selectedIndexes = $editions | ForEach-Object { $_.Index }
}
else {
    $selectedIndexes = $indexInput -split ',' | ForEach-Object { $_.Trim() }
    $validIndexes = $editions | ForEach-Object { $_.Index }
    $selectedIndexes = $selectedIndexes | Where-Object { $validIndexes -contains $_ }
    if ($selectedIndexes.Count -eq 0) {
        Write-Host "No valid edition indexes selected. Exiting." -ForegroundColor Red
        exit 1
    }
}

# process every editions/index in multi edition WIM/ISO file

foreach ($idx in $selectedIndexes) {
    Process-Edition -idx $idx
    $mountPaths += "$env:TEMP\WimMount_${idx}"
}

# After all editions is/are processed, optimize the WIM image
Optimize-WimImage -WimPath $WimPath

if ($choice -eq '4') {
    Write-Host "Starting WIM optimization/export... (credits: abbodi1406)" -ForegroundColor Cyan
    $WimTemp = [IO.Path]::GetDirectoryName($WimPath) + '\temp.wim'
    $count = (Get-WindowsImage -ImagePath $WimPath).Count
    1..$count | foreach {
        Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $_ -DestinationImagePath $WimTemp
    }
    if (Test-Path $WimTemp) {
        Move-Item -Path $WimTemp -Destination $WimPath -Force
        Write-Host "Optimized WIM has replaced original install.wim" -ForegroundColor Green
    }
    else {
        Write-Host "WIM optimization failed: temp.wim not found." -ForegroundColor Red
    }
    # Cleanup and exit after optimization
    Cleanup-WimMounts
    Cleanup-ISOExtracts
    Cleanup-Mountpoints
    $scriptEndTime = Get-Date
    $elapsed = $scriptEndTime - $scriptStartTime
    if ($elapsed.TotalMinutes -ge 1) {
        $elapsedMsg = "Time elapsed for WIM optimization: {0:N2} minutes" -f $elapsed.TotalMinutes
    }
    else {
        $elapsedMsg = "Time elapsed for WIM optimization: {0:N2} seconds" -f $elapsed.TotalSeconds
    }
    Write-Host $elapsedMsg -ForegroundColor Cyan
    exit 0
}

# If input was ISO, save updated ISO before cleanup
if ($isoExtracted -and (Test-Path $tempExtractPath)) {
    $outputIso = Join-Path (Get-Location) 'Updated_Win11.iso'
    Write-Host "Saving updated ISO as $outputIso..." -ForegroundColor Cyan
    try {
        # Use oscdimg if available, otherwise show message
        if (Get-Command oscdimg.exe -ErrorAction SilentlyContinue) {
            # Use recommended flags for optimized ISO: -n (long names), -m (larger than 700MB), -o (optimize storage), -u2 (UDF 2.01), -l (label)
            $isoLabel = "WIN11_UPDATED"
            oscdimg -n -m -o -u2 -l$isoLabel $tempExtractPath $outputIso
            Write-Host "Updated ISO saved as $outputIso (optimized)" -ForegroundColor Green
        }
        else {
            Write-Host "oscdimg.exe not found. Please install Windows ADK to enable ISO creation. Visit the link here: https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to save updated ISO: ${_}" -ForegroundColor Red
    }
}

Cleanup-WimMounts
Cleanup-ISOExtracts
Cleanup-Mountpoints


# Calculate and display elapsed time
$scriptEndTime = Get-Date
$elapsed = $scriptEndTime - $scriptStartTime
if ($elapsed.TotalMinutes -ge 1) {
    $elapsedMsg = "Time elapsed: {0:N2} minutes" -f $elapsed.TotalMinutes
}
else {
    $elapsedMsg = "Time elapsed: {0:N2} seconds" -f $elapsed.TotalSeconds
}

if ($errorsFound) {
    Write-Host "Process completed with errors. Check above for details." -ForegroundColor Red
    Write-Host $elapsedMsg -ForegroundColor Yellow
}
else {
    Write-Host "Process completed successfully!" -ForegroundColor Green
    Write-Host $elapsedMsg -ForegroundColor Cyan
}
