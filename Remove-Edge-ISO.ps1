# Edge Removal Script for Windows 11 ISO (PowerShell 7+)
# Disclaimer: Use at your own risk. Always back up your data before making system changes. Please be advised if you use Edge Browser and WebView components,
# the script is not intended for such use cases as removing them does not make sense

param(
    [Parameter(Mandatory = $true)]
    [string]$IsoOrWimPath,
    [Parameter(Mandatory = $false)]
    [string]$OriginalCwd
)

# Check for Admin Privileges and auto-elevate early
function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if ($isAdmin) { return }
    try {
        $launchDir = (Get-Location).Path
        $psExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
        $elevArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-IsoOrWimPath', $IsoOrWimPath, '-OriginalCwd', $launchDir)
        Start-Process -FilePath $psExe -ArgumentList $elevArgs -WorkingDirectory $launchDir -Verb RunAs | Out-Null
        exit 0
    }
    catch {
        Write-Error "Failed to request elevation: ${_}"
        exit 1
    }
}

Ensure-Admin

if ($OriginalCwd -and (Test-Path $OriginalCwd -PathType Container)) {
    Set-Location -Path $OriginalCwd
}

# Import only required modules
Import-Module -Name Dism -ErrorAction Stop
Import-Module -Name Storage -ErrorAction Stop
Import-Module -Name CimCmdlets -ErrorAction Stop

# Start timer
$scriptStartTime = Get-Date

# Track errors across editions
$script:errorsFound = $false
$script:IsoExtractPreservePath = $null

# Global ISO label
$GlobalIsoLabel = "Custom_Win11"

# Todo (Done): Add Optimized Export Image to Rebuild WIM after edits to reduce size (credits: abbodi1406 from MDL Forums)
# Todo (Done): ESD to WIM conversion option for ESD inputs
# Todo: Improve slow processing time with powershell dism modules

# Function to optimize/rebuild WIM image (credits: abbodi1406)
function Optimize-WimImage {
    param(
        [string]$WimPath,
        [int[]]$Indexes = @()
    )
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
    $allImages = Get-WindowsImage -ImagePath $WimPath
    if ($Indexes -and $Indexes.Count -gt 0) {
        $exportIndexes = $Indexes
    }
    else {
        $exportIndexes = $allImages | ForEach-Object { $_.ImageIndex }
    }
    try {
        foreach ($i in $exportIndexes) {
            Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $i -CheckIntegrity -DestinationImagePath $WimTemp
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

# Function to optimize/rebuild boot.wim image (similar to Optimize-WimImage)
function Optimize-BootWimImage {
    param(
        [string]$BootWimPath,
        [int[]]$Indexes = @()
    )
    Write-Host "Starting boot.wim optimization/export..." -ForegroundColor Cyan
    # Ensure $BootWimPath points to sources\boot.wim under the root folder
    if (!(Test-Path $BootWimPath) -or ($BootWimPath -notmatch "sources\\boot\.wim$")) {
        $possibleBootWim = Join-Path ([IO.Path]::GetDirectoryName($BootWimPath)) 'sources\boot.wim'
        if (Test-Path $possibleBootWim) {
            $BootWimPath = $possibleBootWim
        }
        else {
            Write-Error "Could not locate boot.wim under sources\ in the root folder."
            return
        }
    }
    $bootWimTemp = [IO.Path]::GetDirectoryName($BootWimPath) + '\boot_temp.wim'
    $allImages = Get-WindowsImage -ImagePath $BootWimPath
    if ($Indexes -and $Indexes.Count -gt 0) {
        $exportIndexes = $Indexes
    }
    else {
        $exportIndexes = $allImages | ForEach-Object { $_.ImageIndex }
    }
    try {
        foreach ($i in $exportIndexes) {
            Export-WindowsImage -SourceImagePath $BootWimPath -SourceIndex $i -CheckIntegrity -DestinationImagePath $bootWimTemp
        }
        if (Test-Path $bootWimTemp) {
            Move-Item -Path $bootWimTemp -Destination $BootWimPath -Force
            Write-Host "Optimized boot.wim has replaced original boot.wim" -ForegroundColor Green
        }
        else {
            Write-Host "boot.wim optimization failed: boot_temp.wim not found." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "boot.wim export failed or was interrupted. Cleaning up boot_temp.wim..." -ForegroundColor Red
        if (Test-Path $bootWimTemp) { Remove-Item -Path $bootWimTemp -Force }
    }
}

# Cleanup Old Mounts and Temp Folders
function Cleanup-WimMounts {
    $oldWimMounts = Get-ChildItem -Path $env:TEMP -Directory -Filter 'WimMount_*' -ErrorAction SilentlyContinue
    foreach ($wm in $oldWimMounts) {
        try {
            Remove-Item -Path $wm.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Removed old WimMount folder: $($wm.FullName)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "Could not remove old WimMount folder: $($wm.FullName): ${_}" -ForegroundColor Yellow
            # Try to discard the image if directory removal failed
            try {
                Discard-Image -MountPath $wm.FullName
            }
            catch {
                Write-Host "Could not discard image for $($wm.FullName): ${_}" -ForegroundColor Red
            }
        }
    }
}

# Cleanup ISO Extract Folders if old ones exist and use fresh ones every session
function Cleanup-ISOExtracts {
    $oldIsoExtracts = Get-ChildItem -Path $env:TEMP -Directory -Filter 'ISOExtract_*' -ErrorAction SilentlyContinue
    foreach ($old in $oldIsoExtracts) {
        if ($script:IsoExtractPreservePath) {
            $preservePath = $script:IsoExtractPreservePath
            try {
                if ((Resolve-Path $old.FullName).Path -eq (Resolve-Path $preservePath).Path) {
                    Write-Host "Preserving extracted ISO folder for manual export retry: $($old.FullName)" -ForegroundColor Yellow
                    continue
                }
            }
            catch { }
        }
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


# Define Pause-ForExit at top-level so it is always available
function Pause-ForExit {
    do {
        $resp = Read-Host "Press E or 0 to exit"
    } while ($resp -notmatch '^(?i:e|0)$')
}

function Read-YesNo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    do {
        $resp = [string](Read-Host "$Prompt (Y/N)")
        $resp = $resp.Trim()
    } while ($resp -notmatch '^(?i:y|yes|n|no)$')

    return ($resp -match '^(?i:y|yes)$')
}

function Get-DefaultIsoFileName {
    param([string]$SourcePath)

    $dateCode = (Get-Date).ToString('MMMyy')
    $prefix = 'Windows'
    try {
        if ($SourcePath -and (Test-Path $SourcePath)) {
            $installWim = Join-Path $SourcePath 'sources\install.wim'
            $installEsd = Join-Path $SourcePath 'sources\install.esd'
            $imagePath = $null
            if (Test-Path $installWim) { $imagePath = $installWim }
            elseif (Test-Path $installEsd) { $imagePath = $installEsd }

            if ($imagePath) {
                $img = Get-WindowsImage -ImagePath $imagePath -Index 1 -ErrorAction Stop
                $verText = [string]$img.Version
                if (-not [string]::IsNullOrWhiteSpace($verText)) {
                    $build = ([version]$verText).Build
                    if ($build -ge 22000) { $prefix = 'Win11' }
                    else { $prefix = 'Win10' }
                }
            }
        }
    }
    catch { }

    return ("{0}_{1}.iso" -f $prefix, $dateCode)
}

function Export-UpdatedIsoIfRequested {
    param(
        [bool]$IsoWasExtracted,
        [string]$TempExtractPath,
        [string]$IsoLabel
    )

    if (-not $IsoWasExtracted -or -not (Test-Path $TempExtractPath)) {
        return
    }

    $exportResp = [string](Read-Host "Export updated ISO before cleanup? (Y/N, default: Y)")
    $exportResp = $exportResp.Trim()
    if ([string]::IsNullOrWhiteSpace($exportResp)) { $exportResp = 'Y' }
    if ($exportResp -notmatch '^(?i:y|yes)$') {
        Write-Host "Skipping ISO export by user choice." -ForegroundColor Yellow
        return
    }

    $defaultName = Get-DefaultIsoFileName -SourcePath $TempExtractPath
    $outputResp = [string](Read-Host "Enter output ISO file name or full path (default: $defaultName)")
    $outputResp = $outputResp.Trim()
    if ([string]::IsNullOrWhiteSpace($outputResp)) { $outputResp = $defaultName }
    if (-not ($outputResp.ToLower().EndsWith('.iso'))) { $outputResp = "$outputResp.iso" }
    if ([IO.Path]::IsPathRooted($outputResp)) {
        $outputIsoPath = $outputResp
    }
    else {
        $outputIsoPath = Join-Path (Get-Location) $outputResp
    }

    $outDir = Split-Path -Path $outputIsoPath -Parent
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    Write-Host "Saving updated ISO as $outputIsoPath..." -ForegroundColor Cyan
    $exportResult = New-DualBootIso -SourcePath $TempExtractPath -OutputIso $outputIsoPath -Label $IsoLabel
    if (-not $exportResult) {
        Write-Host "ISO export failed. Keeping extracted ISO folder for manual retry: $TempExtractPath" -ForegroundColor Yellow
        $script:IsoExtractPreservePath = $TempExtractPath
        $script:errorsFound = $true
    }
    else {
        $script:IsoExtractPreservePath = $null
    }
}

# Require elevation before any servicing work
Ensure-Admin

Cleanup-WimMounts
Cleanup-ISOExtracts
Cleanup-Mountpoints

# Function to convert install.esd to install.wim using dism.exe
function convert-ESDWIM {
    param([string]$EsdPath)
    Write-Host "Starting ESD to WIM conversion using dism.exe..." -ForegroundColor Cyan
    if (!(Test-Path $EsdPath) -or ($EsdPath -notmatch "sources\\install\.esd$")) {
        $possibleEsd = Join-Path ([IO.Path]::GetDirectoryName($EsdPath)) 'sources\install.esd'
        if (Test-Path $possibleEsd) {
            $EsdPath = $possibleEsd
        }
        else {
            Write-Error "Could not locate install.esd under sources\ in the root folder."
            return
        }
    }
    $wimPath = [IO.Path]::GetDirectoryName($EsdPath) + '\install.wim'
    $count = (Get-WindowsImage -ImagePath $EsdPath).Count
    try {
        for ($i = 1; $i -le $count; $i++) {
            $dismArgs = @(
                "/Export-Image",
                "/SourceImageFile:$EsdPath",
                "/SourceIndex:$i",
                "/DestinationImageFile:$wimPath",
                "/Compress:max",
                "/CheckIntegrity"
            )
            Write-Host "Running: dism.exe $($dismArgs -join ' ')" -ForegroundColor Cyan
            $proc = Start-Process -FilePath dism.exe -ArgumentList $dismArgs -NoNewWindow -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                Write-Host "dism.exe export failed for index $i with exit code $($proc.ExitCode)" -ForegroundColor Red
                return
            }
        }
        if (Test-Path $wimPath) {
            Write-Host "Converted WIM has been created: $wimPath" -ForegroundColor Green
            # Delete original install.esd after successful WIM conversion
            try {
                Remove-Item -Path $EsdPath -Force
                Write-Host "Deleted original ESD: $EsdPath" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Could not delete original ESD: $EsdPath. ${_}" -ForegroundColor Red
            }
        }
        else {
            Write-Host "WIM export failed: install.wim not found." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "dism.exe export failed or was interrupted." -ForegroundColor Red
    }
}

# Reusable ISO generation function using dual-boot BIOS/UEFI boot files with relative paths
function New-DualBootIso {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$SourcePath,
        [Parameter(Mandatory = $true)] [string]$OutputIso,
        [string]$Label = $GlobalIsoLabel
    )
    # Resolve oscdimg.exe from Windows ADK per architecture, fallback to PATH
    $adkBase = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
    $archMap = @{ 'AMD64' = 'amd64'; 'ARM64' = 'arm64'; 'x86' = 'x86' }
    $procArch = $env:PROCESSOR_ARCHITECTURE
    $archFolder = $archMap[$procArch]
    $candidatePaths = @()
    if ($archFolder) { $candidatePaths += (Join-Path (Join-Path $adkBase $archFolder) 'Oscdimg\oscdimg.exe') }
    $candidatePaths += @(
        (Join-Path (Join-Path $adkBase 'amd64') 'Oscdimg\oscdimg.exe'),
        (Join-Path (Join-Path $adkBase 'arm64') 'Oscdimg\oscdimg.exe'),
        (Join-Path (Join-Path $adkBase 'x86')   'Oscdimg\oscdimg.exe')
    )
    $candidatePaths = @($candidatePaths | Select-Object -Unique)
    $oscdimgPath = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $oscdimgPath) {
        $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
        if ($cmd) { $oscdimgPath = $cmd.Source }
    }
    if (-not $oscdimgPath -and (Test-Path $adkBase)) {
        try {
            $found = Get-ChildItem -Path $adkBase -Filter oscdimg.exe -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) { $oscdimgPath = $found.FullName }
        }
        catch { }
    }
    if (-not $oscdimgPath) {
        Write-Host "oscdimg.exe not found." -ForegroundColor Yellow
        Write-Host "Searched default ADK paths:" -ForegroundColor Yellow
        foreach ($p in $candidatePaths) { Write-Host " - $p" -ForegroundColor DarkYellow }
        Write-Host "Install 'Windows Assessment and Deployment Kit (ADK)' with Deployment Tools, or add oscdimg.exe to PATH." -ForegroundColor Yellow
        Write-Host "If ADK is installed, typical locations are under: $adkBase" -ForegroundColor Yellow

        $customOscdimg = [string](Read-Host "If you have oscdimg.exe in a custom location, enter full path (or press Enter to cancel)")
        $customOscdimg = $customOscdimg.Trim().Trim('"').Trim("'")
        if ([string]::IsNullOrWhiteSpace($customOscdimg)) {
            return $false
        }
        if ((Test-Path $customOscdimg -PathType Leaf) -and ([IO.Path]::GetFileName($customOscdimg) -ieq 'oscdimg.exe')) {
            $oscdimgPath = $customOscdimg
        }
        else {
            Write-Host "Invalid oscdimg.exe path: $customOscdimg" -ForegroundColor Red
            return $false
        }
    }
    if (-not (Test-Path $SourcePath)) {
        Write-Host "Source path '$SourcePath' does not exist." -ForegroundColor Red
        return $false
    }
    Push-Location $SourcePath
    $biosBootRel = "boot\etfsboot.com"
    $uefiBootRel = "efi\microsoft\boot\efisys.bin"
    $oscdimgDir = Split-Path $oscdimgPath -Parent
    if (-not (Test-Path $biosBootRel)) {
        $adkEtfs = Join-Path $oscdimgDir 'etfsboot.com'
        if (Test-Path $adkEtfs) {
            if (-not (Test-Path 'boot')) { New-Item -ItemType Directory -Path 'boot' | Out-Null }
            Copy-Item -Path $adkEtfs -Destination $biosBootRel -Force -ErrorAction SilentlyContinue
        }
    }
    if (-not (Test-Path $uefiBootRel)) {
        $altUefi = "efi\microsoft\boot\efisys_noprompt.bin"
        if (Test-Path $altUefi) { $uefiBootRel = $altUefi }
        else {
            $adkEfi = Join-Path $oscdimgDir 'efisys.bin'
            $adkEfiNoPrompt = Join-Path $oscdimgDir 'efisys_noprompt.bin'
            if (Test-Path $adkEfi -or (Test-Path $adkEfiNoPrompt)) {
                if (-not (Test-Path 'efi\microsoft\boot')) { New-Item -ItemType Directory -Path 'efi\microsoft\boot' -Force | Out-Null }
                if (Test-Path $adkEfi) { Copy-Item -Path $adkEfi -Destination $uefiBootRel -Force -ErrorAction SilentlyContinue }
                if ((-not (Test-Path $uefiBootRel)) -and (Test-Path $adkEfiNoPrompt)) {
                    Copy-Item -Path $adkEfiNoPrompt -Destination $altUefi -Force -ErrorAction SilentlyContinue
                    $uefiBootRel = $altUefi
                }
            }
        }
    }
    try {
        $hasBootData = (Test-Path $biosBootRel) -and (Test-Path $uefiBootRel)
        $oscdimgParams = @("-m", "-o", "-u2")
        if ($hasBootData) {
            $data = "2#p0,e,b$biosBootRel#pEF,e,b$uefiBootRel"
            $oscdimgParams += @("-bootdata:$data", "-u2")
        }
        $oscdimgParams += @("-udfver102", "-l$Label", $SourcePath, $OutputIso)
        Write-Verbose ("oscdimg: {0}" -f $oscdimgPath)
        Write-Verbose ("args: {0}" -f ($oscdimgParams -join ' '))
        $proc = Start-Process -FilePath $oscdimgPath -ArgumentList $oscdimgParams -NoNewWindow -PassThru -Wait
        if ($proc.ExitCode -ne 0) { throw "oscdimg failed with exit code $($proc.ExitCode)" }
        if ($hasBootData) { Write-Host "ISO saved as $OutputIso (dual-boot BIOS/UEFI)" -ForegroundColor Green }
        else { Write-Host "ISO saved as $OutputIso" -ForegroundColor Green }
        return $true
    }
    catch {
        Write-Host "oscdimg execution failed: ${_}" -ForegroundColor Red
        return $false
    }
    finally {
        Pop-Location
    }
}

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

# If install.esd is present under sources, convert it to install.wim before proceeding.
$installDir = [IO.Path]::GetDirectoryName($WimPath)
$esdCandidate = Join-Path $installDir 'install.esd'

if ((-not (Test-Path $WimPath)) -and (Test-Path $esdCandidate)) {
    $convertEsd = Read-YesNo -Prompt "Detected install.esd under sources. Convert ESD to WIM now?"
    if ($convertEsd) {
        Write-Host "Converting install.esd to install.wim..." -ForegroundColor Yellow
        convert-ESDWIM -EsdPath $esdCandidate
        if (-not (Test-Path $WimPath)) {
            Write-Error "ESD to WIM conversion did not produce install.wim. Exiting."
            exit 1
        }
    }
    else {
        Write-Host "Conversion Cancelled! You can still export to ISO using the existing install.esd." -ForegroundColor Yellow
        $exportIsoNow = Read-YesNo -Prompt "Do you want to export to ISO now and skip servicing operations?"
        if ($exportIsoNow) {
            $sourceRoot = Split-Path -Path $installDir -Parent
            if (-not (Test-Path $sourceRoot)) {
                Write-Error "Could not resolve ISO source root from $installDir."
                exit 1
            }

            $defaultName = Get-DefaultIsoFileName -SourcePath $sourceRoot
            $outputIso = Join-Path (Get-Location) $defaultName
            Write-Host "Generating ISO from $sourceRoot ..." -ForegroundColor Cyan
            $isoOk = New-DualBootIso -SourcePath $sourceRoot -OutputIso $outputIso -Label $GlobalIsoLabel
            Cleanup-WimMounts
            Cleanup-ISOExtracts
            Cleanup-Mountpoints
            if (-not $isoOk) {
                Write-Error "ISO export failed."
                exit 1
            }
            Write-Host "ISO export completed." -ForegroundColor Green
            exit 0
        }

        Write-Error "install.wim is missing and conversion was declined. Cannot continue servicing."
        exit 1
    }
}

# Validate file extension
if ($WimPath -notmatch "\.wim$") {
    if ($WimPath -match "sources\\install\.esd$") {
        $targetWimPath = Join-Path ([IO.Path]::GetDirectoryName($WimPath)) 'install.wim'
        $convertEsd = Read-YesNo -Prompt "Detected install.esd under sources. Convert ESD to WIM now?"
        if ($convertEsd) {
            Write-Host "Converting install.esd to install.wim..." -ForegroundColor Yellow
            convert-ESDWIM -EsdPath $WimPath
            $WimPath = $targetWimPath
            if (-not (Test-Path $WimPath)) {
                Write-Error "ESD to WIM conversion failed to make install.wim. Exiting."
                exit 1
            }
        }
        elseif (Test-Path $targetWimPath) {
            Write-Host "Skipping conversion and continuing servicing with existing install.wim." -ForegroundColor Yellow
            $WimPath = $targetWimPath
        }
        else {
            Write-Error "No install.wim is available and conversion was declined. Cannot continue servicing."
            exit 1
        }
    }
    else {
        Write-Error "File '$WimPath' is not a WIM file."
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

function Get-ImageBuildNumber {
    param(
        [string]$WimPath,
        [int]$Index
    )
    try {
        $img = Get-WindowsImage -ImagePath $WimPath -Index $Index
        $verText = [string]$img.Version
        if ([string]::IsNullOrWhiteSpace($verText)) { return $null }
        return ([version]$verText).Build
    }
    catch {
        Write-Host "Could not detect image build for index ${Index}: ${_}" -ForegroundColor Yellow
        return $null
    }
}

function Get-SafeAppxPatterns {
    param([int]$BuildNumber)

    $win10Safe = @(
        'Microsoft.3DBuilder',
        'Microsoft.GetHelp',
        'Microsoft.Getstarted',
        'Microsoft.Microsoft3DViewer',
        'Microsoft.MicrosoftOfficeHub',
        'Microsoft.MicrosoftSolitaireCollection',
        'Microsoft.MixedReality.Portal',
        'Microsoft.OneConnect',
        'Microsoft.People',
        'Microsoft.SkypeApp',
        'Microsoft.Wallet',
        'Microsoft.WindowsAlarms',
        'Microsoft.WindowsFeedbackHub',
        'Microsoft.WindowsMaps',
        'Microsoft.Xbox*',
        'Microsoft.YourPhone',
        'Microsoft.ZuneMusic',
        'Microsoft.ZuneVideo',
        'Microsoft.BingWeather',
        'microsoft.windowscommunicationsapps',
        'Microsoft.WindowsAlarms',
        'Microsoft.Office.OneNote',
        'Microsoft.Windows.Photos'
    )

    $win11Safe = @(
        'Clipchamp.Clipchamp',
        'Microsoft.GetHelp',
        'Microsoft.Getstarted',
        'Microsoft.MicrosoftSolitaireCollection',
        'Microsoft.People',
        'Microsoft.PowerAutomateDesktop',
        'Microsoft.Todos',
        'Microsoft.WindowsAlarms',
        'Microsoft.WindowsFeedbackHub',
        'Microsoft.WindowsMaps',
        'Microsoft.Xbox*',
        'Microsoft.YourPhone',
        'Microsoft.ZuneMusic',
        'Microsoft.ZuneVideo',
        'Microsoft.Windows.DevHome',
        'Microsoft.BingNews',
        'Microsoft.BingWeather',
        'Microsoft.WindowsAlarms',
        'Microsoft.Office.OneNote',
        'MicrosoftWindows.CrossDevice',
        'MSTeams',
        'Microsoft.OutlookForWindows',
        'Microsoft.Windows.Photos',
        'Microsoft.BingSearch',
        'Microsoft.Todos',
        'Microsoft.GamingApp',
        'Microsoft.Copilot',
	'Microsoft.MicrosoftOfficeHub'
    )

    if ($BuildNumber -and $BuildNumber -ge 22000) {
        Write-Host "Detected Windows 11 image build ($BuildNumber). Using Windows 11 safe Appx list." -ForegroundColor Cyan
        return $win11Safe
    }

    if ($BuildNumber) {
        Write-Host "Detected Windows 10 image build ($BuildNumber). Using Windows 10 safe Appx list." -ForegroundColor Cyan
    }
    else {
        Write-Host "Could not detect build number. Defaulting to Windows 10 safe Appx list." -ForegroundColor Yellow
    }
    return $win10Safe
}

function Remove-SafeProvisionedAppx {
    param(
        [string]$MountPath,
        [string[]]$Patterns
    )

    $provisioned = Get-AppxProvisionedPackage -Path $MountPath
    if (-not $provisioned) {
        Write-Host "No provisioned Appx packages found at mount path: $MountPath" -ForegroundColor Yellow
        return
    }

    $targets = foreach ($pkg in $provisioned) {
        $display = [string]$pkg.DisplayName
        $pkgName = [string]$pkg.PackageName
        $matched = $false
        foreach ($pattern in $Patterns) {
            if ($display -like $pattern -or $pkgName -like "$pattern*") {
                $matched = $true
                break
            }
        }
        if ($matched) { $pkg }
    }

    if (-not $targets -or $targets.Count -eq 0) {
        Write-Host "No packages matched the safe Appx removal list." -ForegroundColor Yellow
        return
    }

    Write-Host "Removing $($targets.Count) provisioned Appx package(s)..." -ForegroundColor Cyan
    foreach ($pkg in $targets) {
        try {
            Remove-AppxProvisionedPackage -Path $MountPath -PackageName $pkg.PackageName -ErrorAction Stop | Out-Null
            Write-Host "Removed: $($pkg.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to remove $($pkg.DisplayName): ${_}" -ForegroundColor Red
            $script:errorsFound = $true
        }
    }

    # Run /Optimize-ProvisionedAppxPackages on the image after removal.
    try {
        Write-Host "Optimizing provisioned Appx package state..." -ForegroundColor Cyan
        Optimize-AppxProvisionedPackages -Path $MountPath -ErrorAction Stop | Out-Null
        Write-Host "Provisioned Appx package optimization completed." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to run Optimize-AppxProvisionedPackages: ${_}" -ForegroundColor Red
        $script:errorsFound = $true
    }

}

# Function to discard a mounted WIM image with error handling
function Discard-Image {
    param([string]$MountPath)
    try {
        Dismount-WindowsImage -Path $MountPath -Discard
        Write-Host "Discarded mounted image at $MountPath" -ForegroundColor DarkGray
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'Access to the path.*is denied') {
            Write-Host "Access denied while discarding $MountPath. Retrying with force..." -ForegroundColor Yellow
            try {
                Dismount-WindowsImage -Path $MountPath -Discard -Force
                Write-Host "Force-discarded mounted image at $MountPath" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Could not force-discard mounted image at ${MountPath}: ${_}" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Could not discard mounted image at ${MountPath}: ${_}" -ForegroundColor Red
        }
    }
}

# Function to optimize and export WIM image to ESD using dism.exe
function Optimize-ESD {
    param(
        [string]$WimPath,
        [int[]]$Indexes = @()
    )
    Write-Host "Starting WIM to ESD export using dism.exe..." -ForegroundColor Cyan
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
    $esdPath = [IO.Path]::GetDirectoryName($WimPath) + '\install.esd'
    $allImages = Get-WindowsImage -ImagePath $WimPath
    if ($Indexes -and $Indexes.Count -gt 0) {
        $exportIndexes = $Indexes
    }
    else {
        $exportIndexes = $allImages | ForEach-Object { $_.ImageIndex }
    }
    try {
        foreach ($i in $exportIndexes) {
            $dismArgs = @(
                "/Export-Image",
                "/SourceImageFile:$WimPath",
                "/SourceIndex:$i",
                "/DestinationImageFile:$esdPath",
                "/Compress:recovery",
                "/CheckIntegrity"
            )
            Write-Host "Running: dism.exe $($dismArgs -join ' ')" -ForegroundColor Cyan
            $proc = Start-Process -FilePath dism.exe -ArgumentList $dismArgs -NoNewWindow -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                Write-Host "dism.exe export failed for index $i with exit code $($proc.ExitCode)" -ForegroundColor Red
                return
            }
        }
        if (Test-Path $esdPath) {
            Write-Host "Optimized ESD has been created: $esdPath" -ForegroundColor Green
            # Delete original install.wim after successful ESD conversion
            try {
                Remove-Item -Path $WimPath -Force
                Write-Host "Deleted original WIM: $WimPath" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Could not delete original WIM: $WimPath. ${_}" -ForegroundColor Red
            }
        }
        else {
            Write-Host "ESD export failed: install.esd not found." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "dism.exe export failed or was interrupted." -ForegroundColor Red
    }
}

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



$indexInput = [string](Read-Host "Enter the index number(s) of the edition(s) to modify (e.g. 1,3,5 or * for all editions)")
$indexInput = $indexInput.Trim()
if ([string]::IsNullOrWhiteSpace($indexInput)) {
    Write-Host "No edition index input provided. Exiting." -ForegroundColor Red
    exit 1
}
# Parse selected indexes for use in all operations
$selectedIndexes = @()
if ($indexInput -eq '*') {
    $selectedIndexes = $editions | ForEach-Object { $_.Index }
}
else {
    $inputIndexes = $indexInput -split ',' | ForEach-Object { $_.Trim() }
    $validIndexes = $editions | ForEach-Object { $_.Index }
    $selectedIndexes = $inputIndexes | Where-Object { $validIndexes -contains $_ }
    if ($selectedIndexes.Count -eq 0) {
        Write-Host "No valid edition indexes selected. Exiting." -ForegroundColor Red
        exit 1
    }
}
$selectedIndexes = $selectedIndexes | ForEach-Object { [int]$_ }

Write-Host "Select operation:" -ForegroundColor Cyan
Write-Host "0: Cancel operation"
Write-Host "1: Remove All Edge Components"
Write-Host "2: Remove Edge Browser"
Write-Host "3: Remove Edge WebView"
Write-Host "4: Optimize WIM image for export (credits: abbodi1406)"
Write-Host "5: Generate ISO (dual-boot BIOS/UEFI)"
Write-Host "6: Optimize and Export Image to ESD (dism.exe)"
Write-Host "7: Remove Safe Appx Provisioned Packages (Win10/Win11 auto-detect)"
Write-Host "8: Optimize boot.wim image for export"

$choice = [string](Read-Host "Enter your choice (0/1/2/3/4/5/6/7/8)")
$choice = $choice.Trim()
if ($choice -notin @('0', '1', '2', '3', '4', '5', '6', '7', '8')) {
    Write-Host "Invalid choice '$choice'. Exiting." -ForegroundColor Red
    exit 1
}
if ($choice -eq '4') {
    Optimize-WimImage -WimPath $WimPath -Indexes $selectedIndexes
    Export-UpdatedIsoIfRequested -IsoWasExtracted $isoExtracted -TempExtractPath $tempExtractPath -IsoLabel $GlobalIsoLabel
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
    Pause-ForExit
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
    Pause-ForExit
    exit 0
}
if ($choice -eq '6') {
    # Pass selected indexes to Optimize-ESD
    $indexesToExport = $selectedIndexes | ForEach-Object { [int]$_ }
    Optimize-ESD -WimPath $WimPath -Indexes $indexesToExport
    Export-UpdatedIsoIfRequested -IsoWasExtracted $isoExtracted -TempExtractPath $tempExtractPath -IsoLabel $GlobalIsoLabel
    Cleanup-WimMounts
    Cleanup-ISOExtracts
    Cleanup-Mountpoints
    $scriptEndTime = Get-Date
    $elapsed = $scriptEndTime - $scriptStartTime
    if ($elapsed.TotalMinutes -ge 1) {
        $elapsedMsg = "Time elapsed for ESD export: {0:N2} minutes" -f $elapsed.TotalMinutes
    }
    else {
        $elapsedMsg = "Time elapsed for ESD export: {0:N2} seconds" -f $elapsed.TotalSeconds
    }
    Write-Host $elapsedMsg -ForegroundColor Cyan
    Pause-ForExit
    exit 0
}

if ($choice -eq '8') {
    $bootWimPath = Join-Path (Split-Path $WimPath -Parent) 'boot.wim'
    if (!(Test-Path $bootWimPath)) {
        Write-Host "boot.wim was not found at expected path: $bootWimPath" -ForegroundColor Red
        Cleanup-WimMounts
        Cleanup-ISOExtracts
        Cleanup-Mountpoints
        Pause-ForExit
        exit 1
    }

    Optimize-BootWimImage -BootWimPath $bootWimPath
    Export-UpdatedIsoIfRequested -IsoWasExtracted $isoExtracted -TempExtractPath $tempExtractPath -IsoLabel $GlobalIsoLabel
    Cleanup-WimMounts
    Cleanup-ISOExtracts
    Cleanup-Mountpoints
    $scriptEndTime = Get-Date
    $elapsed = $scriptEndTime - $scriptStartTime
    if ($elapsed.TotalMinutes -ge 1) {
        $elapsedMsg = "Time elapsed for boot.wim optimization: {0:N2} minutes" -f $elapsed.TotalMinutes
    }
    else {
        $elapsedMsg = "Time elapsed for boot.wim optimization: {0:N2} seconds" -f $elapsed.TotalSeconds
    }
    Write-Host $elapsedMsg -ForegroundColor Cyan
    Pause-ForExit
    exit 0
}

# ISO generation operation
if ($choice -eq '5') {
    # Determine source path for ISO creation
    $isoSourcePath = $null
    if ($isoExtracted -and (Test-Path $tempExtractPath)) {
        $isoSourcePath = $tempExtractPath
    }
    elseif ($item -and $item.PSIsContainer) {
        $isoSourcePath = $cleanPath
    }
    elseif ($item -and -not $item.PSIsContainer -and ($cleanPath -match "sources\\install\.(wim|esd)$")) {
        # If input is sources\install.wim or sources\install.esd, use the media root folder.
        $isoSourcePath = Split-Path -Path (Split-Path -Path $cleanPath -Parent) -Parent
    }

    $defaultName = Get-DefaultIsoFileName -SourcePath $isoSourcePath
    $outputIso = Join-Path (Get-Location) $defaultName

    if ($isoSourcePath) {
        Write-Host "Generating ISO from $isoSourcePath ..." -ForegroundColor Cyan
        try {
            New-DualBootIso -SourcePath $isoSourcePath -OutputIso $outputIso -Label $GlobalIsoLabel
        }
        catch {
            Write-Host "ISO generation failed: ${_}" -ForegroundColor Red
        }
    }
    else {
        Write-Host "No valid ISO source folder available. Provide an ISO or an extracted folder path." -ForegroundColor Yellow
    }
    Cleanup-WimMounts
    Cleanup-ISOExtracts
    Cleanup-Mountpoints
    $scriptEndTime = Get-Date
    $elapsed = $scriptEndTime - $scriptStartTime
    if ($elapsed.TotalMinutes -ge 1) { $elapsedMsg = "Time elapsed: {0:N2} minutes" -f $elapsed.TotalMinutes } else { $elapsedMsg = "Time elapsed: {0:N2} seconds" -f $elapsed.TotalSeconds }
    Write-Host $elapsedMsg -ForegroundColor Cyan
    exit 0
}

function Process-Edition {
    param([int]$idx)
    $mountPath = "$env:TEMP\WimMount_${idx}"
    if (!(Test-Path $mountPath)) { New-Item -ItemType Directory -Path $mountPath | Out-Null }
    try {
        Mount-Wim -WimPath $WimPath -Index $idx -MountPath $mountPath
    }
    catch {
        Write-Host "Failed to mount edition index ${idx}: ${_}" -ForegroundColor Red
        $script:errorsFound = $true
        return
    }
    try {
        switch ($choice) {
            '0' { Write-Host "Operation cancelled by user." -ForegroundColor Yellow; exit 0 }
            '1' { Run-DismRemove -MountPath $mountPath -Option "/Remove-Edge" }
            '2' { Run-DismRemove -MountPath $mountPath -Option "/Remove-EdgeBrowser" }
            '3' { Run-DismRemove -MountPath $mountPath -Option "/Remove-EdgeWebView" }
            '7' {
                $buildNumber = Get-ImageBuildNumber -WimPath $WimPath -Index $idx
                $safePatterns = Get-SafeAppxPatterns -BuildNumber $buildNumber
                Remove-SafeProvisionedAppx -MountPath $mountPath -Patterns $safePatterns
            }
            default { Write-Host "Invalid choice."; exit 1 }
        }
        Commit-Wim -MountPath $mountPath
        Write-Host "Changes committed to ${WimPath} for edition index ${idx}." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to commit edition index ${idx}: ${_}" -ForegroundColor Red
    }
}

# process every editions/index in multi edition WIM/ISO file

foreach ($idx in $selectedIndexes) {
    Process-Edition -idx $idx
}

# After all editions is/are processed, optimize the WIM image
Optimize-WimImage -WimPath $WimPath -Indexes $selectedIndexes

# If input was ISO, save updated ISO before cleanup via reusable function
if ($isoExtracted -and (Test-Path $tempExtractPath)) {
    $defaultName = Get-DefaultIsoFileName -SourcePath $tempExtractPath
    $outputIso = Join-Path (Get-Location) $defaultName
    Write-Host "Saving updated ISO as $outputIso..." -ForegroundColor Cyan
    New-DualBootIso -SourcePath $tempExtractPath -OutputIso $outputIso -Label $GlobalIsoLabel
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

# Define Pause-ForExit at top-level so it is always available
if (-not (Get-Command Pause-ForExit -ErrorAction SilentlyContinue)) {
    function Pause-ForExit {
        do {
            $resp = Read-Host "Press E or 0 to exit"
        } while ($resp -notmatch '^(?i:e|0)$')
    }
}

if ($script:errorsFound) {
    Write-Host "Process completed with errors. Check above for details." -ForegroundColor Red
    Write-Host $elapsedMsg -ForegroundColor Yellow
    Pause-ForExit
}
else {
    Write-Host "Process completed successfully!" -ForegroundColor Green
    Write-Host $elapsedMsg -ForegroundColor Cyan
    Pause-ForExit
}
