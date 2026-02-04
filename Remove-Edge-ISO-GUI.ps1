# Hybrid GUI/CMD based Edge Removal Script for Windows 11 ISO (PowerShell 7+)
# Disclaimer: Use at your own risk. Always back up your data before making system changes. Please be advised if you use Edge Browser and WebView components,
# the script is not intended for such use cases as removing them does not make sense

param(
    [Parameter(Mandatory = $false)]
    [string]$IsoOrWimPath,
    [switch]$Gui,
    [switch]$NoElevate
)

# Import only required modules
Import-Module -Name Dism -ErrorAction Stop
Import-Module -Name Storage -ErrorAction Stop
Import-Module -Name CimCmdlets -ErrorAction Stop

# Check for Admin Privileges and auto-elevate early
function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if ($isAdmin) { return }
    if ($NoElevate) {
        Write-Warning "Not running as Administrator. Some operations may fail."
        return
    }
    try {
        # Determine current PowerShell executable and script path
        $psExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
        # Rebuild arguments for elevated relaunch
        $elevArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath)
        if ($Gui) { $elevArgs += '-Gui' }
        if ($IsoOrWimPath) { $elevArgs += @('-IsoOrWimPath', $IsoOrWimPath) }
        if ($NoElevate) { $elevArgs += '-NoElevate' }
        Start-Process -FilePath $psExe -ArgumentList $elevArgs -Verb RunAs | Out-Null
        exit 0
    }
    catch {
        Write-Error "Failed to request elevation: ${_}"
        exit 1
    }
}

# Start timer
$scriptStartTime = Get-Date
$outputIso = $null

# Global ISO label
$GlobalIsoLabel = "Custom_Win11"

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
        1..$count | ForEach-Object {
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

# Minimal WinForms GUI to collect inputs
function Show-MinimalGui {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        Add-Type -AssemblyName System.Drawing | Out-Null
    }
    catch {
        Write-Error "Failed to load WinForms assemblies. GUI mode requires desktop support."
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Remove Edge from ISO/WIM"
    $form.Size = New-Object System.Drawing.Size(600, 360)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true

    # ISO/WIM path
    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Windows Sources Path:"
    $lblPath.Location = New-Object System.Drawing.Point(12, 18)
    $lblPath.AutoSize = $true

    $tbPath = New-Object System.Windows.Forms.TextBox
    $tbPath.Location = New-Object System.Drawing.Point(120, 15)
    $tbPath.Size = New-Object System.Drawing.Size(300, 22)
    $tbPath.AllowDrop = $true
    $tbPath.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' } })
    $tbPath.Add_DragDrop({
            $items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($items -and $items.Length -gt 0) { $tbPath.Text = $items[0]; & $populateEditions }
        })
    if ($IsoOrWimPath) { $tbPath.Text = $IsoOrWimPath }
    # Populate editions when user finishes editing the path
    $tbPath.Add_Leave({ & $populateEditions })

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse File...'
    $btnBrowse.Location = New-Object System.Drawing.Point(430, 14)
    $btnBrowse.Add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = 'Select ISO or WIM'
            $dlg.Filter = 'ISO/WIM (*.iso;*.wim)|*.iso;*.wim|All Files (*.*)|*.*'
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $tbPath.Text = $dlg.FileName
                & $populateEditions
            }
        })

    $btnBrowseFolder = New-Object System.Windows.Forms.Button
    $btnBrowseFolder.Text = 'Folder...'
    $btnBrowseFolder.Location = New-Object System.Drawing.Point(520, 14)
    $btnBrowseFolder.Add_Click({
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            $fbd.Description = 'Select extracted ISO folder (contains sources\\install.wim)'
            $fbd.ShowNewFolderButton = $false
            if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $tbPath.Text = $fbd.SelectedPath
                & $populateEditions
            }
        })

    # Operation selection
    $lblOp = New-Object System.Windows.Forms.Label
    $lblOp.Text = "Operation:"
    $lblOp.Location = New-Object System.Drawing.Point(12, 58)
    $lblOp.AutoSize = $true

    $cbOp = New-Object System.Windows.Forms.ComboBox
    $cbOp.DropDownStyle = 'DropDownList'
    $cbOp.Location = New-Object System.Drawing.Point(120, 55)
    $cbOp.Size = New-Object System.Drawing.Size(320, 24)
    [void]$cbOp.Items.AddRange(@(
            '1 - Remove All Edge Components',
            '2 - Remove Edge Browser',
            '3 - Remove Edge WebView',
            '4 - Optimize WIM image',
            '5 - Generate ISO (dual-boot)',
            '6 - Optimize and Export Image to ESD (dism.exe)'
        ))
    $cbOp.SelectedIndex = 0

    # Edition indexes
    $lblIdx = New-Object System.Windows.Forms.Label
    $lblIdx.Text = "Edition index(es):"
    $lblIdx.Location = New-Object System.Drawing.Point(12, 98)
    $lblIdx.AutoSize = $true

    # Checked list for editions, includes "* - All editions"
    $clbIdx = New-Object System.Windows.Forms.CheckedListBox
    $clbIdx.CheckOnClick = $true
    $clbIdx.Location = New-Object System.Drawing.Point(120, 95)
    $clbIdx.Size = New-Object System.Drawing.Size(320, 70)
    [void]$clbIdx.Items.Add('* - All editions')
    $clbIdx.SetItemChecked(0, $true)
    $script:__idxControlRef = $clbIdx

    # Helper to resolve WIM path from input and populate editions
    $populateEditions = {
        try {
            $clbIdx.BeginUpdate()
            $clbIdx.Items.Clear() | Out-Null
            [void]$clbIdx.Items.Add('* - All editions')
            $pathText = ($tbPath.Text).Trim()
            if ([string]::IsNullOrWhiteSpace($pathText)) { $clbIdx.SetItemChecked(0, $true); return }

            $resolvedWim = $null
            $mountedIsoPath = $null
            if (Test-Path $pathText) {
                $it = Get-Item $pathText -ErrorAction SilentlyContinue
                if ($it -and $it.PSIsContainer) {
                    $candidate = Join-Path $pathText 'sources\install.wim'
                    if (Test-Path $candidate) { $resolvedWim = $candidate }
                }
                else {
                    if ($pathText -match "\.wim$") { $resolvedWim = $pathText }
                    elseif ($pathText -match "\.iso$") {
                        try {
                            $mnt = Mount-DiskImage -ImagePath $pathText -PassThru -ErrorAction Stop
                            $dl = ($mnt | Get-Volume -ErrorAction Stop).DriveLetter
                            if ($dl) {
                                $mountedIsoPath = "$dl`:\\sources\\install.wim"
                                if (Test-Path $mountedIsoPath) { $resolvedWim = $mountedIsoPath }
                            }
                        }
                        catch { }
                        finally {
                            if ($mnt) { Dismount-DiskImage -ImagePath $pathText -ErrorAction SilentlyContinue | Out-Null }
                        }
                    }
                }
            }

            if ($resolvedWim -and (Test-Path $resolvedWim)) {
                try {
                    $imgs = Get-WindowsImage -ImagePath $resolvedWim -ErrorAction Stop
                    foreach ($img in $imgs) {
                        [void]$clbIdx.Items.Add(("{0} - {1}" -f $img.ImageIndex, $img.ImageName))
                    }
                }
                catch { }
            }
            if ($clbIdx.Items.Count -gt 0) { $clbIdx.SetItemChecked(0, $true) }
        }
        finally {
            $clbIdx.EndUpdate()
        }
    }

    # ISO Label (optional)
    $lblIsoLabel = New-Object System.Windows.Forms.Label
    $lblIsoLabel.Text = "ISO Label (optional):"
    $lblIsoLabel.Location = New-Object System.Drawing.Point(12, 138)
    $lblIsoLabel.AutoSize = $true

    $tbIsoLabel = New-Object System.Windows.Forms.TextBox
    $tbIsoLabel.Location = New-Object System.Drawing.Point(160, 135)
    $tbIsoLabel.Size = New-Object System.Drawing.Size(280, 22)
    $tbIsoLabel.Text = $GlobalIsoLabel

    # Output ISO name/path (optional)
    $lblOutIso = New-Object System.Windows.Forms.Label
    $lblOutIso.Text = "Output ISO name/path (optional):"
    $lblOutIso.Location = New-Object System.Drawing.Point(12, 168)
    $lblOutIso.AutoSize = $true

    $tbOutIso = New-Object System.Windows.Forms.TextBox
    $tbOutIso.Location = New-Object System.Drawing.Point(200, 165)
    $tbOutIso.Size = New-Object System.Drawing.Size(240, 22)
    $dateCode = (Get-Date).ToString('MMMyy')
    $tbOutIso.Text = "Win11_{0}.iso" -f $dateCode

    $btnSaveAs = New-Object System.Windows.Forms.Button
    $btnSaveAs.Text = 'Save As...'
    $btnSaveAs.Location = New-Object System.Drawing.Point(450, 164)
    $btnSaveAs.Add_Click({
            $sdlg = New-Object System.Windows.Forms.SaveFileDialog
            $sdlg.Title = 'Choose output ISO name'
            $sdlg.Filter = 'ISO Image (*.iso)|*.iso|All Files (*.*)|*.*'
            $sdlg.FileName = $tbOutIso.Text
            if ($sdlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $tbOutIso.Text = $sdlg.FileName
            }
        })

    # Show ISO fields only for operation 5 (Generate ISO), and adjust layout dynamically
    $setIsoFieldsVisibility = {
        param($visible)
        $idxBottomY = $null
        if ($script:__idxControlRef) { $idxBottomY = $script:__idxControlRef.Location.Y + $script:__idxControlRef.Height } else { $idxBottomY = 120 }
        $baseY = $idxBottomY + 15

        $lblIsoLabel.Visible = $visible
        $tbIsoLabel.Visible = $visible
        $lblOutIso.Visible = $visible
        $tbOutIso.Visible = $visible
        $btnSaveAs.Visible = $visible

        if ($visible) {
            $lblIsoLabel.Location = New-Object System.Drawing.Point(12, $baseY)
            $tbIsoLabel.Location = New-Object System.Drawing.Point(160, $baseY)
            $y2 = $baseY + 30
            $lblOutIso.Location = New-Object System.Drawing.Point(12, $y2)
            $tbOutIso.Location = New-Object System.Drawing.Point(200, $y2)
            $btnSaveAs.Location = New-Object System.Drawing.Point(450, ($y2 - 1))
            $y3 = $y2 + 30
            if ($chkElevate) { $chkElevate.Location = New-Object System.Drawing.Point(120, $y3) }
            $y4 = $y3 + 40
            if ($btnOk) { $btnOk.Location = New-Object System.Drawing.Point(280, $y4) }
            if ($btnCancel) { $btnCancel.Location = New-Object System.Drawing.Point(370, $y4) }
            $form.Height = $y4 + 120
        }
        else {
            if ($chkElevate) { $chkElevate.Location = New-Object System.Drawing.Point(120, $baseY) }
            $yBtn = $baseY + 40
            if ($btnOk) { $btnOk.Location = New-Object System.Drawing.Point(280, $yBtn) }
            if ($btnCancel) { $btnCancel.Location = New-Object System.Drawing.Point(370, $yBtn) }
            $form.Height = $yBtn + 120
        }
    }

    $isIsoGen = ($cbOp.SelectedItem -like '5*')
    & $setIsoFieldsVisibility $isIsoGen

    $cbOp.Add_SelectedIndexChanged({
            $isIsoGenLocal = ($cbOp.SelectedItem -like '5*')
            & $setIsoFieldsVisibility $isIsoGenLocal
        })

    # Elevation option
    $chkElevate = New-Object System.Windows.Forms.CheckBox
    $chkElevate.Text = 'Run as Administrator (UAC)'
    $chkElevate.Location = New-Object System.Drawing.Point(120, 200)
    $chkElevate.AutoSize = $true
    $chkElevate.Checked = $true

    # Buttons
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Start'
    $btnOk.Location = New-Object System.Drawing.Point(280, 240)
    $btnOk.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($tbPath.Text)) {
                $script:IsoOrWimPath = $tbPath.Text.Trim()
                $sel = $cbOp.SelectedItem
                if ($sel) {
                    $script:ChoiceFromGui = ($sel.Split(' ')[0])
                }
                $checkedItems = @()
                for ($i = 0; $i -lt $clbIdx.Items.Count; $i++) { if ($clbIdx.GetItemChecked($i)) { $checkedItems += $clbIdx.Items[$i] } }
                if (($checkedItems | ForEach-Object { $_.ToString() }) -contains '* - All editions') {
                    $script:IndexInputFromGui = '*'
                }
                else {
                    $indices = @()
                    foreach ($it in $checkedItems) { $s = $it.ToString(); if ($s -match '^\s*(\d+)') { $indices += $matches[1] } }
                    if ($indices.Count -gt 0) { $script:IndexInputFromGui = ($indices -join ',') } else { $script:IndexInputFromGui = '*' }
                }
                if (-not [string]::IsNullOrWhiteSpace($tbIsoLabel.Text)) {
                    $script:IsoLabelFromGui = $tbIsoLabel.Text.Trim()
                }
                if (-not [string]::IsNullOrWhiteSpace($tbOutIso.Text)) {
                    $script:OutputIsoFromGui = $tbOutIso.Text.Trim()
                }
                $script:NoElevateFromGui = -not $chkElevate.Checked
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.Close()
            }
            else {
                [System.Windows.Forms.MessageBox]::Show('Please select an ISO or WIM path.', 'Input required', 'OK', 'Warning') | Out-Null
            }
        })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(370, 240)
    $btnCancel.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        })

    $form.Controls.AddRange(@($lblPath, $tbPath, $btnBrowse, $btnBrowseFolder, $lblOp, $cbOp, $lblIdx, $clbIdx, $lblIsoLabel, $tbIsoLabel, $lblOutIso, $tbOutIso, $btnSaveAs, $chkElevate, $btnOk, $btnCancel))
    # Apply layout once controls exist
    & $setIsoFieldsVisibility ($cbOp.SelectedItem -like '5*')
    # Populate editions now if a path was provided via CLI
    & $populateEditions
    [void]$form.ShowDialog()
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

# Elevate early to avoid re-prompting in GUI after restart
Ensure-Admin

# If GUI mode requested or path not supplied, show minimal UI to gather inputs
if ($Gui -or [string]::IsNullOrWhiteSpace($IsoOrWimPath)) {
    Show-MinimalGui | Out-Null
    if (-not $script:IsoOrWimPath) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
    $IsoOrWimPath = $script:IsoOrWimPath
    if ($script:ChoiceFromGui) { $choice = $script:ChoiceFromGui }
    if ($script:IndexInputFromGui) { $indexInput = $script:IndexInputFromGui }
    if ($script:IsoLabelFromGui) { $GlobalIsoLabel = $script:IsoLabelFromGui }
    if ($null -ne $script:NoElevateFromGui) { $NoElevate = $script:NoElevateFromGui }
    if ($script:OutputIsoFromGui) { $outputIso = $script:OutputIsoFromGui }
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

# (moved) Ensure-Admin defined earlier

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
    ) | Select-Object -Unique
    $oscdimgPath = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $oscdimgPath) {
        $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
        if ($cmd) { $oscdimgPath = $cmd.Source }
    }
    if (-not $oscdimgPath) {
        Write-Host "oscdimg.exe not found. Install Windows ADK or add it to PATH." -ForegroundColor Yellow
        return $false
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

# Function to optimize and export WIM image to ESD using dism.exe
function Optimize-ESD {
    param([string]$WimPath)
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
    $count = (Get-WindowsImage -ImagePath $WimPath).Count
    try {
        for ($i = 1; $i -le $count; $i++) {
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
            Write-Host "ESD expo14rt failed: install.esd not found." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "dism.exe export failed or was interrupted." -ForegroundColor Red
    }
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



if (-not $indexInput) {
    $indexInput = Read-Host "Enter the index number(s) of the edition(s) to modify (e.g. 1,3,5 or * for all editions)"
}

if (-not $choice -and -not $Gui) {
    Write-Host "Select operation:" -ForegroundColor Cyan
    Write-Host "0: Cancel operation"
    Write-Host "1: Remove All Edge Components"
    Write-Host "2: Remove Edge Browser"
    Write-Host "3: Remove Edge WebView"
    Write-Host "4: Optimize WIM image for export (credits: abbodi1406)"
    Write-Host "5: Generate ISO (dual-boot BIOS/UEFI)"
    Write-Host "6: Optimize and Export Image to ESD (dism.exe)"
    $choice = Read-Host "Enter your choice (0/1/2/3/4/5/6)"
}
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
if ($choice -eq '6') {
    Optimize-ESD -WimPath $WimPath
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
    exit 0
}

# ISO generation operation
if ($choice -eq '5') {
    $dateCode = (Get-Date).ToString('MMMyy')
    $defaultName = "Win11_{0}.iso" -f $dateCode
    if (-not $outputIso) {
        $resp = Read-Host "Enter output ISO file name or full path (default: $defaultName)"
        if ([string]::IsNullOrWhiteSpace($resp)) { $resp = $defaultName }
        if (-not ($resp.ToLower().EndsWith('.iso'))) { $resp = "$resp.iso" }
        if ([IO.Path]::IsPathRooted($resp)) { $outputIso = $resp } else { $outputIso = (Join-Path (Get-Location) $resp) }
    }
    else {
        if (-not ($outputIso.ToLower().EndsWith('.iso'))) { $outputIso = "$outputIso.iso" }
        if (-not [IO.Path]::IsPathRooted($outputIso)) { $outputIso = (Join-Path (Get-Location) $outputIso) }
    }
    # Ensure output directory exists
    $outDir = Split-Path -Path $outputIso -Parent
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    # Determine source path for ISO creation
    $isoSourcePath = $null
    if ($isoExtracted -and (Test-Path $tempExtractPath)) {
        $isoSourcePath = $tempExtractPath
    }
    elseif ($item -and $item.PSIsContainer) {
        $isoSourcePath = $cleanPath
    }
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
    1..$count | ForEach-Object {
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

# If input was ISO, save updated ISO before cleanup via reusable function
if ($isoExtracted -and (Test-Path $tempExtractPath)) {
    if (-not $outputIso) {
        $dateCode = (Get-Date).ToString('MMMyy')
        $outputIso = Join-Path (Get-Location) ("Win11_{0}.iso" -f $dateCode)
    }
    else {
        if (-not ($outputIso.ToLower().EndsWith('.iso'))) { $outputIso = "$outputIso.iso" }
        if (-not [IO.Path]::IsPathRooted($outputIso)) { $outputIso = (Join-Path (Get-Location) $outputIso) }
    }
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

if ($errorsFound) {
    Write-Host "Process completed with errors. Check above for details." -ForegroundColor Red
    Write-Host $elapsedMsg -ForegroundColor Yellow
}
else {
    Write-Host "Process completed successfully!" -ForegroundColor Green
    Write-Host $elapsedMsg -ForegroundColor Cyan
}
