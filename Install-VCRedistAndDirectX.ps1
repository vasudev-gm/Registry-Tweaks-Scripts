param (
    [switch]$Verbose,
    [switch]$OfflineMode
)

# Check Internet connectivity with a 10-second timeout
Write-Output "Checking for Internet Connectivity by pinging Google DNS 8.8.8.8 (10 sec timeout)"
$pp = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'

# Use Start-Job to implement timeout for connectivity check
$internetAvailable = $false
try {
    $job = Start-Job -ScriptBlock { 
        param($ProgressPref) 
        $ProgressPreference = $ProgressPref
        test-netconnection -ComputerName 8.8.8.8 -InformationLevel Quiet
    } -ArgumentList $ProgressPreference
    
    # Wait up to 10 seconds for the job to complete
    if (Wait-Job -Job $job -Timeout 10) {
        $internetAvailable = Receive-Job -Job $job
        Write-Output "Internet Connectivity Check Completed!"
    } else {
        Write-Output "Internet Connectivity Check timed out after 10 seconds."
        $internetAvailable = $false
    }
    
    # Clean up the job
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
} catch {
    Write-Output "Internet Connectivity Check failed with error: $_"
    $internetAvailable = $false
}

$ProgressPreference = $pp

<#
.SYNOPSIS
This PowerShell script ensures it is run with administrator privileges, then installs the necessary Visual C++ Redistributables and DirectX 9 End-User Runtime silently. It provides progress updates throughout the installation process and handles errors gracefully.

.DESCRIPTION
This script performs the following tasks:
- Checks if it is run with administrator privileges and relaunches with elevated permissions if necessary.
- Checks for offline packages in the WinApps directory at the root of the ISO or in the script directory before attempting to download.
- Downloads and installs the latest Visual C++ Redistributables from a GitHub repository if offline packages don't exist.
- Downloads and installs the DirectX 9 End-User Runtime from Microsoft's official website if offline packages don't exist.
- Provides verbose logging for detailed output, including the URLs being accessed, the files being downloaded, and the installation progress.
- Handles errors gracefully by catching exceptions and providing meaningful error messages.
- Waits for user input before exiting if an error occurs or at the end of the script, ensuring that the user is aware of the script's completion or any issues that occurred.

.PARAMETER Verbose
Enables verbose logging for detailed output. When this switch is used, the script will provide additional information about its progress and actions.

.PARAMETER OfflineMode
Forces the script to use only offline packages and skip internet connectivity checks.

.EXAMPLE
.\Install-VCRedistAndDirectX.ps1 -Verbose
Runs the script with verbose logging enabled, providing detailed output about the script's progress and actions.

.EXAMPLE
.\Install-VCRedistAndDirectX.ps1 -OfflineMode
Runs the script in offline mode, using only locally available packages without checking for internet connectivity.

.EXAMPLE
# Running the script from the Internet
irm https://gist.githubusercontent.com/emilwojcik93/ef790a6b12c8e9358bbc52ed76fb495c/raw/Install-VCRedistAndDirectX.ps1 | iex
Downloads and runs the script directly from the provided URL, ensuring that the latest version is used.

.LINK
https://gist.github.com/emilwojcik93/ef790a6b12c8e9358bbc52ed76fb495c
#>


if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output "This script needs to be run as Administrator. Attempting to relaunch."
    $argList = @()

    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        $argList += if ($_.Value -is [switch] -and $_.Value) {
            "-$($_.Key)"
        } elseif ($_.Value) {
            "-$($_.Key) `"$($_.Value)`""
        }
    }

    
    $script = if ($PSCommandPath) {
        "& { & `"$($PSCommandPath)`" ${argList} }"
    } else {
        "&([ScriptBlock]::Create((irm https://gist.githubusercontent.com/emilwojcik93/ef790a6b12c8e9358bbc52ed76fb495c/raw/Install-VCRedistAndDirectX.ps1))) ${argList}"
    }

    $powershellcmd = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }
    $processCmd = if (Get-Command wt.exe -ErrorAction SilentlyContinue) { "wt.exe" } else { $powershellcmd }

    Start-Process $processCmd -ArgumentList "$powershellcmd -ExecutionPolicy Bypass -NoProfile -Command $script" -Verb RunAs

    break
}

# Function to check if the parent process is explorer.exe
function Is-StartedByExplorer {
    $parentProcess = Get-WmiObject Win32_Process -Filter "ProcessId=$((Get-WmiObject Win32_Process -Filter "ProcessId=$PID").ParentProcessId)"
    return $parentProcess.Name -eq "explorer.exe"
}

# Function to check if the session is likely to close at the end of the script
function Is-SessionLikelyToClose {
    $commandLineArgs = [Environment]::GetCommandLineArgs()
    return ($commandLineArgs -contains "-NoProfile")
}

# Function to check if offline packages exist
function Test-OfflinePackages {
    param (
        [string]$PackageName
    )

    # First, check in the WinApps directory at the root of the ISO
    # Try to find the ISO drive letter
    $isoDrives = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.OperationalStatus -eq 'OK' } | Select-Object -ExpandProperty DriveLetter
    $packagePath = $null
    
    # Check each potential ISO drive for the WinApps directory
    foreach ($driveLetter in $isoDrives) {
        $winAppsPath = "${driveLetter}:\WinApps"
        Write-Verbose "Checking for WinApps directory in drive ${driveLetter}: $winAppsPath"
        
        if (Test-Path $winAppsPath) {
            Write-Verbose "Found WinApps directory in drive ${driveLetter}"
            
            switch ($PackageName) {
                "VCRedist" {
                    $vcRedistFile = Join-Path $winAppsPath "VisualCppRedist_AIO_x86_x64.exe"
                    if (Test-Path $vcRedistFile) {
                        Write-Verbose "Found VCRedist package in ISO WinApps directory"
                        return $true, $vcRedistFile
                    }
                }
                "DirectX" {
                    $directXFile = Join-Path $winAppsPath "DirectX_Redist_Repack_x86_x64.exe"
                    if (Test-Path $directXFile) {
                        Write-Verbose "Found DirectX package in ISO WinApps directory"
                        return $true, $directXFile
                    }
                }
            }
        }
    }
    
    # If not found in ISO, fallback to script directory
    $scriptDirectory = if ($PSCommandPath) {
        Split-Path -Parent $PSCommandPath
    } else {
        $PWD.Path
    }
    
    $offlineDirectory = Join-Path $scriptDirectory "offline-packages"
    
    switch ($PackageName) {
        "VCRedist" {
            $vcRedistFile = Join-Path $offlineDirectory "VisualCppRedist_AIO_x86_x64.exe"
            return (Test-Path $vcRedistFile), $vcRedistFile
        }
        "DirectX" {
            $directXFile = Join-Path $offlineDirectory "DirectX_Redist_Repack_x86_x64.exe"
            return (Test-Path $directXFile), $directXFile
        }
    }
    
    return $false, $null
}

function Get-LatestReleaseUrl {
    param (
        [string]$RepoUrl,
        [string]$FilePattern
    )

    Write-Verbose "Fetching the latest release URL from $RepoUrl..."
    $apiUrl = "$RepoUrl/releases/latest"
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
    } catch {
        Write-Error "Failed to access the URL: $RepoUrl"
        return $null
    }

    $latestAsset = $response.assets | Where-Object { $_.name -like $FilePattern }
    if ($null -eq $latestAsset) {
        Write-Error "No file matching the pattern '$FilePattern' found in the latest release."
        return $null
    }

    $downloadUrl = $latestAsset.browser_download_url
    Write-Verbose "Latest release URL: $downloadUrl"
    return $downloadUrl
}

function Download {
    param (
        [string]$Url,
        [string]$Destination,
        [string]$FileType
    )

    Write-Verbose "Download called with Url: $Url, Destination: $Destination, FileType: $FileType"

    if (Test-Path $Destination) {
        Write-Verbose "Removing existing directory: $Destination"
        Remove-Item -Path $Destination -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Verbose "Creating directory: $Destination"
    New-Item -ItemType Directory -Force -Path $Destination

    $fileName = [System.IO.Path]::GetFileName($Url)
    $filePath = "$Destination\$fileName"
    Write-Verbose "Downloading file from URL: $Url to $filePath"
    try {
        Invoke-WebRequest -Uri $Url -OutFile $filePath -ErrorAction Stop
    } catch {
        Write-Error "Failed to download the file from URL: $Url"
        return $null
    }

    if (-not (Test-Path $filePath)) {
        Write-Error "Downloaded file not found: $filePath"
        return $null
    }

    Write-Verbose "Downloading completed for file: $filePath"
}

function Install-VCRedist {
    Write-Output "Installing Visual C++ Redistributables..."
    
    # Check for offline package
    $offlineExists, $offlinePath = Test-OfflinePackages -PackageName "VCRedist"
    $vcRedistInstaller = $null
    
    if ($offlineExists) {
        Write-Verbose "Found offline VCRedist package: $offlinePath"
        $vcRedistInstaller = $offlinePath
    } else {
        # If Internet check timed out or failed, try to find the installer in WinApps folder on ISO
        if (-not $internetAvailable -or $OfflineMode) {
            # Check for ISO WinApps directory using the same approach as Install-BraveAndIrfanView.cmd
            Write-Output "No internet connection or offline mode enabled, checking ISO for VCRedist package..."
            
            # Create a similar approach to the CMD script with clearly defined variables
            $ISOFound = $false
            $WinAppsDir = $null
            
            # Create a temporary script to find ISO with WinApps directory
            $tempScriptPath = Join-Path $env:TEMP "find_iso.ps1"
            @"
`$isoDrives = Get-Volume | Where-Object { `$_.DriveType -eq 'CD-ROM' -and `$_.OperationalStatus -eq 'OK' } | Select-Object -ExpandProperty DriveLetter
foreach(`$drive in `$isoDrives) {
  if (Test-Path "`${drive}:\WinApps") {
    Write-Host "`$drive"
    exit
  }
}
"@ | Out-File -FilePath $tempScriptPath -Encoding utf8
            
            try {
                # Execute the temporary script to find the ISO
                $isoDrive = & powershell -ExecutionPolicy Bypass -File $tempScriptPath
                
                if ($isoDrive) {
                    $ISOFound = $true
                    $WinAppsDir = "${isoDrive}:\WinApps"
                    Write-Output "Found WinApps directory on drive ${isoDrive}:"
                }
            }
            finally {
                # Clean up the temporary script
                if (Test-Path $tempScriptPath) {
                    Remove-Item -Path $tempScriptPath -Force
                }
            }
            
            # Check for VCRedist in the WinApps directory
            if ($ISOFound) {
                $vcRedistPath = Join-Path $WinAppsDir "VisualCppRedist_AIO_x86_x64.exe"
                if (Test-Path $vcRedistPath) {
                    Write-Output "Found VCRedist package in ISO WinApps directory: $vcRedistPath"
                    $vcRedistInstaller = $vcRedistPath
                } else {
                    throw "VCRedist package not found in WinApps directory on ISO."
                }
            } else {
                throw "No internet connection available and no ISO with WinApps directory found."
            }
        } else {
            # If we have internet, try downloading as usual
            Write-Output "Downloading Visual C++ Redistributables..."
            $vcRedistUrl = Get-LatestReleaseUrl -RepoUrl "https://api.github.com/repos/abbodi1406/vcredist" -FilePattern "VisualCppRedist_AIO_x86_x64.exe"
            if ($null -eq $vcRedistUrl) {
                throw "Failed to get the latest Visual C++ Redistributables URL."
            }

            Write-Verbose "Visual C++ Redistributables URL: $vcRedistUrl"

            $tempDir = "$env:TEMP\VcRedist"
            $result = Download -Url "$vcRedistUrl" -Destination "$tempDir" -FileType "exe"
            if ($null -eq $result) {
                throw "Failed to download Visual C++ Redistributables."
            }

            $vcRedistInstaller = "$tempDir\VisualCppRedist_AIO_x86_x64.exe"
        }
    }
    
    if (-not (Test-Path $vcRedistInstaller)) {
        throw "Installer file not found: $vcRedistInstaller"
    }

    Write-Verbose "Running installer: $vcRedistInstaller with arguments: /ai /gm2"
    # Create a ProcessStartInfo object to set more detailed parameters
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $vcRedistInstaller
    $processInfo.Arguments = "/ai /gm2"
    $processInfo.UseShellExecute = $true
    $processInfo.Verb = "runas"
    
    # Start the process
    $process = [System.Diagnostics.Process]::Start($processInfo)
    $process.WaitForExit()

    Write-Output "Visual C++ Redistributables installation completed."
}

function Install-DirectX {
    Write-Output "Installing DirectX 9 End-User Runtime..."
    $ErrorActionPreference = 'Stop'

    # Check for offline package
    $offlineExists, $offlinePath = Test-OfflinePackages -PackageName "DirectX"
    $directxInstaller = $null
    $tempDir = "$env:TEMP\directx"
    
    if ($offlineExists) {
        Write-Verbose "Found offline DirectX package: $offlinePath"
        $directxInstaller = $offlinePath
        
        # For offline packages, we use the "/ai /gm2" arguments directly
        Write-Verbose "Installing offline DirectX package with arguments: /ai /gm2"
        # Create a ProcessStartInfo object to set more detailed parameters
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $directxInstaller
        $processInfo.Arguments = "/ai /gm2"
        $processInfo.UseShellExecute = $true
        $processInfo.Verb = "runas"
        
        # Start the process
        $process = [System.Diagnostics.Process]::Start($processInfo)
        $process.WaitForExit()
        Write-Output "DirectX installation from offline package completed."
        return
    } else {
        # If Internet check timed out or failed, try to find the installer in WinApps folder on ISO
        if (-not $internetAvailable -or $OfflineMode) {
            # Check for ISO WinApps directory using the same approach as Install-BraveAndIrfanView.cmd
            Write-Output "No internet connection or offline mode enabled, checking ISO for DirectX package..."
            
            # Create a similar approach to the CMD script with clearly defined variables
            $ISOFound = $false
            $WinAppsDir = $null
            
            # Create a temporary script to find ISO with WinApps directory
            $tempScriptPath = Join-Path $env:TEMP "find_iso.ps1"
            @"
`$isoDrives = Get-Volume | Where-Object { `$_.DriveType -eq 'CD-ROM' -and `$_.OperationalStatus -eq 'OK' } | Select-Object -ExpandProperty DriveLetter
foreach(`$drive in `$isoDrives) {
  if (Test-Path "`${drive}:\WinApps") {
    Write-Host "`$drive"
    exit
  }
}
"@ | Out-File -FilePath $tempScriptPath -Encoding utf8
            
            try {
                # Execute the temporary script to find the ISO
                $isoDrive = & powershell -ExecutionPolicy Bypass -File $tempScriptPath
                
                if ($isoDrive) {
                    $ISOFound = $true
                    $WinAppsDir = "${isoDrive}:\WinApps"
                    Write-Output "Found WinApps directory on drive ${isoDrive}:"
                }
            }
            finally {
                # Clean up the temporary script
                if (Test-Path $tempScriptPath) {
                    Remove-Item -Path $tempScriptPath -Force
                }
            }
            
            # Check for DirectX in the WinApps directory
            if ($ISOFound) {
                $directxPath = Join-Path $WinAppsDir "DirectX_Redist_Repack_x86_x64.exe"
                if (Test-Path $directxPath) {
                    Write-Output "Found DirectX package in ISO WinApps directory: $directxPath"
                    $directxInstaller = $directxPath
                    
                    # For offline packages from ISO, we use the "/ai /gm2" arguments directly
                    Write-Verbose "Installing DirectX package from ISO with arguments: /ai /gm2"
                    # Create a ProcessStartInfo object to set more detailed parameters
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $directxInstaller
                    $processInfo.Arguments = "/ai /gm2"
                    $processInfo.UseShellExecute = $true
                    $processInfo.Verb = "runas"
                    
                    # Start the process
                    $process = [System.Diagnostics.Process]::Start($processInfo)
                    $process.WaitForExit()
                    Write-Output "DirectX installation from ISO completed."
                    return
                } else {
                    throw "DirectX package not found in WinApps directory on ISO."
                }
            } else {
                throw "No internet connection available and no ISO with WinApps directory found."
            }
        }
        
        Write-Output "Downloading DirectX 9 End-User Runtime..."
        # Ensure the temporary directory is clean
        if (Test-Path $tempDir) {
            Write-Verbose "Removing existing directory: $tempDir"
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Verbose "Creating directory: $tempDir"
        New-Item -ItemType Directory -Force -Path $tempDir > $null

        $directxUrl = "https://download.microsoft.com/download/8/4/A/84A35BF1-DAFE-4AE8-82AF-AD2AE20B6B14/directx_Jun2010_redist.exe"
        $directxInstaller = "$tempDir\directx_Jun2010_redist.exe"
        Write-Verbose "Downloading DirectX from URL: $directxUrl to $directxInstaller"
        Invoke-WebRequest -Uri $directxUrl -OutFile $directxInstaller
        
        if (-not (Test-Path $directxInstaller)) {
            throw "DirectX installer not found: $directxInstaller"
        }

        Write-Verbose "Extracting DirectX 9 End-User Runtime..."
        # Create a ProcessStartInfo object for extraction
        $extractInfo = New-Object System.Diagnostics.ProcessStartInfo
        $extractInfo.FileName = $directxInstaller
        $extractInfo.Arguments = "/Q /T:$tempDir"
        $extractInfo.UseShellExecute = $true
        $extractInfo.Verb = "runas"
        
        # Start the extraction process
        $extractProcess = [System.Diagnostics.Process]::Start($extractInfo)
        $extractProcess.WaitForExit()

        $dxSetup = "$tempDir\DXSETUP.exe"
        if (-not (Test-Path $dxSetup)) {
            throw "DXSETUP.exe not found in $tempDir"
        }

        Write-Verbose "Installing DirectX 9 End-User Runtime from $dxSetup with arguments: /silent"
        # Create a ProcessStartInfo object for installation
        $installInfo = New-Object System.Diagnostics.ProcessStartInfo
        $installInfo.FileName = $dxSetup
        $installInfo.Arguments = "/silent"
        $installInfo.UseShellExecute = $true
        $installInfo.Verb = "runas"
        
        # Start the installation process
        $installProcess = [System.Diagnostics.Process]::Start($installInfo)
        $installProcess.WaitForExit()
    }

    Write-Output "DirectX 9 End-User Runtime installation completed."
}

function Main {
    # Variable to track if an error occurred
    $errorOccurred = $false
    
    try {
        Write-Output "Starting installation process..."
        
        # Create offline packages directory if it doesn't exist
        if ($PSCommandPath) {
            $offlineDir = Join-Path (Split-Path -Parent $PSCommandPath) "offline-packages"
            if (-not (Test-Path $offlineDir)) {
                Write-Verbose "Creating offline packages directory: $offlineDir"
                New-Item -ItemType Directory -Force -Path $offlineDir > $null
            }
            
            # Check if we're running from an ISO with WinApps directory
            $isoDrives = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.OperationalStatus -eq 'OK' } | Select-Object -ExpandProperty DriveLetter
            foreach ($driveLetter in $isoDrives) {
                $winAppsPath = "${driveLetter}:\WinApps"
                if (Test-Path $winAppsPath) {
                    Write-Verbose "Found WinApps directory in ISO drive ${driveLetter}"
                    break
                }
            }
        }
        
        Install-VCRedist
        Install-DirectX
        Write-Output "Installation process completed successfully."
    } catch {
        $errorOccurred = $true
        Write-Error "An error occurred: $_"
    } finally {
        # Wait for user input before exiting ONLY if an error occurs
        # AND if the script is started by explorer and running as administrator
        # or if the session is likely to close at the end of the script
        if ($errorOccurred -and ((Is-StartedByExplorer) -and ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or (Is-SessionLikelyToClose))) {
            Read-Host "Press Enter to exit"
        }
    }
}

Main -Verbose:$Verbose
Exit