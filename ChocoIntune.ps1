# ==========================================
# Advanced Management Script for Chocolatey and Intune
# Author: Jornada 365
# Date: 19th August 2024
# Version: 4.2
# ==========================================

# Global Configuration
$IntuneDir = "C:\Jornada365\Intune"
$ToolsDir = Join-Path $IntuneDir "Tools"
$AppsDir = Join-Path $IntuneDir "Apps"
$LogsDir = Join-Path $IntuneDir "Logs"
$IntuneWinAppUtilPath = Join-Path $ToolsDir "IntuneWinAppUtil.exe"
$SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
$ProxyUrl = ""
$MaxRetries = 3
$RetryDelay = 5
$CustomLogoPath = Join-Path $IntuneDir "CustomLogos"
$GenericLogoUrl = "https://via.placeholder.com/150?text=App"
$LogFile = Join-Path $LogsDir "activity.log"
$LogLevel = "INFO"

# Configure Security Protocol
[System.Net.ServicePointManager]::SecurityProtocol = $SecurityProtocol

# Logging Function
function Log-Activity {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    # Define the color based on log level
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN" { "Yellow" }
        default { "Green" }
    }

    try {
        Add-Content -Path $LogFile -Value $logEntry
        Write-Host $logEntry -ForegroundColor $color
    } catch {
        Write-Host "Error writing to log file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Directory Setup Function
function Setup-Directories {
    param ([string[]]$Directories)

    foreach ($dir in $Directories) {
        if (-not (Test-Path -Path $dir)) {
            try {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
                Log-Activity "Directory '$dir' created successfully."
            } catch {
                Log-Activity "Error creating directory '$dir': $($_.Exception.Message)" -Level "ERROR"
                return $false
            }
        }
    }
    return $true
}

# Proxy Setup Function
function Setup-Proxy {
    if ($ProxyUrl -ne "") {
        try {
            $proxy = New-Object System.Net.WebProxy($ProxyUrl, $true)
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            [System.Net.WebRequest]::DefaultWebProxy = $proxy
            Log-Activity "Proxy configured: $ProxyUrl"
        } catch {
            Log-Activity "Error configuring proxy: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

# Installation Helper Function
function Install-Utility {
    param (
        [string]$UtilityName,
        [scriptblock]$InstallCommand
    )

    if (-not (Get-Command $UtilityName -ErrorAction SilentlyContinue)) {
        Write-Host "Installing $UtilityName..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        try {
            & $InstallCommand
            if (Get-Command $UtilityName -ErrorAction SilentlyContinue) {
                Log-Activity "$UtilityName installed successfully."
            } else {
                throw "$UtilityName installation failed."
            }
        } catch {
            Log-Activity "Error installing $UtilityName: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    } else {
        Log-Activity "$UtilityName is already installed."
    }
    return $true
}

# Download File Function
function Download-File {
    param (
        [string]$Url,
        [string]$Destination,
        [int]$Retries = 3
    )

    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing
            if (Test-Path -Path $Destination) {
                Log-Activity "Downloaded file successfully to $Destination"
                return $true
            } else {
                throw "File not found after download."
            }
        } catch {
            Log-Activity "Error downloading file (attempt $($i+1)/$Retries): $($_.Exception.Message)" -Level "WARN"
            Start-Sleep -Seconds $RetryDelay
        }
    }
    Log-Activity "Failed to download file after $Retries attempts." -Level "ERROR"
    return $false
}

# Get Installer Path Function
function Get-InstallerPath {
    param ([string]$PackageFolder)

    $possibleInstallers = @("chocolateyinstall.ps1", "*.exe", "*.msi", "*.bat")
    foreach ($installer in $possibleInstallers) {
        try {
            $installerPath = Get-ChildItem -Path $PackageFolder -Filter $installer -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($installerPath) {
                return $installerPath.FullName
            }
        } catch {
            Log-Activity "Error finding installer in $PackageFolder: $($_.Exception.Message)" -Level "WARN"
        }
    }
    Log-Activity "Installer not found in $PackageFolder" -Level "ERROR"
    return $null
}

# Assign Logo Function
function Assign-Logo {
    param (
        [string]$AppName,
        [string]$AppFolder
    )
    
    $logoPath = Join-Path $AppFolder "logo.png"
    $customLogo = Join-Path $CustomLogoPath "$AppName.png"

    try {
        if (Test-Path -Path $customLogo) {
            Copy-Item -Path $customLogo -Destination $logoPath -Force
            Log-Activity "Custom logo assigned to $AppName."
        } else {
            Invoke-WebRequest -Uri $GenericLogoUrl -OutFile $logoPath -UseBasicParsing
            Log-Activity "Generic logo assigned to $AppName."
        }
        return $true
    } catch {
        Log-Activity "Error assigning logo to $AppName: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Create Deployment Scripts Function
function Create-DeploymentScripts {
    param (
        [string]$AppName,
        [string]$InstallerPath,
        [string]$AppVersion = "latest",
        [string]$AppPath = "C:\Program Files\$AppName\$AppName.exe"
    )
    
    $cleanAppName = $AppName -replace '[^\w\s-]', '' -replace '\s+', '_'
    $appFolder = Join-Path $AppsDir $cleanAppName

    $installScriptPath = Join-Path $appFolder "Install.ps1"
    $uninstallScriptPath = Join-Path $appFolder "Uninstall.ps1"
    $detectionScriptPath = Join-Path $appFolder "Detection.ps1"

    if (-not (Test-Path -Path $InstallerPath)) {
        Log-Activity "Installer not found at $InstallerPath. Cannot create deployment scripts." -Level "ERROR"
        return $null
    }

    Set-Content -Path $installScriptPath -Value @"
Start-Process -FilePath "$InstallerPath" -ArgumentList '/silent' -Wait
"@
    Log-Activity "Installation script created at $installScriptPath"

    Set-Content -Path $uninstallScriptPath -Value @"
if (Test-Path "$AppPath") {
    Start-Process -FilePath "$AppPath\uninstall.exe" -ArgumentList '/silent' -Wait
} else {
    Write-Host "$AppName not found for uninstallation."
}
"@
    Log-Activity "Uninstallation script created at $uninstallScriptPath"

    Set-Content -Path $detectionScriptPath -Value @"
if (Test-Path "$AppPath") {
    Write-Host "$AppName $AppVersion is installed."
    exit 0
} else {
    Write-Host "$AppName $AppVersion is not installed."
    exit 1
}
"@
    Log-Activity "Detection script created at $detectionScriptPath"

    return $detectionScriptPath
}

# Generate IntuneWin Package Function
function Generate-IntuneWinPackage {
    param (
        [string]$AppName,
        [string]$InstallerPath,
        [string]$DetectionScriptPath
    )
    
    $cleanAppName = $AppName -replace '[^\w\s-]', '' -replace '\s+', '_'
    $appFolder = Join-Path $AppsDir $cleanAppName
    $outputFolder = Join-Path $appFolder "Output"
    
    if (-not (Test-Path -Path $outputFolder)) {
        try {
            New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
            Log-Activity "Output directory $outputFolder created successfully."
        } catch {
            Log-Activity "Error creating output directory $outputFolder: $($_.Exception.Message)" -Level "ERROR"
            return $null
        }
    }

    if (Test-Path -Path $IntuneWinAppUtilPath) {
        Write-Host "Generating IntuneWin package for $cleanAppName..." -ForegroundColor Yellow
        try {
            & $IntuneWinAppUtilPath -c $appFolder -s $appFolder -o $outputFolder

            $intunewinPackagePath = Join-Path $outputFolder "$cleanAppName.intunewin"
            if (Test-Path -Path $intunewinPackagePath) {
                Log-Activity "IntuneWin package generated at $intunewinPackagePath"
                return $intunewinPackagePath
            } else {
                throw "IntuneWin package was not generated as expected."
            }
        } catch {
            Log-Activity "Error generating IntuneWin package: $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Log-Activity "Error: IntuneWinAppUtil.exe not found at $IntuneWinAppUtilPath." -Level "ERROR"
    }
    return $null
}

# Rollback Function
function Rollback-Installation {
    param (
        [string]$AppName
    )
    $uninstallScriptPath = Join-Path $AppsDir "${AppName}\Uninstall.ps1"
    if (Test-Path -Path $uninstallScriptPath) {
        try {
            Write-Host "Initiating rollback for $AppName..." -ForegroundColor Yellow
            & PowerShell -ExecutionPolicy Bypass -File $uninstallScriptPath
            Log-Activity "Rollback successfully completed for $AppName."
        } catch {
            Log-Activity "Error rolling back $AppName: $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Log-Activity "Uninstallation script not found for $AppName." -Level "ERROR"
    }
}

# Chocolatey Package Installation Function
function Install-ChocoPackage {
    param ([string]$PackageName)

    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            Write-Host "Installing $PackageName via Chocolatey..." -ForegroundColor Yellow
            choco install $PackageName -y
            Log-Activity "$PackageName installed successfully."
            return $true
        } catch {
            Log-Activity "Error installing $PackageName (attempt $($i+1)/$MaxRetries): $($_.Exception.Message)" -Level "WARN"
            Start-Sleep -Seconds $RetryDelay
        }
    }
    Log-Activity "Failed to install $PackageName after $MaxRetries attempts." -Level "ERROR"
    return $false
}

# GUI Creation Function
function Create-GUI {
    try {
        Add-Type -AssemblyName System.Windows.Forms
    } catch {
        Log-Activity "Error loading System.Windows.Forms assembly: $($_.Exception.Message)" -Level "ERROR"
        exit
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Choco Intune | Jornada 365"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

    # Add Logo
    $logo = New-Object System.Windows.Forms.PictureBox
    $logo.ImageLocation = "https://jornada365.cloud/wp-content/uploads/2024/03/Logotipo-Jornada-365-Home.png"
    $logo.SizeMode = "Zoom"
    $logo.Size = New-Object System.Drawing.Size(180, 90)
    $logo.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($logo)

    # Add Search Bar
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Size = New-Object System.Drawing.Size(460, 30)
    $searchBox.Location = New-Object System.Drawing.Point(200, 40)
    $form.Controls.Add($searchBox)

    # Search Button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(670, 40)
    $searchButton.Size = New-Object System.Drawing.Size(80, 30)
    $searchButton.BackColor = [System.Drawing.Color]::DodgerBlue
    $searchButton.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($searchButton)

    # Results List
    $resultsList = New-Object System.Windows.Forms.ListBox
    $resultsList.Size = New-Object System.Drawing.Size(760, 320)
    $resultsList.Location = New-Object System.Drawing.Point(20, 120)
    $resultsList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($resultsList)

    # Deploy Options Group Box
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Deploy Options"
    $groupBox.Location = New-Object System.Drawing.Point(20, 460)
    $groupBox.Size = New-Object System.Drawing.Size(760, 80)
    $groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($groupBox)

    # Deploy Intune Option
    $deployIntune = New-Object System.Windows.Forms.RadioButton
    $deployIntune.Text = "Deploy Intune"
    $deployIntune.Location = New-Object System.Drawing.Point(20, 30)
    $deployIntune.Size = New-Object System.Drawing.Size(120, 30)
    $groupBox.Controls.Add($deployIntune)

    # IntuneWin Package Option
    $intuneWin = New-Object System.Windows.Forms.RadioButton
    $intuneWin.Text = "IntuneWin Package"
    $intuneWin.Location = New-Object System.Drawing.Point(160, 30)
    $intuneWin.Size = New-Object System.Drawing.Size(150, 30)
    $groupBox.Controls.Add($intuneWin)

    # Apply Button
    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = "Apply"
    $applyButton.Location = New-Object System.Drawing.Point(520, 550)
    $applyButton.Size = New-Object System.Drawing.Size(100, 40)
    $applyButton.BackColor = [System.Drawing.Color]::MediumSeaGreen
    $applyButton.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($applyButton)

    # Close Button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(640, 550)
    $closeButton.Size = New-Object System.Drawing.Size(100, 40)
    $closeButton.BackColor = [System.Drawing.Color]::Tomato
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.Add_Click({
        Log-Activity "Closing application."
        $form.Close()
    })
    $form.Controls.Add($closeButton)

    # Search Button Logic
    $searchButton.Add_Click({
        $searchTerm = $searchBox.Text.ToLower()
        if (-not [string]::IsNullOrEmpty($searchTerm)) {
            try {
                $results = & choco search $searchTerm --by-id-only | Select-String -Pattern '^[^|]+'
                $resultsList.Items.Clear()
                if ($results) {
                    $filteredResults = $results | ForEach-Object {
                        $_.Matches.Groups[0].Value.Trim() -replace '[^\w\s-]', '' -replace '\s+', '_'
                    }
                    $resultsList.Items.AddRange($filteredResults)
                    Log-Activity "Search completed successfully for term '$searchTerm'."
                } else {
                    [System.Windows.Forms.MessageBox]::Show("No results found.")
                    Log-Activity "No results found for term '$searchTerm'."
                }
            } catch {
                Log-Activity "Error executing choco search command: $($_.Exception.Message)" -Level "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Error executing choco search command.")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter a search term.")
        }
    })

    # Apply Button Logic
    $applyButton.Add_Click({
        $selectedPackages = $resultsList.SelectedItems
        if ($selectedPackages.Count -gt 0) {
            foreach ($package in $selectedPackages) {
                $packageName = $package.Trim()

                if (-not (Install-ChocoPackage -PackageName $packageName)) {
                    [System.Windows.Forms.MessageBox]::Show("Failed to install package $packageName.")
                    continue
                }

                $packageFolder = "C:\ProgramData\chocolatey\lib\$packageName\tools\"
                $installerPath = Get-InstallerPath -PackageFolder $packageFolder

                if ($installerPath) {
                    $detectionScriptPath = Create-DeploymentScripts -AppName $packageName -InstallerPath $installerPath
                    if ($detectionScriptPath) {
                        $intunewinPackagePath = Generate-IntuneWinPackage -AppName $packageName -InstallerPath $installerPath -DetectionScriptPath $detectionScriptPath
                        if ($intunewinPackagePath) {
                            Assign-Logo -AppName $packageName -AppFolder (Split-Path $installerPath)
                            [System.Windows.Forms.MessageBox]::Show("Deployment of package $packageName completed successfully!")
                            Log-Activity "Deployment of package $packageName completed successfully."
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Failed to generate IntuneWin package for $packageName.")
                        }
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Failed to create deployment scripts for $packageName.")
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Installer not found for $packageName.")
                    Log-Activity "Installer not found for $packageName." -Level "ERROR"
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a package.")
        }
    })

    $form.AutoSize = $true
    $form.AutoSizeMode = "GrowAndShrink"
    $form.ShowDialog()
}

# Main Function
function Main {
    if (-not (Setup-Directories -Directories @($IntuneDir, $ToolsDir, $AppsDir, $LogsDir, $CustomLogoPath))) { exit }
    Setup-Proxy
    if (-not (Install-Utility -UtilityName "choco" -InstallCommand { Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) })) { exit }
    if (-not (Download-File -Url "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" -Destination $IntuneWinAppUtilPath)) { exit }
    Create-GUI
}

# Execute Main Function
Main
