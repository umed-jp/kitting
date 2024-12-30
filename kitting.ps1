function Disable-WindowsWidgets {
    Write-Host "Disabling Windows Widgets..."

    # Define the registry path and key for Widgets
    $regPath = "HKLM:\Software\Policies\Microsoft\Dsh"
    $regName = "AllowNewsAndInterests"

    # Ensure the registry path exists
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set the registry value to disable Widgets
    Set-ItemProperty -Path $regPath -Name $regName -Value 0

    Write-Host "Windows Widgets have been disabled."
}

function Hide-TaskbarSearch {
    Write-Host "Hiding Taskbar Search..."

    # Define the registry path for taskbar search settings
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
    $regName = "SearchboxTaskbarMode"

    # Ensure the registry path exists
    if (!(Test-Path -Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set the registry value to hide Taskbar Search
    # 0 = Hidden, 1 = Show search icon, 2 = Show search box
    Set-ItemProperty -Path $regPath -Name $regName -Value 0

    # Restart Explorer to apply changes
    Write-Host "Restarting Explorer to apply changes..."
    Stop-Process -Name explorer -Force
    Start-Process explorer

    Write-Host "Taskbar Search has been hidden."
}

function Remove-UnnecessaryApps {
    Write-Host "Starting removal of unnecessary applications..."

    # List of App Names or Keywords to Remove
    $appNamesToRemove = @(
        "Xbox",
        "Access",
        "McAfee",
        "Dropbox",
        "ToDo",
        "LinkedIn",
        "Clipchamp",
        "Copilot",
        "Microsoft.Windows.Photos",
        "Cortana",
        "OneNote",
        "Outlook",
        "Dolby",
        "Gamebar",
        "Teams",
        "Realtek",
        "Skype for Business",
        "AMD Software"
        "Solitaire"
    )

    # Loop through each application name and remove
    foreach ($appName in $appNamesToRemove) {
        Write-Host "Processing application: $appName"

        Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$appName*" } | Remove-AppxProvisionedPackage -Online

        # Remove AppxPackage for Current User
        Get-AppxPackage -Name "*$appName*" -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "Removing AppxPackage: $($_.Name)"
            Remove-AppxPackage -Package $_.PackageFullName -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Unnecessary applications removed successfully."
}

# Function: Install Software (Slack, Chrome, etc.)
function Install-Software {
    param (
        [string]$DownloadUrl,
        [string]$InstallerPath,
        [string]$Arguments
    )

    Write-Host "Downloading from $DownloadUrl..."
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -UseBasicParsing
    Write-Host "Downloaded to $InstallerPath"

    Write-Host "Installing $InstallerPath..."
    Start-Process -FilePath $InstallerPath -ArgumentList $Arguments -Wait
    Write-Host "Installation of $InstallerPath completed."

    # Clean up installer
    if (Test-Path -Path $InstallerPath) {
        Remove-Item -Path $InstallerPath -Force
        Write-Host "Deleted installer: $InstallerPath"
    } else {
        Write-Host "Installer not found for deletion: $InstallerPath"
    }
}

# Function: Install Slack
function Install-Slack {
    Write-Host "Installing Slack..."
    $slackUrl = "https://downloads.slack-edge.com/releases_x64/SlackSetup.exe"
    $slackInstaller = "$env:TEMP\SlackSetup.exe"
    Install-Software -DownloadUrl $slackUrl -InstallerPath $slackInstaller -Arguments "/S"
}

# Function: Install Chrome
function Install-Chrome {
    Write-Host "Installing Google Chrome..."
    $chromeUrl = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
    $chromeInstaller = "$env:TEMP\ChromeInstaller.exe"
    Install-Software -DownloadUrl $chromeUrl -InstallerPath $chromeInstaller -Arguments "/silent /install"
}

# Function: Install Office 365 (Word, Excel, PowerPoint)
function Install-Office365 {
    Write-Host "Installing Office 365 Business Apps (Word, Excel, PowerPoint)..."

    # Define Office Deployment Tool URL and paths
    $odtUrl = "https://go.microsoft.com/fwlink/p/?linkid=869426"
    $odtInstaller = "$env:TEMP\OfficeSetup.exe"
    $configPath = "$env:TEMP\configuration.xml"

    # Step 1: Download Office Deployment Tool
    Install-Software -DownloadUrl $odtUrl -InstallerPath $odtInstaller -Arguments ""

    # Step 2: Create Configuration File
    Write-Host "Creating Office configuration file..."
    @"
        <Configuration>
        <Add OfficeClientEdition="64" Channel="Current">
            <Product ID="O365BusinessRetail">
            <Language ID="ja-jp" />
            <ExcludeApp ID="Access" />
            <ExcludeApp ID="Outlook" />
            <ExcludeApp ID="OneDrive" />
            <ExcludeApp ID="Publisher" />
            <ExcludeApp ID="Teams" />
            </Product>
        </Add>
        <RemoveMSI All="True" />
        <Display Level="None" AcceptEULA="TRUE" />
        </Configuration>
"@ | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "Office configuration file created at $configPath."

    # Step 3: Install Office
    Start-Process -FilePath $odtInstaller -ArgumentList "/configure $configPath" -Wait
    Write-Host "Office installation completed."
}


function Update-Windows {
    if (!(Get-Command -Name Install-WindowsUpdate -ErrorAction SilentlyContinue)) {
        Write-Host "Installing PSWindowsUpdate module..."
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
    }

    Write-Host "Checking for available updates..."
    $updates = Get-WindowsUpdate -Verbose

    if ($updates) {
        Write-Host "Installing updates..."
        Install-WindowsUpdate -AcceptAll -AutoReboot -Verbose
    } else {
        Write-Host "No updates available."
    }
}

function Update-Pins {
    Write-Host "Removing pins"
    Remove-Item -Path "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*" -Force -Recurse -ErrorAction SilentlyContinue

    # Remove the Taskband Registry Key to delete taskbar data for recent apps.
    Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -Force -Recurse -ErrorAction SilentlyContinue
    
    #Start-Stop the File Explorer to refresh the taskbar.
    Stop-Process -ProcessName explorer -Force
    Start-Process explorer
}

function Remove-StartupApps {
    Write-Host "Removing startup apps"
    # Define critical programs and services to preserve
    $criticalPrograms = @(
        "SecurityHealth",      # Windows Security health monitor
        "Windows Defender",    # Defender Antivirus
        "Chrome",              # Google Chrome
        "Slack"                # Slack
    )


    $criticalServices = @(
        "WinDefend",           # Windows Defender Antivirus Service
        "SecurityHealthService" # Security Health Service

    )

    # Ensure critical Windows Security services are running and set to automatic
    foreach ($service in $criticalServices) {
        $serviceStatus = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($serviceStatus) {
            if ($serviceStatus.Status -ne "Running") {
                Start-Service -Name $service
                Write-Output "Started service: $service"
            }
            Set-Service -Name $service -StartupType Automatic
            Write-Output "Set service to automatic: $service"
        } else {
            Write-Output "Service not found: $service (may not be present on this system)."
        }
    }

    # Disable non-critical startup items in the registry (Current User)
    Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" |
        ForEach-Object {
            if ($_.PSChildName -notin $criticalPrograms) {
                Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $_.PSChildName -ErrorAction SilentlyContinue
                Write-Output "Disabled registry startup item: $($_.PSChildName) (Current User)"
            }
        }


    # Remove unrelated startup shortcuts from the Startup folder (Current User)
    Get-ChildItem -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup" |
        ForEach-Object {
            $fileName = $_.BaseName
            if ($fileName -notin $criticalPrograms) {
                Remove-Item -Path $_.FullName -ErrorAction SilentlyContinue
                Write-Output "Removed startup shortcut: $fileName (Current User)"
            }
        }
}

function Add-Bookmarks {
    # Define the bookmarks file path
    $bookmarksFilePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"

    # Check if the Bookmarks file exists
    if (Test-Path $bookmarksFilePath) {
    # Read the existing Bookmarks JSON
        $bookmarksJson = Get-Content $bookmarksFilePath -Raw | ConvertFrom-Json
        $newBookmarks = @(
            @{ name = "クリアス"; url = "https://web.clius.jp" },
            @{ name = "カレンダー (Google Calendar)"; url = "https://calendar.google.com" },
            @{ name = "Notion"; url = "https://www.notion.so/umed-group/Group-All-da85f0542882403684771ab796f1ac07" },
            @{ name = "Google Drive"; url = "https://drive.google.com" },
            @{ name = "Microsoft MyApps"; url = "https://myapps.microsoft.com" },
            @{ name = "Gmail"; url = "https://gmail.com" },
            @{ name = "MCS"; url = "https://www.medical-care.net/home?restore_list=1" }
        )

    # Add the new bookmarks under the bookmark bar
        $bookmarksJson.roots.bookmark_bar.children += $newBookmarks

    # Save the updated Bookmarks JSON back to the file
        $bookmarksJson | ConvertTo-Json -Depth 10 | Set-Content -Path $bookmarksFilePath -Force

        Write-Output "Bookmarks updated successfully."
    } else {
        Write-Output "Chrome Bookmarks file not found. Please ensure Chrome is installed and initialized."
    }
}



# Example Main Kitting Script
Write-Host "Starting kitting process..."

# Step1: Update Windows
Update-Windows

# Step2: Clean up unnecessary apps
Remove-UnnecessaryApps

# Step3: Disable unused features
Disable-WindowsWidgets
Hide-TaskbarSearch
# Remove all pinned (Not tested)
Update-Pins

# Step4: Install necessary software
Install-Slack
Install-Chrome
Install-Office365

# Step5: Remove stoped startup apps (Not tested)
Remove-StartupApps

# Step6: Add bookmarks to Chrome
Add-Bookmarks

Write-Host "Kitting process completed successfully."

