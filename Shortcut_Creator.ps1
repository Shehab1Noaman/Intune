# Intune Win32 Shortcut Packaging Script

# Function to display a GUI popup for user input
function Get-UserInput-GUI {
    param (
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultValue = ''
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = $Title
    #$Form.Width = 1000
    $Form.Height = 200
    $Form.AutoSize = $True
    $Form.StartPosition = 'CenterScreen'

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Prompt
    $Label.Left = 10
    $Label.Top = 20
    $Label.AutoSize = $True
    $Label.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12)
    $Form.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Text = $DefaultValue
    $TextBox.Left = 10
    $TextBox.Top = 50
    $TextBox.Width = 900
    $TextBox.Height = 20
    $TextBox.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12,[System.Drawing.FontStyle]::Bold)
    $Form.Controls.Add($TextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "OK"
    $OKButton.Left = 290
    $OKButton.Top = 100
    $OKButton.Width = 80
    $OKButton.Add_Click({ $Form.Close() })
    $Form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Left = 380
    $CancelButton.Top = 100
    $CancelButton.Width = 80
    $CancelButton.Add_Click({
        $TextBox.Text = ''
        $Form.Close()
    })
    $Form.Controls.Add($CancelButton)

    $Form.ShowDialog() | Out-Null
    return $TextBox.Text
}

# Prompt user for required inputs
$ShortcutName = Get-UserInput-GUI 'Shortcut Name' 'Enter the name of the shortcut (e.g., MyAppShortcut):'
$ShortcutVersion = Get-UserInput-GUI 'Shortcut Version' 'Enter the shortcut Version (e.g., 1.0):'
$ShortcutTargetType = (Get-UserInput-GUI 'Shortcut Target Type' 'Enter the target type: URL, FilePath, or MsApp:').ToLower()

switch ($ShortcutTargetType) {
    'msapp' {
        $AppID = Get-UserInput-GUI 'MsApp AppUserModelID' 'Enter the AppUserModelID for the Microsoft Store app (e.g., Microsoft.CompanyPortal_8wekyb3d8bbwe!App):'
        # Ensure correct MsApp target format
        if (-not $AppID.Contains("!App")) {
            Write-Host "The MSAPP does not contain '!App'. It will be appended automatically." -ForegroundColor Yellow
            $AppID += "!App"
        }
        $ShortcutTarget = "shell:AppsFolder\$AppID"
        Write-Host "The MSAPP is set to: $AppID" -ForegroundColor Green
    }
    'filepath' {
        $ShortcutTarget = Get-UserInput-GUI 'File Path Target' 'Enter the path to the app or file (e.g., C:\Program Files\MyApp\app.exe):'
        if (-not (Test-Path $ShortcutTarget)) {
            Write-Host "Target path does not exist: $ShortcutTarget" -ForegroundColor Red
            exit
        }
    }
    'url' {
        $ShortcutTarget = Get-UserInput-GUI 'URL Target' 'Enter the URL (e.g., https://example.com):'
    }
    default {
        Write-Host "Invalid target type. Please choose either URL, FilePath, or MsApp." -ForegroundColor Red
        exit
    }
}

$ShortcutIcon = Get-UserInput-GUI 'Shortcut Icon Path' 'Enter the full path to the icon file (.ico):'
$IntuneWinAppUtilPath = Get-UserInput-GUI 'IntuneWinAppUtil Path' 'Enter the full path to IntuneWinAppUtil.exe (e.g., "C:\Path\To\IntuneWinAppUtil.exe")'
$OutputFolder = Get-UserInput-GUI 'Output Folder' 'Enter the output folder for the .intunewin file:'

# Validate inputs
if (-not ($ShortcutName -and $ShortcutVersion -and $ShortcutTarget -and $IntuneWinAppUtilPath -and $ShortcutIcon -and $OutputFolder)) {
    Write-Host 'All fields are required. Please re-run the script and provide valid inputs.' -ForegroundColor Red
    exit
}

if (-not (Test-Path $IntuneWinAppUtilPath)) {
    Write-Host "IntuneWinAppUtil.exe not found at: $IntuneWinAppUtilPath" -ForegroundColor Red
    exit
}

# Prepare working directories
$WorkingDir = Join-Path -Path $env:TEMP -ChildPath "IntuneWinPackage\$ShortcutName"

# Create necessary folders
New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null

# Generate shortcut (.lnk) file
$ShortcutPath = Join-Path -Path $WorkingDir -ChildPath "$ShortcutName.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
if($ShortcutTargetType -eq 'msapp') {
    $Shortcut.TargetPath = "%windir%\explorer.exe"
    $Shortcut.Arguments = $ShortcutTarget 
    }else {
    $Shortcut.TargetPath = $ShortcutTarget
    }

$Shortcut.IconLocation = "%ALLUSERSPROFILE%\ShortcutsIcons\$ShortcutName.ico"
$Shortcut.Save()

#Copy the shortcut icon to install folder
Copy-Item -Path "$ShortcutIcon" -Destination "$WorkingDir\$ShortcutName.ico" -Force

# Generate Files
$PublicDesktopFolder = "C:\Users\Public\Desktop"
$ShortcutIconPath = "C:\ProgramData\ShortcutsIcons"
$ShortcutVersionPath = "C:\ProgramData\ShortcutsVersion\$ShortcutName"

$InstallScriptPath = Join-Path -Path $WorkingDir -ChildPath "Install_$ShortcutName.ps1"
$UninstallScriptPath = Join-Path -Path $WorkingDir -ChildPath "Uninstall_$ShortcutName.ps1"

$DetectionScriptPath = Join-Path -Path $OutputFolder -ChildPath "Detection_$ShortcutName.ps1"
$ShortcutVersionInstallPath = Join-Path -Path $WorkingDir -ChildPath "$ShortcutVersion.ini"
$InstructionsFilePath = Join-Path -Path $OutputFolder -ChildPath "IntuneInstructions_$ShortcutName.txt"


# Create Version file for Intune ###
@"
# ShortcutVersion.ini
# Shortcut version control
$ShortcutVersion
"@ | Set-Content -Path $ShortcutVersionInstallPath

# Create Install file for Intune
@"
# InstallShortcut.ps1
# Copy shortcut to Public Desktop folder
Copy-Item -Path ".\$ShortcutName.lnk" -Destination "$PublicDesktopFolder" -Force

if (-not (Test-Path "$ShortcutIconPath")){New-Item -ItemType Directory -Path "$ShortcutIconPath" -Force | Out-Null}
Copy-Item -Path ".\$ShortcutName.ico" -Destination "$ShortcutIconPath" -Force

if (-not (Test-Path "$ShortcutVersionPath")){New-Item -ItemType Directory -Path "$ShortcutVersionPath" -Force | Out-Null}
Copy-Item -Path ".\$ShortcutVersion.ini" -Destination "$ShortcutVersionPath" -Force
"@ | Set-Content -Path $InstallScriptPath

# Create Uninstall file for Intune
@"
# UninstallShortcut.ps1
# Remove shortcut from Public Desktop folder
Remove-Item -Path "$PublicDesktopFolder\$ShortcutName.lnk" -Force
Remove-Item -Path "$ShortcutVersionPath\$ShortcutVersion.ini" -Force

"@ | Set-Content -Path $UninstallScriptPath


# Create Detection file for Intune ###
@"
# DetectionShortcut.ps1
# Detection Script for $ShortcutName

if ((Test-Path "$ShortcutVersionPath\$ShortcutVersion.ini") -and (Test-Path "$PublicDesktopFolder\$ShortcutName.lnk")) {
    Write-Host "Version .ini file and Shortcut installation detected."
    exit 0
} else {
    Write-Host "Version .ini file and Shortcut installation not detected."
    exit 1
}
"@ | Set-Content -Path $DetectionScriptPath

# Create instructions file for Intune
@"
Intune Packaging Instructions for $ShortcutName :

Install Command: powershell.exe -ExecutionPolicy Bypass -File .\Install_$ShortcutName.ps1
Uninstall Command: powershell.exe -ExecutionPolicy Bypass -File .\Uninstall_$ShortcutName.ps1
Detection Rules: Check if file exists - $PublicDesktopFolder\$ShortcutName.lnk
"@ | Set-Content -Path $InstructionsFilePath

# Format arguments correctly
$Arguments = "-c `"$WorkingDir`" -s `"$InstallScriptPath`" -o `"$OutputFolder`""

# Log the command for debugging
Write-Host "Executing: $IntuneWinAppUtilPath $Arguments"

# Execute the process
$Process = Start-Process -FilePath $IntuneWinAppUtilPath -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden

# Handle process exit codes
if ($Process.ExitCode -eq 0) {
    Write-Host "Packaging completed successfully! .intunewin file is located at $OutputFolder" -ForegroundColor Green
} else {
    Write-Host "Packaging failed with exit code $($Process.ExitCode). Check the logs for details." -ForegroundColor Red
}

#Clean up instructions file
#Remove-Item -Path $WorkingDir -Force
