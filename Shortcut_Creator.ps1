# Intune Win32 Shortcut Packaging Script

# Ensure the script is running in the correct directory
Set-Location -Path $PSScriptRoot

# Add required assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Logger Function
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile
    )
    $TimeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $LogMessage = "$TimeStamp - $Message"
    
    # Write log to file
    $LogMessage | Out-File -FilePath $LogFile -Append
    
}

# Function to display a GUI popup for user input
function Get-UserInput-GUI {
    param (
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultValue = ''
    )

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = $Title
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
    $TextBox.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
    $Form.Controls.Add($TextBox)

    $global:UserCancelled = $false

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "OK"
    $OKButton.Left = 290
    $OKButton.Top = 100
    $OKButton.Width = 80
    $OKButton.Add_Click({
        $Form.Close() 
    })
    $Form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Left = 380
    $CancelButton.Top = 100
    $CancelButton.Width = 80
    $CancelButton.Add_Click({
        $global:UserCancelled = $true
        $Form.Close()
    })
    $Form.Controls.Add($CancelButton)

    $Form.ShowDialog() | Out-Null
    
    # Return the result based on whether the user cancelled or not
    if ($global:UserCancelled) {
        exit  # Return an empty string if Cancel was clicked
    } else {
        return $TextBox.Text  # Return the text entered in the TextBox if OK was clicked
    }

}

# Define the function to display the error message with an exit button
function Show-ErrorAndExit {
    param (
        [string]$Message,
        [string]$LogMessage,
        [string]$LogFile
    )

    # Display a GUI message box with an Exit button
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    if($LogFile){
    # Log the error message to the log file
    Write-Log -Message $LogMessage -LogFile $LogFile
    }

    # Exit the script
   exit
}

# Check exsting IntuneWinAppUtil
function Delete-FileWithConfirmation {
    param (
        [string]$FilePath,
        [string]$LogFile
    )

    # Show a Yes/No message box asking if the user wants to delete the file
    $result = [System.Windows.Forms.MessageBox]::Show("The $ShortcutName.intunewin file already exists in the destination folder. Do you want to delete the file: $FilePath?", 
                                                      "Delete File Confirmation", 
                                                      [System.Windows.Forms.MessageBoxButtons]::YesNo, 
                                                      [System.Windows.Forms.MessageBoxIcon]::Question)

    # Check the result of the message box (Yes or No)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # If Yes, delete the file
        if (Test-Path $IntuneWinFile) {
            Remove-Item $IntuneWinFile -Force
            Write-Log -Message "File $IntuneWinFile deleted successfully." -LogFile $LogFile
        } else {
            Write-Log -Message "File $IntuneWinFile not found. Nothing to delete." -LogFile $LogFile
        }
    } else {
        # If No, exit the script
        Write-Log -Message "File deletion canceled. Please rerun the script." -LogFile $LogFile
        exit
    }
}


# Prompt user for required inputs
$ShortcutName = Get-UserInput-GUI 'Shortcut Name' 'Enter the name of the shortcut (e.g., MyAppShortcut):'
if (-not ($ShortcutName)) {
    Show-ErrorAndExit -Message 'Shortcut Name is required. Please re-run the script and provide the Shortcut Name.'
         
}else{
    #Creat workingDir
    $WorkingDir = Join-Path -Path $env:TEMP -ChildPath "IntuneWinPackage\$ShortcutName"
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    $LogFile = Join-Path -Path $WorkingDir -ChildPath "PackagingLog.txt"

        
    Write-Log -Message "Copied initial log data to final log file at $LogFile" -LogFile $LogFile

    Write-Log -Message "User provided Shortcut Name: $ShortcutName" -LogFile $LogFile
}

$ShortcutVersion = Get-UserInput-GUI 'Shortcut Version' 'Enter the shortcut Version (e.g., 1.0):'
if (-not ($ShortcutVersion)) {
    Show-ErrorAndExit -Message 'Shortcut Version is required. Please re-run the script and provide the Shortcut Version.'`
                        -LogMessage 'AShortcut Version is  required. Exiting script.' -LogFile $LogFile
}else{
    Write-Log -Message "User provided Shortcut Version: $ShortcutVersion" -LogFile $LogFile
}

$ShortcutTargetType = (Get-UserInput-GUI 'Shortcut Target Type' 'Enter the target type: URL, FilePath, or MsApp:').ToLower()
if (-not ($ShortcutTargetType)) {
    Show-ErrorAndExit -Message 'Please re-run the script and Enter the target type: URL, FilePath, or MsApp'`
                        -LogMessage 'Shortcut Type is required. Exiting script.' -LogFile $LogFile
}else{
    Write-Log -Message "User provided Shortcut Target Type: $ShortcutTargetType" -LogFile $LogFile
}

switch ($ShortcutTargetType) {
    'msapp' {
        $AppID = Get-UserInput-GUI 'MsApp AppUserModelID' 'Enter the AppUserModelID for the Microsoft Store app (e.g., Microsoft.CompanyPortal_8wekyb3d8bbwe!App):'
        Write-Log -Message "User provided MsApp AppUserModelID: $AppID" -LogFile $LogFile

        if (-not ($AppID)) {
            Show-ErrorAndExit -Message 'Please re-run the script and Enter vaild  Application User Model ID (AppID)'`
                        -LogMessage 'Please re-run the script and Enter vaild  Application User Model ID (AppID).' -LogFile $LogFile
        }elseif (-not $AppID.Contains("!App")) {
            Write-Log -Message "The MSAPP does not contain '!App'. It will be appended automatically." -LogFile $LogFile
            $AppID += "!App"
        }
        $ShortcutTarget = "shell:AppsFolder\$AppID"
        Write-Log -Message "The MSAPP is set to: $AppID" -LogFile $LogFile
    }
    'filepath' {
        $ShortcutTarget = Get-UserInput-GUI 'File Path Target' 'Enter the path to the app or file (e.g., C:\Program Files\MyApp\app.exe):'
        
         if (-not ($ShortcutTarget)) {
            Show-ErrorAndExit -Message 'Please re-run the script and Enter vaild  File Path'`
                        -LogMessage 'File Path cant be Empty.' -LogFile $LogFile
        }elseif (-not (Test-Path $ShortcutTarget)) {
            Show-ErrorAndExit -Message "Target path does not exist: $ShortcutTarget" `
                                -LogMessage "Target path does not exist: $ShortcutTarget" -LogFile $LogFile
        }else{
            Write-Log -Message "User provided File Path Target: $ShortcutTarget" -LogFile $LogFile
            }
    }
    'url' {
        $ShortcutTarget = Get-UserInput-GUI 'URL Target' 'Enter the URL (e.g., https://example.com):'
          if (-not ($ShortcutTarget) -or -not ($ShortcutTarget -match '^(https?|ftp|sftp|file)://')) {
                Show-ErrorAndExit -Message 'Invalid URL. Please re-run the script and enter a valid URL (e.g., https://example.com, ftp://example.com, sftp://example.com, file://path).' `
                                  -LogMessage 'URL input was invalid or empty.' -LogFile $LogFile
           } else {
                try {
                    # Parse the URL using [System.Uri] for advanced validation
                    $Uri = [System.Uri]::new($ShortcutTarget)
                    if (-not $Uri.IsAbsoluteUri) {
                        throw [System.Exception]::new("The URL must be an absolute URI.")
                    }
                    # Log the valid URL input
                    Write-Log -Message "User provided a valid URL Target: $ShortcutTarget" -LogFile $LogFile
                } catch {
                        # Log and display an error if [System.Uri] validation fails
                        Show-ErrorAndExit -Message 'Invalid URL. Please ensure the URL is correctly formatted and absolute.' `
                                         -LogMessage "URL validation failed: $_" -LogFile $LogFile
                         }
                   }
    }
    default {
       Show-ErrorAndExit -Message "Invalid target type. Please choose either URL, FilePath, or MsApp."`
                            -LogMessage "Invalid target type entered: $ShortcutTargetType" -LogFile $LogFile
    }
}

$ShortcutIcon = Get-UserInput-GUI 'Shortcut Icon Path' 'Enter the full path to the icon file (.ico):'
if (-not ($ShortcutIcon)) {
    Show-ErrorAndExit -Message 'Please re-run the script and Enter the icon Path'`
                        -LogMessage 'No Icon path added.' -LogFile $LogFile
}elseif (-not (Test-Path $ShortcutIcon)) {
    Show-ErrorAndExit -Message "$ShortcutIcon not found. Please re-run the script and provide valid input."`
                       -LogMessage "$ShortcutIcon not found. Exiting script." -LogFile $LogFile
}else{
    Write-Log -Message "User provided Shortcut Icon Path: $ShortcutIcon" -LogFile $LogFile
    }


$OutputFolder = Get-UserInput-GUI 'Output Folder' 'Enter the output folder for the .intunewin file:'
if (-not ($OutputFolder)) {
    $OutputFolder = ".\"
}elseif (-not (Test-Path $OutputFolder)) {
    Show-ErrorAndExit -Message "$OutputFolder not found. Please re-run the script and provide valid input."`
                       -LogMessage "$OutputFolder not found. Exiting script." -LogFile $LogFile
}else{
    Write-Log -Message "User provided Output Folder: $OutputFolder" -LogFile $LogFile
    }


# Check IntuneWinAppUtil for packaging
$IntuneWinAppUtilPath = ".\IntuneWinAppUtil.exe"
$IntuneWinFile = Join-Path -Path $OutputFolder -ChildPath "Install_$ShortcutName.intunewin"

if (-not (Test-Path $IntuneWinAppUtilPath)) {
    Show-ErrorAndExit -Message "$IntuneWinAppUtilPath not found"`
                        -LogMessage "$IntuneWinAppUtilPath not found. Exiting script." -LogFile $LogFile
}elseif(Test-Path $IntuneWinFile){

Delete-FileWithConfirmation -FilePath $IntuneWinFile -LogFile $LogFile
#Show-ErrorAndExit -Message "The $ShortcutName.intunewin file already exists in the destination folder"`
                       #-LogMessage "The $ShortcutName.intunewin file already exists in the destination folder." -LogFile $LogFile
}

# Generate shortcut (.lnk) file
$ShortcutPath = Join-Path -Path $WorkingDir -ChildPath "$ShortcutName.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
if ($ShortcutTargetType -eq 'msapp') {
    $Shortcut.TargetPath = "%windir%\explorer.exe"
    $Shortcut.Arguments = $ShortcutTarget 
} else {
    $Shortcut.TargetPath = $ShortcutTarget
}
$Shortcut.IconLocation = "%ALLUSERSPROFILE%\ShortcutsIcons\$ShortcutName.ico"
$Shortcut.Save()
Write-Log -Message "Created shortcut: $ShortcutPath" -LogFile $LogFile

# Copy the shortcut icon to install folder
Copy-Item -Path "$ShortcutIcon" -Destination "$WorkingDir\$ShortcutName.ico" -Force
Write-Log -Message "Copied icon to: $WorkingDir\$ShortcutName.ico" -LogFile $LogFile

# Create Version, Install, Uninstall, Detection, Instructions and Shortcut Version path  files for Intune
$InstallScriptPath = Join-Path -Path $WorkingDir -ChildPath "Install_$ShortcutName.ps1"
$UninstallScriptPath = Join-Path -Path $WorkingDir -ChildPath "Uninstall_$ShortcutName.ps1"
$DetectionScriptPath = Join-Path -Path $OutputFolder -ChildPath "Detection_$ShortcutName.ps1"
$InstructionsFilePath = Join-Path -Path $OutputFolder -ChildPath "IntuneInstructions_$ShortcutName.txt"
$ShortcutVersionInstallPath = Join-Path -Path $WorkingDir -ChildPath "$ShortcutVersion.ini"

# Create Version file for Intune
@"
# ShortcutVersion.ini
# Shortcut version control
$ShortcutVersion
"@ | Set-Content -Path $ShortcutVersionInstallPath
Write-Log -Message "Created version file: $ShortcutVersionInstallPath" -LogFile $LogFile

# Generate Files
$PublicDesktopFolder = "C:\Users\Public\Desktop"
$ShortcutIconPath = "C:\ProgramData\ShortcutsIcons"
$ShortcutVersionPath = "C:\ProgramData\ShortcutsVersion\$ShortcutName"

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
Write-Log -Message "Created version file: $InstallScriptPath" -LogFile $LogFile

# Create Uninstall file for Intune
@"
# UninstallShortcut.ps1
# Remove shortcut from Public Desktop folder
Remove-Item -Path "$PublicDesktopFolder\$ShortcutName.lnk" -Force
Remove-Item -Path "$ShortcutVersionPath\$ShortcutVersion.ini" -Force

"@ | Set-Content -Path $UninstallScriptPath
Write-Log -Message "Created version file: $UninstallScriptPath" -LogFile $LogFile

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
Write-Log -Message "Created version file: $DetectionScriptPath" -LogFile $LogFile

# Create instructions file for Intune
@"
Intune Packaging Instructions for $ShortcutName :

Install Command: powershell.exe -ExecutionPolicy Bypass -File ".\Install_$ShortcutName.ps1"
Uninstall Command: powershell.exe -ExecutionPolicy Bypass -File ".\Uninstall_$ShortcutName.ps1"
Detection Rules: Check if file exists - $PublicDesktopFolder\$ShortcutName.lnk
"@ | Set-Content -Path $InstructionsFilePath
Write-Log -Message "Created version file: $InstructionsFilePath" -LogFile $LogFile


$Arguments = "-c `"$WorkingDir`" -s `"$InstallScriptPath`" -o `"$OutputFolder`""
Write-Log -Message "Executing: $IntuneWinAppUtilPath $Arguments" -LogFile $LogFile

$Process = Start-Process -FilePath $IntuneWinAppUtilPath -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
Write-Log -Message "Process exit code: $($Process.ExitCode)" -LogFile $LogFile

# Handle process exit codes
if ($Process.ExitCode -eq 0) {

    Write-Log -Message "Packaging completed successfully! .intunewin file is located at $OutputFolder" -LogFile $LogFile
} else {
    Write-Log -Message "Packaging failed with exit code $($Process.ExitCode)" -LogFile $LogFile
}

# Clean up temporary working directory after packaging
Remove-Item -Path $WorkingDir -Recurse -Force
Write-Log -Message "Cleaned up working directory: $WorkingDir" -LogFile $LogFile

Write-Log -Message "Packaging process completed successfully!" -LogFile $LogFile
