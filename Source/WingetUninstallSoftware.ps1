<#
.SYNOPSIS
    Uninstalls software, manages registry keys, and logs events.

.DESCRIPTION
    This script uninstalls a specified software package using Winget, logs uninstallation events to the Event Viewer, and updates the registry by removing the software version key.

.NOTES
    This script requires administrative privileges to execute due to event viewer logs and HKLM keys. Ensure to run the script with elevated permissions.
    Also some applications don't work well with silent removal from Winget. Ensure to test this before uploading. If a prompt is shown to a user, use the software specific removal to do so
.PARAMETER Winget
    STRING - The Winget package to be uninstalled. Can be found through PowerShell using winget search <name> or through https://winget.run
    Example: Microsoft.VCRedist.2015+.x64

.EXAMPLE
    %SystemRoot%\sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -windowstyle hidden -file ".\WingetUninstallSoftware.ps1 -Winget "Microsoft.VCRedist.2015+.x64"
    This needs to be %SystemRoot% because Intune starts PowerShell as 32-bit when using powershell.exe. This causes the registry keys to be made in: HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Enne instead of: HKEY_LOCAL_MACHINE\SOFTWARE\Enne

.VERSION HISTORY
    v1.0 - 25 January 2025 - Dave Huijten
    - Initial version of uninstall script. Used install script as base to build on
#>
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Winget,

    [Parameter(Mandatory=$true)]
    [int]$PackageVersion,

    [string]$Wingetparam
)

# Function to log events to the Event Viewer
function Log-Event {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$message,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Error", "Success", "Information")]
        [string]$type,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$WingetSoftware
    )

    # Define the source of the event
    $eventSource = "Enne"
    $eventLog = "Application"
    $eventId = 69
    $eventIdSuccess = 045
    $EventIdError = 043 

    # Define the registry path to check and create if necessary
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$eventSource"

    try {
        # Check if the registry key exists
        if (-not (Test-Path $registryPath)) {
            # Create the registry key if it does not exist
            New-Item -Path $registryPath -Force
        }

        # Check if the event source exists and if it's registered in the correct log
        if ([System.Diagnostics.EventLog]::SourceExists($eventSource)) {
            $sourceLog = (Get-EventLog -LogName $eventLog -Newest 1 -Source $eventSource -ErrorAction SilentlyContinue).Log
            if ($null -eq $sourceLog -or $sourceLog -ne $eventLog) {
                [System.Diagnostics.EventLog]::DeleteEventSource($eventSource)
                [System.Diagnostics.EventLog]::CreateEventSource($eventSource, $eventLog)
            }
        } else {
            [System.Diagnostics.EventLog]::CreateEventSource($eventSource, $eventLog)
        }

        # Log the event based on the type
        switch ($type) {
            "Error" {
                Write-EventLog -LogName $eventLog -Source $eventSource -EntryType Error -EventId $eventId -Message "Enne - Tuup | Auw banaan, ${WingetSoftware} gaat mij echt op de wekker: ${message}"
                Write-Host "Enne - Tuup | Auw banaan, ${WingetSoftware} gaat mij echt op de wekker: ${message}"
            }
            "Success" {
                Write-EventLog -LogName $eventLog -Source $eventSource -EntryType Information -EventId $eventIdSuccess -Message "Enne - Tuup | ${message} Tijd om shag te halen!"
                Write-Host "Enne - Tuup | ${message} Tijd om shag te halen!"
            }
            default {
                Write-EventLog -LogName $eventLog -Source $eventSource -EntryType Information -EventId $EventIdError -Message "Enne - Tuup |  ${message}"
                Write-Host "Enne - Tuup | ${message}"
            }
        }
    } catch {
        Write-Error "Enne - Tuup | Auw banaan : $_"
    }
}

# Set the log path and file name
# Extract text before the first period
$WingetSupplier = $winget -split '\.' | Select-Object -First 1

# Extract everything after the first period
$WingetSoftware = $winget.Substring($winget.IndexOf('.') + 1)

# Get the first letter of the supplier and convert it to uppercase
$FirstLetterSoftwareSupplier = $WingetSupplier.Substring(0, 1).ToUpper()
$logPath = "$ENV:ProgramData\Enne\Logs\Packaging\$FirstLetterSoftwareSupplier\$WingetSupplier\$WingetSoftware\Winget"
$transcriptFile = "$logPath\Removal-$WingetSoftware.log"

# Start logging
Start-Transcript -Path $transcriptFile -Force

$StartingMessage = "
 _____                  
| ____|_ __  _ __   ___ 
|  _| | '_ \| '_ \ / _ \
| |___| | | | | | |  __/
|_____|_| |_|_| |_|\___|
                                                                                       
"
# Define the colors for the rainbow
$colors = @("Red", "Yellow", "Green", "Cyan", "Blue", "Magenta")

# Split the text into lines
$lines = $StartingMessage -split "`n"

# Display each line with a different color
for ($i = 0; $i -lt $lines.Length; $i++) {
    $color = $colors[$i % $colors.Length]
    Write-Host $lines[$i] -ForegroundColor $color
}

# Ensure script is run with elevated privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $errorMessage = "You need to run this script as an administrator."
    Write-Host $errorMessage
    Stop-Transcript
    exit 1
}

# Create the log directory if it doesn't exist
if (!(Test-Path -Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
    Log-Event -message "Log directory not found. A new directory has been created: $logPath" -type "Information" -WingetSoftware $WingetSoftware
}

# Set the registry path for the software
$registryPath = "HKLM:\Software\Enne\Packages\$FirstLetterSoftwareSupplier\$WingetSupplier\$WingetSoftware\Winget"
$valueName = "PackageVersion"

try {
    $startMessage = "Starting the removal process for $WingetSoftware with Winget."
    Log-Event -message $startMessage -type "Information" -WingetSoftware $WingetSoftware

    #############################################################
    # Custom Powershell script can be added between these 2 lines
    # This part is before the installation has started-
    #############################################################

    # Since Powershell is started from %SystemRoot%\sysnative\WindowsPowerShell\v1.0\powershell.exe env variables are present. Therefore, looking for Winget location
    try {
        $logMessage = "Trying to find Winget path"
        Log-Event -message  $logMessage -type "Information" -WingetSoftware $WingetSoftware

        $ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
        if ($ResolveWingetPath) {
            $logMessage = "Found it! Winget is located at: $ResolveWingetPath"
            Log-Event -message  $logMessage -type "Information" -WingetSoftware $WingetSoftware
            $WingetPath = $ResolveWingetPath[-1].Path
        }
    }

    catch {
        $errorMessage = "Couldn't find Winget location. Ensure Winget is installed"
        Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
        throw $errorMessage
    }

    # Construct the uninstall command
    $UninstalLine = "& `"$WingetPath\winget.exe`" uninstall --id $Winget --silent --disable-interactivity"

    $logMessage = "Using the following uninstall line: $UninstalLine"
    Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware

    try {
        $logMessage = "Starting the Winget part now!"
        Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware
        Invoke-Expression $UninstalLine
    } catch {
        $errorMessage = "Removal process for $WingetSoftware failed."
        Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
        throw $errorMessage
    }

    $logMessage = "Winget has concluded. Almost done now!"
    Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware


    # Test if the registry path exists and create it if it doesn't
    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Set or update the value of the version key in the registry
    # Check if the registry path exists
    if (Test-Path -Path $registryPath) {
        $property = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction SilentlyContinue
        if ($property) {
            Remove-ItemProperty -Path $registryPath -Name $valueName -Force | Out-Null -ErrorAction SilentlyContinue
            # Check if the registry key was successfully removed
            if (Test-Path $registryPath) {
                $logMessage = "Registry key for ${WingetSoftware} successfully removed at path: ${registryPath}"
                Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware
            } else {
                $errorMessage = "Failed to remove registry key for $WingetSoftware."
                Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
            }
        } else {
            $logMessage = "Property $valueName does not exist at path $registryPath"
            Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware
        }
    } else {
        $logMessage = "Registry path $registryPath does not exist"
        Log-Event -message $logMessage -type "Information" -WingetSoftware $WingetSoftware
    }

    $successMessage = "Removal for $WingetSoftware is done."
    Log-Event -message $successMessage -type "Success" -WingetSoftware $WingetSoftware
}
catch {
    # Log error to the Event Viewer
    $errorMessage = "$_"
    Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
}
finally {
    # Ensure the transcript is stopped even if an error occurs
    Stop-Transcript
}