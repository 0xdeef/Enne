$AppToDetect = ""

# Function to log events to the Event Viewer
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
$WingetSupplier = $AppToDetect -split '\.' | Select-Object -First 1
$WingetSoftware = $AppToDetect.Substring($AppToDetect.IndexOf('.') + 1)
$FirstLetterSoftwareSupplier = $WingetSupplier.Substring(0, 1).ToUpper()
$logPath = "$ENV:ProgramData\Enne\Logs\Packaging\$FirstLetterSoftwareSupplier\$WingetSupplier\$WingetSoftware\Winget"
$transcriptFile = "$logPath\Detection-$WingetSoftware.log"

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
    Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
    Stop-Transcript
    exit 1
}

# Create the log directory if it doesn't exist
if (!(Test-Path -Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
    Log-Event -message "Log directory not found. A new directory has been created: $logPath" -type "Information" -WingetSoftware $WingetSoftware
}

try {
    $startMessage = "Starting the detection process for $WingetSoftware with Winget."
    Log-Event -message $startMessage -type "Information" -WingetSoftware $WingetSoftware

    # Get WinGet Location Function
    function Get-WingetCmd {
        $WingetCmd = $null

        # Get Admin Context Winget Location
        try {
            $WingetInfo = (Get-Item "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe").VersionInfo | Sort-Object -Property FileVersionRaw
            $WingetCmd = $WingetInfo[-1].FileName
        } catch {
            # Get User context Winget Location
            if (Test-Path "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe") {
                $WingetCmd = "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
            }
        }
        $logMessage = "Winget is located at $WingetCmd"
        Log-Event -message  $logMessage -type "Information" -WingetSoftware $WingetSoftware
        return $WingetCmd
    }

    $winget = Get-WingetCmd

    # Set json export file
    $JsonFile = "$env:TEMP\InstalledApps.json"

    # Get installed apps and version in json file
    & $Winget export -o $JsonFile --accept-source-agreements | Out-Null

    # Get json content
    $Json = Get-Content $JsonFile -Raw | ConvertFrom-Json

    # Get apps and version in hashtable
    $Packages = $Json.Sources.Packages

    # Remove json file
    Remove-Item $JsonFile -Force

    # Search for specific app and version
    $Apps = $Packages | Where-Object { $_.PackageIdentifier -eq $AppToDetect }

    if ($Apps) {
        $logMessage = "$WingetSoftware is installed."
        Log-Event -message $logMessage -type "Success" -WingetSoftware $WingetSoftware
        exit 0
    } else {
        $errorMessage = "$WingetSoftware is not installed."
        Log-Event -message $errorMessage -type "Information" -WingetSoftware $WingetSoftware
        exit 1
    }
} catch {
    $errorMessage = "$_"
    Log-Event -message $errorMessage -type "Error" -WingetSoftware $WingetSoftware
    exit 1
} finally {
    # Ensure the transcript is stopped even if an error occurs
    Stop-Transcript
}