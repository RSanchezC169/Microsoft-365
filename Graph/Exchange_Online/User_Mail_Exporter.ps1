##################################################################################################################################################################
##################################################################################################################################################################
<#
                 _..--+~/@-~--.
             _-=~      (  .   "}
          _-~     _.--=.\ \""""
        _~      _-       \ \_\
       =      _=          '--'
      '      =                             .
     :      :       ____                   '=_. ___
___  |      ;                            ____ '~--.~.
     ;      ;                               _____  } |
  ___=       \ ___ __     __..-...__           ___/__/__
     :        =_     _.-~~          ~~--.__
_____ \         ~-+-~                   ___~=_______
     ~@#~~ == ...______ __ ___ _--~~--_
                                                    =
██████╗ ███████╗ █████╗ ███╗   ██╗ ██████╗██╗  ██╗███████╗███████╗ ██████╗ ██╗ ██████╗ █████╗ 
██╔══██╗██╔════╝██╔══██╗████╗  ██║██╔════╝██║  ██║██╔════╝╚══███╔╝██╔════╝███║██╔════╝██╔══██╗
██████╔╝███████╗███████║██╔██╗ ██║██║     ███████║█████╗    ███╔╝ ██║     ╚██║███████╗╚██████║
██╔══██╗╚════██║██╔══██║██║╚██╗██║██║     ██╔══██║██╔══╝   ███╔╝  ██║      ██║██╔═══██╗╚═══██║
██║  ██║███████║██║  ██║██║ ╚████║╚██████╗██║  ██║███████╗███████╗╚██████╗ ██║╚██████╔╝█████╔╝
╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝ ╚═════╝ ╚═╝ ╚═════╝ ╚════╝ 
Script Version: 2
OS Version Script was written on: Microsoft Windows 11 Pro : 10.0.25100 Build 26100
PSVersion 5.1.26100.2161 : PSEdition Desktop : Build Version 10.0.26100.2161
Description of Script: 
NOTE: This script allows you to connect to Microsoft Graph either by delegate or application. When connecting via delegate you can only export emails from the mailbox you authenticated to.
To export emails from any mailbox connect to Microsoft Graph via application.

This PowerShell script is designed to automate email export operations by interacting with Microsoft Graph API. 
It provides an interactive menu system to execute various tasks, including connecting to Microsoft Graph, retrieving emails with various filters, and exporting them to a structured file system. 
It emphasizes logging, validation, and an enhanced user experience through a cleanly defined menu-driven interface.
#>
##################################################################################################################################################################
#==============================Beginning of script======================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================Functions==============================================================================================================================
##################################################################################################################################################################
# Define the global log file path at the start of the script
$Global:LogFile = "C:\Rsanchezc169ScriptLogs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"

Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,                  # The message to log
        [Parameter()]
        [ValidateSet("INFO", "WARNING", "ERROR","DEBUG")]
        [string]$Level = "INFO",           # Log level: INFO, WARNING, DEBUG, or ERROR
        [Parameter()]
        [string]$LogFile = $Global:LogFile # Optional: Specify a log file, default to the global log file
    )

    # Ensure the log directory exists
    try {
        $LogDirectory = Split-Path -Path $LogFile -ErrorAction SilentlyContinue
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
            #Write-Host "Log directory created: $LogDirectory" -ForegroundColor Yellow
        }
    } catch {
        #Write-Warning "Failed to create log directory: $($_.Exception.Message)"
        throw
    }

    # Append the message to the log file with timestamp and level
    try {
        $Timestamp = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt")
        Add-Content -Path $LogFile -Value "$Timestamp [$Level] : $Message"
        if ($Level -eq "ERROR") {
            #Write-Warning "Logged error: $Message"
        } elseif ($Level -eq "WARNING") {
            #Write-Host "Logged warning: $Message" -ForegroundColor Yellow
        } else {
            #Write-Host "Logged message: $Message" -ForegroundColor Green
        }
    } catch {
        #Write-Warning "Failed to write log message: $($_.Exception.Message)"
    }
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Load-Module {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Modules  # Allows processing multiple modules at once
    )

    Write-Log -Message "Started Function Load-Module" -Level "INFO"

    foreach ($Module in $Modules) {
        Write-Progress -Activity "Processing Module: $Module" -Status "Initializing..." -PercentComplete 0

        try {
            # Step 1: Check if the module is installed
            Write-Log -Message "Checking if module '$Module' is installed..." -Level "INFO"
            if ((Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue) -or (Get-Module -Name $Module -ErrorAction SilentlyContinue)) {
                Write-Log -Message "Module '$Module' is already installed." -Level "INFO"
            } else {
                Write-Log -Message "Module '$Module' is not installed. Attempting to install..." -Level "WARNING"
                Install-Module -Name $Module -Force -Scope CurrentUser -ErrorAction Stop
                Write-Log -Message "Module '$Module' installed successfully." -Level "INFO"
            }

            # Step 2: Check for module updates
            Write-Log -Message "Checking for updates to module '$Module'..." -Level "INFO"
            $CurrentVersion = (Find-Module -Name $Module -ErrorAction SilentlyContinue).Version
            $InstalledVersion = (Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue).Version
            if ($CurrentVersion -and $InstalledVersion -and $CurrentVersion -gt $InstalledVersion) {
                Write-Log -Message "Updating module '$Module' to version $CurrentVersion..." -Level "INFO"
                Update-Module -Name $Module -Force -ErrorAction Stop
                Write-Log -Message "Module '$Module' updated successfully to version $CurrentVersion." -Level "INFO"
            } else {
                Write-Log -Message "Module '$Module' is up-to-date." -Level "INFO"
            }

            # Step 3: Import the module
            Write-Log -Message "Importing module '$Module'..." -Level "INFO"
            Import-Module -Name $Module -Force -ErrorAction Stop
            Write-Log -Message "Module '$Module' imported successfully." -Level "INFO"
        } catch {
            # Log any errors that occur while processing the module
            Write-Log -Message "Error processing module '$Module': $($_.Exception.Message)" -Level "ERROR"
        }
    }

    # Clear progress
    Write-Progress -Activity "Module Processing" -Status "Complete" -PercentComplete 100 -Completed

    Write-Log -Message "Ended Function Load-Module" -Level "INFO"
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Set-Environment {
    Write-Log -Message "Started Function Set-Environment" -Level "INFO"

    # Initialize progress bar
    Write-Progress -Activity "Environment Setup" -Status "Starting setup..." -PercentComplete 0 -ErrorAction SilentlyContinue

    try {
        # Step 1: Clear the console
        Write-Progress -Activity "Environment Setup" -Status "Clearing console..." -PercentComplete 10 -ErrorAction SilentlyContinue
        Clear-Host

        # Step 2: Maximize console window
        Write-Progress -Activity "Environment Setup" -Status "Maximizing console window..." -PercentComplete 20 -ErrorAction SilentlyContinue
        Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            public class User32 {
                [DllImport("user32.dll")]
                public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            }
"@
        $handle = (Get-Process -Id $PID).MainWindowHandle
        [User32]::ShowWindow($handle, 3)
        Write-Log -Message "Console window maximized" -Level "INFO"

        # Step 3: Set execution policy
        Write-Progress -Activity "Environment Setup" -Status "Setting execution policy..." -PercentComplete 30 -ErrorAction SilentlyContinue
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force
        Write-Log -Message "Execution policy set to 'Unrestricted'" -Level "INFO"

        # Step 4: Load required modules
        Write-Progress -Activity "Environment Setup" -Status "Loading required modules..." -PercentComplete 50 -ErrorAction SilentlyContinue
        Load-Module -Modules "PowerShellGet", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Mail", "Microsoft.Graph.Beta.Mail"
        Write-Log -Message "Required modules loaded successfully" -Level "INFO"

        # Step 5: Configure console appearance
        Write-Progress -Activity "Environment Setup" -Status "Configuring console appearance..." -PercentComplete 60 -ErrorAction SilentlyContinue
        $Host.UI.RawUI.BackgroundColor = 'Black'
        $Host.UI.RawUI.ForegroundColor = 'Blue'
        $Host.UI.RawUI.WindowTitle = "Export Emails"
        Write-Log -Message "Console appearance configured: Background=Black, Foreground=Blue, Title='Export Emails'" -Level "INFO"

        # Step 6: Set session preferences
        Write-Progress -Activity "Environment Setup" -Status "Configuring session preferences..." -PercentComplete 70 -ErrorAction SilentlyContinue
        $Global:FormatEnumerationLimit = -1
        $Global:ErrorActionPreference = "SilentlyContinue"
        $Global:WarningActionPreference = "SilentlyContinue"
        $Global:InformationActionPreference = "SilentlyContinue"
        Write-Log -Message "Session preferences configured" -Level "INFO"

        # Step 7: Finalize setup
        Write-Progress -Activity "Environment Setup" -Status "Setup complete." -PercentComplete 100 -ErrorAction SilentlyContinue
        Write-Progress -Activity "Environment Setup" -Completed -ErrorAction SilentlyContinue
        Write-Log -Message "Environment setup completed successfully" -Level "INFO"
    } catch {
        Write-Log -Message "Error during environment setup: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Clean up resources
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    Write-Log -Message "Ended Function Set-Environment" -Level "INFO"
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Check-Email {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EmailAddress  # The email address to validate
    )

    Write-Log -Message "Started Function Check-Email" -Level "INFO"

    # Define the email validation regex pattern
    $pattern = '^(?:[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\"[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|

\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f])+)\]

)$'

    # Perform email validation
    try {
        if ($EmailAddress -match $pattern) {
            # Log valid email
            Write-Log -Message "Valid email address: $EmailAddress" -Level "INFO"
            Write-Log -Message "Completed Function Check-Email with success" -Level "INFO"
            return $true
        } else {
            # Log invalid email
            Write-Log -Message "Invalid email address: $EmailAddress" -Level "WARNING"
            Write-Log -Message "Completed Function Check-Email with failure (invalid email)" -Level "INFO"
            return $false
        }
    } catch {
        # Log unexpected errors
        Write-Log -Message "Error validating email address: $($_.Exception.Message)" -Level "ERROR"
        throw "Error validating email address: $($_.Exception.Message)"
    } finally {
        # Clean-up and log function end
        Write-Log -Message "Ended Function Check-Email" -Level "INFO"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Validate-Email {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$EmailAddress  # The email address to validate
    )

    Write-Log -Message "Started Function Validate-Email" -Level "INFO"

    # Define the email validation regex pattern
    $EmailPattern = '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'

    try {
        # Perform email validation
        if ($EmailAddress -match $EmailPattern) {
            Write-Log -Message "Valid email address: $EmailAddress" -Level "INFO"
            return $true
        } else {
            Write-Log -Message "Invalid email address: $EmailAddress" -Level "WARNING"
            Write-Warning "Invalid email address format: $EmailAddress"
            return $false
        }
    } catch {
        # Log any unexpected errors
        Write-Log -Message "Error validating email address: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        Write-Log -Message "Ended Function Validate-Email" -Level "INFO"

        # Clean up
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-DateRange {
    Write-Log -Message "Started Function Get-DateRange" -Level "INFO"

    # Initialize variables
    $StartDate = $null
    $EndDate = $null

    do {
        try {
            # Prompt for and validate the start date
            Write-Log -Message "Prompting for Start Date..." -Level "INFO"
            $StartDateInput = Read-Host "Please enter a valid Start Date (e.g., MM/dd/yyyy)"
            $StartDate = [datetime]::Parse($StartDateInput)

            # Prompt for and validate the end date
            Write-Log -Message "Prompting for End Date..." -Level "INFO"
            $EndDateInput = Read-Host "Please enter a valid End Date (e.g., MM/dd/yyyy)"
            $EndDate = [datetime]::Parse($EndDateInput)

            # Validate that the start date is not after the end date
            if ($StartDate -gt $EndDate) {
                Write-Log -Message "Start date ($StartDate) is after the end date ($EndDate). Prompting user to re-enter dates." -Level "WARNING"
                throw "Start date cannot be later than the end date."
            }

            # Validate that the end date is not in the future
            if ($EndDate -gt (Get-Date).AddDays(1)) {
                Write-Log -Message "End date ($EndDate) is in the future. Prompting user to re-enter dates." -Level "WARNING"
                throw "End date cannot be in the future."
            }

            # Log successful validation
            Write-Log -Message "Successfully validated Start Date: $StartDate and End Date: $EndDate" -Level "INFO"
            break
        } catch {
            # Log and reset values on error
            Write-Log -Message "Error validating date range: $($_.Exception.Message)" -Level "ERROR"
            $StartDate = $null
            $EndDate = $null
        }
    } while (-not ($StartDate -and $EndDate)) # Loop until valid dates are entered

    # Return the validated date range
    $DateRange = @{
        StartDate = $StartDate
        EndDate   = $EndDate
    }

    Write-Log -Message "Completed Function Get-DateRange with StartDate: $($DateRange.StartDate), EndDate: $($DateRange.EndDate)" -Level "INFO"
    return $DateRange
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Draw-Line {
    Write-Log -Message "Started Function Draw-Line" -Level "INFO"

    try {
        # Retrieve the console window width
        Write-Log -Message "Retrieving the console window width..." -Level "INFO"
        $LineWidth = (Get-Host).UI.RawUI.WindowSize.Width

        # Generate the line
        $Line = '=' * $LineWidth
        Write-Log -Message "Line generated with width: $LineWidth" -Level "INFO"

        # Log the line itself
        Write-Log -Message "Generated Line: $Line" -Level "DEBUG"

        # Return the line
        return $Line
    } catch {
        # Log the error if window width retrieval or line generation fails
        Write-Log -Message "Error during Draw-Line execution: $($_.Exception.Message)" -Level "ERROR"
        throw "Failed to generate the line due to an error: $($_.Exception.Message)"
    } finally {
        # Log function completion
        Write-Log -Message "Ended Function Draw-Line" -Level "INFO"

        # Clean up resources
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Validate-Information {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Information  # The information to validate
    )

    Write-Log -Message "Started Function Validate-Information" -Level "INFO"

    do {
        try {
            # Log the provided information
            Write-Log -Message "User provided the following information: $Information" -Level "INFO"

            # Validate the information
            Write-Log -Message "Prompting user to confirm the information is correct..." -Level "INFO"
            $ValidationResponse = Read-Host "The value you inputted is '$Information'. Is this correct? [Y/N]"
            $ValidationResponse = $ValidationResponse.ToUpper()

            Switch ($ValidationResponse) {
                "Y" {
                    Write-Log -Message "User confirmed the information as correct: $Information" -Level "INFO"
                    Write-Log -Message "Completed Function Validate-Information successfully" -Level "INFO"
                    return $Information
                }
                "N" {
                    Write-Log -Message "User rejected the information. Prompting for re-entry..." -Level "WARNING"
                    $Information = Read-Host "Enter new value"
                    
                    # Ensure the new value is not empty
                    while ([string]::IsNullOrWhiteSpace($Information)) {
                        Write-Log -Message "User entered an empty value. Prompting for re-entry..." -Level "ERROR"
                        $Information = Read-Host "Enter new value (cannot be empty)"
                    }

                    Write-Log -Message "User updated the information to: $Information" -Level "INFO"
                }
                Default {
                    Write-Log -Message "Invalid validation response: $ValidationResponse. Prompting user again." -Level "WARNING"
                }
            }
        } catch {
            Write-Log -Message "An error occurred during validation: $($_.Exception.Message)" -Level "ERROR"
            throw "Error in Validate-Information: $($_.Exception.Message)"
        }
    } while ($ValidationResponse -ne "Y")

    Write-Log -Message "Ended Function Validate-Information" -Level "INFO"
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Get-Emails {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$User,        # User's email or ID
        [Parameter()]
        [datetime]$StartDate, # Optional Start Date
        [Parameter()]
        [datetime]$EndDate,   # Optional End Date
        [Parameter()]
        [string]$SubjectLine, # Optional Subject Line
        [Parameter()]
        [string]$Sender,      # Optional Sender's email address
        [Parameter()]
        [string]$Receiver,    # Optional Receiver's email address
        [Parameter()]
        [string]$Domain       # Optional Domain
    )

    Write-Log -Message "Started Function Get-Emails" -Level "INFO"

    try {
        Write-Log -Message "Checking if there are any emails for the user: $User" -Level "INFO"

        # Check for the existence of at least one email
        $FirstEmail = Get-MgUserMessage -UserId $User -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($FirstEmail) {
            Write-Log -Message "Emails found for the user: $User. Proceeding with retrieval." -Level "INFO"

            # Build filter string based on parameters
            $FilterParts = @()
            if ($PSBoundParameters.ContainsKey('StartDate') -and $PSBoundParameters.ContainsKey('EndDate')) {
                $FilterParts += "receivedDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and receivedDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                Write-Log -Message "Filter added: Date range from $StartDate to $EndDate" -Level "INFO"
            } elseif ($PSBoundParameters.ContainsKey('StartDate')) {
                $FilterParts += "receivedDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                Write-Log -Message "Filter added: Start date is $StartDate" -Level "INFO"
            } elseif ($PSBoundParameters.ContainsKey('EndDate')) {
                $FilterParts += "receivedDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                Write-Log -Message "Filter added: End date is $EndDate" -Level "INFO"
            }

            if ($PSBoundParameters.ContainsKey('SubjectLine')) {
                $FilterParts += "contains(subject, '$SubjectLine')"
                Write-Log -Message "Filter added: Subject contains '$SubjectLine'" -Level "INFO"
            }

            if ($PSBoundParameters.ContainsKey('Sender')) {
                $FilterParts += "from/emailAddress/address eq '$Sender'"
                Write-Log -Message "Filter added: Sender is '$Sender'" -Level "INFO"
            }

            if ($PSBoundParameters.ContainsKey('Receiver')) {
                $FilterParts += "toRecipients/any(t: t/emailAddress/address eq '$Receiver')"
                Write-Log -Message "Filter added: Receiver is '$Receiver'" -Level "INFO"
            }

            if ($PSBoundParameters.ContainsKey('Domain')) {
                $FilterParts += "contains(from/emailAddress/address, '$Domain')"
                Write-Log -Message "Filter added: Domain is '$Domain'" -Level "INFO"
            }

            $Filter = $FilterParts -join " and "

            # Retrieve emails with the constructed filter
            Write-Log -Message "Retrieving emails with the filter: $Filter" -Level "INFO"
            $AllEmails = if ($Filter) {
                Get-MgUserMessage -UserId $User -All -Filter $Filter -ErrorAction Stop
            } else {
                Get-MgUserMessage -UserId $User -All -ErrorAction Stop
            }

            if ($AllEmails) {
                Write-Log -Message "Successfully retrieved emails for user: $User" -Level "INFO"
            } else {
                Write-Log -Message "No emails found after applying filters for user: $User" -Level "WARNING"
            }

            return $AllEmails
        } else {
            Write-Log -Message "No emails found for user: $User" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log -Message "Error retrieving emails for user $($User)): $($_.Exception.Message)" -Level "ERROR"
        throw "Failed to retrieve emails for user: $User. Error: $($_.Exception.Message)"
    } finally {
        Write-Log -Message "Ended Function Get-Emails" -Level "INFO"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Export-Email {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [array]$Emails,        # Array of email message objects
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath,   # Folder path where emails will be saved
        [Parameter(Mandatory)]
        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]  # Basic validation for email format
        [string]$User          # The email address of the user whose messages are being exported
    )

    Write-Log -Message "Started Function Export-Email" -Level "INFO"

    try {
        # Validate the provided folder path
        if (-not (Test-Path -Path $FolderPath)) {
            Write-Log -Message "The specified folder path does not exist: $FolderPath" -Level "ERROR"
            throw "The specified folder path does not exist: $FolderPath"
        }

        # Start processing emails
        $TotalEmails = $Emails.Count
        Write-Log -Message "Preparing to export $TotalEmails emails for user $User" -Level "INFO"

        $Counter = 0
        foreach ($Message in $Emails) {
            try {
                $Counter++
                $PercentComplete = [math]::Min((($Counter / $TotalEmails) * 100), 100)

                # Update Progress Bar
                Write-Progress -Activity "Exporting Emails" `
                               -Status "Processing email $Counter of $TotalEmails" `
                               -PercentComplete $PercentComplete

                # Generate a valid file name for the email
                $BaseFileName = ($Message.Subject + ".eml") -replace '[<>:"/\\|?*]', '_'
                $OutFile = Join-Path -Path $FolderPath -ChildPath $BaseFileName
                $Index = 1

                # Handle duplicate file names by appending an index
                while (Test-Path -Path $OutFile) {
                    $IndexedFileName = ($Message.Subject + "_$Index.eml") -replace '[<>:"/\\|?*]', '_'
                    $OutFile = Join-Path -Path $FolderPath -ChildPath $IndexedFileName
                    $Index++
                }

                # Export the email content to the file
                Write-Log -Message "Exporting email to file: $OutFile" -Level "INFO"
                $Global:ProgressPreference = "SilentlyContinue"
                Get-MgUserMessageContent -UserId $User -MessageId $Message.Id -OutFile $OutFile -ErrorAction Stop
                $Global:ProgressPreference = "Continue"
                Write-Log -Message "Email exported successfully to: $OutFile" -Level "INFO"
            } catch {
                # Log any errors for the current email
                Write-Log -Message "Error exporting email with subject '$($Message.Subject)': $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Complete the progress bar
        Write-Progress -Activity "Exporting Emails" `
                       -Status "Export Completed!" `
                       -PercentComplete 100 -Completed

        # Open the folder in File Explorer
        Write-Log -Message "Opening the export folder in File Explorer: $FolderPath" -Level "INFO"
        Start-Process -FilePath "explorer.exe" -ArgumentList $FolderPath
    } catch {
        Write-Log -Message "An error occurred during the email export process: $($_.Exception.Message)" -Level "ERROR"
        throw "Error during email export: $($_.Exception.Message)"
    } finally {
        Write-Log -Message "Completed Function Export-Email" -Level "INFO"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Create-MainFolder {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EmailAddress  # The email address to create a folder for
    )

    Write-Log -Message "Started Function Create-MainFolder" -Level "INFO"

    try {
        # Step 1: Extract folder name from the email address
        Write-Log -Message "Extracting folder name from email address: $EmailAddress" -Level "INFO"
        $FolderName = ($EmailAddress -split "@")[0] + "_Emails"
        Write-Log -Message "Folder name generated: $FolderName" -Level "INFO"

        # Step 2: Get the Downloads folder path
        Write-Log -Message "Retrieving Downloads folder path..." -Level "INFO"
        $MainFolderPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
        Write-Log -Message "Downloads folder path retrieved: $MainFolderPath" -Level "INFO"

        # Step 3: Create the main folder path
        $FullFolderPath = Join-Path -Path $MainFolderPath -ChildPath $FolderName
        Write-Log -Message "Full folder path generated: $FullFolderPath" -Level "INFO"

        # Step 4: Check if the folder already exists or create it
        if (Test-Path -Path $FullFolderPath) {
            Write-Log -Message "Folder already exists: $FullFolderPath" -Level "WARNING"
        } else {
            Write-Log -Message "Folder does not exist. Creating folder: $FullFolderPath" -Level "INFO"
            New-Item -ItemType Directory -Path $FullFolderPath -Force | Out-Null
            Write-Log -Message "Folder created successfully: $FullFolderPath" -Level "INFO"
        }

        # Step 5: Return the folder path
        Write-Log -Message "Returning folder path: $FullFolderPath" -Level "INFO"
        return $FullFolderPath
    } catch {
        # Log and rethrow any errors encountered during the process
        Write-Log -Message "Error occurred in Create-MainFolder: $($_.Exception.Message)" -Level "ERROR"
        throw "Failed to create main folder for email address $EmailAddress. Error: $($_.Exception.Message)"
    } finally {
        Write-Log -Message "Ended Function Create-MainFolder" -Level "INFO"
        # Clean up
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Create-SubFolder {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MainFolderPath,  # The path to the main folder
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SubFolderName    # The name of the sub-folder to create
    )

    Write-Log -Message "Started Function Create-SubFolder" -Level "INFO"

    # Construct the sub-folder path
    $SubFolderPath = Join-Path -Path $MainFolderPath -ChildPath $SubFolderName
    Write-Log -Message "Constructed Sub-Folder Path: $SubFolderPath" -Level "INFO"

    try {
        # Check if the sub-folder already exists
        if (Test-Path -Path $SubFolderPath) {
            Write-Log -Message "Sub-folder already exists: $SubFolderPath" -Level "WARNING"
        } else {
            # Create the sub-folder
            Write-Log -Message "Creating sub-folder: $SubFolderPath" -Level "INFO"
            New-Item -ItemType Directory -Path $SubFolderPath -Force | Out-Null
            Write-Log -Message "Sub-folder created successfully: $SubFolderPath" -Level "INFO"
        }

        # Return the sub-folder path
        Write-Log -Message "Returning Sub-Folder Path: $SubFolderPath" -Level "INFO"
        return $SubFolderPath
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error creating sub-folder: $($_.Exception.Message)" -Level "ERROR"
        throw "Failed to create sub-folder at path: $SubFolderPath. Error: $($_.Exception.Message)"
    } finally {
        # Log completion of the function
        Write-Log -Message "Ended Function Create-SubFolder" -Level "INFO"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-FilterCriteria {
    [CmdletBinding()]
    Param(
        [switch]$User,
        [switch]$StartDate,
        [switch]$EndDate,
        [switch]$SubjectLine,
        [switch]$Sender,
        [switch]$Receiver,
        [switch]$Domain
    )

    Write-Log -Message "Started Function Get-FilterCriteria" -Level "INFO"

    # Initialize a hashtable for criteria
    $Criteria = @{}

    try {
        # User Email Validation with Confirmation and Tenant User Check
        if ($User) {
            do {
                $InputUser = Read-Host "Enter User Email (e.g., user@example.com)"
                Write-Host "You entered: $InputUser" -ForegroundColor Yellow
                $ConfirmUser = Read-Host "Is this correct? Enter Y for Yes, N for No"

                if ($ConfirmUser -eq "Y") {
                    if (Validate-Email -EmailAddress $InputUser) {
                        Write-Log -Message "Validating if the user exists in the tenant: $InputUser" -Level "INFO"
                        try {
                            $TenantUser = Get-MgUser -Filter "userPrincipalName eq '$InputUser'" -ErrorAction Stop
                            if ($TenantUser) {
                                Write-Log -Message "User found in the tenant: $InputUser" -Level "INFO"
                                Write-Host "User exists in the tenant: $InputUser" -ForegroundColor Green
                                $Criteria["User"] = $InputUser
                                break
                            }
                        } catch {
                            Write-Log -Message "User not found in the tenant: $InputUser" -Level "WARNING"
                            Write-Host "The email address $InputUser does not belong to a user in the tenant. Please try again." -ForegroundColor Red
                        }
                    } else {
                        Write-Log -Message "Invalid email format: $InputUser" -Level "WARNING"
                        Write-Host "Invalid email format. Please try again." -ForegroundColor Red
                    }
                } elseif ($ConfirmUser -eq "N") {
                    Write-Host "Please re-enter the email." -ForegroundColor Cyan
                } else {
                    Write-Host "Invalid confirmation input. Please enter Y or N." -ForegroundColor Red
                }
            } while ($true)
        }
        # Start and End Date Validation with Confirmation
        if ($StartDate -and $EndDate) {
            do {
                try {
                    $StartDateInput = Read-Host "Enter Start Date (MM/dd/yyyy)"
                    Write-Host "You entered: $StartDateInput" -ForegroundColor Yellow
                    $ConfirmStartDate = Read-Host "Is this correct? Enter Y for Yes, N for No"

                    $EndDateInput = Read-Host "Enter End Date (MM/dd/yyyy)"
                    Write-Host "You entered: $EndDateInput" -ForegroundColor Yellow
                    $ConfirmEndDate = Read-Host "Is this correct? Enter Y for Yes, N for No"

                    if ($ConfirmStartDate -eq "Y" -and $ConfirmEndDate -eq "Y") {
                        $StartDateObj = [datetime]::Parse($StartDateInput)
                        $EndDateObj = [datetime]::Parse($EndDateInput)

                        if ($StartDateObj -le $EndDateObj) {
                            Write-Log -Message "Valid date range confirmed: $StartDateObj to $EndDateObj" -Level "INFO"
                            $Criteria["StartDate"] = $StartDateObj
                            $Criteria["EndDate"] = $EndDateObj
                            break
                        } else {
                            Write-Log -Message "Invalid date range: StartDate=$StartDateObj is later than EndDate=$EndDateObj" -Level "WARNING"
                            Write-Host "Invalid date range. Start Date must be earlier than End Date. Please try again." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "Please re-enter the date range." -ForegroundColor Cyan
                    }
                } catch {
                    Write-Log -Message "Error parsing dates: $($_.Exception.Message)" -Level "ERROR"
                }
            } while ($true)
        }

        # Subject Line Validation with Confirmation
        if ($SubjectLine) {
            do {
                $InputSubject = Read-Host "Enter Subject Line"
                Write-Host "You entered: $InputSubject" -ForegroundColor Yellow
                $ConfirmSubject = Read-Host "Is this correct? Enter Y for Yes, N for No"

                if ($ConfirmSubject -eq "Y") {
                    if (-not [string]::IsNullOrWhiteSpace($InputSubject)) {
                        Write-Log -Message "Valid subject line confirmed: $InputSubject" -Level "INFO"
                        $Criteria["SubjectLine"] = $InputSubject
                        break
                    } else {
                        Write-Log -Message "Empty subject line entered." -Level "WARNING"
                        Write-Host "Subject line cannot be empty. Please try again." -ForegroundColor Red
                    }
                } elseif ($ConfirmSubject -eq "N") {
                    Write-Host "Please re-enter the subject line." -ForegroundColor Cyan
                } else {
                    Write-Host "Invalid confirmation input. Please enter Y or N." -ForegroundColor Red
                }
            } while ($true)
        }
        # Sender Email Validation with Confirmation
        if ($Sender) {
            do {
                $InputSender = Read-Host "Enter Sender Email (e.g., sender@example.com)"
                Write-Host "You entered: $InputSender" -ForegroundColor Yellow
                $ConfirmSender = Read-Host "Is this correct? Enter Y for Yes, N for No"

                if ($ConfirmSender -eq "Y") {
                    if (Validate-Email -EmailAddress $InputSender) {
                        Write-Log -Message "Valid sender email confirmed: $InputSender" -Level "INFO"
                        $Criteria["Sender"] = $InputSender
                        break
                    } else {
                        Write-Log -Message "Invalid sender email: $InputSender" -Level "WARNING"
                        Write-Host "Invalid email format. Please try again." -ForegroundColor Red
                    }
                } elseif ($ConfirmSender -eq "N") {
                    Write-Host "Please re-enter the sender email." -ForegroundColor Cyan
                } else {
                    Write-Host "Invalid confirmation input. Please enter Y or N." -ForegroundColor Red
                }
            } while ($true)
        }

        # Receiver Email Validation with Confirmation
        if ($Receiver) {
            do {
                $InputReceiver = Read-Host "Enter Receiver Email (e.g., receiver@example.com)"
                Write-Host "You entered: $InputReceiver" -ForegroundColor Yellow
                $ConfirmReceiver = Read-Host "Is this correct? Enter Y for Yes, N for No"

                if ($ConfirmReceiver -eq "Y") {
                    if (Validate-Email -EmailAddress $InputReceiver) {
                        Write-Log -Message "Valid receiver email confirmed: $InputReceiver" -Level "INFO"
                        $Criteria["Receiver"] = $InputReceiver
                        break
                    } else {
                        Write-Log -Message "Invalid receiver email: $InputReceiver" -Level "WARNING"
                        Write-Host "Invalid email format. Please try again." -ForegroundColor Red
                    }
                } elseif ($ConfirmReceiver -eq "N") {
                    Write-Host "Please re-enter the receiver email." -ForegroundColor Cyan
                } else {
                    Write-Host "Invalid confirmation input. Please enter Y or N." -ForegroundColor Red
                }
            } while ($true)
        }
        # Domain Validation with Confirmation
        if ($Domain) {
            do {
                $InputDomain = Read-Host "Enter Domain (e.g., example.com)"
                Write-Host "You entered: $InputDomain" -ForegroundColor Yellow
                $ConfirmDomain = Read-Host "Is this correct? Enter Y for Yes, N for No"

                if ($ConfirmDomain -eq "Y") {
                    if ($InputDomain -match '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        Write-Log -Message "Valid domain confirmed: $InputDomain" -Level "INFO"
                        $Criteria["Domain"] = $InputDomain
                        break
                    } else {
                        Write-Log -Message "Invalid domain: $InputDomain" -Level "WARNING"
                        Write-Host "Invalid domain format. Please try again." -ForegroundColor Red
                    }
                } elseif ($ConfirmDomain -eq "N") {
                    Write-Host "Please re-enter the domain." -ForegroundColor Cyan
                } else {
                    Write-Host "Invalid confirmation input. Please enter Y or N." -ForegroundColor Red
                }
            } while ($true)
        }

        # Finalizing and Returning Criteria
        Write-Log -Message "Filter criteria collected: $($Criteria | Out-String)" -Level "INFO"
        return $Criteria
    } catch {
        # Log errors and rethrow
        Write-Log -Message "Error in Get-FilterCriteria: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Cleanup actions
        Write-Log -Message "Ended Function Get-FilterCriteria" -Level "INFO"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Show-Menu {
    Write-Log -Message "Started Function Show-Menu" -Level "INFO"

    try {
        # Clear the console and prepare the menu
        Clear-Host
        [System.Console]::Clear()
        Write-Log -Message "Console cleared successfully." -Level "INFO"

        # Draw a visual divider for the menu
        Write-Log -Message "Drawing menu separator line..." -Level "INFO"
        Draw-Line

        # Check connection status and display relevant information
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        if ($GraphContext) {
            if ($GraphContext.AuthType -eq "Delegated") {
                Write-Log -Message "Delegated connection detected for account: $($GraphContext.Account)" -Level "INFO"
                Write-Host "Connected to Microsoft Graph!" -ForegroundColor Green
                Write-Host "Account: $($GraphContext.Account)" -ForegroundColor Green
                Write-Host "Auth Type: Delegated" -ForegroundColor Green
            } elseif ($GraphContext.AuthType -eq "AppOnly" -and $GraphContext.TenantId -and $GraphContext.ClientId) {
                Write-Log -Message "App-only connection detected with Tenant ID: $($GraphContext.TenantId)" -Level "INFO"
                Write-Host "Connected via Application (App-Only Authentication)." -ForegroundColor Green
                Write-Host "You can export emails from any user mailbox." -ForegroundColor Green
            }
        } else {
            Write-Log -Message "No active Microsoft Graph connection detected." -Level "WARNING"
            Write-Host "Not connected to Microsoft Graph. Run Option 1: Connect to Microsoft Graph." -ForegroundColor Red
        }

        # Display the menu options
        Write-Log -Message "Displaying menu options to the user." -Level "INFO"
Draw-Line
Write-Host " __  __"  -ForeGroundColor Green                  
Write-Host "|  \/  |"  -ForeGroundColor Green                  
Write-Host "| \  / | ___ _ __  _   _ "  -ForeGroundColor Green 
Write-Host "| |\/| |/ _ \ '_ \| | | |"  -ForeGroundColor Green 
Write-Host "| |  | |  __/ | | | |_| |"  -ForeGroundColor Green 
Write-Host "|_|  |_|\___|_| |_|\__,_|"  -ForeGroundColor Green
Draw-Line
        $menu = @"
1 => Connect to Microsoft Graph
2 => Disconnect from Microsoft Graph
3 => Export All Emails for One User [Dump Format]
4 => Export All Emails in Date Range for One User [Dump Format]
5 => Export All Emails by Sender for One User [Dump Format]
6 => Export All Emails by Subject Line for One User [Dump Format]
7 => Export All Emails by Recipient for One User [Dump Format]
8 => Export All Emails by Domain for One User [Dump Format]
Q => Quit
--------------------------------------------------------------------------------
Select a task by entering its number or Q to quit:
"@
 $choice = Read-Host $menu
    } catch {
        # Handle any unexpected errors in menu initialization
        Write-Log -Message "Error initializing menu: $($_.Exception.Message)" -Level "ERROR"
        throw
    }

Switch ($choice) {
"1" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 1: Connect to Microsoft Graph" -Level "INFO"

    try {
        # Step 1: Check if already connected to Microsoft Graph
        Write-Log -Message "Checking if user is already connected to Microsoft Graph..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "User is already connected to Microsoft Graph with Auth Type: $($GraphContext.AuthType)" -Level "INFO"
            Write-Log -Message "Connected as: $($GraphContext.Account)" -Level "INFO"
            Write-Log -Message "Connection method: $($GraphContext.AuthType)" -Level "INFO"
        } else {
            # Prompt user to select an authentication method
            Write-Log -Message "User not connected to Microsoft Graph. Prompting for authentication method..." -Level "INFO"
            $AuthMethod = Read-Host "Do you want to connect using (A)pplication Authentication[All Mailboxes Export] or (D)elegated Authentication[Only Auth Mailbox Export]? [A/D]"

            Switch ($AuthMethod.ToUpper()) {
                "A" {
                    Write-Log -Message "Application Authentication selected." -Level "INFO"

                    # Gather Application Authentication details
                    $ClientId = Read-Host "Enter the Client ID"
                    $TenantId = Read-Host "Enter the Tenant ID"
                    $CertThumbprint = Read-Host "Enter the Certificate Thumbprint"

                    if (-not [string]::IsNullOrWhiteSpace($ClientId) -and -not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($CertThumbprint)) {
                        Write-Log -Message "Attempting Application Authentication with Client ID: $ClientId, Tenant ID: $TenantId" -Level "INFO"

                        # Connect using application authentication
                        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                        Write-Log -Message "Successfully connected to Microsoft Graph using Application Authentication." -Level "INFO"

                        # Verify Permissions
                        if ((Get-MgContext).Scopes | Where-Object { $_ -eq "Mail.ReadBasic.All" }) {
                            Write-Log -Message "Connection verified: Permission 'Mail.ReadBasic.All' is granted." -Level "INFO"
                        } else {
                            Write-Log -Message "Warning: 'Mail.ReadBasic.All' permission is missing. Please ensure it is added to the application." -Level "WARNING"
                        }
                    } else {
                        Write-Log -Message "Application Authentication failed: One or more required fields are missing." -Level "ERROR"
                    }
                }
                "D" {
                    Write-Log -Message "Delegated Authentication selected." -Level "INFO"

                    # Define Delegated Authentication scopes
                    $Scopes = @(
                        "User.ReadWrite.All",
                        "email",
                        "Mail.ReadBasic"
                    )
                    Write-Log -Message "Attempting Delegated Authentication with scopes: $($Scopes -join ', ')" -Level "INFO"

                    # Connect using delegated authentication
                    Connect-MgGraph -Scope $Scopes -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    Write-Log -Message "Successfully connected to Microsoft Graph using Delegated Authentication." -Level "INFO"
                }
                Default {
                    Write-Log -Message "Invalid selection for authentication method. User input: $AuthMethod" -Level "WARNING"
                }
            }
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error connecting to Microsoft Graph: $($_.Exception.Message)" -Level "ERROR"
    }

    # Prompt user to return to menu
    Write-Log -Message "Connection attempt complete. Prompting user to return to menu." -Level "INFO"
    Read-Host "Press [Enter] to reload the menu"
    Clear-Host
    [System.Console]::Clear()
    Show-Menu
}
"2" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 2: Disconnect from Microsoft Graph" -Level "INFO"

    try {
        # Check if currently connected to Microsoft Graph
        Write-Log -Message "Checking Microsoft Graph connection status..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            # Attempt to disconnect from Microsoft Graph
            Write-Log -Message "Disconnecting from Microsoft Graph..." -Level "INFO"
            Disconnect-MgGraph -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            Write-Log -Message "Successfully disconnected from Microsoft Graph." -Level "INFO"
            Write-Host "Successfully disconnected from Microsoft Graph." -ForegroundColor Green
        } else {
            # Handle scenario where no connection exists
            Write-Log -Message "No active Microsoft Graph connection detected." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Nothing to disconnect." -ForegroundColor Yellow
        }
    } catch {
        # Log errors during disconnection
        Write-Log -Message "Error disconnecting from Microsoft Graph: $($_.Exception.Message)" -Level "ERROR"
        Write-Host "An error occurred while disconnecting from Microsoft Graph. Check logs for more details." -ForegroundColor Red
    } finally {
        # Return to the menu
        Write-Log -Message "Disconnection process complete. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"3" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 3: Export All Emails for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking connection to Microsoft Graph SDK..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve user mailbox
                Write-Log -Message "Prompting user to provide mailbox for email export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User

                if ($UserMailbox) {
                    # Step 3: Create export folder
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails
                        Write-Log -Message "Retrieving emails for $UserMailbox..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully. Beginning export process..." -Level "INFO"

                            # Step 5: Export emails
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox. No export performed." -Level "WARNING"
                            Write-Host "No emails found for the specified user. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "Failed to retrieve user mailbox." -Level "ERROR"
                    throw "User mailbox retrieval failed. Cannot proceed with export."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 3: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 3. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"4" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 4: Export All Emails in Date Range for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking Microsoft Graph SDK connection..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve the user mailbox and date range
                Write-Log -Message "Prompting user to provide mailbox and date range for email export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User
                $DateRange = Get-FilterCriteria -StartDate -EndDate

                if ($UserMailbox -and $DateRange) {
                    # Step 3: Create export folder
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails for the user in the specified date range
                        Write-Log -Message "Retrieving emails for $UserMailbox from $($DateRange.StartDate) to $($DateRange.EndDate)..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox -StartDate $DateRange.StartDate.ToString("MM/dd/yyyy") -EndDate $DateRange.EndDate.ToString("MM/dd/yyyy")

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully for $UserMailbox. Starting export..." -Level "INFO"

                            # Step 5: Export emails
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox in the specified date range." -Level "WARNING"
                            Write-Host "No emails found in the specified date range. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "User mailbox or date range information is missing." -Level "ERROR"
                    throw "Failed to retrieve user mailbox or date range."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 4: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 4. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"5" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 5: Export All Emails by Sender for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking Microsoft Graph SDK connection..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve user mailbox and sender filter
                Write-Log -Message "Prompting user to provide mailbox and sender filter for export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User
                $Sender = (Get-FilterCriteria -Sender).Sender

                if ($UserMailbox -and $Sender) {
                    # Step 3: Create the main folder for email export
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails filtered by sender
                        Write-Log -Message "Retrieving emails for $UserMailbox from sender $Sender..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox -Sender $Sender

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully for $UserMailbox. Starting export..." -Level "INFO"

                            # Step 5: Export emails to the specified folder
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox from sender $Sender." -Level "WARNING"
                            Write-Host "No emails found for the specified sender. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "User mailbox or sender filter information is missing." -Level "ERROR"
                    throw "Failed to retrieve user mailbox or sender information."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 5: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 5. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"6" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 6: Export All Emails by Subject Line for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking Microsoft Graph SDK connection..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve the user mailbox and subject line filter
                Write-Log -Message "Prompting user to provide mailbox and subject line filter for export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User
                $SubjectLine = (Get-FilterCriteria -SubjectLine).SubjectLine

                if ($UserMailbox -and $SubjectLine) {
                    # Step 3: Create the main folder for email export
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails filtered by subject line
                        Write-Log -Message "Retrieving emails for $UserMailbox with subject line containing '$SubjectLine'..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox -SubjectLine $SubjectLine

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully for $UserMailbox. Starting export..." -Level "INFO"

                            # Step 5: Export emails to the specified folder
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox with subject line containing '$SubjectLine'." -Level "WARNING"
                            Write-Host "No emails found matching the specified subject line. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "User mailbox or subject line information is missing." -Level "ERROR"
                    throw "Failed to retrieve user mailbox or subject line filter."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 6: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 6. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"7" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 7: Export All Emails by Recipient for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking Microsoft Graph SDK connection..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve the user mailbox and recipient filter
                Write-Log -Message "Prompting user to provide mailbox and recipient filter for export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User
                $Recipient = (Get-FilterCriteria -Receiver).Receiver

                if ($UserMailbox -and $Recipient) {
                    # Step 3: Create the main folder for email export
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails filtered by recipient
                        Write-Log -Message "Retrieving emails for $UserMailbox with recipient $Recipient..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox -Receiver $Recipient

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully for $UserMailbox. Starting export..." -Level "INFO"

                            # Step 5: Export emails to the specified folder
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox with recipient $Recipient." -Level "WARNING"
                            Write-Host "No emails found matching the specified recipient. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "User mailbox or recipient filter information is missing." -Level "ERROR"
                    throw "Failed to retrieve user mailbox or recipient information."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 7: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 7. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"8" {
    # Log Option Selection
    Clear-Host
    [System.Console]::Clear()
    Write-Log -Message "Processing Option 8: Export All Emails by Domain for One User [Dump Format]" -Level "INFO"

    try {
        # Step 1: Check Microsoft Graph SDK connection
        Write-Log -Message "Checking Microsoft Graph SDK connection..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext) {
            Write-Log -Message "Microsoft Graph SDK is connected." -Level "INFO"

            try {
                # Step 2: Retrieve user mailbox and domain filter
                Write-Log -Message "Prompting user to provide mailbox and domain filter for export..." -Level "INFO"
                $UserMailbox = (Get-FilterCriteria -User).User
                $Domain = (Get-FilterCriteria -Domain).Domain

                if ($UserMailbox -and $Domain) {
                    # Step 3: Create export folder
                    Write-Log -Message "Creating export folder for $UserMailbox..." -Level "INFO"
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox

                    if ($FolderPath) {
                        Write-Log -Message "Export folder created successfully at: $FolderPath" -Level "INFO"

                        # Step 4: Retrieve emails filtered by domain
                        Write-Log -Message "Retrieving emails for $UserMailbox from domain $Domain..." -Level "INFO"
                        $EmailsToExport = Get-Emails -User $UserMailbox -Domain $Domain

                        if ($EmailsToExport) {
                            Write-Log -Message "Emails retrieved successfully for $UserMailbox. Starting export..." -Level "INFO"

                            # Step 5: Export emails
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                            Write-Log -Message "Email export completed for $UserMailbox." -Level "INFO"
                        } else {
                            Write-Log -Message "No emails found for $UserMailbox from domain $Domain." -Level "WARNING"
                            Write-Host "No emails found matching the specified domain. No export performed." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log -Message "Failed to create folder for $UserMailbox." -Level "ERROR"
                        throw "Unable to create export folder for $UserMailbox."
                    }
                } else {
                    Write-Log -Message "User mailbox or domain information is missing." -Level "ERROR"
                    throw "Failed to retrieve user mailbox or domain filter."
                }
            } catch {
                Write-Log -Message "Error during email export process: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Log -Message "Microsoft Graph SDK is not connected. Please run Option 1 to connect to Graph." -Level "WARNING"
            Write-Host "Microsoft Graph SDK is not connected. Please connect to Graph first." -ForegroundColor Red
        }
    } catch {
        # Log any errors encountered during the process
        Write-Log -Message "Error processing Option 8: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        # Return to the menu
        Write-Log -Message "Finished processing Option 8. Returning to the menu..." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    }
}
"Q" {
    # Log the user's choice to quit
    Write-Log -Message "Processing Option Q: User chose to quit the menu." -Level "INFO"
    try {
        Write-Log -Message "Exiting the script. Clearing the console and performing cleanup..." -Level "INFO"

        # Clean up and quit
        Start-Sleep -Seconds 4
        Clear-Host
        [System.Console]::Clear()
    } catch {
        # Log errors during quit
        Write-Log -Message "Error while attempting to quit: $($_.Exception.Message)" -Level "ERROR"
    } finally {
        Write-Log -Message "Script execution has been terminated by the user." -Level "INFO"
        #Return  # Replacing "Break" with "Return" for proper script termination
       #BREAK
    }
}
Default {
    Write-Log -Message "Invalid menu selection: $choice. Prompting the user to try again." -Level "WARNING"
    Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow

    try {
        Write-Log -Message "Reloading the menu for user input." -Level "INFO"
        Read-Host "Press [Enter] to reload the menu"
        Clear-Host
        [System.Console]::Clear()
        Show-Menu
    } catch {
        Write-Log -Message "Error during menu reload: $($_.Exception.Message)" -Level "ERROR"
        Write-Host "An error occurred while reloading the menu. Check logs for more details." -ForegroundColor Red
    }
}
}
}
##################################################################################################################################################################
#=============================End of Functions=========================================================================================================================
##################################################################################################################################################################
#==============================Main================================================================================================================================
##################################################################################################################################################################
Write-Log -Message "Script Started" -Level "INFO"

try {
    # Step 1: Check Operating System Version
    Write-Log -Message "Checking Operating System version..." -Level "INFO"
    if ([System.Environment]::OSVersion.Version.Major -lt 10) {
        Write-Log -Message "Error: This script requires Windows 10 or above. Current OS version does not meet the requirement." -Level "ERROR"
        throw "This script requires Windows 10 or above."
    }
    Write-Log -Message "OS version check passed: Windows 10 or above detected." -Level "INFO"

    # Step 2: Check for Administrative Privileges
    Write-Log -Message "Checking administrative privileges..." -Level "INFO"
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $IsAdmin) {
        Write-Log -Message "Error: Script is not running with administrative privileges. Please run as Administrator." -Level "ERROR"
        throw "Administrative privileges are required to run this script."
    }
    Write-Log -Message "Administrative privileges confirmed." -Level "INFO"

    # Step 3: Check PowerShell Version
    Write-Log -Message "Checking PowerShell version..." -Level "INFO"
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Log -Message "Error: This script requires PowerShell version 5 or above. Current version does not meet the requirement." -Level "ERROR"
        throw "This script requires PowerShell version 5 or above."
    }
    Write-Log -Message "PowerShell version check passed: Version 5 or above detected." -Level "INFO"

    # Step 4: Set Up the Environment
    Write-Log -Message "Clearing console and setting up the environment..." -Level "INFO"
    Clear-Host
    [System.Console]::Clear()
    Set-Environment
    Write-Log -Message "Environment setup completed successfully." -Level "INFO"

    # Step 5: Display Script Details
    Write-Log -Message "Displaying script and system details to the user." -Level "INFO"
    $Title = "=== Export User Emails ==="
    $MenuPrompt = "=" * $Title.Length
    $SystemDetails = @(
        "Operating System: " + (Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption,
        "OS Version: " + (Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).Version,
        "Running as Administrator: " + $IsAdmin,
        "PowerShell Version: " + $PSVersionTable.PSVersion,
        "PowerShell Edition: " + $PSVersionTable.PSEdition,
        "Current Execution Policy: " + (Get-ExecutionPolicy),
        "Computer Name: " + (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue).Name,
        "Computer Owner: " + (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue).PrimaryOwnerName
    )
    Write-Log -Message "System details: $($SystemDetails -join '; ')" -Level "INFO"

    # Step 6: Load the Main Menu
    Read-Host "Press [ENTER] to load the menu"
    Clear-Host
    [System.Console]::Clear()
    Show-Menu

} catch {
    # Log any errors during the main execution flow
    Write-Log -Message "Script terminated due to an error: $($_.Exception.Message)" -Level "ERROR"
    throw
} finally {
    # Clean up resources at the end of the script
    Write-Log -Message "Script execution completed. Performing cleanup..." -Level "INFO"
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Log -Message "Script Ended" -Level "INFO"

    # Open the log file for review
    Start-Process notepad.exe -ArgumentList $Global:LogFile
}
##################################################################################################################################################################
#==============================End of Main===========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script==========================================================================================================================
##################################################################################################################################################################
