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
Script Version: 1
OS Version Script was written on: Microsoft Windows 11 Pro : 10.0.25100 Build 26100
PSVersion 5.1.26100.2161 : PSEdition Desktop : Build Version 10.0.26100.2161
Description of Script: You can only do this with the authenticated account not for all users
#>
##################################################################################################################################################################
#==============================Beginning of script======================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================Functions==============================================================================================================================
##################################################################################################################################################################
# Define the global log file path at the start of the script
$Global:LogFile = "C:\Rsanchezc169ScriptLogs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"

# Global function to handle logging
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,                  # The message to log
        [Parameter()]
        [string]$LogFile = $Global:LogFile # Optional: Specify a log file, default to the global log file
    )

    # Ensure the log file exists
    $LogDirectory = Split-Path -Path $LogFile
    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    # Append the message to the log file with a timestamp
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "$Timestamp : $Message"
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Load-Module {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Modules  # Supports multiple modules
    )

    Foreach ($Module in $Modules) {
        Write-Progress -Activity "Processing Module: $($Module)" -Status "Starting task..." -PercentComplete 0 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        
        Try {
            # Step 1: Check if the module is installed
            Write-Progress -Activity "Processing Module: $($Module)" -Status "Checking installation status..." -PercentComplete 20 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            if ((Get-InstalledModule -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue) -OR (Get-Module -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue)) {
                Write-Warning "Module '$Module' is already installed"
                Write-Log -Message "Module '$Module' is already installed"
            } else {
                Write-Progress -Activity "Processing Module: $($Module)" -Status "Installing module..." -PercentComplete 40 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                Install-Module -Name "$Module*" -Force -Scope CurrentUser -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                Write-Log -Message "Module '$Module' installed successfully"
            }

            # Step 2: Check for updates
            Write-Progress -Activity "Processing Module: $($Module)" -Status "Checking for updates..." -PercentComplete 60 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            $CurrentVersion = (Find-Module -Name "$Module" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).Version
            $InstalledVersion = IF(Get-InstalledModule -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue){(Get-InstalledModule -Name "$Module" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).Version} ELSEIF(Get-Module -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue){(Get-Module -Name "$Module" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).Version}
            
            if ($CurrentVersion -and $InstalledVersion -and $CurrentVersion -gt $InstalledVersion) {
                Write-Progress -Activity "Processing Module: $($Module)" -Status "Updating module..." -PercentComplete 80 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                Update-Module -Name "$Module" -Force -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                Write-Log -Message "Module '$Module' updated successfully"
            } else {
                Write-Warning "Module '$Module' is up-to-date"
                Write-Log -Message "Module '$Module' is already up to date"
            }

            # Step 3: Import the module
            [ARRAY]$CurrentModules = @()
            IF (Get-InstalledModule -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue) {
                $CurrentModules = (Get-InstalledModule -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).Name
            } ELSEIF(Get-Module -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue) {
                $CurrentModules = (Get-Module -Name "$Module*" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).Name
            }

            if ($CurrentModules) {
                Foreach ($CurrentModule in $CurrentModules) {
                    Write-Progress -Activity "Processing Sub-Module: $($CurrentModule)" -Status "Importing module..." -PercentComplete 90 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    
                    Try {
                        Import-Module -Name "$CurrentModule" -Force -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                        Write-Log -Message "Module '$CurrentModule' imported successfully"
                    } Catch {
                        Write-Log -Message "Error importing sub-module '$CurrentModule': $($_.Exception.Message)"
                        Write-Warning "Error importing sub-module '$CurrentModule'. Check logs for details."
                    }
                }
            } else {
                Write-Warning "No modules found matching '$Module'. Ensure the module name is correct."
                Write-Log -Message "No modules found matching '$Module'"
            }
        } Catch {
            # Log any errors
            Write-Log -Message "Error processing module '$Module': $($_.Exception.Message)"
            Write-Warning "Error processing module '$Module'. Check logs for details."
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Module Processing" -Status "Completed all tasks." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
    Write-Progress -Activity "Module Processing" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Set-Environment {
    # Initialize progress bar
    Write-Progress -Activity "Environment Setup" -Status "Starting setup..." -PercentComplete 0 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

    try {
        # Step 1: Clear the console
        Write-Progress -Activity "Environment Setup" -Status "Clearing console..." -PercentComplete 10 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Clear-Host

        # Step 2: Maximize window
        Write-Progress -Activity "Environment Setup" -Status "Maximizing console window..." -PercentComplete 90 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;

            public class User32 {
                [DllImport("user32.dll")]
                public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            }
"@
        $handle = (Get-Process -ID $PID).MainWindowHandle
        [User32]::ShowWindow($handle, 3)  # Maximize window
        Write-Log -Message "Console window maximized"

        # Step 3: Set execution policy
        Write-Progress -Activity "Environment Setup" -Status "Setting execution policy..." -PercentComplete 20 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Set-ExecutionPolicy -Scope "Process" -ExecutionPolicy "Unrestricted" -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Log -Message "Execution policy set to 'Unrestricted'"

        # Step 4: Load required modules
        Write-Progress -Activity "Environment Setup" -Status "Loading PowerShellGet module..." -PercentComplete 30 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Load-Module -Module "PowerShellGet"
        Write-Log -Message "PowerShellGet module loaded successfully"

        Write-Progress -Activity "Environment Setup" -Status "Loading Microsoft.Graph module..." -PercentComplete 40 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Load-Module -Module "Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "Microsoft.Graph.Users", "Microsoft.Graph.Mail"
        Write-Log -Message "Microsoft.Graph module loaded successfully"

        # Step 5: Set window properties
        Write-Progress -Activity "Environment Setup" -Status "Configuring console appearance..." -PercentComplete 60 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        $Host.UI.RawUI.ForegroundColor = 'Blue'
        $Host.UI.RawUI.WindowTitle = "Export Emails"
        Write-Log -Message "Console appearance configured: BackgroundColor=Black, ForegroundColor=Blue, WindowTitle='Export Emails'"

        # Step 6: Configure session preferences
        Write-Progress -Activity "Environment Setup" -Status "Setting session preferences..." -PercentComplete 70 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        $Global:FormatEnumerationLimit = -1 
        $Global:ErrorActionPreference = "SilentlyContinue"
        Write-Log -Message "Session preferences configured: FormatEnumerationLimit=-1, ErrorActionPreference=SilentlyContinue"

        # Clear final progress
        Write-Progress -Activity "Environment Setup" -Status "Setup complete." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Progress -Activity "Environment Setup" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Log -Message "Environment setup completed successfully"
    } catch {
        # Handle errors and log them
        Write-Progress -Activity "Environment Setup" -Status "Setup failed due to an error." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Progress -Activity "Environment Setup" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Log -Message "Error during environment setup: $($_.Exception.Message)"
        throw
    }

    Write-Progress -Activity "Environment Setup" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
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

    # Define the email validation regex pattern
    $pattern = '^(?:[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\

\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9?\.])+[a-z0-9?]|(?:

\[(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f])+)\]

)$'

    # Perform email validation
    try {
        if ($EmailAddress -match $pattern) {
            # Log valid email
            Write-Log -Message "Valid email address: $EmailAddress"
            Write-Progress -Activity "Email Validation" -Status "Validation complete: Email is valid." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            # Clear progress bar 
            Write-Progress -Activity "Email Validation" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            return $true
        } else {
            # Log invalid email
            Write-Log -Message "Invalid email address: $EmailAddress"
            Write-Progress -Activity "Email Validation" -Status "Validation complete: Email is invalid." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            # Clear progress bar
            Write-Progress -Activity "Email Validation" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            return $false
        }
    } catch {
        # Log any unexpected errors
        Write-Log -Message "Error validating email address: $($_.Exception.Message)"
        Write-Progress -Activity "Email Validation" -Status "Validation failed due to an error." -PercentComplete 100 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        # Clear progress bar
        Write-Progress -Activity "Email Validation" -Completed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        throw
    }

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-DateRange {
    # Initialize variables
    $StartDate = $null
    $EndDate = $null

    do {
        try {
            # Prompt user for start date
            Write-Host "Please enter a valid Start Date (e.g., MM/dd/yyyy):" -ForegroundColor Cyan
            $StartDateInput = Read-Host "Start Date"
            $StartDate = [datetime]::Parse($StartDateInput)

            # Prompt user for end date
            Write-Host "Please enter a valid End Date (e.g., MM/dd/yyyy):" -ForegroundColor Cyan
            $EndDateInput = Read-Host "End Date"
            $EndDate = [datetime]::Parse($EndDateInput)

            # Initialize progress
            Write-Progress -Activity "Date Range Validation" -Status "Validating dates..." -PercentComplete 30 -ErrorAction SilentlyContinue

            # Validate the date range
            if ($StartDate -gt $EndDate) {
                Write-Warning "Start date cannot be later than the End date. Please try again."
                Write-Log -Message "Error: Start date ($StartDate) is after End date ($EndDate)"
                continue
            }
            if ($EndDate -gt (Get-Date).AddDays(1)) {
                Write-Warning "End date cannot be in the future. Please try again."
                Write-Log -Message "Error: End date ($EndDate) is in the future"
                continue
            }

            # Validation successful
            Write-Progress -Activity "Date Range Validation" -Status "Validation successful" -PercentComplete 100 -ErrorAction SilentlyContinue
            Write-Log -Message "Success: Start date ($StartDate) and End date ($EndDate) are valid"
            break

        } catch {
            # Log any errors during parsing or validation
            Write-Warning "Invalid date format. Please enter the date in MM/dd/yyyy format."
            Write-Log -Message "Error parsing dates: $($_.Exception.Message)"
        }

    } while (-not ($StartDate -and $EndDate))  # Keep looping until both dates are valid

    # Return the valid date range
    return @{
        StartDate = $StartDate
        EndDate = $EndDate
    }

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
FUNCTION Draw-Line {
    # Get the window width
    $LineWidth = ((get-host -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue).UI.RawUI.WindowSize.Width)
    
    # Generate the line
    $Line += '=' * ($LineWidth)

    # Log the action
    Write-Log -Message "Draw-Line executed. Line width: $LineWidth"

    # Return the line
    RETURN $Line

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Validate-Information {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Information
    )

    # Initialize progress bar
    Write-Progress -Activity "Information Validation" -Status "Starting validation..." -PercentComplete 0   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

    do {
        Write-Progress -Activity "Information Validation" -Status "Displaying input information..." -PercentComplete 20   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        Write-Host "The value you inputted is: $Information"

        # Ask user for validation
        $Validate = Read-Host -Prompt "The above information was inputted. Is this correct [Y/N]"   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        $Validate = $Validate.ToUpper()

        Switch ($Validate) {
            "Y" {
                # Log success and progress
                Write-Log -Message "Information validated successfully: $Information"
                Write-Progress -Activity "Information Validation" -Status "Validation successful." -PercentComplete 100   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                # Clear progress bar
                Write-Progress -Activity "Information Validation" -Completed   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                return $Information
            }
            "N" {
                # Reset and re-enter information
                Write-Progress -Activity "Information Validation" -Status "Re-entering information..." -PercentComplete 50   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                $Information = Read-Host -Prompt "Enter new value ->"   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

                # Ensure value is not empty
                while ([string]::IsNullOrWhiteSpace($Information)) {
                    Write-Warning "You did not put in a value. This cannot be empty."
                    $Information = Read-Host -Prompt "Enter new value ->"   -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                }

                # Log the updated information
                Write-Log -Message "Information updated to: $Information"
            }
            Default {
                # Log invalid validation input
                Write-Log -Message "Invalid validation input: $Validate"
                Write-Warning "Invalid input. Please validate with [Y/N]."
            }
        }
    } while ($Validate -ne "Y")

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-Emails {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$User,  # User's email or ID

        [Parameter()]
        [datetime]$StartDate,  # Optional Start Date

        [Parameter()]
        [datetime]$EndDate,    # Optional End Date

        [Parameter()]
        [string]$SubjectLine,  # Optional Subject Line

        [Parameter()]
        [string]$Sender,       # Optional Sender's email address

        [Parameter()]
        [string]$Receiver,     # Optional Receiver's email address

        [Parameter()]
        [string]$Domain        # Optional Domain
    )

    # Initialize progress bar
    Write-Progress -Activity "Retrieving Emails" -Status "Initializing..." -PercentComplete 0

    try {
        # Check if any emails exist
        Write-Progress -Activity "Retrieving Emails" -Status "Checking for emails for user: $User..." -PercentComplete 30
        $FirstEmail = Get-MgUserMessage -UserId $User | Select-Object -First 1 

        if ($FirstEmail) {
            Write-Log -Message "Emails found for user: $User"

            # Build filter string
            $filterParts = @()
            if ($PSBoundParameters.ContainsKey('StartDate') -and $PSBoundParameters.ContainsKey('EndDate')) {
                $filterParts += "receivedDateTime ge $($StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")) and receivedDateTime le $($EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))"
            } elseif ($PSBoundParameters.ContainsKey('StartDate')) {
                $filterParts += "receivedDateTime ge $($StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))"
            } elseif ($PSBoundParameters.ContainsKey('EndDate')) {
                $filterParts += "receivedDateTime le $($EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))"
            }

            if ($PSBoundParameters.ContainsKey('SubjectLine')) {
                $filterParts += "subject eq '$SubjectLine'"
            }
            if ($PSBoundParameters.ContainsKey('Sender')) {
                $filterParts += "from/emailAddress/address eq '$Sender'"
            }
            if ($PSBoundParameters.ContainsKey('Receiver')) {
                $filterParts += "toRecipients/any(t: t/emailAddress/address eq '$Receiver')"
            }
            if ($PSBoundParameters.ContainsKey('Domain')) {
                $filterParts += "from/emailAddress/address endswith '$Domain'"
            }

            $filter = $filterParts -join " and "

            # Retrieve emails with the filter
            Write-Progress -Activity "Retrieving Emails" -Status "Fetching emails for user: $User..." -PercentComplete 60
            $AllEmails = if ($filter) {
                Get-MgUserMessage -UserId $User -Filter $filter -All
            } else {
                Get-MgUserMessage -UserId $User -All
            }

            Write-Log -Message "Successfully retrieved emails for user: $User"

            # Return emails
            Write-Progress -Activity "Retrieving Emails" -Status "Emails retrieved successfully." -PercentComplete 100 -Completed
            return $AllEmails
        } else {
            Write-Log -Message "No emails found for user: $User"
            Write-Warning "No emails found for user: $User"
            return $null
        }
    } catch {
        # Handle errors and log them
        Write-Progress -Activity "Retrieving Emails" -Status "Failed to retrieve emails for user: $User." -PercentComplete 100 -Completed
        Write-Log -Message "Error retrieving emails for user: $User - $($_.Exception.Message)"
        Write-Warning "Could not retrieve emails for user: $User. Error: $($_.Exception.Message)"
        throw
    }

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Export-Email {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ARRAY]$Emails,  # Array of email message objects
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [STRING]$FolderPath,  # Folder path where emails will be saved
        [Parameter(Mandatory)]
        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]  # Basic validation for email format
        [STRING]$User  # The email address of the user whose messages are being exported
    )

    # Start progress
    Write-Progress -Activity "Email Export" -Status "Initializing..." -PercentComplete 0 -ErrorAction SilentlyContinue

    # Process each email and export
    try {
        $TotalEmails = $Emails.Count
        $Counter = 0

        foreach ($Message in $Emails) {
            $Counter++
            Write-Progress -Activity "Email Export" -Status "Exporting email $Counter of $TotalEmails..." -PercentComplete (($Counter / $TotalEmails) * 100) -ErrorAction SilentlyContinue

            # Generate a valid file name for the email
            try {
                $BaseFileName = ($Message.Subject + ".eml") -replace '[<>:"/\\|?*]', '_'
                $OutFile = Join-Path -Path $FolderPath -ChildPath $BaseFileName
                $Index = 1

                # Check if the file already exists and modify the name if necessary
                while (Test-Path -Path $OutFile) {
                    $IndexedFileName = ($Message.Subject + "_$Index.eml") -replace '[<>:"/\\|?*]', '_'
                    $OutFile = Join-Path -Path $FolderPath -ChildPath $IndexedFileName
                    $Index++
                }

                # Export the email to the file
                Get-MgUserMessageContent -UserId $User -MessageId $Message.Id -OutFile $OutFile
                Write-Log -Message "Email exported successfully: $OutFile"
            } catch {
                Write-Log -Message "Error exporting email: $($_.Exception.Message)"
                Write-Warning "Failed to export email: $($Message.Subject). Check logs for details."
            }
        }
    } catch {
        Write-Log -Message "Error during email export process: $($_.Exception.Message)"
        Write-Warning "An error occurred during the email export process. Check logs for details."
        Write-Progress -Activity "Email Export" -Completed -ErrorAction SilentlyContinue
        throw
    }

    # Finalize progress
    Write-Progress -Activity "Email Export" -Status "Completed" -PercentComplete 100 -ErrorAction SilentlyContinue
    Write-Progress -Activity "Email Export" -Completed -ErrorAction SilentlyContinue

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
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

    # Start progress
    Write-Progress -Activity "Folder Creation" -Status "Initializing..." -PercentComplete 0 -ErrorAction SilentlyContinue

    # Extract folder name from the email address
    $FolderName = $null
    try {
        Write-Progress -Activity "Folder Creation" -Status "Processing email address..." -PercentComplete 20 -ErrorAction SilentlyContinue
        $FolderName = ($EmailAddress -split "@")[0] + "_Emails"
    } catch {
        Write-Log -Message "Error extracting folder name from email address: $($_.Exception.Message)"
        Write-Warning "Failed to extract folder name. Check logs for details."
        Write-Progress -Activity "Folder Creation" -Completed -ErrorAction SilentlyContinue
        throw
    }

    # Get the Downloads folder path
    $MainFolderPath = $null
    try {
        Write-Progress -Activity "Folder Creation" -Status "Retrieving Downloads folder path..." -PercentComplete 40 -ErrorAction SilentlyContinue
        $MainFolderPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
    } catch {
        Write-Log -Message "Error retrieving Downloads folder path: $($_.Exception.Message)"
        Write-Warning "Could not retrieve the Downloads folder path. Check logs for details."
        Write-Progress -Activity "Folder Creation" -Completed -ErrorAction SilentlyContinue
        throw
    }

    # Check if the folder already exists or create it
    try {
        Write-Progress -Activity "Folder Creation" -Status "Creating folder..." -PercentComplete 60 -ErrorAction SilentlyContinue
        $FullFolderPath = "$MainFolderPath\$FolderName"

        if (Test-Path -Path $FullFolderPath) {
            Write-Log -Message "Folder already exists: $FullFolderPath"
            Write-Warning "Folder already exists: $FullFolderPath"
        } else {
            New-Item -ItemType Directory -Path $FullFolderPath -Force | Out-Null
            Write-Log -Message "Folder created successfully: $FullFolderPath"
            Write-Host "Folder created successfully: $FullFolderPath"
        }
    } catch {
        Write-Log -Message "Error creating folder: $($_.Exception.Message)"
        Write-Warning "Failed to create folder. Check logs for details."
        Write-Progress -Activity "Folder Creation" -Completed -ErrorAction SilentlyContinue
        throw
    }

    # Finalize progress
    Write-Progress -Activity "Folder Creation" -Status "Completed" -PercentComplete 100 -ErrorAction SilentlyContinue
    Write-Progress -Activity "Folder Creation" -Completed -ErrorAction SilentlyContinue

    # Return the folder path
    return $FullFolderPath

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
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
        [string]$SubFolderName  # The name of the sub-folder to create
    )

    # Start progress
    Write-Progress -Activity "Sub-Folder Creation" -Status "Initializing..." -PercentComplete 0 -ErrorAction SilentlyContinue

    # Construct the sub-folder path
    $SubFolderPath = "$MainFolderPath\$SubFolderName"

    # Check if the sub-folder already exists or create it
    try {
        Write-Progress -Activity "Sub-Folder Creation" -Status "Creating sub-folder..." -PercentComplete 60 -ErrorAction SilentlyContinue

        if (Test-Path -Path $SubFolderPath) {
            Write-Log -Message "Sub-folder already exists: $SubFolderPath"
            Write-Warning "Sub-folder already exists: $SubFolderPath"
        } else {
            New-Item -ItemType Directory -Path $SubFolderPath -Force | Out-Null
            Write-Log -Message "Sub-folder created successfully: $SubFolderPath"
            Write-Host "Sub-folder created successfully: $SubFolderPath"
        }
    } catch {
        Write-Log -Message "Error creating sub-folder: $($_.Exception.Message)"
        Write-Warning "Failed to create sub-folder. Check logs for details."
        Write-Progress -Activity "Sub-Folder Creation" -Completed -ErrorAction SilentlyContinue
        throw
    }

    # Finalize progress
    Write-Progress -Activity "Sub-Folder Creation" -Status "Completed" -PercentComplete 100 -ErrorAction SilentlyContinue
    Write-Progress -Activity "Sub-Folder Creation" -Completed -ErrorAction SilentlyContinue

    # Return the sub-folder path
    return $SubFolderPath

    # Clean up
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-FilterCriteria {
    [CmdletBinding()]
    Param(
        [switch]$User,        # If set, prompts for a user and validates existence
        [switch]$StartDate,   # If set, prompts for a start date
        [switch]$EndDate,     # If set, prompts for an end date
        [switch]$SubjectLine, # If set, prompts for a subject line
        [switch]$Sender,      # If set, prompts for a sender email
        [switch]$Receiver,    # If set, prompts for a receiver email
        [switch]$Domain      # If set, prompts for a domain
    )

    # Initialize variables for criteria
    $Criteria = @{}
    $EmailPattern = '^(?:[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\

\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9?\.])+[a-z0-9?]|(?:

\[(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f])+)\]

)$'
    $DomainPattern = '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    # Function to validate email addresses
    Function Validate-Email {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]$EmailAddress
        )
        try {
            if ($EmailAddress -match $EmailPattern) {
                Write-Log -Message "Valid email address: $EmailAddress"
                return $true
            } else {
                Write-Log -Message "Invalid email address: $EmailAddress"
                Write-Warning "Invalid email format: $EmailAddress. Please enter a valid email address."
                return $false
            }
        } catch {
            Write-Log -Message "Error validating email address: $($_.Exception.Message)"
            Write-Warning "An error occurred while validating the email address. Please try again."
            return $false
        }
    }

    # Function to validate domains
    Function Validate-Domain {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]$Domain
        )
        if ($Domain -match $DomainPattern) {
            Write-Log -Message "Valid domain: $Domain"
            return $true
        } else {
            Write-Log -Message "Invalid domain: $Domain"
            Write-Warning "Invalid domain format: $Domain. Please enter a valid domain (e.g., example.com)."
            return $false
        }
    }
    # Function to validate date range
    Function Get-DateRangeValidation {
        [CmdletBinding()]
        param()

        # Initialize variables
        $StartDate = $null
        $EndDate = $null

        do {
            try {
                # Prompt user for start date
                Write-Host "Please enter a valid Start Date (e.g., MM/dd/yyyy):" -ForegroundColor Cyan
                $StartDateInput = Read-Host "Start Date"
                $StartDate = [datetime]::Parse($StartDateInput)

                # Prompt user for end date
                Write-Host "Please enter a valid End Date (e.g., MM/dd/yyyy):" -ForegroundColor Cyan
                $EndDateInput = Read-Host "End Date"
                $EndDate = [datetime]::Parse($EndDateInput)

                # Validate the date range
                if ($StartDate -gt $EndDate) {
                    Write-Warning "Start date cannot be later than the End date. Please try again."
                    Write-Log -Message "Error: Start date ($StartDate) is after End date ($EndDate)"
                    continue
                }
                if ($EndDate -gt (Get-Date).AddDays(1)) {
                    Write-Warning "End date cannot be in the future. Please try again."
                    Write-Log -Message "Error: End date ($EndDate) is in the future"
                    continue
                }

                # Validation successful
                Write-Host "Date range is valid!" -ForegroundColor Green
                Write-Log -Message "Success: Start date ($StartDate) and End date ($EndDate) are valid"
                break

            } catch {
                # Handle invalid date format
                Write-Warning "Invalid date format. Please enter the date in MM/dd/yyyy format."
                Write-Log -Message "Error parsing dates: $($_.Exception.Message)"
            }
        } while (-not ($StartDate -and $EndDate))  # Loop until valid

        # Return validated date range
        return @{
            StartDate = $StartDate
            EndDate = $EndDate
        }
    }
    # Main logic for confirmation and re-prompting
    do {
        # Ensure both StartDate and EndDate are provided together
        if ($StartDate -xor $EndDate) {
            Write-Warning "You must provide both -StartDate and -EndDate parameters together. The function will now exit."
            return
        }

        # Validate Date Range
        if ($StartDate -and $EndDate) {
            $DateRange = Get-DateRangeValidation
            $Criteria["StartDate"] = $DateRange.StartDate
            $Criteria["EndDate"] = $DateRange.EndDate
        }

        # Validate SubjectLine
        if ($SubjectLine) {
            do {
                $InputSubject = Read-Host "Enter Subject Line"
                if (-not [string]::IsNullOrWhiteSpace($InputSubject)) {
                    $Criteria["SubjectLine"] = $InputSubject
                    Write-Log -Message "Subject Line: $InputSubject"
                    break
                } else {
                    Write-Warning "Subject line cannot be empty. Please enter a valid subject line."
                }
            } while ($true)
        }
        # Validate User email and existence via Get-MgUser
        if ($User) {
            do {
                $InputUser = Read-Host "Enter User Email (e.g., USER@DOMAIN.COM)"
                if (Validate-Email -EmailAddress $InputUser) {
                    try {
                        $UserExists = Get-MgUser -UserId $InputUser -ErrorAction Stop
                        if ($UserExists) {
                            Write-Host "User exists: $InputUser" -ForegroundColor Green
                            $Criteria["User"] = $InputUser
                            Write-Log -Message "User Email: $InputUser"
                            break
                        }
                    } catch {
                        Write-Warning "The entered user ($InputUser) does not exist in Microsoft Graph. Please try again."
                        Write-Log -Message "Error: Invalid user ($InputUser)"
                    }
                }
            } while ($true)
        }

        # Validate Sender email
        if ($Sender) {
            do {
                $InputSender = Read-Host "Enter Sender Email (e.g., USER@DOMAIN.COM)"
                if (Validate-Email -EmailAddress $InputSender) {
                    $Criteria["Sender"] = $InputSender
                    Write-Log -Message "Sender Email: $InputSender"
                    break
                }
            } while ($true)
        }
        # Validate Receiver email
        if ($Receiver) {
            do {
                $InputReceiver = Read-Host "Enter Receiver Email (e.g., USER@DOMAIN.COM)"
                if (Validate-Email -EmailAddress $InputReceiver) {
                    $Criteria["Receiver"] = $InputReceiver
                    Write-Log -Message "Receiver Email: $InputReceiver"
                    break
                }
            } while ($true)
        }

        # Validate Domain
        if ($Domain) {
            do {
                $InputDomain = Read-Host "Enter Domain (e.g., example.com)"
                if (Validate-Domain -Domain $InputDomain) {
                    $Criteria["Domain"] = $InputDomain
                    Write-Log -Message "Domain: $InputDomain"
                    break
                }
            } while ($true)
        }
        # Display criteria
        Write-Host "Filter Criteria:" -ForegroundColor Green
        $Criteria.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }

        # Confirmation
        $Confirm = Read-Host "Is the above information correct? (Y/N)"
        if ($Confirm -eq "Y") {
            Write-Host "You have confirmed the information as correct." -ForegroundColor Green
            break
        } else {
            Write-Warning "Let's re-enter the information."
            $Criteria.Clear()
            continue
        }
    } while ($true)

    # Log criteria
    Write-Log -Message "Criteria: $($Criteria | Out-String)"

    # Return criteria
    return $Criteria
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Show-Menu {
#Clear-Host
Clear-Host
#[System.Console]::Clear()
# Menu Display Logic
Clear-Host
#[System.Console]::Clear()
Write-Host "------------------------------------" -ForegroundColor Green
Write-Host "    Welcome to the Email Exporter" -ForegroundColor Green
Write-Host "------------------------------------" -ForegroundColor Green
$menu = @"
1 => Connect to Microsoft Graph
2 => Disconnect from Microsoft Graph
3 => Export All Emails for One User [Dump Format]
4 => Export All Emails in Date Range for One User[Dump Format][under development]
5 => Export All Emails by Sender For One Users[Dump Format][under development]
6 => Export All Emails by subject line for One User[Dump Format][under development]
7 => Export All Emails by Recipient for One User[Dump Format][under development]
7 => Export All Emails by Domain for One User[Dump Format][under development]
Q => Quit
--------------------------------------------------------------------------------
Select a task by number or Q to quit ::
"@
        $choice = Read-Host $menu
Switch ($choice) {
            "1" {
                # Connect to Microsoft Graph
                Clear-Host
                Write-Host "Option 1: Connect to Microsoft Graph" -ForegroundColor Cyan
               Write-Log -Message "Processing Option 1: Connect to Microsoft Graph"
                try {
                    $graphContext = Get-MgContext  -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    if ($graphContext -and $graphContext.Account -and $graphContext.TenantId) {
                        Write-Host "User is already connected to Microsoft Graph." -ForegroundColor Green
                        Write-Log -Message "Already connected to Microsoft Graph"  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    } elseif(!($graphContext -and $graphContext.Account -and $graphContext.TenantId)) {
		$Scopes = @(
		"User.ReadWrite.All", 
		"email",
		"Mail.ReadBasic"
		)
		Connect-MgGraph -Scope $Scopes  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                       Write-Log -Message "Successfully connected to Microsoft Graph"  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    }
                } catch {
                    Write-Warning "Could not connect to Microsoft Graph. Error: $($_.Exception.Message)"
                    Write-Log -Message "Error connecting to Microsoft Graph: $($_.Exception.Message)"  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                }
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	     Show-Menu
            }
"2" {
                # Disconnect from Microsoft Graph
                Clear-Host
                Write-Host "Option 2: Disconnect from Microsoft Graph" -ForegroundColor Cyan
                Write-Log -Message "Processing Option 2: Disconnect from Microsoft Graph"
                try {
                    $graphContext = Get-MgContext -ErrorAction Stop
                    if ($graphContext -and $graphContext.Account -and $graphContext.TenantId) {
                        Disconnect-MgGraph
                        Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
                        Write-Log -Message "Disconnected from Microsoft Graph"
                    } else {
                        Write-Host "Microsoft Graph SDK is not connected." -ForegroundColor Yellow
                    }
                } catch {
                    Write-Warning "Could not disconnect from Microsoft Graph. Error: $($_.Exception.Message)"
                    Write-Log -Message "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
                }
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	     Show-Menu
            }
"3" {
    # Export All Emails for One User
    Clear-Host
    Write-Host "Option 3: Export All Emails for One User [Dump Format]" -ForegroundColor Cyan
    Write-Log -Message "Processing Option 3: Export All Emails for One User [Dump Format]"

    try {
        # Check Microsoft Graph SDK connection
        $graphContext = Get-MgContext -ErrorAction Stop
        if ($graphContext -and $graphContext.Account -and $graphContext.TenantId) {
            try {
                # Retrieve the user mailbox
                Write-Host "Retrieving user mailbox..." -ForegroundColor Cyan
                $UserMailbox = (Get-FilterCriteria -User).User

                if ($UserMailbox) {
                    # Create the main folder for email export
                    $FolderPath = Create-MainFolder -EmailAddress $UserMailbox
                    if ($FolderPath) {
                        Write-Host "Exporting emails for $UserMailbox..." -ForegroundColor Cyan
                        
                        # Initialize array to hold emails
                        $EmailsToExport = @()
                        
                        # Retrieve emails for the user
                        $EmailsToExport = Get-Emails -User $UserMailbox

                        if ($EmailsToExport) {
                            # Export emails to the specified folder
                            Export-Email -Emails $EmailsToExport -FolderPath $FolderPath -User $UserMailbox
                        } else {
                            Write-Warning "No emails found for $UserMailbox."
                            Write-Log -Message "No emails found for $UserMailbox"
                        }
                    } else {
                        Write-Warning "Failed to create folder for $UserMailbox."
                        Write-Log -Message "Failed to create folder for $UserMailbox"
                    }

                    # Log export start
                    Write-Log -Message "Export started for user $UserMailbox"
                } else {
                    Write-Warning "Failed to retrieve user information."
                    Write-Log -Message "Failed to retrieve user information"
                }
            } catch {
                Write-Warning "Error while exporting emails for one user. Error: $($_.Exception.Message)"
                Write-Log -Message "Error exporting emails for one user: $($_.Exception.Message)"
            }
        } else {
            Write-Host "Microsoft Graph SDK is not connected. Please run option 1 to connect to Graph." -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Could not process Option 3. Error: $($_.Exception.Message)"
        Write-Log -Message "Error processing Option 3: $($_.Exception.Message)"
    }

    # Wait and return to the menu
    Read-Host "Press [Enter] to reload the menu"
    Start-Sleep -Seconds 1
    [System.Console]::Clear()
    Show-Menu
}
"4" {
                # Placeholder for Export All Emails in Date Range for One User[under development]
                Clear-Host
                Write-Host "Option 4: Export All Emails in Date Range for One User[under development]" -ForegroundColor Cyan
                Write-Host "This option is under development or requires additional logic." -ForegroundColor Yellow
                Write-Log -Message "Option 4 : Export All Emails in Date Range for One User[under development]"
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	     Show-Menu
            }
"5" {
                # Placeholder for Export All Emails by Sender or domain For One Users[under development]
                Clear-Host
                Write-Host "Option 5: Export All Emails by Sender or domain For One Users[under development]" -ForegroundColor Cyan
                Write-Host "This option is under development or requires additional logic." -ForegroundColor Yellow
                Write-Log -Message "Option 5 : Export All Emails by Sender or domain For One Users[under development]"
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	     Show-Menu
            }
"6" {
                # Placeholder for Export All Emails by subject line for One User[under development]
                Clear-Host
                Write-Host "Option 6: Export All Emails by subject line for One User[under development]" -ForegroundColor Cyan
                Write-Host "This option is under development or requires additional logic." -ForegroundColor Yellow
                Write-Log -Message "Option 6 : Export All Emails by subject line for One User[under development]"
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	     Show-Menu
            }
"Q" {
                Write-Host "Quitting..." -ForegroundColor Cyan
                Write-Log -Message "User quit the menu"
                sleep -Seconds 4
	     Clear-Host
	    #[System.Console]::Clear()
                ; BREAK
            }
Default {
                Write-Warning "Invalid selection. Please try again."
                Write-Log -Message "Invalid menu selection ($choice)"
                Read-Host "Press [Enter] to reload the menu"
                Start-Sleep -Seconds 1
	     Clear-Host
	    #[System.Console]::Clear()
	    Show-Menu
            }
        }

    #Clean up
    [System.Console]::Clear()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

}
##################################################################################################################################################################
#=============================End of Functions=========================================================================================================================
##################################################################################################################################################################
#==============================Main================================================================================================================================
##################################################################################################################################################################
Write-Log -Message "Script Started"

# Initialize progress bar
#Write-Progress -Activity "Script Initialization" -Status "Starting pre-checks..." -PercentComplete 0

# Log OS and environment checks
if ([System.Environment]::OSVersion.Version.Major -ge 10) {
    #Write-Progress -Activity "Script Initialization" -Status "Checking administrative privileges..." -PercentComplete 25
    Write-Log -Message "OS version meets the requirement (Windows 10 or above)"
    if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        #Write-Progress -Activity "Script Initialization" -Status "Checking PowerShell version..." -PercentComplete 50
        Write-Log -Message "PowerShell is running with administrative privileges"
        if ($PSVersionTable.PSVersion.Major -ge 5) {
            #Write-Progress -Activity "Script Initialization" -Status "Environment setup in progress..." -PercentComplete 75
            Write-Log -Message "PowerShell version meets the requirement (5 or above)"
            # Clear the screen and set the environment
            Clear-Host
            #[System.Console]::Clear()
            Set-Environment

            # Display script details
            Read-Host "Press [ENTER] to load the Script"
            #Clear-Host
            [System.Console]::Clear()
            $title = "=== Export Emails ==="
            $menuprompt = "=" * $title.Length
            Write-Host $menuprompt -ForegroundColor Yellow
            Write-Host $title -ForegroundColor Red
            Write-Host $menuprompt -ForegroundColor Yellow
            Write-Log -Message "Script title displayed: $title"

            Write-Host "Operating System: " (Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
            Write-Host "OS Version: " (Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).Version
            Write-Host "Running as Administrator: " (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
            Write-Host "PowerShell Version: " ($PSVersionTable.PSVersion)
            Write-Host "PowerShell Edition: " ($PSVersionTable.PSEdition)
            Write-Host "Current Execution Policy: " (Get-ExecutionPolicy)
            Write-Host "Computer Name: " (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
            Write-Host "Computer Owner: " (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue).PrimaryOwnerName
            Write-Log -Message "System and environment details displayed"

            # Display author details
            Write-Host $menuprompt -ForegroundColor Yellow
            Write-Host "Author: Rsanchezc169" -ForegroundColor Red
            Write-Host "Script Version: 1" -ForegroundColor Red
            Write-Host $menuprompt -ForegroundColor Yellow
            Write-Log -Message "Script author and version displayed"

            # Wait for user to load the menu
            Read-Host "Press [ENTER] to load the menu"
            Clear-Host
            #[System.Console]::Clear()
            # Load the menu
            Show-Menu
        } else {
            # PowerShell version too low
            Clear-Host
            #[System.Console]::Clear()
            Write-Warning "You are not using PowerShell 5 or above. Please update your PowerShell version."
            Write-Log -Message "Error: PowerShell version is below 5"
        }
    } else {
        # Not running as administrator
        Clear-Host
        #[System.Console]::Clear()
        Write-Warning "You are not running PowerShell with administrative privileges. Please run as Administrator."
        Write-Log -Message "Error: PowerShell is not running as Administrator"
    }
} else {
    # OS version too low
    Clear-Host
    #[System.Console]::Clear()
    Write-Warning "You are running an old version of Windows. This program requires Windows 10 or above."
    Write-Log -Message "Error: Windows version is below 10"
}

Write-Log -Message "Script Ended"

# Clear final progress
#Write-Progress -Activity "Script Initialization" -Status "Initialization complete." -PercentComplete 100
#Write-Progress -Activity "Script Initialization" -Completed
Clear-Host
#[System.Console]::Clear()

# Clean up
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Notepad $Global:LogFile
##################################################################################################################################################################
#==============================End of Main===========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script==========================================================================================================================
##################################################################################################################################################################
