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
Description of Script: 
This PowerShell script automates the retrieval, organization, and reporting of Microsoft 365 administrative roles and their members. 
It begins by ensuring that the Microsoft Graph SDK is installed and establishes a secure connection to Microsoft Graph. 
The script then retrieves all administrative roles containing "Administrator" in their display name and identifies the members assigned to each role. 
This data is stored in a structured array for further processing. 
The script features modular functions that handle specific tasks, along with comprehensive logging through a global Write-Log function that records key events, errors, and progress updates in a timestamped log file. 
Real-time progress bars provide visibility into the script’s execution.

The highlight of this script is its ability to generate an HTML report that allows users to dynamically view members of each role. 
The report includes a dropdown menu listing all roles, enabling users to select a role and display its associated members. 
Each role is represented as a section in the report, and embedded JavaScript allows seamless toggling between roles. 
CSS styling ensures the report is visually appealing and easy to navigate. 
Validation steps ensure the HTML file is successfully created, and the script automatically opens the report in the system's default browser upon completion.

This script incorporates robust error handling, with try and catch blocks ensuring errors are logged and handled gracefully. 
Its modular design improves readability, maintenance, and scalability, making it a powerful tool for managing and auditing Microsoft 365 administrative accounts. 
By reducing manual effort and providing interactive, well-organized reports, it enhances efficiency and usability for IT administrators.
#>
##################################################################################################################################################################
#==============================Beginning of script================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================Functions==========================================================================================================================
##################################################################################################################################################################
# Define the global log file path at the start of the script
$Global:LogFile = "C:\RSanchezC169_Script_Logs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"

Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,                  # The message to log
        [Parameter()]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
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
        [string[]]$Modules
    )

    Write-Log -Message "Started Function Load-Module" -Level "INFO"

    # Initialize progress variables
    $TotalModules = $Modules.Count
    $ModuleIndex = 0

    foreach ($Module in $Modules) {
        $ModuleIndex++
        $PercentComplete = [math]::Round(($ModuleIndex / $TotalModules) * 100)

        Write-Progress -Activity "Processing Modules" -Status "Processing module '$Module' ($ModuleIndex of $TotalModules)" -PercentComplete $PercentComplete
        Write-Log -Message "Processing module '$Module'..." -Level "INFO"

        try {
            # Step 1: Check if the module is installed
            Write-Progress -Activity "Processing Modules" -Status "Checking if '$Module' is installed..." -PercentComplete $PercentComplete
            if ((Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue) -or (Get-Module -Name $Module -ErrorAction SilentlyContinue)) {
                Write-Log -Message "Module '$Module' is already installed." -Level "INFO"
            } else {
                Write-Log -Message "Module '$Module' is not installed. Attempting to install..." -Level "WARNING"
                Install-Module -Name $Module -Force -Scope CurrentUser -ErrorAction Stop
                Write-Log -Message "Module '$Module' installed successfully." -Level "INFO"
            }

            # Step 2: Check for module updates
            Write-Progress -Activity "Processing Modules" -Status "Checking updates for '$Module'..." -PercentComplete $PercentComplete
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
            Write-Progress -Activity "Processing Modules" -Status "Importing '$Module'..." -PercentComplete $PercentComplete
            Write-Log -Message "Importing module '$Module'..." -Level "INFO"
            Import-Module -Name $Module -Force -ErrorAction Stop
            Write-Log -Message "Module '$Module' imported successfully." -Level "INFO"
        } catch {
            # Log any errors that occur while processing the module
            Write-Log -Message "Error processing module '$Module': $($_.Exception.Message)" -Level "ERROR"
        }
    }

    # Clear progress bar upon completion
    Write-Progress -Activity "Processing Modules" -Status "Completed" -PercentComplete 100 -Completed
    Write-Log -Message "Ended Function Load-Module" -Level "INFO"
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Set-Environment {
    [CmdletBinding()]
    Param()

    Write-Log -Message "Started Function Set-Environment" -Level "INFO"

    try {
        # Initialize progress variables
        $TotalSteps = 7
        $CurrentStep = 0

        # Step 1: Clear the console
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Clearing console..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Clearing console..." -Level "INFO"
        Clear-Host

        # Step 2: Maximize console window
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Maximizing console window..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Maximizing console window..." -Level "INFO"
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
        Write-Log -Message "Console window maximized." -Level "INFO"

        # Step 3: Set execution policy
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Setting execution policy..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Setting execution policy to 'Unrestricted'..." -Level "INFO"
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force
        Write-Log -Message "Execution policy set successfully." -Level "INFO"

        # Step 4: Load required modules
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Loading required modules..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Loading required modules..." -Level "INFO"
        Load-Module -Modules "PowerShellGet", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.Governance"
        Write-Log -Message "Required modules loaded successfully." -Level "INFO"

        # Step 5: Configure console appearance
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Configuring console appearance..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Configuring console appearance: Background=Black, Foreground=Blue, Title='Export Emails'..." -Level "INFO"
        $Host.UI.RawUI.BackgroundColor = 'Black'
        $Host.UI.RawUI.ForegroundColor = 'Blue'
        $Host.UI.RawUI.WindowTitle = "Export Emails"
        Write-Log -Message "Console appearance configured successfully." -Level "INFO"

        # Step 6: Set session preferences
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Configuring session preferences..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Configuring session preferences..." -Level "INFO"
        $Global:FormatEnumerationLimit = -1
        $Global:ErrorActionPreference = "SilentlyContinue"
        $Global:WarningActionPreference = "SilentlyContinue"
        $Global:InformationActionPreference = "SilentlyContinue"
        Write-Log -Message "Session preferences configured successfully." -Level "INFO"

        # Step 7: Finalize setup
        $CurrentStep++
        Write-Progress -Activity "Environment Setup" -Status "Finalizing setup..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Finalizing setup..." -Level "INFO"
        Write-Log -Message "Environment setup completed successfully." -Level "INFO"
    } catch {
        Write-Log -Message "Error during environment setup: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Clean up resources
        Write-Progress -Activity "Environment Setup" -Status "Cleanup in progress..." -PercentComplete 100 -Completed
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Log -Message "Cleaned up resources." -Level "INFO"
    }

    Write-Log -Message "Ended Function Set-Environment" -Level "INFO"
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Connect-MgGraphEnhanced {
    [CmdletBinding()]
    Param()

    Write-Log -Message "Started Function Connect-MgGraphEnhanced" -Level "INFO"

    try {
        # Initialize progress variables
        $TotalSteps = 4
        $CurrentStep = 0

        # Step 1: Check if already connected
        $CurrentStep++
        Write-Progress -Activity "Microsoft Graph Connection" -Status "Checking connection status..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Checking if already connected to Microsoft Graph..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext -and $GraphContext.Account -and $GraphContext.TenantId) {
            Write-Log -Message "Already connected to Microsoft Graph as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -Level "INFO"
        } else {
            # Step 2: Define the scopes for connection
            $CurrentStep++
            Write-Progress -Activity "Microsoft Graph Connection" -Status "Defining connection scopes..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
            $Scopes = @(
                "Directory.Read.All",
	     "RoleManagement.Read.Directory",
	     "User.Read.All"
            )
            Write-Log -Message "Defining connection scopes: $($Scopes -join ', ')" -Level "INFO"

            # Step 3: Attempt connection with retry logic
            $CurrentStep++
            Write-Progress -Activity "Microsoft Graph Connection" -Status "Attempting connection..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
            Write-Log -Message "Attempting connection to Microsoft Graph with scopes: $($Scopes -join ', ')" -Level "INFO"

            $MaxAttempts = 3
            $Attempt = 0
            $Connected = $false

            while (-not $Connected -and $Attempt -lt $MaxAttempts) {
                try {
                    $Attempt++
                    Write-Log -Message "Connection attempt $Attempt of $MaxAttempts..." -Level "INFO"
                    Write-Progress -Activity "Microsoft Graph Connection" -Status "Attempt $Attempt of $MaxAttempts..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
                    Connect-MgGraph -Scope $Scopes -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    $Connected = $true
                    Write-Log -Message "Successfully connected to Microsoft Graph!" -Level "INFO"
                } catch {
                    Write-Log -Message "Connection attempt $Attempt failed. Error: $($_.Exception.Message)" -Level "WARNING"
                    Start-Sleep -Seconds 5
                }
            }

            if (-not $Connected) {
                $ErrorMessage = "All connection attempts to Microsoft Graph failed."
                Write-Log -Message $ErrorMessage -Level "ERROR"
                throw $ErrorMessage
            }

            # Step 4: Retrieve and log connection context
            $CurrentStep++
            Write-Progress -Activity "Microsoft Graph Connection" -Status "Retrieving connection context..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
            $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            Write-Log -Message "Successfully connected to Microsoft Graph as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -Level "INFO"
        }
    } catch {
        # Handle connection errors
        Write-Progress -Activity "Microsoft Graph Connection" -Status "An error occurred!" -PercentComplete 100 -Completed
        $ErrorMessage = $($_.Exception.Message)
        Write-Log -Message "Error connecting to Microsoft Graph: $ErrorMessage" -Level "ERROR"
        throw
    } finally {
        # Finalize progress
        Write-Progress -Activity "Microsoft Graph Connection" -Status "Completed." -PercentComplete 100 -Completed
        Write-Log -Message "Ended Function Connect-MgGraphEnhanced" -Level "INFO"
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-AdministrativeRoles {
    [CmdletBinding()]
    Param()

    # Initialize progress variables
    $TotalSteps = 2
    $CurrentStep = 0

    try {
        # Step 1: Update progress bar for retrieving roles
        $CurrentStep++
        Write-Progress -Activity "Retrieving Administrative Roles" -Status "Retrieving all administrative roles..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Retrieving administrative roles..." -Level "INFO"

        # Retrieve administrative roles
        $adminRoles = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -like "*Administrator" }

        # Check if roles exist
        if (-not $adminRoles) {
            Write-Log -Message "No administrative roles found in the tenant." -Level "WARNING"
            Write-Progress -Activity "Retrieving Administrative Roles" -Status "No roles found" -PercentComplete 100 -Completed
            throw "No administrative roles found."
        }

        Write-Log -Message "Retrieved $($adminRoles.Count) administrative roles." -Level "INFO"

        # Step 2: Finalize progress
        $CurrentStep++
        Write-Progress -Activity "Retrieving Administrative Roles" -Status "Completed" -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        return $adminRoles
    } catch {
        # Handle errors
        Write-Progress -Activity "Retrieving Administrative Roles" -Status "An error occurred" -PercentComplete 100 -Completed
        Write-Log -Message "Failed to retrieve administrative roles. Error: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Ensure progress bar is completed
        Write-Progress -Activity "Retrieving Administrative Roles" -Status "Task Complete" -PercentComplete 100 -Completed
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-RoleMembers {
    [CmdletBinding()]
    Param(
        [array]$AdminRoles
    )

    # Initialize progress variables
    $allRolesMembers = @()
    $totalRoles = $AdminRoles.Count
    $currentRole = 0

    foreach ($role in $AdminRoles) {
        $currentRole++
        $percentComplete = [math]::Round(($currentRole / $totalRoles) * 100)

        # Update progress bar
        Write-Progress -Activity "Retrieving Role Members" -Status "Processing Role $($currentRole) of $($totalRoles): $($role.DisplayName)" -PercentComplete $percentComplete
        Write-Log -Message "Processing role: $($role.DisplayName)" -Level "INFO"

        try {
            # Fetch role assignments
            Write-Log -Message "Retrieving role assignments for role: $($role.DisplayName)..." -Level "INFO"
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "RoleDefinitionId eq '$($role.Id)'"

            if ($assignments) {
                Write-Log -Message "Found $($assignments.Count) assignments for role: $($role.DisplayName)" -Level "INFO"

                # Process each assignment
                foreach ($assignment in $assignments) {
                    $memberId = $assignment.PrincipalId
                    Write-Log -Message "Fetching details for member ID: $memberId" -Level "INFO"

                    $memberDetails = Get-MgUser -UserId $memberId -ErrorAction SilentlyContinue
                    if ($memberDetails) {
                        $allRolesMembers += [PSCustomObject]@{
                            Name         = $memberDetails.DisplayName
                            EmailAddress = $memberDetails.UserPrincipalName
                            Role         = $role.DisplayName
                        }
                        Write-Log -Message "Added member: $($memberDetails.DisplayName) to the list for role: $($role.DisplayName)" -Level "INFO"
                    } else {
                        Write-Log -Message "Failed to retrieve details for member ID: $memberId" -Level "WARNING"
                    }
                }
            } else {
                Write-Log -Message "No members found for role: $($role.DisplayName)" -Level "INFO"
            }
        } catch {
            # Handle any errors during role member retrieval
            Write-Log -Message "Failed to retrieve members for role: $($role.DisplayName). Error: $($_.Exception.Message)" -Level "ERROR"
        }
    }

    # Finalize progress bar
    Write-Progress -Activity "Retrieving Role Members" -Status "Completed" -PercentComplete 100 -Completed
    Write-Log -Message "Completed retrieving role members for all roles." -Level "INFO"

    return $allRolesMembers
}
##################################################################################################################################################################
##################################################################################################################################################################
function Generate-HTMLReport {
    param (
        [array]$AllRolesMembers,
        [string]$OutputHtmlPath
    )

    if ($AllRolesMembers.Count -eq 0) {
        Write-Log "No administrative accounts found to generate an HTML report."
        throw "No administrative accounts found."
    }

    # Group members by role
    $groupedRoles = $AllRolesMembers | Group-Object -Property Role
    Write-Log "Building HTML report with $($AllRolesMembers.Count) accounts."

    # Begin building the HTML content
    $htmlContent = @"
<html>
<head>
    <title>Microsoft 365 Admin Accounts</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; }
        .role { display: none; }
        h1 { text-align: center; }
    </style>
    <script>
        function showRole(roleId) {
            const roles = document.querySelectorAll('.role');
            roles.forEach(role => {
                if (roleId === 'all') {
                    role.style.display = 'block';
                } else if (role.id === roleId) {
                    role.style.display = 'block';
                } else {
                    role.style.display = 'none';
                }
            });
        }
    </script>
</head>
<body>
    <h1>Microsoft 365 Admin Accounts</h1>
    <label for="roleSelect">Select a Role:</label>
    <select id="roleSelect" onchange="showRole(this.value)">
        <option value="all">All Roles</option>
"@

    # Add dropdown options for roles
    foreach ($group in $groupedRoles) {
        $roleId = $group.Name -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $htmlContent += "<option value='$roleId'>$($group.Name)</option>"
    }

    $htmlContent += "</select>"

    # Add sections for each role and its members
    foreach ($group in $groupedRoles) {
        $roleId = $group.Name -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $htmlContent += "<div class='role' id='$roleId'>"
        $htmlContent += @"
        <h2>$($group.Name)</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Email Address</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($member in $group.Group) {
            $htmlContent += @"
                <tr>
                    <td>$($member.Name)</td>
                    <td>$($member.EmailAddress)</td>
                </tr>
"@
        }

        $htmlContent += @"
            </tbody>
        </table>
        </div>
"@
    }

    # Finalize HTML content
    $htmlContent += @"
</body>
</html>
"@

    # Ensure the output directory exists
    $OutputDirectory = Split-Path -Path $OutputHtmlPath -Parent
    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created missing output directory: $OutputDirectory"
    }

    # Write HTML content to file
    try {
        $htmlContent | Set-Content -Path $OutputHtmlPath -Force
        Write-Log "HTML report successfully written to: $OutputHtmlPath"
    } catch {
        Write-Log "Failed to write HTML report. Error: $_"
        throw
    }

    # Validate if the file was created
    if (-not (Test-Path -Path $OutputHtmlPath)) {
        Write-Log "Error: HTML file does not exist at the specified path: $OutputHtmlPath"
        throw "HTML file was not created."
    } else {
        Write-Log "HTML file exists at: $OutputHtmlPath"
    }

    # Automatically open the file
    Write-Log "Opening the HTML file: $OutputHtmlPath"
    Start-Process -FilePath $OutputHtmlPath -Verb Open
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Validate-AndOpenFile {
    [CmdletBinding()]
    Param(
        [string]$FilePath
    )

    # Initialize progress variables
    $TotalSteps = 3
    $CurrentStep = 0

    try {
        # Step 1: Validate if the file exists
        $CurrentStep++
        Write-Progress -Activity "File Validation" -Status "Validating file path..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        if (-not (Test-Path -Path $FilePath)) {
            Write-Progress -Activity "File Validation" -Status "Validation failed." -PercentComplete 100 -Completed
            Write-Log -Message "Error: File does not exist at the specified path: $FilePath" -Level "ERROR"
            throw "File was not created."
        }
        Write-Log -Message "File exists and is ready to open: $FilePath" -Level "INFO"

        # Step 2: Log that the file is being opened
        $CurrentStep++
        Write-Progress -Activity "File Validation" -Status "Preparing to open file..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Opening the file: $FilePath" -Level "INFO"

        # Step 3: Open the file
        $CurrentStep++
        Write-Progress -Activity "File Validation" -Status "Opening file..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Start-Process -FilePath $FilePath -Verb Open
        Write-Log -Message "File successfully opened: $FilePath" -Level "INFO"
    } catch {
        # Handle errors during file opening
        Write-Progress -Activity "File Validation" -Status "Error encountered." -PercentComplete 100 -Completed
        Write-Log -Message "Failed to open the file: $FilePath. Error: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Finalize progress
        Write-Progress -Activity "File Validation" -Status "Completed" -PercentComplete 100 -Completed
    }
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
Write-Log -Message "Main execution started." -Level "INFO"

# Initialize progress variables
$CurrentStep = 0
$TotalSteps = 8  # Total steps in the main execution

try {
    # Step 1: Check Windows Version
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Checking Windows version..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Checking Operating System version..." -Level "INFO"
    if ([System.Environment]::OSVersion.Version.Major -lt 10) {
        Write-Log -Message "Error: This script requires Windows 10 or above. Current OS version does not meet the requirement." -Level "ERROR"
        throw "This script requires Windows 10 or above."
    }
    Write-Log -Message "OS version check passed: Windows 10 or above detected." -Level "INFO"
    Clear-Host

    # Step 2: Check PowerShell Version
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Checking PowerShell version..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Checking PowerShell version..." -Level "INFO"
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Log -Message "Error: This script requires PowerShell version 5 or above. Current version does not meet the requirement." -Level "ERROR"
        throw "This script requires PowerShell version 5 or above."
    }
    Write-Log -Message "PowerShell version check passed: Version 5 or above detected." -Level "INFO"
    Clear-Host

    # Step 3: Check Administrative Privileges
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Checking administrative privileges..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Checking administrative privileges..." -Level "INFO"
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $IsAdmin) {
        Write-Log -Message "Error: Script is not running with administrative privileges. Please run as Administrator." -Level "ERROR"
        throw "Administrative privileges are required to run this script."
    }
    Write-Log -Message "Administrative privileges confirmed." -Level "INFO"
    Clear-Host

    # Step 4: Set up the environment
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Setting up the environment..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Setting up the environment..." -Level "INFO"
    Set-Environment
    Write-Log -Message "Environment setup completed successfully." -Level "INFO"
    Clear-Host

    # Step 5: Connect to Microsoft Graph
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Connecting to Microsoft Graph..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Connecting to Microsoft Graph..." -Level "INFO"
    Connect-MgGraphEnhanced
    Write-Log -Message "Microsoft Graph connection established successfully." -Level "INFO"
    Clear-Host

    # Step 6: Retrieve administrative roles
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving administrative roles..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Retrieving administrative roles..." -Level "INFO"
    $adminRoles = Get-AdministrativeRoles
    Clear-Host

    # Step 7: Retrieve role members
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving role members..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Retrieving role members..." -Level "INFO"
    $allRolesMembers = Get-RoleMembers -AdminRoles $adminRoles
    Clear-Host

    # Step 8: Generate HTML report
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Generating HTML report..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Generating HTML report..." -Level "INFO"
    $outputPath = "C:\RSanchezC169_Reports\AdminAccounts.html"
    Generate-HTMLReport -AllRolesMembers $allRolesMembers -OutputHtmlPath $outputPath
    Write-Log -Message "HTML report generated successfully: $outputPath" -Level "INFO"
    Clear-Host

    # Mark main execution as complete
    Write-Log -Message "Main execution completed successfully." -Level "INFO"
} catch {
    # Log any errors encountered during the main execution
    Write-Progress -Activity "Main Execution" -Status "An error occurred" -PercentComplete 100 -Completed
    $ErrorMessage = $($_.Exception.Message)
    Write-Log -Message "An error occurred during main execution: $ErrorMessage" -Level "ERROR"
    throw
} finally {
    Write-Progress -Activity "Main Execution" -Status "Cleaning up..." -PercentComplete 100 -Completed
    Write-Log -Message "Main execution ended." -Level "INFO"

    # Clean up resources at the end of the script
    Write-Log -Message "Script execution completed. Performing cleanup..." -Level "INFO"
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Log -Message "Script Ended" -Level "INFO"

    # Open the log file for review
    Start-Process notepad.exe -ArgumentList $Global:LogFile

    Clear-Host
}
##################################################################################################################################################################
#==============================End of Main========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script======================================================================================================================
##################################################################################################################################################################
