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
        Load-Module -Modules "PowerShellGet", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement"
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
                "User.Read.All",
	     "Directory.Read.All",
	    "LicenseAssignment.Read.All"
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
Function Get-Licenses {
    [CmdletBinding()]
    Param()

    # Initialize progress variables
    $TotalSteps = 3
    $CurrentStep = 0

    try {
        # Step 1: Start progress for license retrieval
        $CurrentStep++
        Write-Progress -Activity "Retrieving Licenses" -Status "Fetching available licenses..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Retrieving all available licenses in the tenant..." -Level "INFO"

        # Retrieve all available licenses
        $licenses = Get-MgSubscribedSku
        if (-not $licenses) {
            Write-Log -Message "No licenses found in the tenant." -Level "WARNING"
            Write-Progress -Activity "Retrieving Licenses" -Status "No licenses found" -PercentComplete 100 -Completed
            throw "No licenses found in the tenant."
        }

        # Step 2: Process licenses with usage details
        $CurrentStep++
        Write-Progress -Activity "Retrieving Licenses" -Status "Processing license details..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Processing license details..." -Level "INFO"

        # Include license counts (available and assigned)
        $licensesWithDetails = $licenses | ForEach-Object {
            [PSCustomObject]@{
                SkuPartNumber  = $_.SkuPartNumber
                SkuId          = $_.SkuId
                PrepaidUnits   = $_.PrepaidUnits.Enabled # Total available licenses
                ConsumedUnits  = $_.ConsumedUnits        # Assigned licenses
            }
        }

        Write-Log -Message "Retrieved $($licensesWithDetails.Count) licenses with usage details." -Level "INFO"

        # Step 3: Finalize progress
        $CurrentStep++
        Write-Progress -Activity "Retrieving Licenses" -Status "Completed retrieval of license data." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))

        return $licensesWithDetails
    } catch {
        # Handle errors during license retrieval
        Write-Progress -Activity "Retrieving Licenses" -Status "An error occurred during license retrieval." -PercentComplete 100 -Completed
        Write-Log -Message "Failed to retrieve licenses. Error: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Ensure progress bar completion
        Write-Progress -Activity "Retrieving Licenses" -Status "Task Complete" -PercentComplete 100 -Completed
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-UsersByLicense {
    [CmdletBinding()]
    Param(
        [array]$Licenses
    )

    # Initialize progress variables
    $allLicenseUsers = @()
    Write-Log -Message "Retrieving all users and their assigned licenses..." -Level "INFO"

    try {
        # Fetch all users in the tenant
        Write-Log -Message "Fetching all users in the tenant..." -Level "INFO"
        $allUsers = Get-MgUser -All
        if (-not $allUsers) {
            Write-Log -Message "No users found in the tenant." -Level "WARNING"
            throw "No users found."
        }

        # Initialize progress bar
        $totalUsers = $allUsers.Count
        $currentUser = 0

        # Iterate through each user to retrieve license details
        foreach ($user in $allUsers) {
            $currentUser++
            $percentComplete = [math]::Round(($currentUser / $totalUsers) * 100)

            # Update progress bar
            Write-Progress -Activity "Retrieving Users by License" -Status "Processing user $currentUser of $($totalUsers): $($user.DisplayName)" -PercentComplete $percentComplete
            Write-Log -Message "Processing licenses for user: $($user.DisplayName)" -Level "INFO"
            try {
                # Fetch license details for the user
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
                foreach ($license in $licenseDetails) {
                    $skuPartNumber = ($Licenses | Where-Object { $_.SkuId -eq $license.SkuId }).SkuPartNumber
                    if ($skuPartNumber) {
                        $allLicenseUsers += [PSCustomObject]@{
                            UserName      = $user.DisplayName
                            EmailAddress  = $user.UserPrincipalName
                            License       = $skuPartNumber
                        }
                        Write-Log -Message "User $($user.DisplayName) assigned license: $skuPartNumber" -Level "INFO"
                    }
                }
            } catch {
                Write-Log -Message "Failed to retrieve license details for user: $($user.DisplayName). Error: $($_.Exception.Message)" -Level "WARNING"
            }
        }
    } catch {
        Write-Progress -Activity "Retrieving Users by License" -Status "An error occurred" -PercentComplete 100 -Completed
        Write-Log -Message "Failed to retrieve users. Error: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Finalize progress bar
        Write-Progress -Activity "Retrieving Users by License" -Status "Completed processing all users." -PercentComplete 100 -Completed
    }

    if ($allLicenseUsers.Count -eq 0) {
        Write-Log -Message "No license-user assignments found." -Level "INFO"
    }

    return $allLicenseUsers
}
##################################################################################################################################################################
##################################################################################################################################################################
function Generate-HTMLReport {
    param (
        [array]$AllLicenseUsers,
        [array]$Licenses,
        [string]$OutputHtmlPath
    )

    if ($AllLicenseUsers.Count -eq 0) {
        Write-Log "No license-user data found to generate an HTML report."
        throw "No license-user data available."
    }

    # Group users by license
    $groupedLicenses = $AllLicenseUsers | Group-Object -Property License
    Write-Log "Building HTML report with $($AllLicenseUsers.Count) user-license assignments."

    # Begin building the HTML content
    $htmlContent = @"
<html>
<head>
    <title>Microsoft 365 License Assignments</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; }
        .license { display: none; }
        h1 { text-align: center; }
    </style>
    <script>
        function showLicense(licenseId) {
            const licenses = document.querySelectorAll('.license');
            licenses.forEach(license => {
                if (licenseId === 'all') {
                    license.style.display = 'block';
                } else if (license.id === licenseId) {
                    license.style.display = 'block';
                } else {
                    license.style.display = 'none';
                }
            });
        }
    </script>
</head>
<body>
    <h1>Microsoft 365 License Assignments</h1>
    <label for="licenseSelect">Select a License:</label>
    <select id="licenseSelect" onchange="showLicense(this.value)">
        <option value="all">All Licenses</option>
"@

    # Add dropdown options for licenses and their usage details
    foreach ($license in $Licenses) {
        $licenseId = $license.SkuPartNumber -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $availableLicenses = $license.PrepaidUnits - $license.ConsumedUnits
        $htmlContent += "<option value='$licenseId'>$($license.SkuPartNumber) (Available: $availableLicenses)</option>"
    }

    $htmlContent += "</select>"

    # Add sections for each license and its users
    foreach ($group in $groupedLicenses) {
        $licenseId = $group.Name -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $licenseDetails = $Licenses | Where-Object { $_.SkuPartNumber -eq $group.Name }
        $availableLicenses = $licenseDetails.PrepaidUnits - $licenseDetails.ConsumedUnits

        $htmlContent += "<div class='license' id='$licenseId'>"
        $htmlContent += @"
        <h2>$($group.Name)</h2>
        <p><strong>Total Licenses:</strong> $($licenseDetails.PrepaidUnits) | <strong>Assigned:</strong> $($licenseDetails.ConsumedUnits) | <strong>Available:</strong> $availableLicenses</p>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Email Address</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($user in $group.Group) {
            $htmlContent += @"
                <tr>
                    <td>$($user.UserName)</td>
                    <td>$($user.EmailAddress)</td>
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
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
Write-Log -Message "Main execution started." -Level "INFO"

# Initialize progress variables
$CurrentStep = 0
$TotalSteps = 8  # Total steps in this main execution flow

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

    # Step 6: Retrieve all licenses
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving licenses..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Retrieving all licenses with their details..." -Level "INFO"
    $licenses = Get-Licenses
    Clear-Host

    # Step 7: Retrieve users assigned to licenses
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving users for licenses..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Retrieving users assigned to licenses..." -Level "INFO"
    $licenseUsers = Get-UsersByLicense -Licenses $licenses
    Clear-Host

    # Step 8: Generate the HTML report
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Generating HTML report..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Generating the HTML report with license and user details..." -Level "INFO"
    $outputPath = "C:\RSanchezC169_Reports\LicenseAssignments.html"
    Generate-HTMLReport -AllLicenseUsers $licenseUsers -Licenses $licenses -OutputHtmlPath $outputPath
    Write-Log -Message "HTML report generated successfully: $outputPath" -Level "INFO"
    Clear-Host

    Write-Log -Message "Main execution completed successfully." -Level "INFO"
} catch {
    # Log any errors encountered during the main execution
    Write-Progress -Activity "Main Execution" -Status "An error occurred" -PercentComplete 100 -Completed
    $ErrorMessage = $($_.Exception.Message)
    Write-Log -Message "An error occurred during main execution: $ErrorMessage" -Level "ERROR"
    throw
} finally {
    Write-Progress -Activity "Main Execution" -Status "Cleaning up resources..." -PercentComplete 100 -Completed
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
