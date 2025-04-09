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
Description of Script: This script is to remove a user from all conditional access polices exclusion list.
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
        Load-Module -Modules "PowerShellGet", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.SignIns"
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
	              "Policy.Read.All",
	              "Policy.ReadWrite.ConditionalAccess",
	               "Policy.ReadWrite.SecurityDefaults"
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
Function Remove-AdminToCAPolicyExclusion {
    [CmdletBinding()]
    param (
        [string]$ExclusionInput # Accepts either an email address or a domain name
    )

    Write-Log -Message "Started Function Add-AdminToCAPolicyExclusion" -Level "INFO"

    try {
        # Validate that ExclusionInput is a valid Graph user
        $validUser = $false
        $attempts = 0
        while (-not $validUser) {
            try {
                Write-Log -Message "Validating ExclusionInput: $ExclusionInput..." -Level "INFO"
                $ExclusionUser = Get-MgUser -UserId $ExclusionInput
                if ($ExclusionUser) {
                    Write-Log -Message "ExclusionInput '$ExclusionInput' validated successfully as a Graph user." -Level "INFO"
                    $validUser = $true
                } else {
                    throw "Invalid user data provided."
                }
            } catch {
                $attempts++
                if ($attempts -ge 3) {
                    Write-Log -Message "Error: Exceeded maximum attempts to validate user details for ExclusionInput." -Level "ERROR"
                    throw "Failed to validate user after multiple attempts."
                }

                Write-Log -Message "Error: Failed to retrieve user details for ExclusionInput '$ExclusionInput'. Exception: $($_.Exception.Message)" -Level "ERROR"
                Write-Log -Message "Reprompting user for a valid email address." -Level "WARNING"
                $ExclusionInput = Read-Host "Please enter a valid email address for the exclusion removal"
            }
        }

        # Retrieve all conditional access policies
        Write-Log -Message "Retrieving all conditional access policies..." -Level "INFO"
        $policies = Get-MgIdentityConditionalAccessPolicy

        if (-not $policies -or $policies.Count -eq 0) {
            Write-Log -Message "No conditional access policies found in the tenant." -Level "WARNING"
            return
        }

        # Initialize progress variables
        $TotalPolicies = $policies.Count
        $CurrentPolicy = 0

        # Loop through each policy and update exclusions
        foreach ($policy in $policies) {
            $CurrentPolicy++

            # Update progress bar
            $ProgressPercentage = ($CurrentPolicy / $TotalPolicies) * 100
            Write-Progress -Activity "Updating Conditional Access Policies" `
                -Status "Processing policy $CurrentPolicy of $TotalPolicies" `
                -PercentComplete $ProgressPercentage

            Write-Log -Message "Updating policy: $($policy.DisplayName)" -Level "INFO"

            try {
                # Retrieve the current exclusions
                $currentExclusions = @()
                if ($policy.Conditions.Users.ExcludeUsers) {
                    $currentExclusions = $policy.Conditions.Users.ExcludeUsers
                }
	
                # Remove the validated Graph user as an exclusion
                $updatedExclusions = @($currentExclusions)
                $updatedExclusions = $updatedExclusions | Where-Object {$_  -NE $ExclusionUser.Id}
	    
	
                # Update the policy to include the updated exclusions
                $bodyParameters = @{
                    conditions = @{
                        users = @{
                            excludeusers = $updatedExclusions
                        }
                    }
                }
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $bodyParameters

                Write-Log -Message "Policy '$($policy.DisplayName)' has been updated with the exclusion successfully." -Level "INFO"
            } catch {
                Write-Log -Message "Failed to update policy '$($policy.DisplayName)'. Error: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Complete progress
        Write-Progress -Activity "Updating Conditional Access Policies" -Status "Complete" -Completed

        Write-Log -Message "All applicable policies have been updated with the specified exclusion." -Level "INFO"

    } catch {
        Write-Log -Message "An error occurred during execution: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        Write-Log -Message "Ended Function Add-AdminToCAPolicyExclusion" -Level "INFO"
    }
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
Write-Log -Message "Main execution started." -Level "INFO"

try {
    # Step 1: Check Windows Version
    Write-Log -Message "Checking Operating System version..." -Level "INFO"
    if ([System.Environment]::OSVersion.Version.Major -lt 10) {
        Write-Log -Message "Error: This script requires Windows 10 or above. Current OS version does not meet the requirement." -Level "ERROR"
        throw "This script requires Windows 10 or above."
    }
    Write-Log -Message "OS version check passed: Windows 10 or above detected." -Level "INFO"
    Clear-Host

    # Step 2: Check PowerShell Version
    Write-Log -Message "Checking PowerShell version..." -Level "INFO"
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Log -Message "Error: This script requires PowerShell version 5 or above. Current version does not meet the requirement." -Level "ERROR"
        throw "This script requires PowerShell version 5 or above."
    }
    Write-Log -Message "PowerShell version check passed: Version 5 or above detected." -Level "INFO"
    Clear-Host

    # Step 3: Check Administrative Privileges
    Write-Log -Message "Checking administrative privileges..." -Level "INFO"
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $IsAdmin) {
        Write-Log -Message "Error: Script is not running with administrative privileges. Please run as Administrator." -Level "ERROR"
        throw "Administrative privileges are required to run this script."
    }
    Write-Log -Message "Administrative privileges confirmed." -Level "INFO"
    Clear-Host

    # Step 4: Set up the environment
    Write-Log -Message "Setting up the environment..." -Level "INFO"
    Set-Environment
    Write-Log -Message "Environment setup completed successfully." -Level "INFO"
    Clear-Host

    # Step 5: Connect to Microsoft Graph
    Write-Log -Message "Connecting to Microsoft Graph..." -Level "INFO"
    Connect-MgGraphEnhanced
    Write-Log -Message "Microsoft Graph connection established successfully." -Level "INFO"
    Clear-Host

    # Step 6: Remove user from all conditional access policies exclusion list
    Write-Log -Message "Removing user from Conditional Access Policies exclusion list..." -Level "INFO"
    try {
        #$ExclusionInput = Read-Host "Enter the email address of the admin user to add to exclusions"
        Remove-AdminToCAPolicyExclusion -ExclusionInput $(Read-Host "Enter the email address of the admin user to add to exclusions")
        Write-Log -Message "User successfully remove from Conditional Access Policies exclusion list." -Level "INFO"
    } catch {
        Write-Log -Message "Error while removing user from exclusions: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
    Clear-Host

    Write-Log -Message "Main execution completed successfully." -Level "INFO"
} catch {
    # Log any errors encountered during the main execution
    $ErrorMessage = $($_.Exception.Message)
    Write-Log -Message "An error occurred during main execution: $ErrorMessage" -Level "ERROR"
    throw
} finally {
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
