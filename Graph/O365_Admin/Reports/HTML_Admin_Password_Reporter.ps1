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
        Load-Module -Modules "PowerShellGet", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users","Microsoft.GraphIdentity.Governance"
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
               "RoleManagement.Read.Directory"
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
Function Get-UsersByRole {
    [CmdletBinding()]
    Param()

    Write-Log -Message "Started Function: Get-UsersByRole" -Level "INFO"

    # Initialize progress variables
    $TotalSteps = 3
    $CurrentStep = 0

    try {
        # Step 1: Retrieve all user accounts in the tenant
        $CurrentStep++
        Write-Progress -Activity "Get Users by Role" -Status "Fetching all user accounts..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Fetching all user accounts from the tenant..." -Level "INFO"
        $AllUsers = Invoke-WithRetry -ScriptBlock {
            Get-MgUser -All -Select "Id,DisplayName,UserPrincipalName" -ErrorAction Stop
        }

        if (-not $AllUsers -or $AllUsers.Count -eq 0) {
            Write-Log -Message "No users found in the tenant." -Level "WARNING"
            Write-Host "No users found in the tenant." -ForegroundColor Yellow
            Write-Progress -Activity "Get Users by Role" -Status "No users found" -PercentComplete 100 -Completed
            return @{}
        }

        Write-Log -Message "Total users retrieved: $($AllUsers.Count)" -Level "INFO"

        # Step 2: Retrieve all administrative roles and their assignments
        $CurrentStep++
        Write-Progress -Activity "Get Users by Role" -Status "Fetching administrative role assignments..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        Write-Log -Message "Fetching all administrative role assignments..." -Level "INFO"
        $AdminAssignments = Invoke-WithRetry -ScriptBlock {
            Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop
        }

        $AdminGroups = @{} # Hashtable to hold admin groups and their members
        $AdminUserIds = @{} # Dictionary to identify admin users by their IDs

        if ($AdminAssignments -and $AdminAssignments.Count -gt 0) {
            Write-Log -Message "Total admin assignments retrieved: $($AdminAssignments.Count)" -Level "INFO"

            # Retrieve details of role definitions
            $RoleDefinitions = Invoke-WithRetry -ScriptBlock {
                Get-MgRoleManagementDirectoryRoleDefinition -All -ErrorAction Stop
            }

            # Map role definitions with their display names and include UserPrincipalName
            foreach ($Assignment in $AdminAssignments) {
                $RoleDefinition = $RoleDefinitions | Where-Object { $_.Id -eq $Assignment.RoleDefinitionId }
                if ($RoleDefinition) {
                    if (-not $AdminGroups.ContainsKey($RoleDefinition.DisplayName)) {
                        $AdminGroups[$RoleDefinition.DisplayName] = @()
                    }

                    # Find user details for the PrincipalId
                    $User = $AllUsers | Where-Object { $_.Id -eq $Assignment.PrincipalId }
                    if ($User) {
                        $AdminGroups[$RoleDefinition.DisplayName] += @{
                            PrincipalId       = $User.Id
                            UserPrincipalName = $User.UserPrincipalName
                        }
                        $AdminUserIds[$User.Id] = $true
                    }
                }
            }
        } else {
            Write-Log -Message "No admin role assignments found. Assuming no admins in the tenant." -Level "INFO"
        }

        # Step 3: Classify users into Admin and Non-Admin arrays
        $CurrentStep++
        Write-Progress -Activity "Get Users by Role" -Status "Classifying admin and non-admin users..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
        $AdminUsers = @()
        $NonAdminUsers = @()

        foreach ($User in $AllUsers) {
            $UserDetails = [PSCustomObject]@{
                Id                 = $User.Id
                DisplayName        = $User.DisplayName
                UserPrincipalName  = $User.UserPrincipalName
            }

            if ($AdminUserIds.ContainsKey($User.Id)) {
                $AdminUsers += $UserDetails
            } else {
                $NonAdminUsers += $UserDetails
            }
        }

        Write-Log -Message "Total Admin Users: $($AdminUsers.Count). Total Non-Admin Users: $($NonAdminUsers.Count)." -Level "INFO"
        Write-Log -Message "Total Admin Groups Found: $($AdminGroups.Keys.Count)" -Level "INFO"

        # Finalize progress
        Write-Progress -Activity "Get Users by Role" -Status "Completed" -PercentComplete 100 -Completed

        return @{
            AdminUsers    = $AdminUsers
            NonAdminUsers = $NonAdminUsers
            AdminGroups   = $AdminGroups
        }
    } catch {
        # Handle errors gracefully
        Write-Progress -Activity "Get Users by Role" -Status "An error occurred" -PercentComplete 100 -Completed
        Write-Log -Message "Error in Get-UsersByRole: $($_.Exception.Message)" -Level "ERROR"
        Write-Warning "An error occurred: $($_.Exception.Message)"
        throw
    } finally {
        Write-Log -Message "Ended Function: Get-UsersByRole" -Level "INFO"
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Get-PasswordDetailsForUsers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$UsersByRole  # Output from Get-UsersByRole with AdminUsers, NonAdminUsers, and AdminGroups
    )

    Write-Log -Message "Started Function: Get-PasswordDetailsForUsers" -Level "INFO"

    try {
        # Combine AdminUsers and NonAdminUsers into a single array for processing
        $AllUsers = $UsersByRole.AdminUsers + $UsersByRole.NonAdminUsers | Sort-Object -Property UserPrincipalName -Unique
        Write-Log -Message "Total users to process: $($AllUsers.Count)" -Level "INFO"

        if (-not $AllUsers -or $AllUsers.Count -eq 0) {
            Write-Log -Message "No users provided to process." -Level "WARNING"
            Write-Progress -Activity "Retrieving Password Details" -Status "No users to process" -PercentComplete 100 -Completed
            return
        }

        # Initialize an array to store password details
        $UserPasswordDetails = @()
        $TotalUsers = $AllUsers.Count
        $CurrentUser = 0

        # Loop through each user and retrieve password-related details
        foreach ($User in $AllUsers) {
            $CurrentUser++
            $PercentComplete = [math]::Round(($CurrentUser / $TotalUsers) * 100)

            # Update progress bar
            Write-Progress -Activity "Retrieving Password Details" `
                           -Status "Processing User $CurrentUser of $TotalUsers ($($User.UserPrincipalName))" `
                           -PercentComplete $PercentComplete

            Write-Log -Message "Processing user: $($User.DisplayName) ($($User.UserPrincipalName))" -Level "INFO"

            try {
                # Fetch password-related details
                $PasswordDetails = Invoke-WithRetry -ScriptBlock {
                    Get-MgUser -UserId $User.Id -Select "id,displayName,userPrincipalName,passwordPolicies,LastPasswordChangeDateTime,accountEnabled"
                }

                if ($PasswordDetails) {
                    # Calculate time since last password change
                    $TimeSinceLastPasswordChange = "N/A"  # Default if no password set date exists
                    if ($PasswordDetails.LastPasswordChangeDateTime -ne $null) {
                        $PasswordLastSetDate = [datetime]$PasswordDetails.LastPasswordChangeDateTime
                        $Today = Get-Date
                        $TimeSpan = $Today - $PasswordLastSetDate

                        $Years = [math]::Floor($TimeSpan.TotalDays / 365)
                        $Months = [math]::Floor(($TimeSpan.TotalDays % 365) / 30)
                        $Days = [math]::Floor(($TimeSpan.TotalDays % 365) % 30)
                        $TimeSinceLastPasswordChange = "${Years} Years, ${Months} Months, ${Days} Days"
                    }

                    # Determine if the user is part of an admin group
                    $GroupMemberships = @()
                    foreach ($Group in $UsersByRole.AdminGroups.Keys) {
                        if ($UsersByRole.AdminGroups[$Group] -contains $User.Id) {
                            $GroupMemberships += $Group
                        }
                    }

                    # Add retrieved details to the results list
                    $UserPasswordDetails += [PSCustomObject]@{
                        DisplayName                 = $PasswordDetails.DisplayName
                        UserPrincipalName           = $PasswordDetails.UserPrincipalName
                        PasswordLastSetDate         = $PasswordDetails.LastPasswordChangeDateTime
                        TimeSinceLastPasswordChange = $TimeSinceLastPasswordChange
                        AccountEnabled              = $PasswordDetails.AccountEnabled
                        PasswordPolicies            = $PasswordDetails.PasswordPolicies
                        AdminGroups                 = ($GroupMemberships -join ", ")
                    }
                    Write-Log -Message "Password details retrieved for user: $($PasswordDetails.DisplayName)" -Level "INFO"
                } else {
                    Write-Log -Message "No password details found for user: $($User.UserPrincipalName)" -Level "WARNING"
                }
            } catch {
                # Handle and log errors for the specific user
                Write-Log -Message "Error retrieving password details for user $($User.UserPrincipalName): $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Complete progress bar
        Write-Progress -Activity "Retrieving Password Details" -Status "Completed" -PercentComplete 100 -Completed

        # Log final results
        Write-Log -Message "Completed processing all users. Total users processed: $($UserPasswordDetails.Count)" -Level "INFO"
        Write-Host "Password details retrieved for $($UserPasswordDetails.Count) users." -ForegroundColor Green

        return $UserPasswordDetails
    } catch {
        # Handle unexpected errors
        Write-Progress -Activity "Retrieving Password Details" -Status "An error occurred" -PercentComplete 100 -Completed
        Write-Log -Message "Error in Get-PasswordDetailsForUsers: $($_.Exception.Message)" -Level "ERROR"
        Write-Warning "An error occurred: $($_.Exception.Message)"
        throw
    } finally {
        Write-Log -Message "Ended Function: Get-PasswordDetailsForUsers" -Level "INFO"
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Create-PasswordDetailsHtml {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [array]$PasswordDetails,  # Array of user password details

        [Parameter(Mandatory)]
        [hashtable]$AdminGroups,  # Hashtable of admin groups and their members

        [Parameter(Mandatory)]
        [string]$FolderPath,      # Path to save the HTML file

        [Parameter(Mandatory)]
        [string]$FileName,        # Base name for the HTML file

        [Parameter(Mandatory)]
        [array]$AllUsers          # Array of all users for filtering by display name
    )

    Write-Log -Message "Started Function: Create-PasswordDetailsHtml" -Level "INFO"

    try {
        # Validate and create folder if necessary
        if (-not (Test-Path -Path $FolderPath)) {
            Write-Log -Message "Folder does not exist. Creating it: $FolderPath" -Level "INFO"
            New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
        }

        # Generate a unique file name if it exists
        $FullFilePath = Join-Path -Path $FolderPath -ChildPath "$FileName.html"
        $FileIndex = 1
        while (Test-Path -Path $FullFilePath) {
            $FullFilePath = Join-Path -Path $FolderPath -ChildPath "$FileName`_$FileIndex.html"
            $FileIndex++
        }

        Write-Log -Message "Generating HTML file at: $FullFilePath" -Level "INFO"
        # Build HTML header with styles and scripts
        $HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Password Details</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { width: 80%; margin: auto; }
        h1 { text-align: center; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; }
        .group { display: none; }
        h2 { background-color: #e8e8e8; }
    </style>
    <script>
        function showGroup(groupId) {
            const groups = document.querySelectorAll('.group');
            groups.forEach(group => {
                if (groupId === 'All') {
                    group.style.display = 'block';
                } else if (group.id === groupId) {
                    group.style.display = 'block';
                } else {
                    group.style.display = 'none';
                }
            });
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>Password Details</h1>
        <label for="groupSelect">Select a Role Group:</label>
        <select id="groupSelect" onchange="showGroup(this.value)">
            <option value="All">All Groups</option>
"@

        # Populate dropdown dynamically with admin group names
        foreach ($Group in $AdminGroups.Keys) {
            $GroupId = $Group -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
            $HtmlHeader += "<option value='$GroupId'>$Group</option>"
        }

        $HtmlHeader += @"
        </select>
"@
        # Group users by admin role groups
        $HtmlBody = ""
        foreach ($Group in $AdminGroups.Keys) {
            $GroupId = $Group -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
            $HtmlBody += "<div class='group' id='$GroupId'>"
            $HtmlBody += "<h2>$Group</h2>"
            $HtmlBody += @"
            <table>
                <thead>
                    <tr>
                        <th>Display Name</th>
                        <th>User Principal Name</th>
                        <th>Password Last Set Date</th>
                        <th>Days Since Last Password Change</th> <!-- New Column -->
                        <th>Account Enabled</th>
                        <th>Password Policies</th>
                    </tr>
                </thead>
                <tbody>
"@

            foreach ($User in $PasswordDetails) {
                # Calculate Days Since Last Password Change
                $DaysSinceLastChange = "N/A"
                if ($User.PasswordLastSetDate -ne $null) {
                    $LastPasswordChangeDate = [datetime]$User.PasswordLastSetDate
                    $Today = Get-Date
                    $TimeSpan = $Today - $LastPasswordChangeDate

                    $Years = [math]::Floor($TimeSpan.TotalDays / 365)
                    $Months = [math]::Floor(($TimeSpan.TotalDays % 365) / 30)
                    $Days = [math]::Floor(($TimeSpan.TotalDays % 365) % 30)
                    $DaysSinceLastChange = "$Years Years, $Months Months, $Days Days"
                }

                # Check if user belongs to this group
                if ($AdminGroups[$Group] | Where-Object { $_.UserPrincipalName -eq $User.UserPrincipalName }) {
                    $HtmlBody += @"
                    <tr>
                        <td>$($User.DisplayName)</td>
                        <td>$($User.UserPrincipalName)</td>
                        <td>$($User.PasswordLastSetDate)</td>
                        <td>$DaysSinceLastChange</td> <!-- New Column -->
                        <td>$($User.AccountEnabled)</td>
                        <td>$($User.PasswordPolicies)</td>
                    </tr>
"@
                }
            }

            $HtmlBody += @"
                </tbody>
            </table>
            </div>
"@
        }
        $HtmlFooter = @"
</div>
</body>
</html>
"@

        # Combine all parts into the final HTML content
        $FullHtmlContent = $HtmlHeader + $HtmlBody + $HtmlFooter

        # Save the HTML content to the file
        Write-Log -Message "Writing HTML content to file..." -Level "INFO"
        $FullHtmlContent | Out-File -FilePath $FullFilePath -Encoding UTF8

        Write-Log -Message "HTML report successfully written to: $FullFilePath" -Level "INFO"
        Write-Host "HTML file created successfully at $FullFilePath" -ForegroundColor Green

        # Open the file automatically
        Start-Process -FilePath $FullFilePath
    } catch {
        Write-Log -Message "Error generating HTML report: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
Function Invoke-WithRetry {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,             # The operation to retry
        [int]$MaxAttempts = 3,                 # Maximum number of retry attempts
        [int]$RetryDelay = 2                   # Delay in seconds between retries
    )

    # Initialize retry variables
    $attempts = 0

    while ($attempts -lt $MaxAttempts) {
        try {
            $attempts++
            $percentComplete = [math]::Round(($attempts / $MaxAttempts) * 100)

            # Update progress bar
            Write-Progress -Activity "Retry Operation" `
                           -Status "Attempt $attempts of $MaxAttempts in progress..." `
                           -PercentComplete $percentComplete

            Write-Log -Message "Attempt $attempts of $MaxAttempts. Executing script block..." -Level "INFO"

            # Execute the operation
            $result = & $ScriptBlock
            Write-Log -Message "Operation succeeded on attempt $attempts." -Level "INFO"

            # Complete progress bar upon success
            Write-Progress -Activity "Retry Operation" -Status "Operation succeeded." -PercentComplete 100 -Completed
            return $result
        } catch {
            # Log failure for this attempt
            Write-Log -Message "Attempt $attempts failed. Error: $($_.Exception.Message)" -Level "WARNING"

            if ($attempts -ge $MaxAttempts) {
                Write-Log -Message "Max attempts reached. Throwing error..." -Level "ERROR"

                # Finalize progress bar upon failure
                Write-Progress -Activity "Retry Operation" -Status "Operation failed after $MaxAttempts attempts." -PercentComplete 100 -Completed
                throw $_
            }

            Write-Log -Message "Retrying operation in $RetryDelay seconds..." -Level "INFO"
            Start-Sleep -Seconds $RetryDelay
        }
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
$TotalSteps = 9  # Total steps including setup, user retrieval, password details, and report generation

try {
    # Step 1: Check Windows Version
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Checking Windows version..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Checking Operating System version..." -Level "INFO"
    if ([System.Environment]::OSVersion.Version.Major -lt 10) {
        Write-Log -Message "Error: This script requires Windows 10 or above. Current OS version does not meet the requirement." -Level "ERROR"
        Write-Progress -Activity "Main Execution" -Status "Failed - OS not compatible." -PercentComplete 100 -Completed
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
        Write-Progress -Activity "Main Execution" -Status "Failed - PowerShell version not compatible." -PercentComplete 100 -Completed
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
        Write-Progress -Activity "Main Execution" -Status "Failed - Admin privileges required." -PercentComplete 100 -Completed
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

    # Step 6: Retrieve Users and Admin Groups
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving users and categorizing roles..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Step 6: Retrieving users and admin groups..." -Level "INFO"
    $UsersByRole = Get-UsersByRole
    if (-not $UsersByRole.AdminUsers -and -not $UsersByRole.NonAdminUsers) {
        Write-Log -Message "No users found to process." -Level "WARNING"
        Write-Progress -Activity "Main Execution" -Status "No users found." -PercentComplete 100 -Completed
        return
    }
    Write-Log -Message "Users retrieved: Admin Users: $($UsersByRole.AdminUsers.Count), Non-Admin Users: $($UsersByRole.NonAdminUsers.Count), Admin Groups: $($UsersByRole.AdminGroups.Keys.Count)" -Level "INFO"
    Clear-Host

    # Step 7: Retrieve Password Details
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Retrieving password details..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Step 7: Retrieving password details..." -Level "INFO"
    try {
        $PasswordDetails = Get-PasswordDetailsForUsers -UsersByRole $UsersByRole
        if (-not $PasswordDetails -or $PasswordDetails.Count -eq 0) {
            Write-Log -Message "No password details retrieved." -Level "WARNING"
            Write-Progress -Activity "Main Execution" -Status "Password details missing." -PercentComplete 100 -Completed
            return
        }
        Write-Log -Message "Password details retrieved for $($PasswordDetails.Count) users." -Level "INFO"
    } catch {
        Write-Log -Message "Error retrieving password details: $($_.Exception.Message)" -Level "ERROR"
        Write-Progress -Activity "Main Execution" -Status "Password details retrieval failed." -PercentComplete 100 -Completed
        throw
    }
    Clear-Host

    # Step 8: Generate HTML Report
    $CurrentStep++
    Write-Progress -Activity "Main Execution" -Status "Generating HTML report..." -PercentComplete ([math]::Round(($CurrentStep / $TotalSteps) * 100))
    Write-Log -Message "Step 8: Generating HTML report..." -Level "INFO"
    $OutputFolderPath = "C:\RSanchezC169_Reports"
    $HtmlFileName = "Password_Report.html"
    try {
        Create-PasswordDetailsHtml -PasswordDetails $PasswordDetails `
                                   -AdminGroups $UsersByRole.AdminGroups `
                                   -AllUsers ($UsersByRole.AdminUsers + $UsersByRole.NonAdminUsers) `
                                   -FolderPath $OutputFolderPath `
                                   -FileName $HtmlFileName
        Write-Log -Message "HTML report created successfully at $OutputFolderPath\$HtmlFileName.html" -Level "INFO"
        Write-Host "HTML report generated successfully and opened in your browser!" -ForegroundColor Green
    } catch {
        Write-Log -Message "Error generating HTML report: $($_.Exception.Message)" -Level "ERROR"
        Write-Progress -Activity "Main Execution" -Status "HTML report generation failed." -PercentComplete 100 -Completed
        throw
    }
    Clear-Host

    Write-Log -Message "Main execution completed successfully." -Level "INFO"

} catch {
    Write-Progress -Activity "Main Execution" -Status "An error occurred." -PercentComplete 100 -Completed
    $ErrorMessage = $($_.Exception.Message)
    Write-Log -Message "An error occurred during main execution: $ErrorMessage" -Level "ERROR"
    throw
} finally {
    Write-Log -Message "Main execution ended." -Level "INFO"

    # Clean up resources
    Write-Progress -Activity "Main Execution" -Status "Cleaning up resources..." -PercentComplete 100 -Completed
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
