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
$Global:LogFile = "C:\Rsanchezc169ScriptLogs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"

Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,                  # The message to log
        [Parameter()]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",           # Log level: INFO, WARNING, DEBUG, or ERROR
        [Parameter()]
        [string]$LogFile = $Global:LogFile # Optional: Specify a log file, fallback to global log file
    )

    try {
        # Ensure the log directory exists
        $LogDirectory = Split-Path -Path $LogFile -ErrorAction SilentlyContinue
        if (-not (Test-Path -Path $LogDirectory)) {
            #Write-Host "Creating log directory: $LogDirectory" -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
        }
    } catch {
        #Write-Error "Failed to create log directory: $($_.Exception.Message)"
        throw
    }

    try {
        # Append the message to the log file with timestamp and level
        $Timestamp = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt")
        Add-Content -Path $LogFile -Value "$Timestamp [$Level] : $Message"

        # Optionally display log messages to the console
        if ($Level -eq "ERROR") {
            #Write-Host "[ERROR] $Message" -ForegroundColor Red
        } elseif ($Level -eq "WARNING") {
            #Write-Host "[WARNING] $Message" -ForegroundColor Yellow
        } else {
            #Write-Host "[INFO] $Message" -ForegroundColor Green
        }
    } catch {
        #Write-Warning "Failed to write log message: $($_.Exception.Message)"
    }
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Connect-MgGraphEnhanced {
    [CmdletBinding()]
    Param()

    Write-Log -Message "Started Function Connect-MgGraphEnhanced" -Level "INFO"

    try {
        # Check if already connected
        Write-Log -Message "Checking if already connected to Microsoft Graph..." -Level "INFO"
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

        if ($GraphContext -and $GraphContext.Account -and $GraphContext.TenantId) {
            # Already connected
            Write-Host "You are already connected to Microsoft Graph!" -ForegroundColor Green
            Write-Host "Connected as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -ForegroundColor Green
            Write-Log -Message "Already connected to Microsoft Graph as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -Level "INFO"
        } else {
            # Define the scopes for connection
            $Scopes = @(
                "AuditLog.Read.All",
                "Directory.Read.All"
            )
            Write-Log -Message "Defining connection scopes: $($Scopes -join ', ')" -Level "INFO"

            # Attempt to connect to Microsoft Graph
            Write-Host "Connecting to Microsoft Graph with specified scopes..." -ForegroundColor Cyan
            Write-Log -Message "Attempting connection to Microsoft Graph with scopes: $($Scopes -join ', ')" -Level "INFO"

            # Retry connection logic
            $MaxAttempts = 3
            $Attempt = 0
            $Connected = $false

            while (-not $Connected -and $Attempt -lt $MaxAttempts) {
                try {
                    $Attempt++
                    Write-Log -Message "Connection attempt $Attempt of $MaxAttempts..." -Level "INFO"
                    Connect-MgGraph -Scope $Scopes -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                    $Connected = $true
                    Write-Log -Message "Successfully connected to Microsoft Graph!" -Level "INFO"
                } catch {
                    Write-Warning "Connection attempt $Attempt failed. Error: $($_.Exception.Message)"
                    Write-Log -Message "Connection attempt $Attempt failed. Error: $($_.Exception.Message)" -Level "WARNING"
                    Start-Sleep -Seconds 5  # Wait before retrying
                }
            }

            if (-not $Connected) {
                throw "All connection attempts to Microsoft Graph failed."
            }

            # Retrieve and log connection context
            $GraphContext = Get-MgContext -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
            Write-Host "Successfully connected to Microsoft Graph as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -ForegroundColor Green
            Write-Log -Message "Successfully connected to Microsoft Graph as $($GraphContext.Account) (Tenant: $($GraphContext.TenantId))" -Level "INFO"
        }
    } catch {
        # Handle connection errors
        $ErrorMessage = $($_.Exception.Message)
        Write-Warning "Could not connect to Microsoft Graph. Error: $ErrorMessage"
        Write-Log -Message "Error connecting to Microsoft Graph: $ErrorMessage" -Level "ERROR"
    } finally {
        # Finalize connection attempt
        Write-Log -Message "Ended Function Connect-MgGraphEnhanced" -Level "INFO"
    }
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Get-UsersByRole {
    [CmdletBinding()]
    Param()

    Write-Log -Message "Started Function: Get-UsersByRole" -Level "INFO"

    try {
        # Step 1: Retrieve all user accounts in the tenant
        Write-Log -Message "Fetching all user accounts from the tenant..." -Level "INFO"
        $AllUsers = Invoke-WithRetry -ScriptBlock {
            Get-MgUser -All -Select "Id,DisplayName,UserPrincipalName" -ErrorAction Stop
        }

        if (-not $AllUsers -or $AllUsers.Count -eq 0) {
            Write-Log -Message "No users found in the tenant." -Level "WARNING"
            Write-Host "No users found in the tenant." -ForegroundColor Yellow
            return @{}
        }

        Write-Log -Message "Total users retrieved: $($AllUsers.Count)" -Level "INFO"

        # Step 2: Retrieve all administrative roles and their assignments
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

        return @{
            AdminUsers    = $AdminUsers
            NonAdminUsers = $NonAdminUsers
            AdminGroups   = $AdminGroups
        }
    } catch {
        # Handle errors gracefully
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
        Write-Host "Total users to process: $($AllUsers.Count)" -ForegroundColor Cyan

        if (-not $AllUsers -or $AllUsers.Count -eq 0) {
            Write-Log -Message "No users provided to process." -Level "WARNING"
            Write-Host "No users provided to process." -ForegroundColor Yellow
            return
        }

        # Initialize an array to store password details
        $UserPasswordDetails = @()
        $TotalUsers = $AllUsers.Count
        $CurrentUser = 0

        # Loop through each user and retrieve password-related details
        foreach ($User in $AllUsers) {
            $CurrentUser++
            $PercentComplete = [math]::Min((($CurrentUser / $TotalUsers) * 100), 100)

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
                    $TimeSinceLastPasswordChange = "N/A"  # Default value if no password set date exists
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
                        TimeSinceLastPasswordChange = $TimeSinceLastPasswordChange # New field
                        AccountEnabled              = $PasswordDetails.AccountEnabled
                        PasswordPolicies            = $PasswordDetails.PasswordPolicies
                        AdminGroups                 = ($GroupMemberships -join ", ") # List admin groups the user belongs to
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
        Write-Progress -Activity "Retrieving Password Details" -Status "Completed" -Completed

        # Log final results
        Write-Log -Message "Completed processing all users. Total users processed: $($UserPasswordDetails.Count)" -Level "INFO"
        Write-Host "Password details retrieved for $($UserPasswordDetails.Count) users." -ForegroundColor Green

        return $UserPasswordDetails
    } catch {
        # Handle unexpected errors
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

    $attempts = 0

    while ($attempts -lt $MaxAttempts) {
        try {
            $attempts++
            Write-Log -Message "Attempt $attempts of $MaxAttempts. Executing script block..." -Level "INFO"
            $result = & $ScriptBlock
            Write-Log -Message "Operation succeeded on attempt $attempts." -Level "INFO"
            return $result
        } catch {
            Write-Log -Message "Attempt $attempts failed. Error: $($_.Exception.Message)" -Level "WARNING"

            if ($attempts -ge $MaxAttempts) {
                Write-Log -Message "Max attempts reached. Throwing error..." -Level "ERROR"
                throw $_
            }

            Write-Log -Message "Retrying operation in $RetryDelay seconds..." -Level "INFO"
            Start-Sleep -Seconds $RetryDelay
        }
    }
}

##################################################################################################################################################################
##################################################################################################################################################################
Function Main-Execution {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$OutputFolderPath,  # Path where the HTML file will be saved
        [Parameter(Mandatory)]
        [string]$HtmlFileName       # Name of the output HTML file
    )

    Write-Log -Message "Started Main Execution Logic" -Level "INFO"

    try {
        # Step 1: Connect to Microsoft Graph
        Write-Log -Message "Step 1: Connecting to Microsoft Graph..." -Level "INFO"
        Connect-MgGraphEnhanced
        Write-Log -Message "Successfully connected to Microsoft Graph." -Level "INFO"

        # Step 2: Retrieve Users and Admin Groups
        Write-Log -Message "Step 2: Retrieving users and categorizing them by role (Admin/Non-Admin) and fetching admin groups..." -Level "INFO"
        $UsersByRole = Get-UsersByRole
        if (-not $UsersByRole.AdminUsers -and -not $UsersByRole.NonAdminUsers) {
            Write-Log -Message "No users found to process." -Level "WARNING"
            Write-Host "No users found in the tenant. Please ensure your Azure AD has users to process." -ForegroundColor Yellow
            return
        }
        Write-Log -Message "Successfully retrieved users. Admin Users: $($UsersByRole.AdminUsers.Count), Non-Admin Users: $($UsersByRole.NonAdminUsers.Count), Admin Groups: $($UsersByRole.AdminGroups.Keys.Count)" -Level "INFO"

        # Step 3: Retrieve Password Details
        Write-Log -Message "Step 3: Retrieving password details for all users..." -Level "INFO"
        $PasswordDetails = $null
        try {
            $PasswordDetails = Get-PasswordDetailsForUsers -UsersByRole $UsersByRole
            if (-not $PasswordDetails -or $PasswordDetails.Count -eq 0) {
                Write-Log -Message "No password details could be retrieved." -Level "WARNING"
                Write-Host "Failed to retrieve password details. Please ensure users have valid data." -ForegroundColor Yellow
                return
            } else {
                Write-Log -Message "Password details retrieved for $($PasswordDetails.Count) users." -Level "INFO"
                Write-Host "Password details retrieved for $($PasswordDetails.Count) users." -ForegroundColor Green
            }
        } catch {
            Write-Log -Message "Error retrieving password details: $($_.Exception.Message)" -Level "ERROR"
            Write-Warning "Failed to retrieve password details: $($_.Exception.Message)"
            return
        }

        # Step 4: Generate HTML Report
        Write-Log -Message "Step 4: Generating HTML report..." -Level "INFO"
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
            Write-Warning "Failed to generate HTML report: $($_.Exception.Message)"
        }
    } catch {
        # Handle unexpected errors in execution
        Write-Log -Message "Unexpected error in Main Execution Logic: $($_.Exception.Message)" -Level "ERROR"
        Write-Warning "An unexpected error occurred: $($_.Exception.Message)"
    } finally {
        # Final cleanup
        Write-Log -Message "Completed Main Execution Logic" -Level "INFO"
    }
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
# Example Usage:
Main-Execution -OutputFolderPath "C:\Reports" -HtmlFileName "PasswordDetailsReport"
##################################################################################################################################################################
#==============================End of Main========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script======================================================================================================================
##################################################################################################################################################################
