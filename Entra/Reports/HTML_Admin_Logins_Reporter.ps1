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
# Define the global log file path
$Global:LogFile = "C:\Rsanchezc169ScriptLogs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"
# Global logging function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter()]
        [string]$LogFile = $Global:LogFile
    )
$LogDirectory = Split-Path -Path $LogFile
    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
$Timestamp = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt")
    Add-Content -Path $LogFile -Value "$Timestamp : $Message"
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to retrieve administrative roles and their members
Function Get-AdminRolesAndMembers {
    Write-Log -Message "Retrieving all administrative roles..."
    try {
        # Step 1: Fetch all administrative roles
        $adminRoles = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -like "*Administrator" }
        if (-not $adminRoles -or $adminRoles.Count -eq 0) {
            throw "No administrative roles found in the tenant."
        }

        $adminUsers = @()
        $totalRoles = $adminRoles.Count
        $currentRole = 0

        if ($totalRoles -gt 0) {
            # Step 2: Process each role and retrieve its members
            foreach ($role in $adminRoles) {
                $currentRole++
                $percentComplete = [math]::Min((($currentRole / $totalRoles) * 100), 100) # Ensure the percentage doesn't exceed 100

                Write-Progress -Activity "Retrieving Admin Role Members" `
                               -Status "Processing Role $currentRole of $totalRoles" `
                               -PercentComplete $percentComplete
                Write-Log -Message "Processing role: $($role.DisplayName)"

                # Step 3: Fetch assignments for the role
                $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "RoleDefinitionId eq '$($role.Id)'"
                if ($assignments) {
                    foreach ($assignment in $assignments) {
                        $memberId = $assignment.PrincipalId
                        try {
                            # Fetch user details
                            $memberDetails = Get-MgUser -UserId $memberId -ErrorAction SilentlyContinue
                            if ($memberDetails) {
                                $adminUsers += $memberDetails.UserPrincipalName
                            } else {
                                Write-Log -Message "Principal with ID $memberId is not a user."
                            }
                        } catch {
                            Write-Log -Message "Error fetching details for PrincipalId $($memberId): $_"
                        }
                    }
                } else {
                    Write-Log -Message "No role assignments found for role: $($role.DisplayName)"
                }
            }
        } else {
            Write-Log -Message "No roles available for processing."
            Write-Host "No roles found in the tenant."
        }

        Write-Progress -Activity "Retrieving Admin Role Members" -Status "Completed" -Completed
        Write-Log -Message "Total unique admin users retrieved: $($adminUsers.Count)"
        return $adminUsers
    } catch {
        Write-Log -Message "Error retrieving admin roles: $_"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to fetch sign-in logs for all admin accounts
Function Get-AdminSignInLogs {
    Param(
        [array]$Admins
    )

    Write-Log -Message "Fetching sign-in logs for administrators..."
    $signInLogs = @()
    $totalAdmins = $Admins.Count
    $currentAdmin = 0

    if ($totalAdmins -gt 0) {
        foreach ($admin in $Admins) {
            $currentAdmin++
            $percentComplete = [math]::Min((($currentAdmin / $totalAdmins) * 100), 100) # Ensure percentage doesn't exceed 100

            Write-Progress -Activity "Fetching Sign-In Logs" `
                           -Status "Processing Admin $currentAdmin of $totalAdmins" `
                           -PercentComplete $percentComplete

            Write-Log -Message "Fetching sign-in logs for admin: $($admin)"
            try {
                $logs = Get-MgAuditLogSignIn -All | Where-Object {$_.UserPrincipalName -eq $admin}
                foreach ($log in $logs) {
                    $signInLogs += [PSCustomObject]@{
                        UserPrincipalName = $log.UserPrincipalName
                        IPAddress = $log.IpAddress
                        Location = if ($log.Location.City -and $log.Location.CountryOrRegion) {
                            $log.Location.City + ", " + $log.Location.CountryOrRegion
                        } else {
                            "Unknown Location"
                        }
                        SignInTime = $log.CreatedDateTime
                        Status = if ($log.Status.ErrorCode -eq 0) { "Success" } else { "Failure" }
                    }
                }
            } catch {
                Write-Log -Message "Failed to fetch sign-in logs for: $($admin). Error: $_"
            }
        }

        Write-Progress -Activity "Fetching Sign-In Logs" -Status "Completed" -Completed
        Write-Log -Message "Total sign-in logs collected: $($signInLogs.Count)"
    } else {
        Write-Log -Message "No administrators to process."
        Write-Host "No admin users found."
    }

    return $signInLogs
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to generate an HTML report with a dropdown for users
Function GenerateAdminSignInHtml {
    Param(
        [array]$SignInLogs,
        [array]$Admins,
        [string]$OutputHtmlPath = "C:\Temp\AdminSignInLogs.html"
    )

    Write-Log -Message "Generating HTML content for the report..."
    try {
        $htmlContent = @"
<html>
<head>
    <title>Admin Sign-In Logs</title>
    <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; }
        .unusual { background-color: #ffcccc; font-weight: bold; } /* Highlight unusual logins */
        h1 { text-align: center; }
    </style>
    <script>
        // JavaScript to filter logs based on dropdown selection
        function filterLogs() {
            var selectedUser = document.getElementById("userDropdown").value;
            var rows = document.getElementsByClassName("logRow");
            for (var i = 0; i < rows.length; i++) {
                rows[i].style.display = (selectedUser === "All" || rows[i].dataset.user === selectedUser) ? "" : "none";
            }
        }
    </script>
</head>
<body>
    <h1>Admin Sign-In Logs</h1>
    <label for="userDropdown">Select User:</label>
    <select id="userDropdown" onchange="filterLogs()">
        <option value="All">All Users</option>
"@

        # Add options to the dropdown for each admin
        foreach ($admin in $Admins) {
            $htmlContent += "<option value='$admin'>$admin</option>"
        }

        # Start the table for logs
        $htmlContent += "</select><br><br><table><tr><th>Sign-In Time</th><th>User</th><th>IP Address</th><th>Location</th><th>Status</th></tr>"

        $totalLogs = $SignInLogs.Count
        $currentLog = 0

        if ($totalLogs -gt 0) {
            foreach ($log in $SignInLogs) {
                $currentLog++
                $percentComplete = [math]::Min((($currentLog / $totalLogs) * 100), 100) # Cap percentage at 100

                Write-Progress -Activity "Generating HTML Report" `
                               -Status "Processing Log $currentLog of $totalLogs" `
                               -PercentComplete $percentComplete

                $htmlContent += "<tr class='logRow' data-user='$($log.UserPrincipalName)'>" +
                                "<td>$($log.SignInTime)</td>" +
                                "<td>$($log.UserPrincipalName)</td>" +
                                "<td>$($log.IPAddress)</td>" +
                                "<td>$($log.Location)</td>" +
                                "<td>$($log.Status)</td></tr>"
            }

            Write-Progress -Activity "Generating HTML Report" -Status "Completed" -Completed
        } else {
            Write-Log -Message "No sign-in logs available to include in the report."
            $htmlContent += "<tr><td colspan='5'>No sign-in logs available.</td></tr>"
        }

        # Close table and HTML tags
        $htmlContent += "</table></body></html>"

        # Write the HTML content to the output file
        $htmlContent | Set-Content -Path $OutputHtmlPath -Force
        Write-Log -Message "HTML report saved to: $OutputHtmlPath"
    } catch {
        Write-Log -Message "Error generating HTML report: $_"
        throw
    }
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
# Main Execution Logic
try {
    Write-Log -Message "Script execution started."
    Write-Host "Retrieving tenant administrators and their roles..."

    # Step 1: Retrieve all admin users
    $admins = Get-AdminRolesAndMembers
    if (-not $admins -or $admins.Count -eq 0) {
        Write-Log -Message "No administrators found in the tenant."
        Write-Host "No administrators were found. Exiting script."
        return
    }
    Write-Log -Message "Admins retrieved: $($admins -join ', ')"

    # Step 2: Fetch sign-in logs for the retrieved admin users
    Write-Host "Fetching sign-in logs for administrators..."
    $signInLogs = Get-AdminSignInLogs -Admins $admins
    if (-not $signInLogs -or $signInLogs.Count -eq 0) {
        Write-Log -Message "No sign-in logs found for the administrators."
        Write-Host "No sign-in logs were found. Exiting script."
        return
    }
    Write-Log -Message "Total sign-in logs retrieved: $($signInLogs.Count)"

    # Step 3: Generate the HTML report with dropdown filtering
    Write-Host "Generating the HTML report..."
    $outputPath = "C:\Temp\AdminSignInLogs.html"
    GenerateAdminSignInHtml -SignInLogs $signInLogs -Admins $admins -OutputHtmlPath $outputPath
    Write-Log -Message "The admin sign-in logs report has been successfully generated and saved to $outputPath."

    # Final Success Message
    Write-Host "Script execution completed successfully. HTML report is available at: $outputPath"
} catch {
    Write-Log -Message "An error occurred during script execution: $_"
    Write-Error "An error occurred: $_"
}
##################################################################################################################################################################
#==============================End of Main========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script======================================================================================================================
##################################################################################################################################################################
