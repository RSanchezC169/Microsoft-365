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
Flipper Zero FirmWare mntm-001 https://momentum-fw.dev/update/
Description of Script: 
This script is to be used in tangent wtih my other script Create_Global_Admin.ps1. 
First once you have access to the tenant via a global admin account use Create_Global_Admin.ps1 to create your own global admin account. 
Once you have creaeted your own global admin account in the tenant use Disconnect-MgGraph to disconnect so you can then run this script as the new global admin account.
After this has been done use this script to remove every admin from the tenant except for your new global admin account so that you are the only admin of the tenant.
NOTE:
Make sure to replace -KeepAccountUPN "admin@yourdomain.com" with the new global admins smtp address.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function CleanUp-M365AdminAccounts {
    param (
        [string]$KeepAccountUPN
    )

    try {
        # Ensure the Microsoft Graph SDK is installed
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }

        # Connect to Microsoft Graph with the required permissions
        Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory Directory.ReadWrite.All User.Read.All"

        # Validate that the account to keep exists
        $keepAccount = Get-MgUser -Filter "userPrincipalName eq '$KeepAccountUPN'"
        if (-not $keepAccount) {
            Write-Warning "The specified account '$KeepAccountUPN' was not found in the tenant. Exiting script."
            return
        }

        # Retrieve the Global Administrator role
        $globalAdminRole = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" }
        if (-not $globalAdminRole) {
            throw "Global Administrator role not found. Ensure it exists in your tenant."
        }

        # Check if the specified account is a current Global Administrator
        $globalAdminMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id
        $isKeepAccountGlobalAdmin = $globalAdminMembers | Where-Object { $_.Id -eq $keepAccount.Id }

        if (-not $isKeepAccountGlobalAdmin) {
            Write-Warning "The specified account '$KeepAccountUPN' is not a Global Administrator. Exiting script."
            return
        }

        Write-Output "The account '$KeepAccountUPN' exists and is confirmed as a Global Administrator."

        # Retrieve all administrative roles
        Write-Output "Retrieving all administrative roles..."
        $adminRoles = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -like "*Administrator" }

        if (-not $adminRoles) {
            throw "No administrative roles found in the tenant."
        }

        # Loop through each admin role to retrieve assignments and clean up
        foreach ($role in $adminRoles) {
            Write-Output "Processing role: $($role.DisplayName)"

            # Get role assignments for this role
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "RoleDefinitionId eq '$($role.Id)'"

            if ($assignments -and $assignments.Count -gt 0) {
                foreach ($assignment in $assignments) {
                    # Fetch the member assigned to this role
                    $memberId = $assignment.PrincipalId
                    $memberDetails = Get-MgUser -UserId $memberId -ErrorAction SilentlyContinue

                    if ($memberDetails) {
                        # Skip the account that needs to be retained
                        if ($memberDetails.UserPrincipalName -eq $KeepAccountUPN) {
                            Write-Output "Skipping retention of account: $KeepAccountUPN in role: $($role.DisplayName)"
                            continue
                        }

                        # Remove the user from the role
                        try {
                            Write-Output "Removing account: $($memberDetails.UserPrincipalName) from role: $($role.DisplayName)"
                            Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $assignment.Id
                        } catch {
                            Write-Warning "Failed to remove account: $($memberDetails.UserPrincipalName) from role: $($role.DisplayName). Error: $_"
                        }
                    } else {
                        Write-Output "Skipping invalid or non-user principal with ID: $memberId"
                    }
                }
            } else {
                Write-Output "No users found in $($role.DisplayName) role."
            }
        }

        Write-Output "Admin account cleanup completed. Retained only: $KeepAccountUPN"

    } catch {
        Write-Error "An error occurred: $_"
    }
}

# Example usage of the function
CleanUp-M365AdminAccounts -KeepAccountUPN "admin@yourdomain.com"
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
