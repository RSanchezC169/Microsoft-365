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
This PowerShell script, titled Get-M365AdminAccounts, is designed to retrieve and display all administrative accounts and their assigned roles within a Microsoft 365 tenant. It accomplishes this by connecting to Microsoft Graph, fetching all administrative roles, and iterating through those roles to identify and list their members.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function Get-M365AdminAccounts {
   #Set execution Policy
   TRY{
	Set-ExecutionPolicy -Scope "Process" -ExecutionPolicy "Unrestricted" -Force  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
     }CATCH{
            Write-Warning "Could not set execution policy"
    }

    try {
        # Ensure the Microsoft Graph SDK is installed
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Write-Output "Microsoft Graph SDK not found. Installing it now..."
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }

        # Connect to Microsoft Graph
        Write-Output "Connecting to Microsoft Graph..."
        TRY {
             Connect-MgGraph -Scopes "RoleManagement.Read.Directory User.Read.All"
        }CATCH{
              Write-Warning "Could not connect to Microsoft Graph"
       }
        # Get all administrative roles
        Write-Output "Retrieving all administrative roles..."
        $adminRoles = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -like "*Administrator" }

        if (-not $adminRoles) {
            throw "No administrative roles found in the tenant."
        }

        # Initialize a master array to store members for all roles
        $allRolesMembers = @()

        # Total count of roles to process
        $totalRoles = $adminRoles.Count
        $currentRoleIndex = 0

        # Loop through each admin role to retrieve assignments
        foreach ($role in $adminRoles) {
            # Increment the current role index for progress
            $currentRoleIndex++

            # Update progress bar
            Write-Progress -Activity "Processing Admin Roles" `
                           -Status "Processing role: $($role.DisplayName)" `
                           -PercentComplete (($currentRoleIndex / $totalRoles) * 100)

            # Initialize an array to hold members of the current role
            $roleMembers = @()

            # Get role assignments for this role
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "RoleDefinitionId eq '$($role.Id)'"

            if ($assignments -and $assignments.Count -gt 0) {
                foreach ($assignment in $assignments) {
                    # Fetch the member assigned to this role
                    $memberId = $assignment.PrincipalId
                    $memberDetails = Get-MgUser -UserId $memberId -ErrorAction SilentlyContinue

                    if ($memberDetails) {
                        # Add valid member details to the roleMembers array
                        $roleMembers += [PSCustomObject]@{
                            Name         = $memberDetails.DisplayName
                            EmailAddress = $memberDetails.UserPrincipalName
                        }
                        Write-Output "Found member: $($memberDetails.DisplayName) for Role: $($role.DisplayName)"
                    } else {
                        Write-Output "Skipping invalid or non-user principal with ID: $memberId"
                    }
                }

                # If the roleMembers array is not empty, output the results for this role
                if ($roleMembers.Count -gt 0) {
                    $allRolesMembers += [PSCustomObject]@{
                        RoleName = $role.DisplayName
                        Members  = $roleMembers
                    }
                    Write-Output "Role: $($role.DisplayName) has members:"
                    $roleMembers | ForEach-Object {
                        Write-Output "Name: $($_.Name), Email: $($_.EmailAddress)"
                    }
                }
            } else {
                Write-Output "No users found in $($role.DisplayName) role."
            }
        }

        # Complete the progress bar
        Write-Progress -Activity "Processing Admin Roles" -Status "Completed" -PercentComplete 100

        # Return all roles with their members if any exist
        if ($allRolesMembers.Count -gt 0) {
            Write-Output "Summary of all roles and members:"
            return $allRolesMembers
        } else {
            Write-Output "No administrative accounts with permissions were found in the tenant."
        }
    } catch {
        Write-Error "An error occurred: $_"
    }
}
#Clear PowerShell Host
Clear-Host

Get-M365AdminAccounts
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
