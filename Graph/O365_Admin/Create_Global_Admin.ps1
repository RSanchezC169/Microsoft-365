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
The Create-M365GlobalAdmin function is designed to automate the creation of a new Microsoft 365 Global Administrator user account using the Microsoft Graph API. 
It handles critical steps such as user creation, default domain retrieval, and role assignment, making the process efficient and repeatable.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function Create-M365GlobalAdmin {
    param (
        [string]$DisplayName,
        [string]$MailNickname,
        [string]$Password
    )

    # Set execution policy
    try {
        Set-ExecutionPolicy -Scope "Process" -ExecutionPolicy "Unrestricted" -Force `
            -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
    } catch {
        Write-Warning "Could not set execution policy: $_"
    }

    try {
        # Ensure the Microsoft Graph SDK is installed
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }

        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes "User.ReadWrite.All RoleManagement.ReadWrite.Directory"

        # Retrieve the default .onmicrosoft.com domain
        $defaultDomain = (Get-MgDomain | Where-Object {$_.Id -LIKE "*.onmicrosoft.com"}).Id
        if (-not $defaultDomain -or $defaultDomain -notlike "*.onmicrosoft.com") {
            throw "Default .onmicrosoft.com domain not found in the tenant. Ensure it exists."
        }

        # Construct the UserPrincipalName (SMTP address) dynamically
        $UserPrincipalName = "$MailNickname@$defaultDomain"

        # Debugging output for parameter values
        Write-Output "Creating user with the following details:"
        Write-Output "DisplayName: $DisplayName"
        Write-Output "UserPrincipalName: $UserPrincipalName"
        Write-Output "MailNickname: $MailNickname"
        Write-Output "Password: $Password"

        # Create the user
        $user = New-MgUser -DisplayName $DisplayName `
                            -UserPrincipalName $UserPrincipalName `
                            -MailNickname $MailNickname `
                            -PasswordProfile @{Password = $Password} `
                            -AccountEnabled

        # Get the Global Administrator role
        $globalAdminRole = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" }

        if (-not $globalAdminRole) {
            throw "Global Administrator role not found. Ensure it exists in your tenant."
        }

        # Assign the role to the new user
        $DirObject = @{
           "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"
        }

        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $globalAdminRole.Id `
                                       -BodyParameter $DirObject

        Write-Output "User '$DisplayName' with UPN '$UserPrincipalName' created and assigned Global Administrator role successfully."

    } catch {
        Write-Error "An error occurred: $_"
    }
}

# Example usage of the function
Create-M365GlobalAdmin -DisplayName "$(Read-Host "Enter in new admins display name -> ")" -MailNickname "$(Read-Host "Enter in new admins mail nickname no spaces -> ")" -Password "$(Read-Host "Enter in new admins password -> ")"
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
