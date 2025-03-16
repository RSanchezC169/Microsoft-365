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
# Function to ensure Microsoft Graph SDK
function Ensure-MicrosoftGraphSDK {
    try {
        Write-Progress -Activity "Setup" -Status "Checking Microsoft Graph SDK" -PercentComplete 10
        Write-Log "Checking Microsoft Graph SDK..."
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
            Write-Log "Microsoft Graph SDK installed successfully."
        } else {
            Write-Log "Microsoft Graph SDK is already installed."
        }
    } catch {
        Write-Log "Failed to ensure Microsoft Graph SDK is installed. Error: $_"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    try {
        Write-Progress -Activity "Setup" -Status "Connecting to Microsoft Graph" -PercentComplete 20
        Write-Log "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "RoleManagement.Read.Directory User.Read.All"
        Write-Log "Connected to Microsoft Graph successfully."
    } catch {
        Write-Log "Failed to connect to Microsoft Graph. Error: $_"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to retrieve admin roles
function Get-AdministrativeRoles {
    try {
        Write-Progress -Activity "Retrieving Roles" -Status "Retrieving all administrative roles" -PercentComplete 30
        Write-Log "Retrieving administrative roles..."
        $adminRoles = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -like "*Administrator" }
        if (-not $adminRoles) {
            Write-Log "No administrative roles found in the tenant."
            throw "No administrative roles found."
        }
        Write-Log "Retrieved $($adminRoles.Count) administrative roles."
        return $adminRoles
    } catch {
        Write-Log "Failed to retrieve administrative roles. Error: $_"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
# Function to retrieve role members
function Get-RoleMembers {
    param (
        [array]$AdminRoles
    )

    $allRolesMembers = @()
    $totalRoles = $AdminRoles.Count
    $currentRole = 0

    foreach ($role in $AdminRoles) {
        $currentRole++
        Write-Progress -Activity "Retrieving Role Members" -Status "Processing Role $currentRole of $totalRoles" -PercentComplete (($currentRole / $totalRoles) * 100)
        Write-Log "Processing role: $($role.DisplayName)"
        try {
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "RoleDefinitionId eq '$($role.Id)'"
            if ($assignments) {
                foreach ($assignment in $assignments) {
                    $memberId = $assignment.PrincipalId
                    $memberDetails = Get-MgUser -UserId $memberId -ErrorAction SilentlyContinue
                    if ($memberDetails) {
                        $allRolesMembers += [PSCustomObject]@{
                            Name         = $memberDetails.DisplayName
                            EmailAddress = $memberDetails.UserPrincipalName
                            Role         = $role.DisplayName
                        }
                    }
                }
                Write-Log "Retrieved $($assignments.Count) members for role: $($role.DisplayName)"
            } else {
                Write-Log "No members found for role: $($role.DisplayName)"
            }
        } catch {
            Write-Log "Failed to retrieve members for role: $($role.DisplayName). Error: $_"
        }
    }

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
# Function to validate and open the file
function Validate-AndOpenFile {
    param (
        [string]$FilePath
    )

    if (-not (Test-Path -Path $FilePath)) {
        Write-Log "Error: File does not exist at the specified path: $FilePath"
        throw "File was not created."
    }

    Write-Log "File exists and is ready to open: $FilePath"

    try {
        Write-Log "Opening the file: $FilePath"
        Start-Process -FilePath $FilePath -Verb Open
    } catch {
        Write-Log "Failed to open the file: $FilePath. Error: $_"
        throw
    }
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
try {
    Write-Progress -Activity "Execution" -Status "Starting script" -PercentComplete 0
    Write-Log "Script execution started."

    Ensure-MicrosoftGraphSDK
    Connect-ToMicrosoftGraph
    $adminRoles = Get-AdministrativeRoles
    $allRolesMembers = Get-RoleMembers -AdminRoles $adminRoles
    $outputPath = "C:\Temp\AdminAccounts.html"
    Generate-HTMLReport -AllRolesMembers $allRolesMembers -OutputHtmlPath $outputPath

    Write-Progress -Activity "Execution" -Status "Script completed successfully" -PercentComplete 100
    Write-Log "Script execution completed successfully."
} catch {
    Write-Log "An error occurred during execution: $_"
    Write-Error "An error occurred: $_"
}
##################################################################################################################################################################
#==============================End of Main========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script======================================================================================================================
##################################################################################################################################################################
