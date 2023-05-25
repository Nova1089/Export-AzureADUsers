<#
This script exports the following info for Azure AD users:
- User Principal Name (UPN)
- Display name
- User type
- Account enabled
- Title
- Department
- Manager
- Licenses (Use this article to translate the SKU IDs: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference)
- Created date & time
- Last sign in date & time

You are prompted to filter users with the following methods:
[1] Get all users in Azure AD.
[2] Enter list of users to export by txt file.
    (Users listed by full name or email, separated by new line.)
[3] Enter list of users one by one.
[4] Get all users with missing title or department.
[5] Get all users whose UPN does not follow the standard convention.
    (xxx.xxx@xxx.com - where x could be any number of characters)
[6] Get all users that have not signed in for a specified amount of days.
#>

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:failColor = "Red"
    # warning color is yellow, but that is built into Write-Warning
}

function Show-Introduction
{
    Write-Host "This script exports Azure AD user info." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}


function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-MgGraph
{
    $connected = Test-ConnectedToMgGraph

    while(-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor
        Connect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        $connected = Test-ConnectedToMgGraph

        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
        else
        {
            Write-Host "Successfully connected!" -ForegroundColor $successColor
        }
    }    
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function Prompt-UserFilterMethod
{
    Write-Host "How would you like to filter users to export?`n"
    Write-Host "[1] Get all users in Azure AD."
    Write-Host ("[2] Enter list of users to export by txt file.`n" +
                    "`t(Users listed by full name or email, separated by new line.)")
    Write-Host "[3] Enter list of users one by one."
    Write-Host "[4] Get all users with missing title or department."
    Write-Host ("[5] Get all users whose UPN does not follow the standard convention.`n" +
                    "`t(xxx.xxx@xxx.com - where x could be any number of characters)")
    Write-Host "[6] Get all users that have not signed in for a specified amount of days."

    while ($true)
    {
        $response = Read-Host

        if ($response -imatch "^\s*[1-6]\s*$") # regex matches a 1 through 6 but allows spaces
        {
            break
        }
        Write-Warning "Please enter a number 1-6."
    }
    
    return [int]($response.Trim())
}

function New-DesktopPath($fileName, $fileExt)
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $timeStamp = (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
    return "$desktopPath\$fileName $timeStamp.$fileExt"
}

function Export-UsersFromFilter($whereFilter, $path)
{
    Get-MGUser -All -Property "UserPrincipalName, DisplayName, JobTitle, UserType, AccountEnabled, Department, CreatedDateTime, SignInActivity" | 
        Write-ProgressInPipeline -activity "Exporting users..." -status "users processed" | 
        Where-Object $whereFilter |
        Get-UserData | 
        Export-CSV -Path $path -Append -NoTypeInformation

    Write-Host "Finished exporting to $path" -ForegroundColor $successColor
}

function Write-ProgressInPipeline
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory, ValueFromPipeline)]
        [object[]] $inputObjects,
        [string] $activity = "Processing items...",
        [string] $status = "items processed"
    )

    Begin 
    { 
        $itemsProcessed = 0 
    }

    Process
    {
        Write-Progress -Activity $activity -Status "$itemsProcessed $status"
        $itemsProcessed++
        return $_
    }
}

function Get-UserData
{
    [CmdletBinding()]
    Param 
    (
        [Parameter(Position=0, Mandatory, ValueFromPipeline)]
        $mgUser
    )

    Process
    {
        [PSCustomObject]@{
            UPN = $mgUser.UserPrincipalName
            Name = $mgUser.DisplayName            
            UserType = $mgUser.UserType
            AccountEnabled = $mgUser.AccountEnabled
            Title = $mgUser.JobTitle
            Department = $mgUser.Department
            Manager = (Get-MgUserManager -UserId $mgUser.UserPrincipalName -ErrorAction SilentlyContinue).AdditionalProperties.mail
            Licenses = Get-Licenses $mgUser
            'CreatedDateTime (UTC)' = $mgUser.CreatedDateTime
            'LastSignIn (UTC)'= $mgUser.SignInActivity.LastSignInDateTime
        }
    }
}

function Get-Licenses($user)
{
    $licenses = (Get-MgUserLicenseDetail -UserID $user.UserPrincipalName).SkuPartNumber
    $combinedLicenses = ""

    foreach ($license in $licenses)
    {
        $combinedLicenses += $license + ", "
    }

    return $combinedLicenses
}

function Get-UsersFromTXT
{
    do
    {
        $path = Read-Host "Enter path to txt file. (i.e. C:\UserList.txt)"
        $path = $path.Trim('"') # trims quotes if they were entered on path
        $userList = Get-Content -Path $path -ErrorAction SilentlyContinue

        if ($null -eq $userList)
        {
            Write-Warning "File not found or contents are empty."
            $keepGoing = $true
            continue
        }
        else
        {
            $keepGoing = $false
        }        
    }
    while ($keepGoing)

    return $userList
}

function Export-UsersFromList($userList)
{
    Write-Host "Exporting users..." -ForegroundColor $infoColor

    $csvPath = New-DesktopPath -fileName "AzureAD User Export" -fileExt "csv"

    $totalUsersSearched = 0 
    foreach ($user in $userList)
    {
        Write-Progress -Activity "Exporting users..." -Status "$totalUsersSearched users processed"
        $totalUsersSearched++

        $mgUser = TryGet-User $user
        if ($null -eq $mgUser) { continue }

        Get-UserData $mgUser | Export-CSV -Path $csvPath -Append -NoTypeInformation
    }

    Write-Host "Finished exporting to $csvPath" -ForegroundColor $successColor
}

function Get-UsersManually
{
    $userList = New-Object -TypeName System.Collections.Generic.List[string]

    while ($true)
    {
        $response = Read-Host "Enter a user (full name or email) or type `"done`""
        $response = $response.Trim()
        if ($response -ieq "done") { break }
        if ($null -eq (TryGet-User $response -tellWhenFound)) { continue }
        $userList.Add($response)
    }

    return $userList
}

function TryGet-User($userId, [switch]$tellWhenFound)
{
    $userId = $userId.Trim()
    $userProperties = "UserPrincipalName, DisplayName, JobTitle, UserType, AccountEnabled, Department, CreatedDateTime, SignInActivity"
    $mgUser = Get-MGUser -Filter "(UserPrincipalName eq '$userId') or (DisplayName eq '$userId')" -Property $userProperties -ErrorAction SilentlyContinue

    if ($null -eq $mgUser)
    {
        Write-Warning "User not found: $user"
        return $null
    }
    elseif($mgUser.Count -gt 1)
    {
        Write-Warning "Multiple users found with identifier: $userId. Please use their UPN instead."
        $mgUser | Format-Table -Property DisplayName, UserPrincipalName, JobTitle, Department | Out-Host
        return $null
    }

    if ($tellWhenFound)
    {
        Write-Host "Found user:" -ForegroundColor $successColor
        $mgUser | Format-Table -Property DisplayName, UserPrincipalName, JobTitle, Department | Out-Host
    }

    return $mgUser
}

function Prompt-NumDays
{
    while ($true)
    {
        $response = Read-Host "How many days?"

        if ($response -notmatch '^\s*\d+\s*$')
        {
            Write-Warning "Please enter an integer."
            continue
        }
        return [int]$response
    }    
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Users"
TryConnect-MgGraph
$userFilterMethod = Prompt-UserFilterMethod
switch ($userFilterMethod)
{
    1 # export all users in AzureAD
    {
        $whereFilter = { $true }
        $path = New-DesktopPath -fileName "All AzureAD users" -fileExt "csv"
        Export-UsersFromFilter -whereFilter $whereFilter -path $path
    }
    2 # export users provided in text file
    { 
        $userList = Get-UsersFromTXT 
        Export-UsersFromList $userList
    }
    3 # export users entered one by one
    { 
        $userList = Get-UsersManually 
        Export-UsersFromList $userList
    }
    4 # export users with missing title or department
    {
        $whereFilter = { ($null -eq $_.JobTitle) -or ($null -eq $_.Department) }
        $path = New-DesktopPath -fileName "AzureAD Users Missing Org Info" -fileExt "csv"
        Export-UsersFromFilter -whereFilter $whereFilter -path $path
    }
    5 # export users that have an irregular UPN
    {
        $regex = '.+\..+@.+\.com' # regex matches an email like x.x@x.com where x can be 1 or more chars
        $whereFilter = { $_.UserPrincipalName -inotmatch $regex }
        $path = New-DesktopPath -fileName "AzureAD Users Irregular UPN" -fileExt "csv"
        Export-UsersFromFilter -whereFilter $whereFilter -path $path
    }
    6 # export users that haven't signed in for x days
    {
        $numDays = Prompt-NumDays
        $calculatedDate = (Get-Date).AddDays(-$numDays)
        $whereFilter = { $_.SignInActivity.LastSignInDateTime -le $calculatedDate } 
        $path = New-DesktopPath -fileName "AzureAD Users No Sign In $numDays Days" -fileExt "csv"
        Export-UsersFromFilter -whereFilter $whereFilter -path $path
    }
}
Read-Host "Press Enter to exit"