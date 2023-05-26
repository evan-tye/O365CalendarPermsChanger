# Evan Tye - May 2023

Clear-Host
Write-Host " ==================================================="
Write-Host "  O365 Calendar Permissions Changer - Evan Tye 2023 "
Write-Host " ==================================================="
Write-Host ""

# Check EXO module is installed, if not, exit

Write-Host ""
Set-ExecutionPolicy RemoteSigned
Write-Host " >> Checking if Exchange Online Management module is installed" -NoNewline
Start-Sleep -Seconds 0.5
Write-Host "." -NoNewline
Start-Sleep -Seconds 0.5
Write-Host "." -NoNewline
Start-Sleep -Seconds 0.5
Write-Host "." -NoNewline
Write-Host ""
Start-Sleep -Seconds 1.5

if(-not(Get-Module -ListAvailable -Name ExchangeOnlineManagement))
{
    Write-Host " >> Exchange Online Management module is not installed! Run: Install-Module -Name ExchangeOnlineManagement and then rerun the script!"
    Exit
}
else
{
    Write-Host " >> EXO module check passed! "
}

# Take input for mailbox calendar and user mailbox

Write-Host ""
[string]$identity = Read-Host " >> Calendar Owner Email Address"
Write-Host ""
[string]$user = Read-Host " >> Delegate Email Address"
Write-Host ""
Write-Host " >> Enter a numeric value for the permission level you wish to grant:"
Write-Host ""
Write-Host "1: Owner"
Write-Host "2: Publishing Editor"
Write-Host "3: Editor"
Write-Host "4: Publishing Author"
Write-Host "5: Author"
Write-Host "6: Non-Editing Author"
Write-Host "7: Reviewer"
Write-Host "8: Contributor"
Write-Host "9: Limited Details"
Write-Host "10: Availability"
Write-Host "11: Remove Permissions (None)"
Write-Host ""
[int]$permLevel = Read-Host "Value"
Write-Host ""
Start-Sleep -Seconds 1.0
Write-Host ""
Write-Host " >> Enter Global Administrator credentials in the browser window..."

# Connect to EXO

Connect-ExchangeOnline
Start-Sleep -Seconds 1.5

# Process user input and Add permissions accordingly

switch($permLevel)
{
    1 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights Owner}
    2 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights PublishingEditor}
    3 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights Editor}
    4 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights PublishingAuthor}
    5 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights Author}
    6 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights NonEditingAuthor}
    7 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights Reviewer}
    8 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights Contributor}
    9 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights LimitedDetails}
    10 {Add-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user -AccessRights AvailabilityOnly}
    11 {Remove-MailboxFolderPermission -Identity "$($identity):\calendar" -User $user}
}

# Display new permission Add for selected user

Write-Host ""
Write-Host " >> Updated permissions list for calendar of $($identity)"
Write-Host ""
Get-MailboxFolderPermission -Identity "$($identity):\calendar"
Exit