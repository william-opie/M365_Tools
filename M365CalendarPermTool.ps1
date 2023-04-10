<#
.SYNOPSIS
    This script provides a convenient front-end to view & modify calendar permissions in M365 from Powershell.
.DESCRIPTION
    This tool provides the following options:
        1. Single Calendar Permissions Tool
            - View a Calendar's Current Permissions: View a single mailbox calendar's permissions at a glance.
            - Add Permissions to a Single Calendar: Add permissions to a user's mailbox calendar. (Ex. "Give Bob editor-level access to Tom's Outlook calendar.")
            - Remove Permissions from a Single Calendar: Remove permissions from a user's mailbox calendar. (Ex. "Remove Bob's access to Tom's Outlook calendar.")
        2. Tenant-Wide Default User Permission Change Tool: Change the Default user permissions for all users' calendars.
        3. Export Tenant-Wide Calendar Permissions Report: Exports a .CSV with the permissions for every calendar in the Exchange tenant.

.NOTES
    @Author: Will Opie
    @Initial Date: 2022-11-25
    @Version: 2023-03-31
#>
#region Functions ------------------------------------------------------------------------------------------------
Function pause ($message)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $null = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

#Below function adapted from Office 365 Reports "GetMailboxCalendarPermissions" script 
#(Link: https://o365reports.com/2021/11/02/get-calendar-permissions-report-for-office365-mailboxes-powershell/)
Function Get-CalendarPermsReport
{
    Get-Mailbox -ResultSize Unlimited | ForEach-Object {
        $CurrUserData = $_
        $global:MailboxCount = $global:MailboxCount+1
        $EmailAddress = $CurrUserData.PrimarySmtpAddress
        $CalendarFolders=@()
        $CalendarStats = Get-MailboxFolderStatistics -Identity $EmailAddress -FolderScope Calendar
    
        #Processing the calandar folder path
        ForEach($LiveCalendarFolder in $CalendarStats){
            if (($LiveCalendarFolder.FolderType) -eq "Calendar"){
                $CurrCalendarFolder = $EmailAddress + ":\Calendar"
            }
            else{
                $CurrCalendarFolder = $EmailAddress + ":\Calendar\" + $LiveCalendarFolder.Name
            }
            $CalendarFolders += $CurrCalendarFolder  
        }
        foreach($CalendarFolder in $CalendarFolders){
            $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
            Write-Progress "Checking calendar permission in: $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
            Get-MailboxFolderPermission -Identity $CalendarFolder | ForEach-Object {
                $CurrCalendarData=$_
                $Identity = $CurrUserData.Identity
                $global:ReportSize = $global:ReportSize + 1
                $MailboxType = $CurrUserData.RecipientTypeDetails
                $CalendarName = $CalendarName
                $SharedToMB=$CurrCalendarData.User.DisplayName
                if ($SharedToMB.StartsWith("ExchangePublishedUser.")){
                    $AllowedUser = $SharedToMB -replace ("ExchangePublishedUser.", "")      
                    $UserType = "External/Unauthorized"
                }
                else{
                    $AllowedUser = $SharedToMB
                    $UserType = "Member"
                }
                $AccessRights = $CurrCalendarData.AccessRights -join ","
                if ($Empty -ne ($CurrCalendarData.SharingPermissionFlags)){
                    $PermissionFlag = $CurrCalendarData.SharingPermissionFlags -join ","
                }
                else{
                    $PermissionFlag = "-"
                }
                $ExportResult = @{ 
                    'Mailbox Name'             = $Identity;
                    'Email Address'            = $EmailAddress;
                    'Mailbox Type'             = $MailboxType; 
                    'Calendar Name'            = $CalendarName;
                    'Shared To'                = $AllowedUser;
                    'User Type'                = $UserType;
                    'Access Rights'            = $AccessRights;
                    'Sharing Permission Flags' = $PermissionFlag;
                }
                
                $ExportResults = New-Object PSObject -Property $ExportResult
                $ExportResults | Select-object 'Mailbox Name', 'Email Address', 'Mailbox Type', 'Calendar Name', 'Shared To', 'Access Rights', 'Sharing Permission Flags', 'User Type' | Export-csv -path $ExportCSV -NoType -Append
            }
        }
    }
    
}

#endregion Functions --------------------------------------------------------------------------------------------------

$global:MailboxCount = 0
$global:ReportSize = 0

#Verifies ExchangeOnlineManagement module is installed before proceeding
if ($null -eq (Get-Module -ListAvailable -Name ExchangeOnlineManagement)){
    Write-Host "ExchangeOnlineManagement module is not installed. Script requires this module to function." -ForegroundColor Yellow -BackgroundColor Red
    Write-Host "Please install the ExchangeOnlineManagement module and try again." -ForegroundColor Yellow
    Pause "`nPress any key to exit."
    Exit
}

#Set current PS Session to TLS 1.2
$TLS12Protocol = [System.Net.SecurityProtocolType] 'Ssl3 , Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $TLS12Protocol

#Removes any active PS Sessions before starting script
$SessionCheck = Get-PSSession
If($Null -ne $SessionCheck){
    Get-PSSession | Remove-PSSession; Disconnect-ExchangeOnline
}
Import-Module ExchangeOnlineManagement
Clear-Host

Do{
    Write-Host "M365 Calendar Perms Tool`n" -ForegroundColor Green
    $GlobalAdmin = Read-Host "Please enter a Global Admin email address"
    Try{Connect-ExchangeOnline -UserPrincipalName $GlobalAdmin}
    Catch [System.AggregateException] {
        Clear-Host
        Write-Host "Invalid entry for Global Admin account. Please try again." -ForegroundColor Red -BackgroundColor Yellow
        $GlobalAdmin = $Null
    }
    $ActiveSession = Get-PSSession
} Until ($ActiveSession)
Clear-Host

While($true){
    Clear-Host
    Write-Host "M365 Calendar Perms Tool`n" -ForegroundColor Green
    Write-Host "`nModes:`n1. Single Calendar Permissions Tool`n2. Tenant-Wide Default User Permission Change Tool`n3. Export Tenant-Wide Calendar Permissions Report`n4. Exit"
    Write-Host "`nLogged in as Global Admin account: " -NoNewLine
    Write-Host "$GlobalAdmin" -ForegroundColor Green
    $InitialMenuSelection = Read-Host "`nPlease enter selection"

    switch($InitialMenuSelection){
        1{
            $SingleCalSelection = $Null
            Clear-Host
            Do{
                Write-Host "Single Calendar Perms Tool`n" -ForegroundColor Cyan
                Write-Host "`nPlease choose from the following options:`n"
                Write-Host "1. View a Calendar's Current Permissions"
                Write-Host "2. Add Permissions to a Single Calendar"
                Write-Host "3. Remove Permissions from a Single Calendar"
                $SingleCalSelection = Read-Host "`nEnter your selecion here"
                switch($SingleCalSelection){
                    1{
                        $Mailbox = $Null
                        $MailboxCheck = $Null
                        Clear-Host
                        While($true){
                            Write-Host "View Single Calendar's Current Permissions`n" -ForegroundColor Cyan
                            Do{
                                $Mailbox = Read-Host "Enter the email address for the user whose calendar permissions you would like to view"
                                $MailboxCheck = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
                                If($Null -eq $MailboxCheck){
                                    Write-Host "`nInvalid entry. Please enter a valid email address.`n" -BackgroundColor Red -ForegroundColor Yellow
                                }
                            }Until($MailboxCheck)
                            Write-Host "`nBelow are the currently set calendar permissions for $Mailbox :`n"
                            $Cal = ($Mailbox+":\Calendar")
                            Get-MailboxFolderPermission -Identity $Cal | Format-Table FolderName,User,AccessRights
                            Do{
                                $SearchAgain = Read-Host "Would you like to search another mailbox's calendar permissions? Y/N"
                                switch($SearchAgain){
                                    "Y"{Clear-Host}
                                    "N"{
                                        Break
                                    }
                                    Default{
                                        Write-Host "`nInvalid input. Please try again.`n" -ForegroundColor Red
                                    }
                                }
                            }Until($SearchAgain -like "[Y/N]")
                            If($SearchAgain -eq "N"){
                                Break
                            }
                        }
                        Pause "`nPress any key to return to the main menu."
                    }
                    2{
                        Clear-Host
                        Do{
                            Write-Host "Add Permissions to a Single Calendar`n" -ForegroundColor Cyan
                            $Mailbox = Read-Host "Enter the email address for the user mailbox calendar whose permissions you would like to modify"
                            $MailboxCheck = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
                            If($Null -eq $MailboxCheck){
                                Clear-Host
                                Write-Host "`nInvalid entry. Please enter a valid email address.`n" -BackgroundColor Red -ForegroundColor Yellow
                            }
                        }Until($MailboxCheck)
                        $MailboxCalendar = $Mailbox+":\Calendar"
                        Do{
                            Write-Host "`n`nEnter the email address for the user who you want to give access permissions to the mailbox calendar for " -NoNewLine
                            Write-Host "$Mailbox" -ForegroundColor Yellow -NoNewline
                            Write-Host ":"
                            $AddPermsMailbox = Read-Host
                            $AddPermsMailboxCheck = Get-Mailbox -Identity $AddPermsMailbox -ErrorAction SilentlyContinue
                            If($Null -eq $AddPermsMailboxCheck){
                                Clear-Host
                                Write-Host "`nInvalid entry. Please enter a valid email address.`n" -BackgroundColor Red -ForegroundColor Yellow
                            }
                            Else{
                                $AddPermsMailboxName = (Get-Mailbox -Identity $AddPermsMailbox).DisplayName
                            }
                        }Until($AddPermsMailboxCheck)
                        
                        Write-Host "`n`nChecking current access level " -NoNewLine
                        Write-Host "$AddPermsMailbox" -ForegroundColor Yellow -NoNewline
                        Write-Host " has to " -NoNewLine
                        Write-Host "$Mailbox" -ForegroundColor Green -NoNewLine
                        Write-Host "'s calendar....`n`n"
                        $CurrentAccessCheck = (Get-MailboxFolderPermission -identity "$MailboxCalendar" | Where-Object -Property User -match "$AddPermsMailboxName").AccessRights
                        If($Null -ne $CurrentAccessCheck){
                            Write-Host "$AddPermsMailbox" -ForegroundColor Yellow -NoNewline
                            Write-Host " has the following access to " -NoNewline
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar: " -NoNewline
                            Write-Host "$CurrentAccessCheck" -ForegroundColor Cyan
                        }
                        Else{
                            Write-Host "$AddPermsMailbox" -ForegroundColor Yellow -NoNewline
                            Write-Host " currently has no explicitly defined access to " -NoNewline
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar."
                        }
                        Do{
                            Write-Host "`n`nPlease select the desired permissions level you want to give " -NoNewLine
                            Write-Host "$AddPermsMailbox" -ForegroundColor Yellow -NoNewLine
                            Write-Host " to access " -NoNewLine
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar:`n"
                            Write-Host "1. Owner: Allows read, create, modify and delete all items and folders. Also allows manage items permissions."
                            Write-Host "2. PublishingEditor: Allows read, create, modify and delete items/subfolders."
                            Write-Host "3. Editor: Allows read, create, modify and delete items."
                            Write-Host "4. PublishingAuthor: Allows read, create all items/subfolders. User can modify and delete only items they have created."
                            Write-Host "5. Author: Allows create and read items; edit and delete own items."
                            Write-Host "6. NonEditingAuthor: Allows full read access and create items. User can delete only their own items."
                            Write-Host "7. Reviewer: Read only."
                            Write-Host "8. Contributor: Allows user to create items and folders."
                            Write-Host "9. AvailabilityOnly: Allows read free/busy information from calendar."
                            Write-Host "10. None: No permissions to access folder and files."
                            $PermRoleSelection = Read-Host "`nEnter desired perms level"
                            switch($PermRoleSelection){
                                1{$PermRole = "Owner"}
                                2{$PermRole = "PublishingEditor"}
                                3{$PermRole = "Editor"}
                                4{$PermRole = "PublishingAuthor"}
                                5{$PermRole = "Author"}
                                6{$PermRole = "NonEditingAuthor"}
                                7{$PermRole = "Reviewer"}
                                8{$PermRole = "Contributor"}
                                9{$PermRole = "AvailabilityOnly"}
                                10{$PermRole = "None"}
                                Default{
                                    Clear-Host
                                    Write-Host "Invalid selection. Please try again.`n`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($PermRoleSelection -like "[1/2/3/4/5/6/7/8/9/10]")
                        If(!$CurrentAccessCheck){
                            Add-MailboxFolderPermission -Identity $MailboxCalendar -User $AddPermsMailboxName -AccessRights $PermRole
                        }
                        Else{
                            Set-MailboxFolderPermission -Identity $MailboxCalendar -User $AddPermsMailboxName -AccessRights $PermRole
                        }
                        Clear-Host
                        Write-Host "Checking permissions for " -NoNewline
                        Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                        Write-Host "'s calendar after change..."
                        $PostAccessCheck = (Get-MailboxFolderPermission -identity "$MailboxCalendar" | Where-Object -Property User -match "$AddPermsMailboxName").AccessRights
                        If($PostAccessCheck -ne $PermRole){
                            Write-Host "`nERROR: Permissions change failed!" -BackgroundColor Red -ForegroundColor Yellow
                            Write-Host "`n$AddPermsMailbox" -ForegroundColor Yellow -NoNewLine
                            Write-Host " has the following access  to " -NoNewLine
                            Write-Host "$Mailbox" -ForegroundColor Red -NoNewline
                            Write-Host "'s calendar: " -NoNewline
                            Write-Host "$PostAccessCheck" -ForegroundColor Cyan
                        }
                        Else{
                            Write-Host "`nSUCCESS: Permissions change succeeded!" -BackgroundColor Green -ForegroundColor White
                            Write-Host "`n$AddPermsMailbox" -ForegroundColor Yellow -NoNewLine
                            Write-Host " has the following access  to " -NoNewLine
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar: " -NoNewline
                            Write-Host "$PostAccessCheck`n" -ForegroundColor Cyan
                        }
                        Pause "`nPress any key to return to the main menu."
                    }
                    3{
                        Clear-Host
                        Do{
                            Write-Host "Remove Permissions from a Single Calendar`n" -ForegroundColor Cyan
                            $Mailbox = Read-Host "Enter the email address for the user mailbox calendar whose permissions you would like to modify"
                            $MailboxCheck = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
                            If($Null -eq $MailboxCheck){
                                Clear-Host
                                Write-Host "`nInvalid entry. Please enter a valid email address.`n" -BackgroundColor Red -ForegroundColor Yellow
                            }
                        }Until($MailboxCheck)
                        $MailboxCalendar = $Mailbox+":\Calendar"
                        Do{
                            Write-Host "`nEnter the email address for the user who you want to remove access permissions from the mailbox calendar for " -NoNewLine 
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewLine
                            Write-Host ":"
                            $DeletePermsMailbox = Read-Host
                            $DeletePermsMailboxCheck = Get-Mailbox -Identity $DeletePermsMailbox -ErrorAction SilentlyContinue
                            If($Null -eq $DeletePermsMailboxCheck){
                                Clear-Host
                                Write-Host "`nInvalid entry. Please enter a valid email address.`n" -BackgroundColor Red -ForegroundColor Yellow
                            }
                            Else{
                                $DeletePermsMailboxName = (Get-Mailbox -Identity $DeletePermsMailbox).DisplayName
                            }
                        }Until($DeletePermsMailboxCheck)
                        Write-Host "`n`nChecking current access level " -NoNewLine
                        Write-Host "$DeletePermsMailbox" -ForegroundColor Yellow -NoNewline
                        Write-Host " has to " -NoNewLine
                        Write-Host "$Mailbox" -ForegroundColor Green -NoNewLine
                        Write-Host "'s calendar....`n`n"
                        $CurrentAccessCheck = (Get-MailboxFolderPermission -identity "$MailboxCalendar" | Where-Object -Property User -match "$DeletePermsMailboxName").AccessRights
                        If($Null -ne $CurrentAccessCheck){
                            Do{
                                Write-Host "$DeletePermsMailbox" -ForegroundColor Yellow -NoNewline
                                Write-Host " has the following access to " -NoNewline
                                Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                                Write-Host "'s calendar: " -NoNewLine
                                Write-Host "$CurrentAccessCheck" -ForegroundColor Cyan
                                Write-Host "`n`nWould you like to move forward with removing " -NoNewline
                                Write-Host "$DeletePermsMailbox" -NoNewline -ForegroundColor Yellow
                                Write-Host "'s access to " -NoNewline
                                Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                                Write-Host "'s calendar? Y/N`n`n"
                                $DeleteMenuSelection = Read-Host "Enter your selection"
                                switch($DeleteMenuSelection){
                                    "Y"{
                                        Clear-Host
                                        Write-Host "Removing calendar permissions..." -ForegroundColor Cyan
                                        Remove-MailboxFolderPermission -Identity "$MailboxCalendar" -User $DeletePermsMailbox -Confirm:$false
                                        Write-Host "`n`nChecking permissions for " -NoNewline
                                        Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                                        Write-Host " after removal..."
                                        $PostAccessCheck = (Get-MailboxFolderPermission -identity "$MailboxCalendar" | Where-Object -Property User -match "$DeletePermsMailboxName").AccessRights
                                        If(!$PostAccessCheck){
                                            Write-Host "`nSUCCESS: Permissions change succeeded!" -BackgroundColor Green -ForegroundColor White
                                            Write-Host "`n$DeletePermsMailbox" -ForegroundColor Yellow -NoNewLine
                                            Write-Host " no longer has explicit access to " -NoNewLine
                                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                                            Write-Host "'s calendar."
                                        }
                                        Else{
                                            Write-Host "`nERROR: Permissions change failed!" -BackgroundColor Red -ForegroundColor Yellow
                                            Write-Host "`n$DeletePermsMailbox" -ForegroundColor Yellow -NoNewLine
                                            Write-Host " still has the following access to " -NoNewLine
                                            Write-Host "$Mailbox" -ForegroundColor Red -NoNewline
                                            Write-Host "'s calendar: " -NoNewLine
                                            Write-Host "$PostAccessCheck" -ForegroundColor Cyan
                                        }

                                    }
                                    "N"{
                                        Clear-Host
                                        Write-Host "Returning to main menu without modifying calendar permissions..." -ForegroundColor Cyan
                                    }
                                    Default{
                                        Clear-Host
                                        Write-Host "Invalid selection. Please try again.`n`n" -BackgroundColor Yellow -ForegroundColor Red
                                    }
                                }
                            }Until($DeleteMenuSelection -like "[Y/N]")
                        }
                        Else{
                            Write-Host "$DeletePermsMailbox" -ForegroundColor Yellow -NoNewline
                            Write-Host " currently has no explicitly defined access to " -NoNewline
                            Write-Host "$Mailbox" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar. No need to delete user access perms."
                        }
                        Pause "`nPress any key to return to the main menu."
                    }
                    Default{
                        Clear-Host
                        Write-Host "Invalid selection. Please try again.`n`n" -BackgroundColor Yellow -ForegroundColor Red
                    }
                }
            }Until($SingleCalSelection)
            
        }
        2{
            $DefaultPermSelection = $Null
            $PermRole = $Null
            $PermRoleSelection = $Null
            Do{
                Clear-Host
                Write-Host "Tenant-Wide Default User Permissions Change`n" -ForegroundColor Cyan
                Write-Host "Please select from the following options:`n"
                Write-Host "1. View Default User access for All User Mailbox Calendars"
                Write-Host "2. Set Default User Permissions for All User Mailbox Calendars"
                Write-Host "3. Return to main menu"
                $DefaultPermSelection = Read-Host "`nEnter your selection here"
                switch($DefaultPermSelection){
                    1{
                        Clear-Host
                        Write-Host "Default User currently has the following access to the below users' mailbox calendars:`n`n"
                        $MailboxList = @()
                        $MailboxList += (Get-Mailbox -RecipientTypeDetails UserMailbox)
                        ForEach($User in Get-Mailbox -RecipientTypeDetails UserMailbox){
                            $cal = $user.alias+":\Calendar"
                            $perms = (Get-MailboxFolderPermission -Identity $cal -User Default).AccessRights
                            Write-Host "Default" -NoNewline -ForegroundColor Cyan
                            Write-Host " user has " -NoNewLine
                            Write-Host "$perms" -ForegroundColor Yellow -NoNewline 
                            Write-Host " access to " -NoNewLine
                            Write-Host "$user" -ForegroundColor Green -NoNewline
                            Write-Host "'s calendar"
                        }
                    }
                    2{
                        Clear-Host
                        Do{
                            Write-Host "Set Default User Permissions for All User Tenant Mailboxes" -ForegroundColor Yellow
                            Write-Host "`nPlease select the desired Default User permissions level for all users calendars:`n"
                            Write-Host "1. Owner: Allows read, create, modify and delete all items and folders. Also allows manage items permissions."
                            Write-Host "2. PublishingEditor: Allows read, create, modify and delete items/subfolders."
                            Write-Host "3. Editor: Allows read, create, modify and delete items."
                            Write-Host "4. PublishingAuthor: Allows read, create all items/subfolders. User can modify and delete only items they have created."
                            Write-Host "5. Author: Allows create and read items; edit and delete own items."
                            Write-Host "6. NonEditingAuthor: Allows full read access and create items. User can delete only their own items."
                            Write-Host "7. Reviewer: Read only."
                            Write-Host "8. Contributor: Allows user to create items and folders."
                            Write-Host "9. AvailabilityOnly: Allows read free/busy information from calendar."
                            Write-Host "10. None: No permissions to access folder and files."
                            $PermRoleSelection = Read-Host "`nEnter desired access level for Default user"
                            switch($PermRoleSelection){
                                1{$PermRole = "Owner"}
                                2{$PermRole = "PublishingEditor"}
                                3{$PermRole = "Editor"}
                                4{$PermRole = "PublishingAuthor"}
                                5{$PermRole = "Author"}
                                6{$PermRole = "NonEditingAuthor"}
                                7{$PermRole = "Reviewer"}
                                8{$PermRole = "Contributor"}
                                9{$PermRole = "AvailabilityOnly"}
                                10{$PermRole = "None"}
                                Default{
                                    Clear-Host
                                    Write-Host "Invalid selection. Please try again.`n`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($PermRoleSelection -like "[1/2/3/4/5/6/7/8/9/10]")
                        ForEach($User in Get-Mailbox -RecipientTypeDetails UserMailbox){
                            $cal = $user.alias+":\Calendar"
                            Set-MailboxFolderPermission -Identity $cal -User Default -AccessRights $PermRole
                        }
                        Write-Host "`n`nChecking access permissions for " -NoNewline
                        Write-Host "Default user" -ForegroundColor Green -NoNewline
                        Write-Host " after change..."
                        $PostAccessCheck = (Get-MailboxFolderPermission -identity ($MailboxList.Name[0] + ":\Calendar") -User Default).AccessRights
                        If($PermRole -eq $PostAccessCheck){
                            Write-Host "`nSUCCESS: Permissions change succeeded!" -BackgroundColor Green -ForegroundColor White
                            Write-Host "Default user" -ForegroundColor Green -NoNewLine
                            Write-Host " now has " -NoNewline
                            Write-Host "$PermRole" -ForegroundColor Cyan -NoNewLine
                            Write-Host " access to all user calendars." -ForegroundColor Green -NoNewline
                        }
                        Else{
                            Write-Host "`nERROR: Permissions change failed!" -BackgroundColor Red -ForegroundColor Yellow
                            Write-Host "`nDefault user" -ForegroundColor Yellow -NoNewLine
                            Write-Host " has the following access to user mailboxes: " -NoNewLine
                            Write-Host "$PostAccessCheck" -ForegroundColor Red -NoNewline
                        }
                    }
                    3{
                        Clear-Host
                    }
                    Default{
                        Write-Host "Invalid selection. Please try again." -BackgroundColor Yellow -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        Clear-Host
                    }
                }
            }Until($DefaultPermSelection -like "[1/2/3]")
            Pause "`nPress any key to return to the main menu."
        }
        3{
            Clear-Host
            Write-Host "Export Tenant-Wide Calendar Permissions Report`n`n" -ForegroundColor Cyan
            Do{
                Do{
                    $ExportCSV = Read-Host "Enter file path for exporting the tenant-wide calendar permissions report. (Use .csv for file extension)"
                    If(-Not ($ExportCSV.Contains(".csv"))){
                        Write-Host "`n$ExportCSV" -NoNewLine -ForegroundColor Red
                        Write-Host " must have a .csv file extension.`n"
                    }
                }Until($ExportCSV.Contains(".csv"))
                Set-Content -Path $ExportCSV -Value "Mailbox Name,Email Address,Mailbox Type,Calendar Name,Shared To,Access Rights,Sharing Permission Flags,User Type"
                If(-Not(Test-Path -Path $ExportCSV)){
                    Write-Host "`nInvalid path " -NoNewLine 
                    Write-Host "$ExportCSV" -NoNewLine -ForegroundColor Red
                    Write-Host ". Please enter a valid file path.`n" 
                }
            }Until(Test-Path -Path $ExportCSV)
            Write-Host "`nGenerating tenant-wide calendar permissions report...`n" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Get-CalendarPermsReport       
            Write-Progress -Activity "Checking calendar permission in: $CalendarFolder" -Completed
            Start-Sleep -Seconds 1
            Clear-Host
            If(Test-Path -Path $ExportCSV){
                Write-Host "Tenant-wide calendar permissions report exported to .CSV successfully: " -NoNewLine
                Write-Host "$ExportCSV" -ForegroundColor Green
            }
            Else{
                Write-Host "Failed to export tenant-wide calendar permissions report successfully." -ForegroundColor Yellow -BackgroundColor Red
            }
            Pause "Press any key to return to the main menu."
        }
        4{Exit}
        default{
            Write-Host "`nInvalid selection; please try again.`n" -ForegroundColor Yellow -BackgroundColor Red
            Start-Sleep -Seconds 3
        }
    }
}
