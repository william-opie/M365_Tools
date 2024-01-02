<#
.SYNOPSIS
    This script provides a convenient front-end for creating new content searches in M365 Compliance/Purview from Powershell.
.DESCRIPTION
    This tool provides the following options:
    - Start a single user search, retreiving all emails and/or Teams messages the user sent and received (during a specified timeframe, or without a specified date range).
    - Start a multi-user search, retreiving all emails and/or Teams messages sent and received between the users specified in the search (during a specified timeframe, or without a specified date range).
    - Check on status of existing content searches & start export job from existing content searches
    - Check on status of existing content search export jobs

    The tool also allows you to input KQL search queries when creating new content searches (Note: There is no input validation on KQL search queries).
.NOTES
    @Author: Will Opie
    @Initial Date: 2022-09-30
    @Version: 2024-01-01

    This script requires the Exchange Online Powershell module to function (https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
#>

#region Functions ------------------------------------------------------------------------------------------------
Function Exit-Prompt ($message)
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
#endregion Functions --------------------------------------------------------------------------------------------------

#Verifies ExchangeOnlineManagement module is installed before proceeding
if ($null -eq (Get-Module -ListAvailable -Name ExchangeOnlineManagement)){
    Write-Host "ExchangeOnlineManagement module is not installed. Script requires this module to function." -ForegroundColor Yellow -BackgroundColor Red
    Write-Host "Please install the ExchangeOnlineManagement module and try again." -ForegroundColor Yellow
    Exit-Prompt "`nPress any key to exit."
    Exit
}

#Removes any active ExchangeOnline connections before starting script
$SessionCheck = Get-ConnectionInformation | Where-Object {$_.Name -match 'ExchangeOnline' -and $_.state -eq 'Connected'}
If($Null -ne $SessionCheck){
    Disconnect-ExchangeOnline
}
Import-Module ExchangeOnlineManagement
Clear-Host

Do{
    Write-Host "M365 Content Search Tool`n" -ForegroundColor Green
    $GlobalAdmin = Read-Host "Enter global admin email address"
    Connect-IPPSSession -UserPrincipalName $GlobalAdmin -ShowBanner:$false
    $ActiveSession = Get-ConnectionInformation | Where-Object {$_.Name -match 'ExchangeOnline' -and $_.state -eq 'Connected'}
} Until ($ActiveSession)
Connect-ExchangeOnline -UserPrincipalName $GlobalAdmin -ShowBanner:$false
Clear-Host

While($true){
    Clear-Host
    Write-Host "M365 Content Search Tool`n" -ForegroundColor Green
    Write-Host "`nModes:`n1. Start New Search`n2. Check Existing Cases & Start New Export Job`n3. Check Existing Export Jobs`n4. Exit"
    Write-Host "`nLogged in as Global Admin account: " -NoNewLine
    Write-Host "$GlobalAdmin" -ForegroundColor Green
    $InitialMenuSelection = Read-Host "`nPlease enter selection"

    switch($InitialMenuSelection){
        1{
            Clear-Host
            Do{
                Write-Host "New Content Search`n" -ForegroundColor Yellow
                Write-Host "Search Options`n1. Single User Search`n2. Multi-User Search`n3. Return to main menu"
                $NewSearchMenu = Read-Host "Enter selection"
                Switch($NewSearchMenu){
                    1{
                        $Date = $Null
                        $DateKQL = $Null
                        $DateQuerySuccess = $False
                        $Participant = $Null
                        $Mailbox = $Null
                        $ParticipantKQL = $Null
                        $RawKQL = $Null
                        $KQLQuery = $Null
                        $SearchSuccess = $True
                        Clear-Host
                        Write-Host "Single User Search`n" -ForegroundColor Yellow
                        $NewSearchName = Read-Host "Please enter a name for the new Content Search"
                        While($true){
                            Write-Host "`nPlease enter a date range for the search.`nWrite date range in YYYY-MM-DD format, with two periods between dates. (ex. 2021-07-01..2022-09-15)"
                            Write-Host "If you would like to search without a date range, enter N/A`n"
                            $Date = Read-Host "Please enter date range"
                            switch -Regex ($Date) {
                                '(20[0-9]{2})-(0[1-9]|[1][0-2])-([3][0-1]|[1-2][0-9]|0[1-9])\.\.(20[0-9]{2})-(0[1-9]|[1][0-2])-([3][0-1]|[1-2][0-9]|0[1-9])'{
                                    $DateKQL = ("(Date=$Date)")
                                    $DateQuerySuccess = $True
                                }
                                '[nN]\/[aA]'{
                                    $DateQuerySuccess = $True
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "Invalid input. Try again."
                                }
                            }
                            If($DateQuerySuccess){Break}
                        }
                        Do{
                            Do{
                                $Participant = Read-Host "Please enter the email address of the user whose messages you would like to search"
                                Try{$Mailbox = Get-EXOMailbox -Identity $Participant -ErrorAction Stop}
                                Catch {
                                    Clear-Host
                                    Write-Host "No mailbox found with this address: $Participant`n`n"
                                    $Participant = $Null
                                    $Mailbox = $Null
                                }
                            }Until($Null -ne $Mailbox)
                            $ParticipantKQL = ("(Participants:$Participant)")
                            Write-Host "Is this the correct email address: " -NoNewLine
                            Write-Host "$Participant" -ForegroundColor Yellow -NoNewLine
                            $Confirmation = Read-Host "? Y/N"
                        }Until($Confirmation -eq 'y')
                        $KQLQuery = ($DateKQL + $ParticipantKQL)
                        Do{
                            $EmailSearchCheck = Read-Host "`nWould you like to filter your search for email messages? Y/N"
                            switch($EmailSearchCheck){
                                'Y'{
                                    $KQLQuery += "(kind:email)"
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($EmailSearchCheck -like "[y|n]")
                        Do{
                            $TeamsSearchCheck = Read-Host "`nWould you like to filter your search for Teams messages? Y/N"
                            switch($TeamsSearchCheck){
                                'Y'{
                                    $KQLQuery += "(kind:im)"
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($TeamsSearchCheck -like "[y|n]")
                        Do{
                            $RawKQLCheck = Read-Host "`nWould you like to input a custom KQL query? Y/N"
                            switch($RawKQLCheck){
                                'Y'{
                                    Write-Host "NOTE: Custom KQL queries do not have input validation." -BackgroundColor Yellow
                                    $RawKQL = Read-Host "`nEnter your custom KQL query here"
                                    $KQLQuery += $RawKQL
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($RawKQLCheck -like "[y|n]")
                        New-ComplianceSearch -Name "$NewSearchName" -ExchangeLocation "$Participant" -ContentMatchQuery "$KQLQuery" | Format-Table
                        Start-Sleep -Seconds 3
                        Clear-Host
                        Start-ComplianceSearch -Identity $NewSearchName
                        Try{Get-ComplianceSearch -Identity $NewSearchName -ErrorAction Stop | Format-Table}
                        Catch {$SearchSuccess = $False}
                        If($SearchSuccess){
                            Write-Host "Content Search started for job " -NoNewLine
                            Write-Host "$NewSearchName" -NoNewLine -ForegroundColor Green 
                            Write-Host ". Please check on job progress later today."
                        }
                        Else{
                            Write-Host "Content Search failed to start for job " -NoNewline
                            Write-Host "$NewSearchName" -NoNewLine -ForegroundColor Red
                            Write-Host ". Please try running Content Search manually through the Compliance portal."
                        }
                        Exit-Prompt "`nPress any key to return to the main menu."
                    }
                    2{
                        $Date = $Null
                        $DateKQL = $Null
                        $DateQuerySuccess = $False
                        $Response = $Null
                        $Mailbox = $Null
                        $MailboxList = @()
                        $SenderKQL = $Null
                        $RecipientKQL = $Null
                        $RawKQL = $Null
                        $KQLQuery = $Null
                        $SearchSuccess = $True
                        Clear-Host
                        Write-Host "Multi-User Search`n" -ForegroundColor Yellow
                        $NewSearchName = Read-Host "Please enter a name for the new Content Search"
                        While($true){
                            Write-Host "`nPlease enter a date range for the search.`nWrite date range in YYYY-MM-DD format, with two periods between dates. (ex. 2021-07-01..2022-09-15)"
                            Write-Host "If you would like to search without a date range, enter N/A`n"
                            $Date = Read-Host "Please enter date range"
                            switch -Regex ($Date) {
                                '(20[0-9]{2})-(0[1-9]|[1][0-2])-([3][0-1]|[1-2][0-9]|0[1-9])\.\.(20[0-9]{2})-(0[1-9]|[1][0-2])-([3][0-1]|[1-2][0-9]|0[1-9])'{
                                    $DateKQL = ("(Date=$Date)")
                                    $DateQuerySuccess = $True
                                }
                                '[nN]\/[aA]'{
                                    $DateQuerySuccess = $True
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "Invalid input. Try again."
                                }
                            }
                            If($DateQuerySuccess){Break}
                        }
                        Do{
                            Do{
                                Do{
                                    $Mailbox = Read-Host "Please enter the email address for the mailbox you would like to search"
                                    Try{$MailboxCheck = Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop}
                                    Catch {
                                        Clear-Host
                                        Write-Host "No mailbox found with this address: $Mailbox`n`n"
                                        $Mailbox = $Null
                                        $MailboxCheck = $Null
                                    }
                                }Until($Null -ne $MailboxCheck)
                                $MailboxList += $Mailbox
                                $SenderKQL += ("(From:$Mailbox)")
                                $RecipientKQL += ("(recipients:$Mailbox)")
                                $Response = Read-Host "`nAre there any additional mailboxes you would like to search? Y/N"
                            } Until($Response -eq 'n')
                            Clear-Host
                            Write-Output $MailboxList
                            $Confirmation = Read-Host "`nAre these all of the mailboxes you would like to search? Y/N"
                        }Until($Confirmation -eq 'y')
                        $KQLQuery = ($SenderKQL + $RecipientKQL + $DateKQL)
                        Do{
                            $EmailSearchCheck = Read-Host "`nWould you like to filter your search for email messages? Y/N"
                            switch($EmailSearchCheck){
                                'Y'{
                                    $KQLQuery += "(kind:email)"
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($EmailSearchCheck -like "[y|n]")
                        Do{
                            $TeamsSearchCheck = Read-Host "`nWould you like to filter your search for Teams messages? Y/N"
                            switch($TeamsSearchCheck){
                                'Y'{
                                    $KQLQuery += "(kind:im)"
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($TeamsSearchCheck -like "[y|n]")
                        Do{
                            $RawKQLCheck = Read-Host "`nWould you like to input a custom KQL query? Y/N"
                            switch($RawKQLCheck){
                                'Y'{
                                    Write-Host "NOTE: Custom KQL queries do not have input validation." -BackgroundColor Yellow
                                    $RawKQL = Read-Host "`nEnter your custom KQL query here"
                                    $KQLQuery += $RawKQL
                                }
                                'N'{
                                    Break
                                }
                                Default{
                                    Clear-Host
                                    Write-Host "`nInvalid input. Please enter y or n.`n" -BackgroundColor Yellow -ForegroundColor Red
                                }
                            }
                        }Until($RawKQLCheck -like "[y|n]")
                        New-ComplianceSearch -Name "$NewSearchName" -ExchangeLocation $MailboxList -ContentMatchQuery "$KQLQuery"
                        Start-Sleep -Seconds 3
                        Start-ComplianceSearch -Identity "$NewSearchName"
                        Try{Get-ComplianceSearch -Identity $NewSearchName -ErrorAction Stop | Format-Table}
                        Catch {$SearchSuccess = $False}
                        If($SearchSuccess){
                            Write-Host "Content Search started for job " -NoNewLine
                            Write-Host "$NewSearchName" -NoNewLine -ForegroundColor Green 
                            Write-Host ". Please check on job progress later today."
                        }
                        Else{
                            Write-Host "Content Search failed to start for job " -NoNewline
                            Write-Host "$NewSearchName" -NoNewLine -ForegroundColor Red
                            Write-Host ". Please try running Content Search manually through the Compliance portal."
                        }
                        Exit-Prompt "`nPress any key to return to the main menu."
                    }
                    3{
                        Write-Host "`nReturning to main menu...`n" -ForegroundColor Green
                        Exit-Prompt "`nPress any key to return to the main menu."
                    }
                    default{
                        Clear-Host
                        Write-Host "Invalid selection. Please try again..."
                    }
                }
            }Until($InitialMenuSelection -like "[1/2/3]")
        }
        2{
            Clear-Host
            Do{
                Write-Host "Existing Case Search/Start Export Job`n" -ForegroundColor Yellow
                Write-Host "Below are all existing Content Search cases for this tenant:`n"
                Get-ComplianceSearch | Format-Table
                $SearchMenuSelection = Read-Host "`nWould you like to start an export for one of the above content searches? Y/N"
                    Switch($SearchMenuSelection){
                        'Y' {
                            $ExportSearchSuccess = $True
                            $ContentSearches = Get-ComplianceSearch
                            $i=1;
                            Write-Host "`nContent Search Job Export Selection`n" -ForegroundColor Yellow
                            Get-ComplianceSearch | ForEach-Object{"$($i).) $($_.name)"; $i++}
                            $ExportJob = Read-Host "`nPlease select which content search you would like to export. Otherwise, input 'exit' to return to main menu."
                            If($ExportJob -eq "exit"){Break}
                            $ExportJobSearch = ($ContentSearches.name[($ExportJob - 1)])
                            Clear-Host
                            Do{
                                Write-Host "`nPlease select the format for your exported search results:`n"
                                Write-Host "1. PerUserPst: One PST file for each mailbox."
                                Write-Host "2. SinglePst: One PST file that contains all exported messages."
                                Write-Host "3. SingleFolderPst: One PST file with a single root folder for the entire export."
                                Write-Host "4. IndividualMessage: Export each message as an .msg message file."
                                Write-Host "5. PerUserZip: One ZIP file for each mailbox. Each ZIP file contains the exported .msg message files from the mailbox."
                                Write-Host "6. SingleZip: One ZIP file for all mailboxes. The ZIP file contains all exported .msg message files from all mailboxes."
                                Write-Host "7. Cancel & return to main menu"
                                Write-Host "`nExporting the following Content Search case: " -NoNewLine
                                Write-Host "$ExportJobSearch" -ForegroundColor Green
                                $ExportFormatSelection = Read-Host "`n`nPlease enter your selection"
                                switch($ExportFormatSelection){
                                    1{$ExportFormat = "PerUserPst"}
                                    2{$ExportFormat = "SinglePst"}
                                    3{$ExportFormat = "SingleFolderPst"}
                                    4{$ExportFormat = "IndividualMessage"}
                                    5{$ExportFormat = "PerUserZip"}
                                    6{$ExportFormat = "SingleZip"}
                                    7{Break}
                                    Default{
                                        Clear-Host
                                        Write-Host "Invalid selection. Please try again." -BackgroundColor Yellow -ForegroundColor Red
                                    }
                                }
                            }Until($ExportFormatSelection -like "[1/2/3/4/5/6/7]")
                            If($ExportFormatSelection -eq "7"){Break}
                            Try{
                                New-ComplianceSearchAction -SearchName $ExportJobSearch -EnableDedupe $true -Export -ExchangeArchiveFormat $ExportFormat -Scope BothIndexedAndUnindexedItems -Force -ErrorAction Stop | Format-Table
                            }
                            Catch {
                                Write-Host "Unable to start export job successfully. Please try starting export job " -NoNewLine -BackgroundColor Red
                                Write-Host "$ExportJobSearch" -NoNewLine -BackgroundColor Red -ForegroundColor Yellow
                                Write-Host " manually from the M365 Compliance/Purview portal." -BackgroundColor Red
                                $ExportSearchSuccess = $False
                            }
                            If($ExportSearchSuccess){
                                Write-Host "Export job started for search " -NoNewLine
                                Write-Host "$ExportJobSearch" -NoNewLine -ForegroundColor Green 
                                Write-Host ". Please check on export job progress later today."
                            }
                        }
                        'N'{
                            Break
                        }
                        Default{
                            Clear-Host
                            Write-Host "Invalid selection. Please try again." -ForegroundColor Red -BackgroundColor Yellow
                        }
                    }
            }Until($SearchMenuSelection -like "[y|n]")
            Exit-Prompt "`nPress any key to return to the main menu."
        }
        3{
            Clear-Host
            Write-Host "List of Content Search Export Jobs`n" -ForegroundColor Yellow
            Get-ComplianceSearchAction | Format-Table
            Exit-Prompt "`nPress any key to return to the main menu."
        }
        4{Exit}
        default{
            Write-Host "`nInvalid selection; please try again.`n" -ForegroundColor Yellow -BackgroundColor Red
            Start-Sleep -Seconds 3
        }
    }
}