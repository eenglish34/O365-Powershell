# POWERSHELL MENU v1.0
# Menu system using functions
#
# Purpose: This script is used to log into O365 with an admin account, and quarantine an email from one, multiple, or all mailboxes. 
# Author: Eric English
#
 
Function Menu ($MenuStart, $MenuEnd, $MenuName, $MenuFunctionName, $MenuOption1, $MenuFunction1, $MenuOption2, $MenuFunction2, $MenuOption3, $MenuFunction3, $MenuOption4, $MenuFunction4, $MenuOption5, $MenuFunction5, $MenuOption6, $MenuFunction6, $MenuOption7, $MenuFunction7, $MenuOption8, $MenuFunction8, $MenuOption9, $MenuFunction9, $MenuOption10, $MenuFunction10)
{
        Write-Host “`n$MenuName`n” -Fore Gray;
        if($MenuOption1 -ne $null){Write-Host “`t`t$MenuOption1” -Fore White;}
        if($MenuOption2 -ne $null){Write-Host “`t`t$MenuOption2” -Fore White;}
        if($MenuOption3 -ne $null){Write-Host “`t`t$MenuOption3” -Fore White;}
        if($MenuOption4 -ne $null){Write-Host “`t`t$MenuOption4” -Fore White;}
        if($MenuOption5 -ne $null){Write-Host “`t`t$MenuOption5” -Fore White;}
        if($MenuOption6 -ne $null){Write-Host “`t`t$MenuOption6” -Fore Green;}
        if($MenuOption7 -ne $null){Write-Host “`t`t$MenuOption7” -Fore Green;}
        if($MenuOption8 -ne $null){Write-Host “`t`t$MenuOption8” -Fore Green;}
        if($MenuOption9 -ne $null){Write-Host “`t`t$MenuOption9” -Fore Green;}
        if($MenuOption10 -ne $null){Write-Host “`t`t$MenuOption9” -Fore Green;}
       [int]$MenuOption = Read-Host “`n`t`tPlease select an option”
            
            if(($MenuOption -lt $MenuStart) -or ($MenuOption -gt $MenuEnd))
            {
            Write-Host “`t`nPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
            Invoke-Expression $MenuFunctionName
            }
            else
            {
                if($MenuOption -eq $MenuStart) {Invoke-Expression $MenuFunction1}
                if($MenuOption -eq ($MenuStart + "1")) {Invoke-Expression $MenuFunction2}
                if($MenuOption -eq ($MenuStart + "2")) {Invoke-Expression $MenuFunction3}   
                if($MenuOption -eq ($MenuStart + "3")) {Invoke-Expression $MenuFunction4} 
                if($MenuOption -eq ($MenuStart + "4")) {Invoke-Expression $MenuFunction5} 
                if($MenuOption -eq ($MenuStart + "5")) {Invoke-Expression $MenuFunction6} 
                if($MenuOption -eq ($MenuStart + "6")) {Invoke-Expression $MenuFunction7} 
                if($MenuOption -eq ($MenuStart + "7")) {Invoke-Expression $MenuFunction8}
                if($MenuOption -eq ($MenuStart + "8")) {Invoke-Expression $MenuFunction9} 
                if($MenuOption -eq ($MenuStart + "9")) {Invoke-Expression $MenuFunction10}
            }    
}
#MENU OPTIONS: Space seperated list - 
#Menu option start number, Menu option end number, Menus descriptive name, Menus function name,
#Menu option 1, Option 1 action (specify exisiting menu or other function), option 2, option 2 action and so on up to 10 options.
#
#E.G 
# function MENU_HOME #Wrap menu in function so that you can call it from anywhere to get back
# {
# Menu 1 3 "Menu Home" MENU_HOME "1. Go to menu home one" MENU_HOME_ONE "2. Go to menu home two" MENU_HOME_TWO "3. Do stuff" MENU_HOME_DOSTUFF #Invokes menu builder function and passes variables
# } 
# MENU_HOME #Runs function
#
#Remember that powershell is processed line by line so functions need to be above where they are called. This gets trickier and trickier the more complicated your menu gets :)
#
 
 
###MENU_HOME
###############################################################################################################################################################################################################################
 
function MENU_HOME
{
Menu 1 5 "###########################################################################################################
######################################## Email Removal Menu ###############################################
###########################################################################################################
############################## Remove Malicious Emails from O365 Mailboxes ################################
###########################################################################################################
####################################### Author: Eric English ##############################################
###########################################################################################################
!!!!!!!!Make sure you run this script as admin! If needed close and reopen this menu script as admin!!!!!!!
###########################################################################################################
" MENU_HOME "1. Sign into O365 with admin account" MENU_HOME_ONE "2. Export CSV of Emails" MENU_HOME_TWO "3. Remove email from single mailbox" MENU_HOME_THREE "4. Remove email from multiple mailboxes" MENU_HOME_FOUR "5. Remove email from all mailboxes" MENU_HOME_FIVE
}   
 
###MENU_HOME_ONE
###############################################################################################################################################################################################################################
 
function MENU_HOME_ONE
{
Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection 
Import-PSSession $Session
Write-Host “`t`tDone Logging into O365” -Fore Green;
pause
MENU_HOME
}   
 
###MENU_HOME_TWO
###############################################################################################################################################################################################################################
 
function MENU_HOME_TWO
{
$email = Read-Host "Please enter the full email address of the sender or from address"
$dateStart = Read-Host "Please enter the start date of your search in this format mm/dd/yyyy"
$dateEnd = Read-Host "Please enter the end date of your search in this format mm/dd/yyyy"
Get-MessageTrace -SenderAddress $email -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID | Export-Csv -Path C:\Users_Affected.csv -Encoding ascii -NoTypeInformation
Write-Host “`t`tDone exporting to CSV. You can find the file under c:\Users_Affected.csv” -Fore Green;
pause
MENU_HOME
}
 
###MENU_HOME_THREE
###############################################################################################################################################################################################################################
 
function MENU_HOME_THREE
{
$email = Read-Host "Please enter the full email address of the sender or from address"
$email2 = Read-Host "Please enter full email address of the users mailbox to search"
$subject = Read-Host "Please enter the subject of the email"
$date = Read-Host "Please enter the date the email was sent in this format mm/dd/yyyy"
Search-Mailbox -Identity "$email2" -SearchQuery 'From:"$email"','Subject:"$subject"','Sent:"$date"' -DeleteContent -TargetMailbox "yourtargetmailbox@example.com" -TargetFolder "Junk Email" -loglevel full
Write-Host “`t`tDone removing the email from a single mailbox” -Fore Green;
pause
MENU_HOME
}  
###MENU_HOME_FOUR
###############################################################################################################################################################################################################################
 
function MENU_HOME_FOUR
{
$email = Read-Host "Please enter the full email address of the sender or from address"
$subject = Read-Host "Please enter the subject of the email"
$date = Read-Host "Please enter the date the email was sent in this format mm/dd/yyyy"
Write-Host "Before running this command make sure to put your text file with the email addresses, one on each line, in c:\ and name it Users_Affected.txt"
pause
Get-Content c:\Users_Affected.txt | Get-Mailbox -ResultSize unlimited | Search-Mailbox -SearchQuery '"From:$email"','"Subject:$subject"','"Sent:$date"' -DeleteContent -TargetMailbox "yourtargetmailbox@example.com" -TargetFolder "Junk Email" -loglevel full
Write-Host “`t`tDone removing the email from the list provided.” -Fore Green;
pause
MENU_HOME
} 
###MENU_HOME_FOUR
###############################################################################################################################################################################################################################
 
function MENU_HOME_FIVE
{
$email = Read-Host "Please enter the full email address of the sender or from address"
$subject = Read-Host "Please enter the subject of the email"
$date = Read-Host "Please enter the date the email was sent in this format mm/dd/yyyy"
Get-Mailbox -ResultSize unlimited | Search-Mailbox -SearchQuery '"Subject:$subject"','"From:$email"','"Sent:$date"' -DeleteContent -TargetMailbox "yourtargetmailbox@example.com" -TargetFolder "Junk Email" -loglevel full
Write-Host “`t`tDone removing the email from all mailboxes.” -Fore Green;
pause
MENU_HOME
} 
MENU_HOME
