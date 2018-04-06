This script can only be run as administrator in the Microsoft Azure Active Directory Module for Windows Powershell. You can download a copy of it from here: https://azure.microsoft.com/en-us/downloads/

Before running the script, please edit options 3-5 and change the example email address at the end of the line. I have mine going to an IT Security mailbox in the Junk Email folder. This is up to you where you want the quarantined emails to go to. 

This script gives you 3 options for removing an email from one, multiple, or all mailboxes in O365. 

The first option you have to run is option 1. This logs you into O365 and saves your credentials for which ever option you select next. 

Option 3 lets you remove an email from one mailbox. It will prompt you for some information about what to look for and will remove based on your input. 

Option 4 is used to remove an email from specific mailboxes. First you need to run option 2 and get a CSV export of the emails that were sent by that user. From within the CSV, sort and copy the email addresses for people that received the message in your organization. Copy those email addresses into a text file and put it on the root of your C drive and name it Users_Affected.txt. Youc an also change this in the script to suit your needs. After you've got your text file created and saved, then you can run option 4. 

Option 5 is used to remove an email from all mailboxes in your O365 environment. 

I am not responsible for any mistakes you make and there is no warranty or guarantee this will work for you and your environment. All responsibility is with the end user of this script. 
