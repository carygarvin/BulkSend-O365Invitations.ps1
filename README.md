# BulkSend-O365Invitations.ps1

Author       : Cary GARVIN  
Contact      : cary(at)garvin.tech  
LinkedIn     : [https://www.linkedin.com/in/cary-garvin](https://www.linkedin.com/in/cary-garvin)  
GitHub       : [https://github.com/carygarvin/](https://github.com/carygarvin/)  


Script Name  : [BulkSend-O365Invitations.ps1](https://github.com/carygarvin/BulkSend-O365Invitations.ps1/)  
Script URL   : [https://carygarvin.github.io/BulkSend-O365Invitations.ps1/](https://carygarvin.github.io/BulkSend-O365Invitations.ps1)  
Version      : 1.0  
Release date : 14/06/2019 (CET)  
History      : The present script has been developed to send mass invitations for the Microsoft Teams product b.  
Purpose      : 

# Script history
This Script was used to send invitations to a SharePoint Online site with an hypothetical URL to a list of external Guests compiled in an Excel file.  

# Script usage
There are 2 user configurable parameters, one with the Admin account under which the script is to be executed and the other one being the actual URL of the SharePoint Online site to attach in the invitation.  
The present Script relies on 2 input files: an Excel file containing the list of external Guests to invite to Teams and a text file names 'invitation.txt' containing the body of the message to send to these potential Teams guests.  
Both files are to be present in the same directory as the present Script.  
When executed, the script will update the Excel file with the status on the invitations for each Guest listed, either "Sent" or "Error" in the 8th column (H) titled 'Invitation sent'.  
Obviously, the script relies on Microsoft Excel to be installed in order to run  using COM Objects as it reads the input Excel file (see sample in repository) containing the names of Guests to which invitations need to be sent.  
Column F, G and I (columns 6, 7 and 9) can be used at the user's discretion as they are unused by the present Script.  
