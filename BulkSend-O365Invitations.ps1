# ***************************************************************************************************
# ***************************************************************************************************
#
#  Author       : Cary GARVIN
#  Contact      : cary(at)garvin.tech
#  LinkedIn     : https://www.linkedin.com/in/cary-garvin
#  GitHub       : https://github.com/carygarvin/
#
#
#  Script Name  : BulkSend-O365Invitations.ps1
#  Script URL   : https://carygarvin.github.io/BulkSend-O365Invitations.ps1/
#  Version      : 1.0
#  Release date : 14/06/2019 (CET)
#  History      : The present script has been developed to send mass invitations for the Microsoft Teams product based on a list compiled in an Excel file.
#  Purpose      : The present Script relies on 2 input files: an Excel file containing the list of external Guests to invite to Teams and a text file names 'invitation.txt' containing the body of the message to send to these potential Teams guests.
#                 Both files are to be present in the same directory as the present Script.
#                 When executed, the script will update the Excel file with the status on the invitations for each Guest listed, either "Sent" or "Error" in the 8th column (H) titled 'Invitation sent'.
#
#
#  This Script was use to send invitations to a SharePoint Online site and thus contains an hypothetical URL to the SharePoint site in question.
#  There are 2 user configurable parameters, one with the Admin account under which the script is to be executed and the other one being the actual URL of the SharePoint Online site to attach in the invitation. 
#  Obviously, the script relies on Microsoft Excel to be installed in order to run  using COM Objects as it reads the input Excel file containing the names of Guests to which invitations need to be sent.
#  Column F, G and I (columns 6, 7 and 9) can be used at the user's discretion as they are unused by the present Script.
#
#



####################################################################################################
#                                     User configurable parameter                                  #
####################################################################################################

$AzureAdmin = "superman@contoso.com"
$messageURL = "https://contoso.sharepoint.com/sites/ContosoHQ"




####################################################################################################
#                                          Script Main                                             #
####################################################################################################


$ExpectedExcelHeaderRow = @("Last Name","First Name","Company","Title","E-mail Address")


$ScriptDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)


If (test-path "Invitation.txt")
	{
	$InvitationMessageBody = Get-Content "Invitation.txt"
	write-host "The 'Invitation.txt' file currently containing the following text:`r`n" -foregroundcolor "gray"
	$InvitationMessageBody | ForEach {Write-Output $_}
	write-host "`r`nIs the above contents of the 'Invitation.txt' file OK to send in the script's execution - [yes] or [no]? " -NoNewLine -foregroundcolor "gray"
	$MessageToSendIsOK = read-host
	While("yes","no" -notcontains $MessageToSendIsOK)
		{
		write-host "`r`nIs the above contents of the 'Invitation.txt' file OK to send in the script's execution - [yes] or [no]? " -NoNewLine -foregroundcolor "gray"
		$MessageToSendIsOK = read-host
		}
	If ($MessageToSendIsOK -eq "no")
		{
		write-host "Exiting script to allow the 'Invitation.txt' to be amended."
		Break
		}
	}
Else
	{
	write-host "The 'Invitation.txt' file containing the body of the message to send could not be found in the script's '$ScriptDir' directory." -foregroundcolor "red"
	write-host "Create this 'Invitation.txt' file and relaunch the script. The script will now exit." -foregroundcolor "red"
	Break
	}

<#
$Cred = Get-Credential -Message "Please enter the Administrator password to use to connect to AzureAD" -UserName $AzureAdmin
Connect-AzureAD -Credential $Cred | out-null
Connect-MsolService -Credential $Cred | out-null



$messageInfo = New-Object Microsoft.Open.MSGraph.Model.InvitedUserMessageInfo
$messageInfo.customizedMessageBody = $InvitationMessageBody
#>


write-host "`r`nList of Excel files found in script's '$ScriptDir' directory:`r`n" -foregroundcolor "white"
$ExcelFiles = Get-ChildItem | Where-Object {$_.Extension -like ".xls*"}
For ($i=0;$i-le $ExcelFiles.length-1;$i++) {"`tExcel file #{0} -> {1} [Last modified: {2}]" -f $($i+1),$ExcelFiles[$i],$ExcelFiles[$i].LastWriteTime}

$ExcelFileNrToOpen = read-host -Prompt "`r`nSpecify the number of the Excel file containing the users to which the inviations need to be sent"
$ExcelFileToOpen =  $($ExcelFiles[$ExcelFileNrToOpen-1])



$xlCellTypeLastCell = 11 
$FirstEntry = 2
$ColumnWithLastName = 1
$ColumnWithFirstName = 2
$ColumnWithCompany = 3
$ColumnWithJobTitle = 4
$ColumnWithSMTPAddress = 5

$ColumnWithInvitationResult = 8


$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$WorkBook = $Excel.Workbooks.Open("$ScriptDir\$ExcelFileToOpen")
$WorkSheet = $WorkBook.Sheets.Item(1)


$EncounteredExcelHeaderRow = $WorkSheet.Range("A1","E1").Value2
If (!(Compare-Object $EncounteredExcelHeaderRow $ExpectedExcelHeaderRow))
	{
	$LastEntry = $WorkSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row
	Write-host "$($LastEntry - 1) inviations to process:" -foregroundcolor "white"
	
	For ($i = $FirstEntry; $i -le $LastEntry; $i++)
		{
		$LastName = $WorkSheet.Cells.Item($i, $ColumnWithLastName).Text.Trim()
		$FirstName = $WorkSheet.Cells.Item($i, $ColumnWithFirstName).Text.Trim()
		$Company = $WorkSheet.Cells.Item($i, $ColumnWithCompany).Text.Trim()
		$JobTitle = $WorkSheet.Cells.Item($i, $ColumnWithJobTitle).Text.Trim()
		$SMTPAddress = $WorkSheet.Cells.Item($i, $ColumnWithSMTPAddress).Text.Trim()

		$CurrentUserInfo = "Sending Azure Teams invite to '$LastName $FirstName' [$SMTPAddress]".PadRight(105,[char]32)
		Write-host $CurrentUserInfo -NoNewLine -foregroundcolor "yellow"

		If ($SMTPAddress)
			{
			Try
				{
				$Invitation = New-AzureADMSInvitation -InvitedUserDisplayName "$LastName $FirstName" -InvitedUserEmailAddress $SMTPAddress -InviteRedirectUrl $messageURL -SendInvitationMessage $TRUE -InvitedUserMessageInfo $messageInfo 
				start-sleep -m 200

				$NewGuest = Get-AzureADUser -ObjectId $Invitation.InvitedUser.Id

				Set-AzureADUser -ObjectId $Invitation.InvitedUser.Id -Surname $FirstName -GivenName $LastName -JobTitle $JobTitle -Department $Company -DisplayName $NewGuest.DisplayName
				write-host "Sent!" -foregroundcolor "green"
				$WorkSheet.Cells.Item($i, $ColumnWithInvitationResult) = "Sent"
				}
			Catch
				{
				write-host "Failed!" -foregroundcolor "red"
				$WorkSheet.Cells.Item($i, $ColumnWithInvitationResult) = "Error"
				}
			}
		}

	$WorkBook.Save()
	start-sleep -s 2
	}
Else
	{
	write-host "The header row in Excel file '$ScriptDir\$ExcelFileToOpen' doesn't mach expected header row format!" -foregroundcolor "red"
	write-host "Encountered header row: " -NoNewLine
	Write-host """$($EncounteredExcelHeaderRow -join '","')""" -foregroundcolor "magenta"
	write-host "Expected header row   : " -NoNewLine
	Write-host """$($ExpectedExcelHeaderRow -join '","')""" -foregroundcolor "green"
	Write-host "The scipt will now abort..." -foregroundcolor "red"
	}

$Excel.Workbooks.Close()
$Excel.Quit()


