<#
IMPORTANT TO DO: 
Open the file that contains the Get-OutlookInbox function and run it once inside the Windows PowerShell ISE.
This places the function onto the function drive and makes it available to me within my Windows PowerShell ISE session.

https://blogs.technet.microsoft.com/heyscriptingguy/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox/
https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Core/Where-Object?view=powershell-5.1
#>
Function Send-Email
{
  Param($arg1_Subject, $arg2_Body)

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = "eemannysoto@gmail.com; elmannysoto@gmail.com"
	$Mail.Subject = $arg1_Subject
    $Mail.Body = $arg2_Body
    $Mail.Send()
}


$Subject_MS = "This is automated"
$Body_MS = "I'm Henry he 8th I am, Henry the 8th I am I am. I got married to the widow next door, she'd been married seven times before."

Send-Email -arg1_Subject $Subject_MS -arg2_Body $Body_MS


