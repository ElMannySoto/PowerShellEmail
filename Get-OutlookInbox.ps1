<#
IMPORTANT TO DO: 
Open the file that contains the Get-OutlookInbox function and run it once inside the Windows PowerShell ISE.
This places the function onto the function drive and makes it available to me within my Windows PowerShell ISE session.

https://blogs.technet.microsoft.com/heyscriptingguy/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox/
https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Core/Where-Object?view=powershell-5.1
#>


Function Get-OutlookInBox
{


 Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

 $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]

 $outlook = new-object -comobject outlook.application

 $namespace = $outlook.GetNameSpace("MAPI")

 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

 $folder.items |

 Select-Object -Property Subject, ReceivedTime, Importance, SenderName

} 

<#

.Example: Stores Outlook InBox items into the $InBox variable for further "offline" processing.

    $InBox = Get-OutlookInbox
    
    $manny = ($InBox | where-object { $_.SenderName -match 'soto, manuel' } )
    $manny = ($InBox | where-object { $_.subject -match 'Azure' -AND $_.SenderName -match 'soto, manuel' } )
    $manny = ($InBox | where-object { $_.subject -like "*Azure" } )
    $manny = ($InBox | where-object { $_.subject -like "*Azure*" -AND $_.SenderName -like "*microsoft*" } )
    $manny = ($InBox | where-object { $_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com' } | Measure-object).count
    $manny = ($InBox | where-object { $_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com' } | Measure-object).count
    $manny = ($InBox | where-object { $_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com' } | Measure-object).count
    $manny = ($InBox | where-object { ($_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com') -and ($_.ReceivedTime -gt [datetime]"2/13/18" -AND $_.ReceivedTime -lt [datetime]"2/14/18") } | Measure-object).count
    $manny = ($InBox | where-object { ($_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com') -and ($_.ReceivedTime -gt [datetime]"11/21/17" -AND $_.ReceivedTime -lt [datetime]"2/14/18") } | Measure-object).count
    $manny = ($InBox | where-object { ($_.subject -match 'azure' -AND $_.SenderName -match 'manuel.soto@hpe.com') -and ($_.ReceivedTime -gt [datetime]"11/23/17" -AND $_.ReceivedTime -lt [datetime]"1/4/18") } | Measure-object).count




.Example: Displays Subject, ReceivedTime, Importance, SenderName for all InBox items that 
          are in InBox between 5/5/11 and 5/10/11 and sorts by importance of the email.


    Get-OutlookInbox | where-object { $_.ReceivedTime -gt [datetime]"5/5/11" -AND $_.ReceivedTime -lt [datetime]"5/10/11" } | sort importance


.Example: Displays Count, SenderName and grouping information for all InBox items. The most frequently used contacts appear at bottom of list.


    Get-OutlookInbox | Group-Object -Property SenderName | sort-Object Count

  
.Example: Displays the number of messages in InBox Items

    ($InBox | Measure-Object).count

    
.Example

    $InBox | where-object { $_.subject -match '2011 Scripting Games' } | sort ReceivedTime -Descending | select subject, ReceivedTime -last 5

    NOTE: Uses $InBox variable (previously created) and searches subject field
        for the string '2011 Scripting Games' it then sorts by the date InBox.
        This sort is descending which puts the oldest messages at bottom of list.
        The Select-Object cmdlet is then used to choose only the subject and ReceivedTime
        properties and then only the last five messages are displayed. These last
        five messages are the five oldest messages that meet the string.

 #Requires -Version 2.0

 #>