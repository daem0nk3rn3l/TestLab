
# create an object 
$OL = New-Object -ComObject outlook.application


Start-Sleep 5

<#
olAppointmentItem
olContactItem
olDistributionListItem
olJournalItem
olMailItem
olNoteItem
olPostItem
olTaskItem
#>



# Create Item
$mItem = $OL.CreateItem("olMailItem")


$mItem.To = "xyz000@gmail.com"
$mItem.Subject = "PowerMail"
$mItem.Body = "Sent from PowerShell"

$mItem.Send()



