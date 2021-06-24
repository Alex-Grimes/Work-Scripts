
Add-Type -AssemblyName Microsoft.VisualBasic
$userID = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter your ADM username")
$GroupName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the AD group to get membership")
Get-ADGroupmember -identity "$GroupName" | select name | Export-Csv -path C:\Users\$userID\Desktop\Results.csv
$FileLocation = [Microsoft.VisualBasic.Interaction]::MsgBox("C:Users\$userID\Desktop\Results.csv ")
