# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# outlook-simpleload.ps1
#
# Script to generate minimal load using Outlook client to allow evaluation of the Microsoft 365 informed network routing feature.
#
# This script assumes that the Outlook client is running and configured to connect to a Microsoft 365 (Exchange Online) account.
# Usage of an isolated test account is recommended, as the script does delete messages.
#

$olFolderInbox = 6
$olMailItem = 0

$subject = "M365 INR Test Message"
$subjectFilter = '[Subject] = '''+ $subject + ''''

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"                              
Add-Type -AssemblyName "System.Runtime.InteropServices"                                

$outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')   
$namespace = $outlook.GetNamespace("MAPI")                                             

$inbox = $outlook.Session.GetDefaultFolder($olFolderInbox)

while ($true)
{
    $table = $inbox.GetTable($subjectFilter)
    while (-not $table.EndOfTable)
    {
        $row = $table.GetNextRow()
        $item = $namespace.GetItemFromID($row["EntryID"])
        Write-Host "$(Get-Date): Found received item: '$($item.Subject)' received '$($item.ReceivedTime)' - deleting item..."
        $item.Delete() | Out-Null
    }
    
    Write-Host "$(Get-Date): Sending test message to $($outlook.Session.CurrentUser.Name)"
    $message = $outlook.CreateItem($olMailItem)
    $message.Subject = $subject
    $message.To = $outlook.Session.CurrentUser.Address
    $message.Recipients.ResolveAll() | Out-Null
    $message.Body = "This is a simple message to trigger network activity while evaluating the Microsoft 365 informed network routing feature."
    $message.Send() | Out-Null

    Start-Sleep -Seconds 60
}
