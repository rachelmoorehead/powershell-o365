Function Show-Total {
<#
.SYNOPSIS
Function to list number of items in folders
.DESCRIPTION
Show-Total uses Get-MailboxfolderStatistics to retrieve information about the total number of items in folders in Mailbox.  Developed by Lewis Noles 20120307, and updated by Lewis Noles on 20120628.
.PARAMETER MOSID
Customer's Microsoft Online Services ID
.EXAMPLE
 Show-Total MOSID | Format-List | Out-File C:\Temp\MOSID.txt
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Customer's Microsoft Online Services ID")]
        [string]$MOSID
    )
Get-Mailbox $MOSID | Get-MailboxFolderStatistics | Select-Object Identity, ItemsInFolder, FolderSize, FolderType | Sort-Object -Property ItemsInFolder -Descending
}
