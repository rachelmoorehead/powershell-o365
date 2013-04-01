<#

    .Description Enable Advanced Features: Audit Logging and Single Item Recovery

#>


function Enable-Logging {
    param([string]$EmailAddress)
    try{
    
        Set-Mailbox $EmailAddress -AuditEnabled $true -AuditLogAgeLimit 365.00:00:00 -AuditAdmin Update,Copy,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf,MessageBind -AuditDelegate Update,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf -SingleItemRecoveryEnabled $true 
        Write-Output "Successfully enabled logging and item recovery."
    }
    catch {
        Write-Output "Could not enable logging and item recovery."
    }

}