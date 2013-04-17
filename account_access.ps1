<#

    .Description 
    
    Functions to Lock and Unlock Accounts in Exchange Online to be used in coordination with regular account disablements.
    Functions to Impersonate and Un-impersonate a user.

    Lock-User
    Unlock-User

    Add-Impersonation
    Remove-Impersonation

#>

function Lock-User {

    [cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    param(
            [parameter(Mandatory=$true,HelpMessage="Email address or UPN of user to be locked out.")]
            [string]$EmailAddress
         )
    
    Process
    {
        if ($PSCmdlet.ShouldProcess($EmailAddress, "Lockout User"))
        {
            Set-CasMailbox -Identity $EmailAddress -MAPIEnabled $false -POPEnabled $false -IMAPEnabled $false -ActiveSyncEnabled $false -OWAEnabled $false
            Write-Output "Mailbox Access Protocols disabled. [MAPI, POP, IMAP, ActiveSync, OWA]"
            
            Get-InboxRule -Mailbox $EmailAddress | Disable-InboxRule
            Write-Output "All Inbox Rules disabled."
            
            Set-Mailbox $EmailAddress -ForwardingAddress $null -ForwardingSMTPAddress $null
            Write-Output "Forwarding via Connected Accounts is disabled."
            
            Get-ActiveSyncDevice -Mailbox $EmailAddress | Remove-ActiveSyncDevice
            Write-Output "Removed All ActiveSync devices to kill active sessions."

            # Generate Dummy Password
	        $length = 15
	        $numberOfNonAlphanumericCharacters = 6
	        Add-Type -Assembly System.Web 
	        $pw = [Web.Security.Membership]::GeneratePassword($length,$numberOfNonAlphanumericCharacters)

	        $password = ConvertTo-SecureString "$pw" -AsPlainText -Force
            Set-Mailbox $EmailAddress -Password $password
            Write-Output "UGAMail password reset."

        }
    }


 
}

function Unlock-User {
 param(
        [parameter(Mandatory=$true,HelpMessage="Email address or UPN of user to be re-enabled.")]
        [string]$EmailAddress
      )

          Set-CasMailbox -Identity $EmailAddress -MAPIEnabled $true -POPEnabled $true -IMAPEnabled $true -ActiveSyncEnabled $true -OWAEnabled $true
          Write-Output "Mailbox Access Protocols enabled. [MAPI, POP, IMAP, ActiveSync, OWA]."
          Get-InboxRule -Mailbox $EmailAddress | Enable-InboxRule
          Write-Output "All Inbox Rules enabled."
          Write-Output "Perform a Password Re-Sync, if necessary. User will need to resetup ActiveSync devices and Connected Accounts, if necessary."
}

function Add-Impersonation{
    param(
            [parameter(Mandatory=$true,HelpMessage="Email address or UPN of user to be impersonated.")][string]$EmailAddress,
            [parameter(Mandatory=$true,HelpMessage="Email address or UPN of user impersonating, if not yourself.")][string]$DelegateAddress
    )

        try{

            Add-MailboxPermission -AccessRights FullAccess -Identity $EmailAddress -User $DelegateAddress -Confirm
            [Environment]::SetEnvironmentVariable("O365Impersonate","$EmailAddress,$DelegateAddress","User")
        }

        catch{
        
            Write-Host "Sorry there was an error.  Please try again."
        }
        
}

function Remove-Impersonation{
    param(
            [parameter(ParameterSetName="Supplied",HelpMessage="Email address or UPN of mailbox")][string]$EmailAddress,
            [parameter(ParameterSetName="Supplied",HelpMessage="Email address or UPN of delegate")][string]$DelegateAddress
        )

        if(($EmailAddress.Length -eq '0') -and ($DelegateAddress.Length -eq '0')){
        
            try{
                $checkEnvVariable = [Environment]::GetEnvironmentVariable("O365Impersonate","User")
                $temp = $checkEnvVariable.Split(',')
                $EmailAddress = $temp[0]
                $DelegateAddress = $temp[1]
                [Environment]::GetEnvironmentVariable("O365Impersonate",$null,"User")
            
                Remove-MailboxPermission -AccessRights "FullAccess" -User $DelegateAddress -Identity $EmailAddress -Confirm
            }
            catch{
            
                Write-Host "Sorry there was an error.  Please try again."
            }
        }
        else{
        
            try{
                            
                Remove-MailboxPermission -AccessRights "FullAccess" -User $DelegateAddress -Identity $EmailAddress -Confirm
            }
            catch{
            
                Write-Host "Sorry there was an error.  Please try again."
            }
        }
}