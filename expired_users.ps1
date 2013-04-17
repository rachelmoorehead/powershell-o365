<#

    Generate Account Disablement Listing
    
#>

$log = "C:\PowerShell\account_disable_log.txt"
if(Test-Path $log){  Remove-Item $log }
$results = "C:\PowerShell\account_disable_results.txt"
if(Test-Path $results){  Remove-Item $results }

<### FUNCTIONS ###>

Function Check-Expiration {

    $check_users = $args[0]
    $check_date = $args[1]

    $number = $check_users.Length
    Write-Output "Checking $number users against logon time before $check_date" >> $log

    $e_users = @()

    $check_users | foreach{
    
    
        try{
            
            Write-Output "Checking Mailbox Statistics for $_" >> $log

            $logon_data = Get-MailboxStatistics -Identity $_.UserPrincipalName | Select-Object DisplayName, LastLogonTime, LastLogoffTime, TotalItemSize, LastLoggedOnUserAccount
            
            Write-Output $logon_data >> $log
            
            $user = $logon_data | Where-Object {(($_.LastLogonTime -lt $expired_date) -or ($_.LastLogonTime -eq $null)) -and (($_.LastLogoffTime -lt $expired_date) -or ($_.LastLogoffTime -eq $null))} 

            $v1 = $user.DisplayName
            if($v1.Length -gt '0') {Write-Output "Meets expiration criteria: $v1"  >> $log}
            
                if($v1.Length -gt '0'){
                    $user_data = @{
            
                       'User' = $_.UserPrincipalName;
                       'WhenCreated' = $_.WhenCreated;
                       'LastLogonTime' = $user.LastLogonTime;
                       'LastLogoffTime' = $user.LastLogoffTime;
                       'TotalItemSize' = $user.TotalItemSize;
                       'LastLoggedOnUserAccount' = $user.LastLoggedOnUserAccount;   
                     }      

                     $temp = New-Object PSObject -Property $user_data
                     $e_users += $temp   

                     $v2 = $e_users.Length
                     Write-Output "Number of Expired Users: $v2" >> $log
                }
                else{
                
                    $var = $_.UserPrincipalName
                    Write-Output "Skipped, does not match expiration criteria: $var" >> $log
                
                }

        }
        catch{
        
            $temp = $_.UserPrincipalName
            Write-Output "Error with $temp. Skipping."  >> $log
        
        }

    }
   
    Return $e_users

}

Function Check-Rules {

    $users = $args[0]

    $c_users = @()

    foreach ($u in $users)
    {
        $v3 = $u.User
        Write-Output "Checking Inbox Rules for: $v3" >> $log

        $temp = Get-InboxRule -Mailbox $v3 | where {(($_.Enabled -eq $true) -AND (($_.RedirectTo -ne $null) -OR ($_.ForwardTo -ne $null)))}

        if($temp.Length -eq '0'){  Write-Output "No InboxRules found for: $v3" >> $log  }

        if($temp.Length -eq '0'){
        
            $c_users += $u

        }
        else{
       
             Write-Output "Skipped, Has Inbox Rules Forwarding or Redirecting: $v3" >> $log
        
        }
    }

    Return $c_users
}

<### MAIN ROUTINE ###>

$date = Get-Date
Write-Host "Start: $date"  >> $log

$expired_date = $date.AddDays(-365)

$users = Get-User -Filter "(RecipientType -eq 'UserMailbox') -and (WhenCreated -gt '$expired_date')" -ResultSize unlimited
#$users

$expired_users = Check-Expiration $users $expired_date

Write-Host "Expired Users Returned"
#$expired_users 

$clean_users = Check-Rules $expired_users

Write-Host "Clean Users Returned"
#TODO Handle List Output for Action
$clean_users |Format-Table | Out-File $results

$date = Get-Date
Write-Host "Completed: $date" >> $log
