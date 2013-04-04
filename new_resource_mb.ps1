<#

    .Description Create a new Resource Mailbox and Associated Delegation Security Group

    Please change the naming structure to suite your organizations needs.

#>

function New-Resource {
    param(
            [parameter(Mandatory=$true,HelpMessage="Display Name Used for Resource and Group.")][string]$DisplayName,
            [parameter(Mandatory=$true,HelpMessage="Email Address Base in this format: 'dept.name'")][string]$AddressBase,
            [parameter(Mandatory=$true,HelpMessage="Requester's Email Address.")][string]$OwnerEmail,
		    [parameter(ParameterSetName="Room",HelpMessage="Room Resource.")][switch]$Room,
            [parameter(ParameterSetName="Equipment",HelpMessage="Equipment Resource.")][switch]$Equipment
         )
    
    ##TODO Auto look up domain with optional parameter
    #$domain
    $domain = "business.com"

    $windowsLiveID = "res."+ $AddressBase +"@" + $domain
    Write-Output "WindowsLiveID: "  $windowsLiveID

    $groupLiveID = "grp."+ $AddressBase +"@" + $domain
    Write-Output "Group WindowsLiveID: "  $groupLiveID

    $groupAlias = "grp." + $AddressBase
    Write-Output "Group Alias: "  $groupAlias

    Write-Output "Display Name: "  $DisplayName
    $groupDisplayName = "Group " + $DisplayName
    Write-Output "Group Display Name: "  $groupDisplayName

    $name = $DisplayName.Replace(" ","")
    Write-Output "Name Parameter: " $name

    Read-Host "Is this correct?  If so, hit Enter to Process.  Else hit Ctrl-C or Stop to break the script."
	
	# random string
	$length = 15
	$numberOfNonAlphanumericCharacters = 6
	Add-Type -Assembly System.Web 
	$pw = [Web.Security.Membership]::GeneratePassword($length,$numberOfNonAlphanumericCharacters)
	
	# generate secure string  
	$password = ConvertTo-SecureString "$pw" -AsPlainText -Force
    
    New-Mailbox -Name $name -DisplayName $DisplayName -WindowsLiveID $windowsLiveID -Password $password
    if($Room){  Set-Mailbox -Identity $windowsLiveID -Type Room  }
    if($Equipment){  Set-Mailbox -Identity $windowsLiveID -Type Equipment  }

    New-DistributionGroup -Name $groupDisplayName -Alias $groupAlias -PrimarySmtpAddress $groupLiveID -CopyOwnerToMember -ManagedBy $OwnerEmail -Type Security -ModerationEnabled $false

    $i=0
    while($i -lt 10){
    try
    {

        Set-CalendarProcessing -Identity $windowsLiveID -ResourceDelegates $groupLiveID -BookInPolicy $groupLiveID -AutomateProcessing AutoAccept -ForwardRequestsToDelegates $true -TentativePendingApproval $true -AllowRecurringMeetings $true -AllowConflicts $true -AllBookInPolicy $false -AllRequestInPolicy $true -AllRequestOutOfPolicy $false  -DeleteSubject $false -AddOrganizerToSubject $false -ErrorAction "Stop"
        Write-Output "Calendar Processing Set"
        $i = 10;
    }
    catch
    {
        Start-Sleep -Milliseconds 15000
        Write-Output "Calendar Processing - 15 Second Sleep Timer Started - On $i th iteration of 10."
        $i++;
        $retry_calproc = "Set-CalendarProcessing -Identity $windowsLiveID -ResourceDelegates $groupLiveID -BookInPolicy $groupLiveID -AutomateProcessing AutoAccept -ForwardRequestsToDelegates $true -TentativePendingApproval $true -AllowRecurringMeetings $true -AllowConflicts $true -AllBookInPolicy $false -AllRequestInPolicy $true -AllRequestOutOfPolicy $false  -DeleteSubject $false -AddOrganizerToSubject $false";
    }
    }


    $i=0
    while($i -lt 10){
    try
    {
        Add-MailboxPermission -Identity $windowsLiveID -AccessRights FullAccess -User $groupLiveID -ErrorAction "Stop"
        Write-Output "FullAccess Granted."
        $i = 10;
    }
    catch
    {
        Start-Sleep -Milliseconds 15000
        Write-Output "Mailbox Permissions - 15 Second Sleep Timer Started - On $i th iteration of 10."
        $i++;
        $retry_mbperms = "Add-MailboxPermission -Identity $windowsLiveID -AccessRights FullAccess -User $groupLiveID";
    }
    }

    
    $i=0
    while($i -lt 10){
    try
    {
        $calendar = $windowsLiveID + ":\Calendar"
        Add-MailboxFolderPermission -Identity $calendar -AccessRights Reviewer -User $groupLiveID -ErrorAction "Stop"
        Write-Output "Reviewer Access Granted."
        $i=10
    }
    catch
    {
        Start-Sleep -Milliseconds 15000
        Write-Output "Calendar Permissions - 15 Second Sleep Timer Started - On $i th iteration of 10."
        $i++;
        $retry_calperms = "Add-MailboxFolderPermission -Identity $calendar -AccessRights Reviewer -User $groupLiveID";
    }
    }
    Write-Output "All done!  Review any errors."
    Write-Output "$retry_calproc"
    Write-Output "$retry_mbperms"
    Write-Output "$retry_calperms"
}