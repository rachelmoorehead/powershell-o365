<#

    Script to Determine Userbase Application Profile

    Code:

    Program      - Client=XXX; UserAgent= XXX; ExchangeWebServices/Action

    Outlook 2011 - Client=WebServices;UserAgent=MacOutlook/14.3.2.130206 (Intel Mac OS X 10.8.3)
    Thunderbird? - Client=WebServices;UserAgent=ExchangeWebServicesProxy/CrossSite/EXCH/14.16.0287.008/Mozilla/4.0 
    MacMail      - Client=WebServices;UserAgent=ExchangeWebServicesProxy/CrossSite/EXCH/14.15.0129.007/Mac OS X/10.6.8
                   Client=WebServices;UserAgent=Mac OS X/10.8.2 (12C2034); ExchangeWebServices/3.0 (157); Mail/6.2 (1499) 
                   Client=WebServices;UserAgent=Mac OS X/10.8.3 (12D78); ExchangeWebServices/3.0 (157); Mail/6.3 (1503)
    OWA          - Client=WebServices;UserAgent=OwaProxy
    Outlook MAPI - Client=WebServices;UserAgent=Microsoft Office/14.0 (Windows NT 6.1; Microsoft Outlook 14.0.6129; Pro)

    iPhone 3     - Client=ActiveSync;UserAgent=Apple-iPhone3C3/1002.329;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Apple-iPhone3C1/1001.523;Action=/Microsoft-Server-ActiveSync/default.eas
    iPhone 4     - Client=ActiveSync;UserAgent=Apple-iPhone4C1/1002.142;Action=/Microsoft-Server-ActiveSync/default.eas
    iPhone 5     - Client=ActiveSync;UserAgent=Apple-iPhone5C1/1002.143;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Apple-iPhone5C2/1002.329;Action=/Microsoft-Server-ActiveSync/default.eas
    iPad 1       - Client=ActiveSync;UserAgent=Apple-iPad1C1/902.206;Action=/Microsoft-Server-ActiveSync/default.eas
    iPad 2       - Client=ActiveSync;UserAgent=Apple-iPad2C1/1002.146;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Apple-iPad2C2/1002.329;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Apple-iPad2C5/1002.329;Action=/Microsoft-Server-ActiveSync/default.eas
    Android      - Client=ActiveSync;UserAgent=Android/4.2.1-EAS-1.3;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Android/4.2.2-EAS-1.3;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=Android-EAS/0.1;Action=/Microsoft-Server-ActiveSync/default.eas
    TouchDown    - Client=ActiveSync;UserAgent=TouchDown(MSRPC)/8.1.00020/;Action=/Microsoft-Server-ActiveSync/default.eas
                   Client=ActiveSync;UserAgent=TouchDown(MSRPC)/7.2.00016/;Action=/Microsoft-Server-ActiveSync/default.eas
    DroidRazr    - Client=ActiveSync;UserAgent=motorola-DROIDRAZR/1.0;Action=/Microsoft-Server-ActiveSync/default.eas
    Samsung      - Client=ActiveSync;UserAgent=SAMSUNGSCHS720C/2.3.6-EAS-1.2;Action=/Microsoft-Server-ActiveSync/default.eas

    Outlook RPC  - Client=MSExchangeRPC

    Passive DAG  - Client=CI (Content Indexing)

    OWA          - Client=OWA;Action=ViaProxy

    IMAP         - Client=IMAP4
    

    COW=Delegate
#>

Function Find-ConnectedApplications {

    # Results
    $results = "C:\PowerShell\results.csv"

    # Log
    $log = "C:\PowerShell\log.csv"

    # Get all users
    #Check for Session
            $status = (Get-PSSession).State
            if($status -eq 'Opened'){
                $users = Get-User -Filter "RecipientType -eq 'UserMailbox'" -ResultSize unlimited
            }
            else {
            
                Write-Host "The Remote Session either has been closed or no longer exists.  Please re-connect."
                Start-Session -LiveAtEdu

                $users = Get-User -Filter "RecipientType -eq 'UserMailbox'" -ResultSize unlimited
            
            }

    # Get all users ApplicationIds
    $users | foreach { 

        $user = $_.Name; 
        $windowsliveid = $_.WindowsLiveID

        try{
    
                $status = (Get-PSSession).State

                if($status -eq 'Opened'){
                    
                    # Array of Sessions
                    $li_stats = Get-LogonStatistics $windowsliveid ;
                    Write-Output "Processing User: $user, $windowsliveid"
                }
                else {
            
                        Write-Host "The Remote Session either has been closed or no longer exists.  Please re-connect."
                        Start-Session -LiveAtEdu

                        # Array of Sessions
                        $li_stats = Get-LogonStatistics $windowsliveid ;
                        Write-Output "Processing User: $user, $windowsliveid"
                }

           
        }
        catch [ParameterArgumentValidationErrorNullNotAllowed] {
            

            Write-Output "Parameter Binding Exception: $user, $windowsliveid" | Out-File -FilePath $log -Append 

        }
        catch {
        
            Write-Output "Something has occurred with $user." | Out-File -FilePath $log -Append
        
        }

        # Get number of clients connected
        $length = $li_stats.Length;
        Write-Host "Length for li_stats: $length"

        # Check for no clients
        if( -not ($length -eq '0')){
        
            Write-Host "Processing Clients for $user"
            $length = $li_stats.Length;
            Write-Host "Length for li_stats: $length"
            Process-Clients $user $li_stats | Format-Table -HideTableHeaders | Out-File -FilePath $results -Append 

        }
        
    }
}
        
Function Process-Clients {

    $user = $args[0]
    $li_stats = $args[1]

    # Get number of clients connected
    $length = $li_stats.Length;
    Write-Host "Length for li_stats: $length"

    #loop clients
            do{
       
                #advance counter, becomes index 
                $length--;
                #Write-Host "Index is $length"

                #parse
                $appId = $li_stats[$length].ApplicationId

                #determine if multi-valued
                $sections = $appId.IndexOf(";")
                if($sections -eq '-1'){
                
                    #process single
                    $client = $appId.Substring($appId.IndexOf("=") + 1, $appId.Length - $appId.IndexOf("=") - 1)
                    $useragent = ""

                }
                else{

                    #process values

                    #grab the client
                    $client = $appId.Substring($appId.IndexOf("=") + 1, $appId.IndexOf(";") - $appId.IndexOf("=") - 1)

                    #grab the useragent
                    $remainder = $appId.Substring($appId.IndexOf(";") + 1)

                    $sections = $remainder.IndexOf(";")
                    if($sections -eq '-1'){
                    
                        $useragent = $remainder
                    
                    }
                    else{
                    
                        $useragent = $remainder.Substring($remainder.IndexOf("=") + 1, $remainder.IndexOf(";") - $remainder.IndexOf("=") - 1)
                    
                    }
                
                }

                
                $data = @{'User'=$user; 'Client'=$client; 'UserAgent'=$useragent }
                Write-Host "The Data collected: $data"

            }
            while(($length -gt '0') -or ($length.ToString() -gt '0')) 
            
            Write-Host "The Data collected:" + $data.GetEnumerator()

            #create object
            $object = New-Object PSObject -Property $data

            #record
            Write-Output $object

}

<# Remove comment if you do not have these in your profile #>

<# 

Function Start-Session {
 param(
        [parameter(ParameterSetName="LiveAtEdu",HelpMessage="Use to connect to a Live@EDU environment.")][switch]$LiveAtEdu,
        [parameter(ParameterSetName="Office365",HelpMessage="Use to connect to an Office 365 environment.")][switch]$Office365,
        [parameter(HelpMessage="Use if connecting through the Internet Proxy service.")][switch]$UseProxy)


 if($LiveAtEdu){
    

    if($UseProxy){

        #Grabs Proxy Information
        $SessionOption = New-PSSessionOption -ProxyAccessType IEConfig 

        try{

            #Creates server-side connection
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection -SessionOption $SessionOption -EA SilentlyContinue

            #Imports LIVE Powershell commands
            Import-PSSession $Session -EA SilentlyContinue
        }

        catch {

            for( $i = 0; $i -lt 10; $i=$i+1 ){

                Start-Sleep -Milliseconds 10000 
                Write-Output "Retry $i - 10 seconds - Without Proxy Settings"
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection 
                Import-PSSession $Session 
    
            }

         }
    
    }
    else{

        #Creates server-side connection
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection -EA SilentlyContinue

        #Imports LIVE Powershell commands
        Import-PSSession $Session -EA SilentlyContinue
    
    }



 }

 if($Office365){

    Write-Host "Note: You must have the MSOnline Cmdlets installed for this to connect properly."
    
    Import-Module MSOnline

    if($UseProxy){

        #Grabs Proxy Information
        $SessionOption = New-PSSessionOption -ProxyAccessType IEConfig

        try{

            #Creates server-side connection
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection -SessionOption $SessionOption -EA SilentlyContinue

            #Imports LIVE Powershell commands
            Import-PSSession $Session -EA SilentlyContinue

        }

        catch {

            for( $i = 0; $i -lt 10; $i=$i+1 ){

                Start-Sleep -Milliseconds 10000 
                Write-Output "Retry $i - 10 seconds - Without Proxy Settings"
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection 
                Import-PSSession $Session 
    
            }
        }
    
    }
    else{

        #Creates server-side connection
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection -EA SilentlyContinue

        #Imports LIVE Powershell commands
        Import-PSSession $Session -EA SilentlyContinue
    
    }
    

    Connect-MsolService -Credential $cred

 }
}




Function End-Session {
    Get-PSSession | Remove-PSSession
    exit
}

#>