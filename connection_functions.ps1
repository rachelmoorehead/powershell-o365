<# Function Declarations

Note: Ideally added to Windows PowerShell Profile.

Start-Session : Initiate a PS connection to Live@EDU or O365 and import in the O365 API (if required)
End-Session   : End Current Session
Switch-Tenant : Switch between Microsoft Environments (if multiple tenants, testing or production)

#>

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




#TODO Add Proxy Consideration
Function Switch-Tenant {
    param (
        [parameter(DataSet="O365",HelpMessage="Connect to Office 365")][switch]$Office365,
        [parameter(DataSet="LiveAtEdu",HelpMessage="Connect to Live@EDU")][switch]$LiveAtEDU
        )

    Get-PSSession | Remove-PSSession 
    Clear-Host
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential ($cred = Get-Credential) -Authentication Basic -AllowRedirection 
    Import-PSSession $Session
    if($Office365){
        Import-Module MSOnline
        Write-Output "Connecting to O365."
        Connect-MsolService -Credential $cred
    }

}
