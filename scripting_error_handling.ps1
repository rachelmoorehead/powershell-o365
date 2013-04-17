<#

    Various O365 Specific Error Handling

#>


### Checking for an Active Session ###

$status = (Get-PSSession).State

if($status -eq 'Opened'){ <# Do Something #> } else { <# Start a New Session #> }


### Try - Catch Statements ###

try{  <# Your Remote Call #>  }

<# Specific Catch Statements #>

catch [ParameterArgumentValidationErrorNullNotAllowed] { <# Log It #> }

catch [System.Exception] { <# Log It #> }

catch { <# Generic Catch - Log It #> }