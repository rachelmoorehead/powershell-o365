# Primary Licensing Cmdlets

Get-MsolAccountSku
Get-MsolSubscription
Set-MsolUserLicense

## Find Your Licensing

## Licensing Manipulation
Get-MsolAccountSku | Format-Table AccountSkuId, SkuPartNumber
$ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "[SkuPartNumber]"}
# i.e. $ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "[ENTERPRISEPACK]"}

$ServicePlans.ServiceStatus
$MyO365Sku = New-MsolLicenseOptions -AccountSkuId [tenantname:AccountSkuId] -DisabledPlans Comma_Seperated_List_From_ServicePlan_Output

# In the STANDARDWOFFPACK_STUDENT and STANDARDWOFFPACK_FACULTY plans:  
#  SHAREPOINTWAC_EDU, MCOSTANDARD, SHAREPOINTSTANDARD_EDU, EXCHANGE_S_STANDARD
$STUDENT_SKU = New-MsolLicenseOptions -AccountSkuId tenantname:STANDARDWOFFPACK_STUDENT -DisabledPlans SHAREPOINTWAC_EDU,MCOSTANDARD,SHAREPOINTSTANDARD_EDU
$FACULTY_SKU = New-MsolLicenseOptions -AccountSkuId tenantname:STANDARDWOFFPACK_FACULTY -DisabledPlans SHAREPOINTWAC_EDU,MCOSTANDARD,SHAREPOINTSTANDARD_EDU

# Set User Location and Licensing
Set-MsolUser -UserPrincipalName user1@tenant.onmicrosoft.com -UsageLocation US
Set-MsolUserLicense -UserPrincipalName user1@tenant.onmicrosoft.com -AddLicenses tenantname:STANDARDWOFFPACK_STUDENT -LicenseOptions $STUDENT_SKU
## End

#TODO Scheduled Task script to auto-license on provision action with correct options