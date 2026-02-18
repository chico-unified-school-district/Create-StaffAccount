<#
.SYNOPSIS
Creates or updates an Active Directory user object using values from the pipeline object.

.DESCRIPTION
New-ADUserObject accepts a PSCustomObject via the pipeline (the script's processing object) and uses properties
on that object to construct a New-ADUser call and then apply additional AD attribute updates such as proxyAddresses,
targetAddress, department number and optional account expiration. The function respects the global -WhatIf switch
passed to the calling script and avoids making persistent changes when -WhatIf is set.

.INPUTS
PSCustomObject from the pipeline. Expected properties (not all required): name, fn, ln, mi, new (original row), samid,
empid, emailWork, pw1 (temporary password), targetOU, company, new.siteCode, new.jobType

.OUTPUTS
The same PSCustomObject passed in (object is emitted to the pipeline after operations complete).

.NOTES
This function depends on AD cmdlets (New-ADUser, Set-ADUser) being available and the caller setting $WhatIf if a dry
run is required.
#>
function New-ADUserObject {
 process {
  $securePw = ConvertTo-SecureString -String $_.pw1 -AsPlainText -Force
  $attributes = @{
   Name                  = $_.name
   DisplayName           = $_.name
   GivenName             = $_.fn
   SurName               = $_.ln
   Title                 = $_.new.jobType
   Description           = ($_.new.siteDescr + ' ' + $_.new.jobType).Trim()
   Office                = $_.new.siteDescr
   SamAccountName        = $_.samid
   UserPrincipalName     = $_.emailWork
   EmployeeID            = $_.empid
   EMailAddress          = $_.emailWork
   HomePage              = $_.gsuite
   Company               = $_.company
   Country               = 'US'
   Path                  = $_.targetOU
   Enabled               = $True
   AccountPassword       = $securePw
   CannotChangePassword  = $False
   ChangePasswordAtLogon = $False
   PasswordNotRequired   = $True
   WhatIf                = $WhatIf
  }

  New-ADUser @attributes -ErrorAction Stop | Out-Null

  Write-Verbose ('{0},{2},Setting Extra User Attributes...' -f $MyInvocation.MyCommand.Name)

  if (!$WhatIf -and ($_.mi -match '\w')) {
   $middleName = $_.mi
   Set-ADUser -Identity $_.samid -Replace @{middleName = "$middleName"; Initials = $($middleName.substring(0, 1)) }
  }

  # Main Proxy address has 'SMTP' in UPPER case. Alternate Proxy Addresses use lowercase 'smpt'
  # $proxyAddresses = "SMTP:$samid@chicousd.org", "smtp:$samid@mail.chicousd.org", "smtp:$samid@chicousd.mail.onmicrosoft.com"
  # foreach ( $address in $proxyAddresses ) {
  #  if (!$WhatIf) { Set-ADUser -Identity $samid -Add @{proxyAddresses = "$address" } }
  # }
  # $targetAddress = "SMTP:$samid@chicousd.mail.onmicrosoft.com"
  if (!$WhatIf) {
   Set-ADUser -Identity $_.samid -Replace @{
    # targetAddress              = "$targetAddress"
    # msExchRecipientDisplayType = 0
    co          = 'United States'
    countryCode = 840
   }
  }
  if (!$WhatIf -and ($_.new.siteCode -match '\d')) { Set-ADUser $_.samid -Replace @{DepartmentNumber = $_.new.siteCode } }
  # AD Sync Delay
  if (!$WhatIf) { Start-Sleep 30 }
 } # End Process
}
