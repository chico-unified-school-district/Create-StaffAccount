function Create-ADUserObject {
 [cmdletbinding()]
 param(
  $SamId,
  $Name,
  $EmpId,
  $Pw,
  [switch]$WhatIf
 )

 $ou = 'OU=Substitue Teachers,OU=Teachers,OU=Employees,OU=Users,OU=Domain_Root,DC=chico,DC=usd'
 $securepw

 $attributes = @{
  Name                  = $Name
  SamAccountName        = $SamId
  UserPrincipalName     = $SamId + '@chicousd.org'
  EmployeeID            = $EmpId
  EMailAddress          = $Samid + '@chicousd.org'
  HomePage              = $Samid + '@chicousd.net'
  Company               = 'Chico Unified School District'
  Country               = 'US'
  Path                  = $ou
  Enabled               = $True
  AccountPassword       = $pw
  CannotChangePassword  = $False
  ChangePasswordAtLogon = $False
  PasswordNotrequired   = $True
  Whatif                = $WhatIf 
 }

 New-ADUser @attributes | Out-Null

 Get-ADUser -Identity $SamId -Properties *
}