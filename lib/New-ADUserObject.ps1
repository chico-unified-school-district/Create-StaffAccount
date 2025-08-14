function New-ADUserObject {
 begin {
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $_.empId, $_.samid)
  $shortTermTypes = 'student teacher', 'coach', 'volunteer', 'student worker', 'intern'
 }
 process {
  $filter = "employeeId -eq `'{0}`'" -f $_.empId
  $obj = Get-ADUser -Filter $filter -Properties *
  if (-not$obj) {
   $securePw = ConvertTo-SecureString -String $_.pw1 -AsPlainText -Force
   $attributes = @{
    Name                  = $_.name
    DisplayName           = $_.name
    GivenName             = $_.fn
    SurName               = $_.ln
    Title                 = $_.jobType
    Description           = ($_.siteDescr + ' ' + $_.jobType).Trim()
    Office                = $_.siteDescr
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

   Write-Verbose ('{0},Setting Extra User Attributes...' -f $MyInvocation.MyCommand.Name)
   $samid = $_.samid
   if ( ($_.empid -match '\d{7,}') -or ($shortTermTypes -match $_.jobType) ) {
    # ♥ If current month is greater than 6 (June), set AccountExpirationDate to after the end of the current school term. ♥
    $year = "{0:yyyy}" -f $(if ([int](Get-Date -f MM) -gt 6) { (Get-Date).AddYears(1) } else { Get-Date })
    $accountExpirationDate = Get-Date "July 30 $year"
    Write-Host '{0},{1},Setting Account Expiration to: {1}' -f $MyInvocation.MyCommand.Name, $samid, $accountExpirationDate
    Set-ADUser -Identity $samid -AccountExpirationDate $AccountExpirationDate -WhatIf:$WhatIf
   }

   if ( $_.mi -match '\w') {
    $middleName = $_.mi
    Set-ADUser $samid -replace @{middleName = "$middleName"; Initials = $($middleName.substring(0, 1)) } -WhatIf:$WhatIf
   }

   # Main Proxy address has 'SMTP' in UPPER case. Alternate Proxy Addresses use lowercase 'smpt'
   $proxyAddresses = "SMTP:$samid@chicousd.org",
   "smtp:$samid@mail.chicousd.org", "smtp:$samid@chicousd.mail.onmicrosoft.com"
   foreach ( $address in $proxyAddresses ) { Set-ADUser $samid -add @{proxyAddresses = "$address" } -WhatIf:$WhatIf }
   $targetAddress = "SMTP:$samid@chicousd.mail.onmicrosoft.com"
   Set-ADUser $samid -replace @{
    targetAddress              = "$targetAddress"
    msExchRecipientDisplayType = 0
    co                         = 'United States'
    countryCode                = 840
   } -WhatIf:$WhatIf
   if ($_.siteCode -gt 0) { Set-ADUser $samid -replace @{DepartmentNumber = $_.siteCode } }
   # AD Sync Delay
   if (-not$WhatIf) { Start-Sleep 7 }
  } # End New-ADUser
  Get-ADUser -Filter $filter
 } # End Process
}
