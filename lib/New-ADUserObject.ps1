function New-ADUserObject {
 begin {
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $_.empId, $_.samid)
  $shortTermTypes = 'student teacher', 'coach', 'volunteer', 'student worker'
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
    PasswordNotrequired   = $True
    WhatIf                = $WhatIf
   }

   New-ADUser @attributes -ErrorAction Stop | Out-Null

   Write-Verbose 'Setting Extra User Attributes...'
   $samid = $_.samid
   if (-not$WhatIf) {
    if ( ($_.empid -match '\d{7,}') -or ($shortTermTypes -match $_.jobType) ) {
     # ♥ If current month is greater than 6 (June)
     # Set AccountExpirationDate to after the end of the current school term. ♥
     $year = "{0:yyyy}" -f $(if ([int](Get-Date -f MM) -gt 6) { (Get-Date).AddYears(1) } else { Get-Date })
     $accountExpirationDate = Get-Date "July 30 $year"
     Write-Host '{0},{1}, Account Expiration set: {1}' -f $MyInvocation.MyCommand.Name, $samid, $accountExpirationDate
     Set-ADUser -Identity $samid -AccountExpirationDate $AccountExpirationDate
    }

    if ( -not([DBNull]::Value).Equals($_.mi) -and ($_.mi.length -gt 0)) {
     $middleName = $_.mi
     Set-ADUser $samid -replace @{middleName = "$middleName"; Initials = $($middleName.substring(0, 1)) }
    }
    #TODO
    # if ( $BargUnitId ) {
    #  Set-ADUser $samid -add @{extensionAttribute1 = "$bargUnitID" }
    # }
    # Main Proxy address has 'SMTP' in UPPER case. Alternate Proxy Addresses use lowercase 'smpt'
    $proxyAddresses = "SMTP:$samid@chicousd.org",
    "smtp:$samid@mail.chicousd.org", "smtp:$samid@chicousd.mail.onmicrosoft.com"
    foreach ( $address in $proxyAddresses ) {
     Set-ADUser $samid -add @{proxyAddresses = "$address" }
    }
    $targetAddress = "SMTP:$samid@chicousd.mail.onmicrosoft.com"
    Set-ADUser $samid -replace @{targetAddress = "$targetAddress" }
    Set-ADUser $samid -replace @{msExchRecipientDisplayType = 0 }
    Set-ADUser $samid -replace @{co = 'United States' }
    Set-ADUser $samid -replace @{countryCode = 840 }
    if ($_.siteCode -gt 0) { Set-ADUser $samid -replace @{DepartmentNumber = $_.siteCode } }
   }
   # AD Sync Delay
   if (-not$WhatIf) { Start-Sleep 7 }
  } # End New-ADUser
  Get-ADUser -Filter $filter
 }
}
