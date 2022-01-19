function Create-ADUserObject {
 begin {
  Write-Host 'Begin Create-ADUserObject'
  $shortTermTypes = 'student teacher', 'coach', 'volunteer'
 }
 process {
  # $filter = " samAccountName -eq `'{0}`' -or employeeId -eq `'{1}`'" -f $_.samid, $_.empId
  $filter = "employeeId -eq `'{0}`'" -f $_.empId
  $obj = Get-ADUser -Filter $filter -Properties *
  if (-not$obj) {
   $securePw = ConvertTo-SecureString -String $_.pw1 -AsPlainText -Force
   $attributes = @{
    Name                  = $_.name
    GivenName             = $_.fn
    SurName               = $_.ln
    Title                 = $_.jobType
    Description           = $_.jobType
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
    Whatif                = $WhatIf
   }

   New-ADUser @attributes -ErrorAction Stop | Out-Null

   Write-Verbose 'Setting Extra User Attributes...'
   $samid = $_.samid
   if (-not$WhatIf) {
    if ( ([long]$_.empid -ge 100000000) -or ($shortTermTypes -match $_.jobType) ) {
     # ♥ If current month is greater than 6 (June) Set AccountExpirationDate to after the end of the current school term. ♥
     $year = "{0:yyyy}" -f $(if ([int](Get-Date -f MM) -gt 6) { (Get-Date).AddYears(1) } else { Get-Date })
     $accountExpirationDate = Get-Date "July 30 $year"
     '[{0}]Account Expiration set to [{1}]' -f $samid, $accountExpirationDate
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
    $proxyAddresses = "SMTP:$samid@chicousd.org", "smtp:$samid@mail.chicousd.org", "smtp:$samid@chicousd.mail.onmicrosoft.com"
    foreach ( $address in $proxyAddresses ) {
     Set-ADUser $samid -add @{proxyAddresses = "$address" }
    }
    $targetAddress = "SMTP:$samid@chicousd.mail.onmicrosoft.com"
    Set-ADUser $samid -replace @{targetAddress = "$targetAddress" }
    Set-ADUser $samid -replace @{msExchRecipientDisplayType = 0 }
    Set-ADUser $samid -replace @{co = 'United States' }
    Set-ADUser $samid -replace @{countryCode = 840 }
   }
  } # End New-ADUser
  # AD Sync Delay
  if (-not$WhatIf) { Start-Sleep 7 }
  Write-Verbose 'Gettings AD user afte robject creation and extra attributes applied'
  Get-ADUser -Filter $filter -Properties * | Select-Object name, proxyAddresses, LastLogonDate
 }
 end { Write-Host 'End Create-ADUserObject' }
}
