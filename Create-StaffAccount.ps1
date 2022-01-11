[cmdletbinding()]
param(
 [Alias('DC')]
 [string]$DomainController,
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [Alias('MSCred')]
 [System.Management.Automation.PSCredential]$O365Credential,
 [Alias('FSCred')]
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [Alias('ESCServer')]
 [string]$EscapeServer,
 [Alias('ESCDB')]
 [string]$EscapeDatabase,
 [Alias('ESCCred')]
 [System.Management.Automation.PSCredential]$EscapeCredential,
 [Alias('IntServer')]
 [string]$IntermediateSqlServer,
 [Alias('IntDB')]
 [string]$IntermediateDatabase,
 [Alias('Table')]
 [string]$NewAccountsTable,
 [Alias('IntCred')]
 [System.Management.Automation.PSCredential]$IntermediateCredential,
 [Alias('OU')]
 [string]$TargetOrgUnit,
 [string]$Organization,
 [Alias('wi')]
 [switch]$WhatIf
)


# Imported Functions
. .\lib\Create-ADUserObject.ps1
. .\lib\Create-Name.ps1
. .\lib\Create-O365PSSession.ps1
. .\lib\Create-PassPhrase.ps1
. .\lib\Create-SamID.ps1
. .\lib\Create-StaffHomeDir.ps1
. .\lib\Load-Module.ps1
. .\lib\Show-TestRun.ps1

# Script Functions

function cleanUp {
 Get-Module -name *tmp* | Remove-Module -Confirm:$false -Force
 Get-PSSession | Remove-PSSession -Confirm:$false
}

function Get-EscapeData {
 process {
  $sql = 'SELECT * FROM vwHREmployementList WHERE empId = {0}' -f $_
  Invoke-SqlCmd @escapeDBParams -Query $sql
 }
}

function Get-UserData {
 process {
  $scriptBlock = [scriptblock]::Create($_)
  if ($script:countThis -eq 0) { Write-Host ('{0} Running: {1}' -f (Get-Date), $_) }
  if (-not($WhatIf)) {
   $obj = & $scriptBlock
   if ($obj) {
    $script:countThis = 0
    Write-Host ('{0} Success: {1}' -f (Get-Date), $_)
    $obj
   }
   elseif ($script:countThis -lt 90) {
    $script:countThis++
    # Write-Host $script:countThis
    Start-Sleep 60
    $_ | Get-UserData
   }
   else {
    Write-Host ('{0} Failed: {1}' -f (Get-Date), $_)
    $script:countThis = 0
    # break
   }
  }
 }
}

function New-UserPropObject {
 process {
  $newName = Create-Name -First $_.nameFirst -Middle $_.nameMiddle -Last $_.nameLast
  $samId = $_ | Set-SamId
  $empId = $_ | Set-EmpId
  # $siteData = Set-Site -siteCode $_.siteCode
  # $siteData
  $hash = @{
   fn         = $_.nameFirst
   ln         = $_.nameLast
   mi         = $_.nameMiddle
   name       = $newName
   samid      = $samId
   empid      = $empId
   emailHome  = $_.emailHome
   emailWork  = $samid + '@chicousd.org'
   gsuite     = $samid + '@chicousd.net'
   siteDescr  = $_.SiteDescr
   siteCode   = $_.siteCode
   company    = $Organization
   jobType    = $_.jobType
   pw1        = Create-PassPhrase
   pw2        = Create-PassPhrase
   groups     = ''
   fileServer = ''
   targetOU   = $TargetOrgUnit
  }
  New-Object PSObject -Property $hash
 }
}

function Set-EmpId {
 process {
  # create a large random number when employeeid is DBNULL
  # A large int is used to ensure no overlap with the current Escape empId's scope
  Write-Host 'trying to set EMPID'
  if ( ([DBNull]::Value).Equals($_.empId) -or ($_.empId -eq 0)) { $id = Get-Random -Min 100000000 -Max 1000000000 }
  else { $id = $_.empId }
  Write-Host ('Empid set to [{0}]' -f $id)
  $id
 }
}

function Set-SamId {
 process {
  if ( ([DBNull]::Value).Equals($_.empId) ) {
   Create-Samid -First $_.nameFirst -Middle $_.nameMIddle -Last $_.nameLast
   return
  }
  $id = $_.empId
  $obj = Get-ADUser -Filter "employeeID -eq `'$id`'" -ErrorAction SilentlyContinue
  if ($obj) {
   $obj.samAccountName
   return
  }
  else {
   Create-Samid -First $_.nameFirst -Middle $_.nameMIddle -Last $_.nameLast
  }
 }
}

function Set-Site {
 process {
  $sc = $_.siteCode
  Write-Host ('{0} {1} {2} {3} Determining site data' -f $_.empid, $_.fn, $_.ln, $sc)
  $lookupTable | Where-Object { [int]$_.siteCode -eq [int]$sc }
 }
}

function Update-ADGroupMemberships {
 process {
  # Add user to various groups
  $groups = 'Staff_Filtering', 'staffroom', 'Employee-Password-Policy'
  if ( $_.groups ) { $groups += $_.groups.Split(",") }
  Write-Host ('Adding {0} to {1}' -f $_.samid, ($groups -join ','))
  if ( -not$WhatIf ) { Add-ADPrincipalGroupMembership -Identity $_.samid -MemberOf $groups }
 }
}

function Update-EscapeEmailWork {
 process {
  $checkEscapeUserSql = 'SELECT empId FROM HREmployment WHERE empID = {0}' -f $_.empid
  Write-Verbose $checkEscapeUserSql
  $escapeResult = Invoke-Sqlcmd @escapeDBParams -Query $checkEscapeUserSql
  if ($escapeResult) {
   $updateEscapeEmailSql = "UPDATE HREmployment SET EmailWork = `'{0}`' WHERE EmpID = {1}" -f $_.emailWork, $_.empid
   Write-Host $updateEscapeEmailSql
   if (-not$WhatIf) {
    Invoke-SqlCmd @escapeDBParams -Query $updateEscapeEmailSql
   }
  }
  else {
   Write-Host ('{0} {1} not found in Escape' -f $_.empid, $_.emailWork)
  }
 }
}

function Update-IntDBEmailWork {
 process {
  $sql = "UPDATE {0} SET emailWork = `'{1}`', dts = CURRENT_TIMESTAMP WHERE emailHome = `'{2}`'" -f $NewAccountsTable, $_.emailWork, $_.emailHome
  Write-Verbose $sql
  if (-not$WhatIf) { Invoke-SqlCmd @intermediateDBparams -Query $sql }
 }
}

function Update-IntDBEmpID {
 process {
  $sql = "UPDATE new_employee_accounts SET empId = {0} WHERE emailHome = `'{1}`'" -f [long]$_.empid, $_.emailHome
  Write-Verbose $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}

function Update-IntDBTempPw {
 process {
  $sql = "UPDATE new_employee_accounts SET tempPw = `'{0}`' WHERE empId = {1}" -f $_.pw2, $_.empid
  Write-Verbose $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}
function Update-IntDBSamAccountName {
 process {
  $sql = "UPDATE new_employee_accounts SET samAccountName = `'{0}`' WHERE empId = {1}" -f $_.samid, $_.empid
  Write-Verbose $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}
function Update-IntDBSrcSys {
 process {
  $sql = "UPDATE new_employee_accounts SET sourceSystem = `'{0}`' WHERE empId = {1}" -f $ENV:COMPUTERNAME, $_.empid
  Write-Verbose $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}

function Update-MsolLicense {
 begin { $targetLicense = 'chicousd:STANDARDWOFFPACK_FACULTY' }
 process {
  Write-Host ('[{0}] Assigning Msol Region [US] and License [{1}]' -f $_.emailWork, $targetLicense)
  if (-not$WhatIf) {
   Set-MsolUser -UserPrincipalName $_.emailWork -UsageLocation US -ErrorAction SilentlyContinue
   Set-MsolUserLicense -UserPrincipalName $_.emailWork -AddLicenses $targetLicense -ErrorAction SilentlyContinue
  }
 }
}

function Update-PW {
 process {
  $securePw = ConvertTo-SecureString -String $_.pw2 -AsPlainText -Force
  Write-Host ( '{0} Updating Password' -f $_.samid )
  if (-not$WhatIf) {
   Set-ADAccountPassword -Identity $_.samid -NewPassword $securePw -Confirm:$false
  }
 }
}

# ==================================================================
# ==================================================================

$gam = '.\bin\gam-64\gam.exe'
Write-Verbose ( 'gam path: {0}' -f $gam )
$escapeDBParams = @{
 Server     = $EscapeServer
 Database   = $EscapeDatabase
 Credential = $EscapeCredential
}

$intermediateDBparams = @{
 Server     = $IntermediateSqlServer
 Database   = $IntermediateDatabase
 Credential = $IntermediateCredential
}

# ==================================================================
# ==================================================================

$stopTime = Get-Date "11:00pm"
'Process looping once a minute until {0}' -f $stopTime
do {
 if ($WhatIf) { Show-TestRun }
 cleanUp

 $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json

 $newAccountSql = 'SELECT * FROM {0} WHERE emailWork IS NULL' -f $NewAccountsTable
 $newAccountData = Invoke-Sqlcmd @intermediateDBparams -Query $newAccountSql
 if ($newAccountData) {
  'MSOnline', 'SqlServer' | Load-Module

  $adSession = New-PSSession -ComputerName $DomainController -Credential $ActiveDirectoryCredential
  Import-PSSession -Session $adSession -Module ActiveDirectory -AllowClobber | Out-Null

  Connect-MsolService -Credential $O365Credential -ErrorAction Stop
  Create-O365PSSession -Credential $O365Credential
 }

 # Create New User Data Variables
 $varList = @()
 foreach ($row in $newAccountData) {
  $personData = $row | New-UserPropObject
  # $site = $lookupTable | Where-Object { ($_.SiteDescr -eq $row.siteDescr) -or ($_.SiteCode -eq $row.siteId) | Select-Object -First 1 }
  $site = $personData | Set-Site
  $personData.groups = $site.groups
  $personData.fileServer = $site.fileServer
  $varName = $personData.samId
  $varList += $varName
  New-Variable -Name $varName -Value $personData -Scope Script
 }

 foreach ($var in $varList) {
  "+++++++++++++++++++++Create AD Accounts and Home Directories+++++++++++++++++++"
  $userData = Get-Variable -Name $var -ValueOnly
  $userData
  Write-Host ( '{0} {1} Phase I' -f $userData.empid , $userData.emailWork )
  Write-Verbose ( $userData | Out-String )
  Write-Debug 'User settings ok?'
  $userData | Create-ADUserObject
  $userData | Update-ADGroupMemberships
  $userData | Update-IntDBEmpID
  Create-StaffHomeDir -userData $userData -ServerCredential $FileServerCredential -WhatIf:$WhatIf
 }

 foreach ($var in $varList) {
  "===============Wait for Azure sync and assign Microsoft licensing=================="
  $userData = Get-Variable -Name $var -ValueOnly
  Write-Host ( '{0} {1} Phase II' -f $userData.empid , $userData.emailWork )
  $script:countThis = 0

  $msolBlock = "Get-MsolUser -SearchString {0} -All" -f $userData.emailWork
  $msolUser = $msolBlock | Get-UserData
  if (-not($msolUser)) { if (-not($WhatIf)) { continue } }
  # # Add MS license if needed
  if ($msolUser.IsLicensed -eq $false ) {
   $userData | Update-MsolLicense
   $msolUser = $msolBlock | Get-UserData
   if ($msolUser.IsLicensed -eq $false) {
    $errorMsg = '{0} {1} Licensing Failed. Skipping' -f $userData.empid, $userData.emailWork
    Write-Error $errorMsg
    if (-not($WhatIf)) { continue }
   }
   else {
    '{0} {1} Licensing Succeeded.' -f $userData.empid, $userData.emailWork
   }
  }
  $mailBoxBlock = "Get-Mailbox -Identity {0} -ErrorAction SilentlyContinue" -f $userData.emailWork
  if (-not($mailBoxBlock | Get-UserData)) { if (-not($WhatIf)) { continue } }

  $gsuiteBlock = "(`$guser = .`$gam print users query `"email:{0}`" | ConvertFrom-Csv)*>`$null;`$guser" -f $userData.gsuite
  $gsuiteData = $gsuiteBlock | Get-UserData
  if (-not($gsuiteData)) { if (-not($WhatIf)) { continue } } else { Start-Sleep 60; $gsuiteData }

  Write-Host ( '{0} {1} Phase III' -f $userData.empid , $userData.emailWork )

  $userData | Update-EscapeEmailWork
  $userData | Update-IntDBSamAccountName
  $userData | Update-PW
  $userData | Update-IntDBTempPw
  $userData | Update-IntDBSrcSys
  $userData | Update-IntDBEmailWork
  '{0} {1} Account creation complete' -f $userData.empid, $userData.emailWork
 }

 foreach ($var in $varList) {
  Get-Variable -Name $var | Remove-Variable -Confirm:$false
 }

 cleanUp
 if ($WhatIf) { Show-TestRun }
 if (-not$WhatIf) {
  # Loop once a minute
  Start-Sleep 60
 }
} until ($WhatIf -or ((Get-Date) -ge $stopTime))