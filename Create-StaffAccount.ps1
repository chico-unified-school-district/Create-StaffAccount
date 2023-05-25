<#
 .Synopsis
 This script monitors an external employee database for new entries. when a new entry is detected
 the process performs various activities to prepare an account for use in the organization:
 - Active Directory Account
 - Home Directory
 - Office 365 Account
 - GSuite Account
 .DESCRIPTION
 This process creates accounts and home directories.
 This process is meant to be run every morning and run at intervals until a specified time every evening.
 .EXAMPLE
 .EXAMPLE
 .INPUTS
 .OUTPUTS
 .NOTES
#>
[cmdletbinding()]
param(
 [Alias('DCs')]
 [string[]]$DomainControllers,
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [Alias('MSCred')]
 [System.Management.Automation.PSCredential]$O365Credential,
 # [Alias('MgToken')]
 # [System.Management.Automation.PSCredential]$MGGraphToken,
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
 begin {
  function Format-FirstLetter ($str) {
   # This capitalizes the 1st letter of each word in a string.
   if ($str -and $str.length -gt 0) {
    $strArr = $str -split ' '
    $newArr = @()
    $strArr.foreach({ $newArr += $_.substring(0, 1).ToUpper() + $_.substring(1) })
    $newArr -join ' '
   }
  }
 }
 process {
  $filter = "employeeId -eq `'{0}`'" -f $_.empId
  $obj = Get-ADUser -Filter $filter -Properties *
  if ($obj) {
   $newName = $obj.name
  }
  else {
   $newName = Create-Name -First $_.nameFirst -Middle $_.nameMiddle -Last $_.nameLast
  }
  $samId = $_ | Set-SamId
  $empId = $_ | Set-EmpId
  # $siteData = Set-Site -siteCode $_.siteCode
  # $siteData
  $hash = @{
   id         = $_.id
   fn         = Format-FirstLetter $_.nameFirst
   ln         = Format-FirstLetter $_.nameLast
   mi         = Format-FirstLetter $_.nameMiddle
   name       = $newName
   samid      = $samId
   empid      = $empId
   bargUnitId = $_.BargUnitId
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
 begin {
  function New-RandomEmpId {
   $id = Get-Random -Min 100000000 -Max 1000000000
   $filter = " employeeId -eq `'{0}`' " -f $id
   $obj = Get-ADuser -Filter $filter
   if (-not$obj ) { $id }
   else { New-RandomEmpId }
  }
 }
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
  $siteData = $lookupTable | Where-Object { [int]$_.siteCode -eq [int]$sc }
  if (-not$siteData) { return }
  Write-Host ('{0} {1} {2} {3} {4} Site' -f $_.empid, $_.fn, $_.ln, $sc, $siteData.SiteDescr)
  $siteData
 }
}

function Update-ADGroupMemberships {
 begin {
  $A5 = 'Office365_A5_Faculty' # Microsoft 365 License for admin
  $A1 = 'Office365_A1_Faculty'
 }
 process {
  $azureGroup = if ($_.BargUnitId -eq 'CUMA') { $A5 } else { $A1 }
  # Add user to various groups
  $groups = 'Staff_Filtering', 'staffroom', 'Employee-Password-Policy', $azureGroup
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

function Update-IntDB {
 begin {
  $baseSql = "UPDATE {0}
   SET emailWork = '{1}'
   SET gsuite = '{2}'
   SET empId = {3}
   SET tempPw = '{4}'
   SET samAccountName = '{5}'
   SET sourceSystem = '{6}'
   dts = CURRENT_TIMESTAMP
  WHERE id = {7};"
 }
 process {
  $updateVars = @($NewAccountsTable, $_.emailWork, $_.gsuite, [long]$_.empid, $_.pw2, $_.samid, $ENV:COMPUTERNAME, $_.id)
  $sql = $baseSql -f $updateVars
  Write-Host $sql
  if (-not$WhatIf) { Invoke-SqlCmd @intermediateDBparams -Query $sql }
 }
}

function Update-IntDBEmailWork {
 process {
  $sql = "UPDATE {0} SET emailWork = `'{1}`', dts = CURRENT_TIMESTAMP WHERE id = {2}" -f $NewAccountsTable, $_.emailWork, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-SqlCmd @intermediateDBparams -Query $sql }
 }
}
function Update-IntDBGsuite {
 process {
  $sql = "UPDATE {0} SET gsuite = `'{1}`' WHERE id = {2}" -f $NewAccountsTable, $_.gsuite, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-SqlCmd @intermediateDBparams -Query $sql }
 }
}

function Update-IntDBEmpID {
 process {
  $sql = "UPDATE {0} SET empId = {1} WHERE id = `'{2}`'" -f $NewAccountsTable, [long]$_.empid, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}

function Update-IntDBTempPw {
 process {
  $sql = "UPDATE {0} SET tempPw = `'{1}`' WHERE id = {2}" -f $NewAccountsTable, $_.pw2, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}
function Update-IntDBSamAccountName {
 process {
  $sql = "UPDATE {0} SET samAccountName = `'{1}`' WHERE id = {2}" -f $NewAccountsTable, $_.samid, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}
function Update-IntDBSrcSys {
 process {
  $sql = "UPDATE {0} SET sourceSystem = `'{1}`' WHERE id = {2}" -f $NewAccountsTable, $ENV:COMPUTERNAME, $_.id
  Write-Host $sql
  if (-not$WhatIf) { Invoke-Sqlcmd @intermediateDBparams -Query $sql }
 }
}

function Update-AzureLicense {
 begin {
  $skuIdA1 = Get-MgSubscribedSku -all | Where-Object skupartnumber -eq 'STANDARDWOFFPACK_FACULTY'
  $skuIdA5 = Get-MgSubscribedSku -all | Where-Object skupartnumber -eq 'M365EDU_A5_FACULTY'
 }
 process {
  $sku = if ($_.BargUnitId -eq 'CUMA') { $skuIdA5 } else { $skuIdA1 }
  Write-Host ('{0},UsageLocation: [US],License: [{1}]' -f $MyInvocation.MyCommand.Name, $sku.SkuPartNumber)
  # get-mguser -UserId RHammerplamp@chicousd.org -Property UsageLocation | select UsageLocation
  Update-MgUser -UserId $_.mgUserId -UsageLocation "US"
  # Add license uses a hash table. Remove uses an array. But why though?!?
  Set-MgUserLicense -UserId $_.mgUserId -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @() -WhatIf:$WhatIf
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

# Imported Functions
. .\lib\Create-ADUserObject.ps1
. .\lib\Create-Name.ps1
. .\lib\Create-PassPhrase.ps1
. .\lib\Create-SamID.ps1
. .\lib\Create-StaffHomeDir.ps1
. .\lib\Load-Module.ps1
. .\lib\New-ADSession.ps1
. .\lib\Select-DomainController.ps1
. .\lib\Show-TestRun.ps1

$gam = '.\bin\gam-64\gam.exe'
Write-Host ( 'gam path: {0}' -f $gam )
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

$stopTime = Get-Date "11:00pm"
$delay = 60
'Process looping every {0} seconds until {1}' -f $delay, $stopTime
do {
 if ($WhatIf) { Show-TestRun }
 cleanUp

 $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json

 $newAccountSql = 'SELECT * FROM {0} WHERE emailWork IS NULL' -f $NewAccountsTable
 $newAccountData = Invoke-Sqlcmd @intermediateDBparams -Query $newAccountSql
 if ($newAccountData) {
  # 'Microsoft.Graph', 'SqlServer', 'ExchangeOnlineManagement' | Load-Module
  'SqlServer', 'ExchangeOnlineManagement' | Load-Module

  $dc = Select-DomainController $DomainControllers
  New-ADSession -dc $dc -Cred $ActiveDirectoryCredential
  Connect-ExchangeOnline -Credential $O365Credential -ShowBanner:$false
  # Connect-MgGraph -AccessToken $MGGraphToken
 }

 # Create New User Data Variables
 $varList = @()
 foreach ($row in $newAccountData) {
  $personData = $row | New-UserPropObject
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
  $checkUser = $userData | Create-ADUserObject
  if ($checkUser.LastLogonDate -is [datetime]) {
   Write-Host ('{0}, LastLogonDate already present. User account exists and is in use. Skipping Phase I' -f $checkUser.mail)
   continue
  }
  $userData | Update-ADGroupMemberships
  $userData | Update-IntDBEmpID
  Create-StaffHomeDir -userData $userData -ServerCredential $FileServerCredential -WhatIf:$WhatIf
 }

 foreach ($var in $varList) {
  "===============Wait for Azure sync and assign Microsoft licensing=================="
  $userData = Get-Variable -Name $var -ValueOnly
  $checkUser = $userData | Create-ADUserObject
  if ($checkUser.LastLogonDate -isnot [datetime]) {
   Write-Host ( '{0} {1} Phase II' -f $userData.empid , $userData.emailWork )
   $script:countThis = 0

   $mailBoxBlock = "Get-Mailbox -Identity {0} -ErrorAction SilentlyContinue" -f $userData.emailWork
   if (-not($mailBoxBlock | Get-UserData)) { if (-not($WhatIf)) { continue } }

   $gsuiteBlock = "(`$guser = .`$gam print users query `"email:{0}`" | ConvertFrom-Csv)*>`$null;`$guser" -f $userData.gsuite
   $gsuiteData = $gsuiteBlock | Get-UserData
   if (-not($gsuiteData)) { if (-not($WhatIf)) { continue } } else { Start-Sleep 60; $gsuiteData }
  }
  else {
   Write-Host ('{0}, LastLogonDate already present. User account exists and is in use. Skipping Phases II and III' -f $checkUser.mail)
   $userData.pw2 = 'Account Already Active. Password Not Changed.'
   $userData | Update-EscapeEmailWork
   #TODO Consollidate the IntDB Updates
   $userData | Update-IntDBSamAccountName
   $userData | Update-IntDBTempPw
   $userData | Update-IntDBSrcSys
   $userData | Update-IntDBGsuite
   $userData | Update-IntDBEmailWork
   '{0} {1} Account creation complete' -f $userData.empid, $userData.emailWork
   continue
  }

  Write-Host ( '{0} {1} Phase III' -f $userData.empid , $userData.emailWork )

  $userData | Update-EscapeEmailWork
  $userData | Update-IntDBSamAccountName
  $userData | Update-PW
  $userData | Update-IntDBTempPw
  $userData | Update-IntDBSrcSys
  $userData | Update-IntDBGsuite
  $userData | Update-IntDBEmailWork
  '{0} {1} Account creation complete' -f $userData.empid, $userData.emailWork
 }

 foreach ($var in $varList) {
  Get-Variable -Name $var | Remove-Variable -Confirm:$false
 }

 cleanUp
 if ($WhatIf) { Show-TestRun }
 if (-not$WhatIf) {
  # Loop delay
  Start-Sleep $delay
 }
} until ($WhatIf -or ((Get-Date) -ge $stopTime))