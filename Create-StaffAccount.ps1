[cmdletbinding()]
param(
 [Alias('EmpID')]
 [int]$EmployeeID,
 [Alias('fn', 'NameFirst')]
 [string]$FirstName,
 [Alias('mn', 'NameMiddle')]
 [string]$MiddleName,
 [Alias('ln', 'NameLast')]
 [string]$LastName,
 [ValidateSet(500, 510, 520, 540, 550, 560, 570, 580, 600, 620, 640, 660, 740,
  010, 020, 030, 050, 060, 070, 080, 090, 111, 110, 120, 130, 160, 180, 190,
  200, 210, 230, 240, 250, 260, 270, 280, 380, 191, 330, 430, 440, 999)]
 [Alias('sc')]
 [int]$SiteCode,
 [switch]$SubAccount,
 [System.Management.Automation.PSCredential]$O365Credential,
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [string]$EscapeServer,
 [string]$EscapeDatabase,
 [System.Management.Automation.PSCredential]$EscapeCredential,
 [string]$IntermediateSqlServer,
 [string]$IntermediateDatabase,
 [string]$NewAccountsTable,
 [System.Management.Automation.PSCredential]$IntermediateCredential,
 [Alias('wi')]
 [switch]$WhatIf
)

$lookupTable = Get-Content -Path .\config\lookup-table.json | ConvertFrom-Json
$config = Get-Content -Path .\config\config.json | ConvertFrom-Json

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

function userPropsObj {
 $hash = @{
  name  = Create-Name -First $FirstName -Middle $MiddleName -Last $LastName
  fn    = $FirstName
  sam   = setSamid
  empid = setEmpId
  pw1   = Create-PassPhrase
  pw2   = Create-PassPhrase
 }
 New-Object PSObject -Property $hash
}
function setEmpId ($empId) {
 process {
  # create a large random number when employeeid is omitted
  # A large int is used to ensure no overlap with the current Escape empId's scope
  if (-not($empId)) { Get-Random -Min 100000000 -Max 1000000000 }
  else { $empId }
 }
}
function setSamid {
 if ($EmployeeID) {
  $userObj = Get-ADUser -LDAPFilter "(employeeID=$EmployeeID)"
  if ($userObj) {
   $userObj.samAccountName
   return
  }
 }
 Create-Samid -First $FirstName -Middle $MiddleName -Last $LastName
}

cleanUp
Show-TestRun

'SqlServer' | Load-Module
# 'ActiveDirectory', 'MSOnline', 'SqlServer' | Load-Module

# Connect-MsolService -Credential $O365Credential -ErrorAction Stop

# Create-O365PSSession -Credential $jcred
if ($EmployeeId) { $userObj = Get-ADUser -LDAPFilter "(employeeID=$EmployeeID)" }
if (-not($userObj)) {
 userPropsObj
}

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

$newAccountSql = 'SELECT * FROM {0} WHERE emailWork IS NULL' -f $NewAccountsTable
$newAccountData = Invoke-Sqlcmd @intermediateDBparams -Query $newAccountSql
$newAccountData 
# Create-ADUserObject
# Create-HomeDir
# Wait-DirSync
# Assign-License
# Wait-MailBoxCreate
# Update-EscapeEmail
# Wait-GSuiteSync
# Update-PW
# Format-Emails
# Send-Emails

cleanUp
Show-TestRun