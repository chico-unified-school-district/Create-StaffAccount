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
 [Alias('ADCred')][System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [string[]]$DefaultStaffGroups,
 [Alias('MSCred')] [System.Management.Automation.PSCredential]$O365Credential,
 [Alias('FSCred')] [System.Management.Automation.PSCredential]$FileServerCredential,
 [string[]]$FullAccess,
 [string]$EmployeeServer,
 [string]$EmployeeDatabase,
 [string]$EmployeeTable,
 [System.Management.Automation.PSCredential]$EmployeeCredential,
 [Alias('IntServer')][string]$IntermediateSqlServer,
 [Alias('IntDB')][string]$IntermediateDatabase,
 [Alias('Table')][string]$NewAccountsTable,
 [Alias('IntCred')][System.Management.Automation.PSCredential]$IntermediateCredential,
 [Alias('OU')][string]$TargetOrgUnit,
 [string]$Organization,
 [string]$Domain1,
 [string]$Domain2,
 [Alias('wi')][switch]$WhatIf
)

function Add-AccountStatus {
 process {
  if ( $_.ad -and $_.ad.WhenCreated -lt ((Get-Date).AddDays(-3)) ) {
   $_.status = 'Account Already Exists'
  }
  $_
 }
}

function Add-ADData {
 process {
  $filter = "EmployeeID -eq '{0}'" -f $_.empId
  $_.ad = Get-ADUser -Filter $filter -Properties *
  $_
 }
}

function Add-Info {
 process {
  # $_.info = $_.empId, $_.samid, ($_.fn + ' ' + $_.ln) -join ','
  $_.info = '[{0},{1},{2}]' -f $_.empId, $_.samid, ($_.fn + ' ' + $_.ln)
  $_
 }
}

function Add-ADName {
 begin {
  . .\lib\Format-Name.ps1
  . .\lib\New-Name.ps1
 }
 process {
  $_.fn = Format-Name $_.new.nameFirst
  $_.ln = Format-Name $_.new.nameLast
  $_.mi = Format-Name $_.new.nameMiddle
  $_.name = if ($_.ad) { $_.ad.name } else { New-Name -first $_.fn -middle $_.mi -last $_.ln }
  $_
 }
}

function Add-ADSamId {
 begin {
  . .\lib\New-SamID.ps1
 }
 process {
  $_.samid = if ($_.ad) { $_.ad.SamAccountName } else { New-SamID -F $_.fn -M $_.mi -L $_.ln }
  $_
 }
}

function Add-EmailAddresses ($dom1, $dom2) {
 process {
  $_.emailWork = $_.samid + $dom1
  $_.gSuite = $_.samid + $dom2
  $_
 }
}

function Add-EmpId {
 begin {
  function randomEmpId {
   do { $id = Get-Random -Min 10000000 -Max 99999999; Write-Verbose $id }
   until ( !(Get-ADUser -Filter " EmployeeId -eq '$id'"))
   $id
  }
 }
 process {
  $_.empId = if ($_.new.empId -and ($_.new.empId -ne 0)) { $_.new.empId; $_.empIdExists = $true } else { randomEmpId }
  if ($_.empId -lt 1) {
   # Remove by Feb 2026 if no hits
   Write-Host ($_.new | Out-String)
   Write-Error ('{0}, BAD Employee ID' -f $MyInvocation.MyCommand.Name)
   return
  }
  $_
 }
}

function Add-SiteData ($tableData) {
 process {
  # Skip blank or null site codes
  $sc = $_.new.siteCode
  $sd = $_.new.siteDescr
  $_.site = $tableData.Where({ [int]$_.SiteCode -eq [int]$sc })
  if (!$_.site) { $_.site = $tableData.Where({ $_.SiteDescr -eq $sd }) }
  if (!$_.site) { Write-Verbose ('{0},{1},No Site match: [{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $sc) }
  $_
 }
}

function Complete-Processing {
 process {
  if (!$_.new) { return $_ }
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, ($_ | Out-String))
  $msg = $MyInvocation.MyCommand.Name, $_.info, (Get-Date -Format G), ('+' * (20 - $str.length))
  Write-Host ('{0},{1},{2} <{3}' -f $msg) -F DarkMagenta
 }
}

function Confirm-GSuite {
 process {
  ($gUser = & $gam print users query "email: $($_.gSuite)" | ConvertFrom-Csv)*>$null
  if ($gUser.PrimaryEmail -ne $_.gSuite) { return }
  Write-Host ('{0},[{1}],Gsuite Found!' -f $MyInvocation.MyCommand.Name, $_.gSuite) -F Green
  $_
 }
}

function Confirm-OrgEmail {
 process {
  Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.emailWork)
  $upn = $_.ad.UserPrincipalName
  $params = @{
   filter      = "UserPrincipalName -eq '$upn'"
   ErrorAction = 'SilentlyContinue'
  }
  $mailBox = Get-EXOMailbox @params
  # Stop processing until mailbox exists in the cloud
  if ($mailBox.UserPrincipalName -ne $_.ad.UserPrincipalName) { return }
  Write-Host ('{0},[{1}],Mailbox found!' -f $MyInvocation.MyCommand.Name, $_.emailWork) -F Green
  $_.mailbox = $mailBox
  $_
 }
}

function Enable-EmailForwarding {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F DarkYellow
  $_.mailbox | Set-Mailbox -ForwardingSmtpAddress $_.gSuite -DeliverToMailboxAndForward $true -WhatIf:$WhatIf
  $_
 }
}

function Format-Object {
 process {
  [PSCustomObject]@{
   ad          = $null
   company     = $Organization
   empId       = $_.empId
   fn          = $fn
   emailWork   = $null
   empIdExists = $null
   mailbox     = $null
   gSuite      = $null
   info        = $null
   ln          = $ln
   mi          = $mn
   name        = $newName
   new         = $_
   pw1         = New-PassPhrase
   pw2         = New-PassPhrase
   pwReset     = $null
   samid       = $samId
   site        = $null
   status      = $null
   targetOU    = $TargetOrgUnit
  }
 }
}

function New-HomeDir ($fsUser, $full) {
 process {
  # Begin Switch to 'New HDrive Location' AD group
  $_ | New-StaffHomeDir $fsUser $full
  $_
 }
}

function New-UserADObj {
 process {
  if ($_.ad) { return $_ }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Yellow
  $_ | New-ADUserObject
 }
}

function Remove-TmpEXOs {
 Start-Sleep 10
 Write-Host ('{0}' -f $MyInvocation.MyCommand.Name)
 $cutOff = (Get-Date).AddDays(-1)
 $myDir = Get-Location
 Set-Location $ENV:Temp
 $tmpFolders = Get-ChildItem -Filter tmpEXO* -Recurse -Force
 $tmpFolders | Where-Object { $_.LastWriteTime -lt $cutOff } |
  Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
 Set-Location $myDir
}

function Update-ADGroups {
 begin {
  $A5 = 'Office365_A5_Faculty' # Microsoft 365 License for admin and office staff
  $A1 = 'Office365_A1_Faculty' # Microsoft 365 License for general staff
 }
 process {
  if ($_.ad.memberof) { return $_ } # if groups present then, likely already ran
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.samid)
  $licenseGroup = if ($_.emp.BargUnitId -eq 'CUMA') { $A5 } else { $A1 } # HR rarely has this data correct, so A1 is default.
  # Add user to various groups
  $groups = $DefaultStaffGroups + $licenseGroup
  if ( $_.site.Groups ) { $groups += $_.site.Groups.Split(',') }

  $msg = $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, ($groups -join ',')
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F Yellow
  Add-ADPrincipalGroupMembership -Identity $_.ad.ObjectGUID -MemberOf $groups -WhatIf:$WhatIf
  $_
 }
}

function Update-ADPW {
 process {
  if (($_.status -eq 'Account Already Exists')) { $_.pw2 = $_.status } else {
   Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info ) -F Yellow
   $securePw = ConvertTo-SecureString -String $_.pw2 -AsPlainText -Force
   # Updating the password activates the GSuite account
   Set-ADAccountPassword -Identity $_.ad.ObjectGUID -NewPassword $securePw -Confirm:$false -WhatIf:$WhatIf
   $_.pwReset = $true
  }
  $_
 }
}

function Update-EmpEmailWork ($sqlInstance, $table) {
 begin {
  $sql = "UPDATE $table SET EmailWork = @mail WHERE empId = @empId"
 }
 process {
  $sqlVars = @{mail = $_.emailWork; empId = $_.empId }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}

function Update-IntDB ($sqlInstance, $table) {
 begin {
  $sql = "UPDATE $table SET
  empId = @empId,
  samAccountName = @samid,
  emailWork = @mail,
  gSuite = @gmail,
  tempPw = @pw
  WHERE id = @id;"
 }
 process {
  $sqlVars = @{id = $_.new.id; empId = $_.empId; samid = $_.ad.SamAccountName; mail = $_.emailWork; gmail = $_.gSuite; pw = $_.pw2 }
  Write-Verbose ('{0},{1},[{2}],[{3}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, ($sqlVars.Values -join ','))
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}

function Update-IntDBEmpId ($sqlInstance, $table) {
 begin { $sql = "UPDATE $table SET empId = @empId WHERE id = @id;" }
 process {
  if ($_.empIdExists) { return $_ }
  $sqlVars = @{empId = $_.empId; id = $_.new.id }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}


Import-Module -Name ExchangeOnlineManagement -Cmdlet Connect-ExchangeOnline, Get-EXOMailBox, Disconnect-ExchangeOnline, Set-Mailbox
Import-Module -Name dbatools -Cmdlet Set-DbatoolsConfig, Invoke-DbaQuery, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module -Name CommonScriptFunctions -Cmdlet Show-TestRun, New-SqlOperation, Clear-SessionData, New-RandomPassword

Show-BlockInfo Main
if ($WhatIf) { Show-TestRun }

# Imported Functions
. .\lib\New-ADUserObject.ps1
. .\lib\New-PassPhrase.ps1
. .\lib\New-StaffHomeDir.ps1

$gam = '.\bin\gam.exe'

Disconnect-ExchangeOnline -Confirm:$false

$empSQLInstance = Connect-DbaInstance -SqlInstance $EmployeeServer -Database $EmployeeDatabase -SqlCredential $EmployeeCredential
$intSQLInstance = Connect-DbaInstance -SqlInstance $IntermediateSqlServer -Database $IntermediateDatabase -SqlCredential $IntermediateCredential

$cmdlets = 'Get-ADUser', 'New-ADuser', 'Set-ADUser', 'Add-ADPrincipalGroupMembership' , 'Set-ADAccountPassword'

$lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json

$newAccountSql = 'SELECT * FROM {0} WHERE emailWork IS NULL' -f $NewAccountsTable

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date '5:00pm' }
$delay = if ($WhatIf) { 0 } else { 180 }

do {
 $newAccounts = Invoke-DbaQuery -SqlInstance $intSQLInstance -Query $newAccountSql |
  ConvertTo-Csv | ConvertFrom-Csv
 if ($newAccounts) {
  Connect-ExchangeOnline -Credential $O365Credential -ShowBanner:$false
  Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdlets -Cred $ActiveDirectoryCredential
 }

 $newAccounts |
  Format-Object |
   Add-EmpId |
    Add-ADData |
     Add-ADName |
      Add-ADSamId |
       Add-Info |
        Update-IntDBEmpId $intSQLInstance $NewAccountsTable |
         Add-EmailAddresses $Domain1 $Domain2 |
          Add-AccountStatus |
           Add-SiteData $lookupTable |
            New-UserADObj |
             Update-ADGroups |
              New-HomeDir $FileServerCredential $FullAccess |
               Confirm-OrgEmail |
                Confirm-GSuite |
                 Update-ADPW |
                  Enable-EmailForwarding |
                   Update-IntDB $intSQLInstance $NewAccountsTable |
                    Update-EmpEmailWork $empSQLInstance $EmployeeTable |
                     Complete-Processing

 Clear-SessionData
 Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
 if (!$WhatIf) { Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay))) }
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )

if (!$WhatIf) { Remove-TmpEXOs }
if ($WhatIf) { Show-TestRun }