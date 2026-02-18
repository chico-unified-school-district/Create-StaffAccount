[cmdletbinding()]
<#
.SYNOPSIS
Creates and provisions staff accounts by monitoring an intermediate database table for new employees
and performing AD, file-server and cloud mailbox related actions.

.DESCRIPTION
Create-StaffAccount.ps1 polls an intermediate SQL table for new employee rows (rows where EmailWork is NULL). For
each new employee it will collect and normalize name data, allocate or verify an EmployeeID, generate a SamAccountName
(SAM ID) and temporary password, create or update an Active Directory user object, create a home directory on the
appropriate file server, and perform mailbox and forwarding configuration with Exchange Online and GSuite (via GAM).

This script is designed to be run periodically (for example once each morning and then at intervals throughout the
day until a configured stop time). It supports a -WhatIf switch to preview actions without making changes.

.PARAMETER DomainControllers
An array of Active Directory Domain Controller hostnames to use when connecting to AD. Passed through to
Connect-ADSession / AD cmdlets.

.PARAMETER ActiveDirectoryCredential
A PSCredential object used to authenticate to Active Directory when creating or updating user objects.

.PARAMETER DefaultStaffGroups
An array of AD group names that new staff accounts should be added to by default. The script will also add licensing
groups based on HR fields and site-specific groups from the lookup table.

.PARAMETER O365Credential
A PSCredential used to connect to ExchangeOnline (Connect-ExchangeOnline) to manage mailboxes and retention/forwarding.

.PARAMETER FileServerCredential
Credentials used when creating home directories on remote file servers (New-StaffHomeDir uses these credentials).

.PARAMETER FullAccess
Array of group or user identifiers that should receive full access ACLs on newly created home directories.

.PARAMETER EmployeeServer
The SQL Server (instance name/address) hosting the authoritative employee database.

.PARAMETER EmployeeDatabase
The database name on the EmployeeServer that contains the authoritative data.

.PARAMETER EmployeeTable
The table name in the employee database containing employee rows (not frequently used directly — connection helper uses
these values to read authoritative data if necessary).

.PARAMETER EmployeeCredential
PSCredential used to authenticate to the EmployeeServer when querying employee data.

.PARAMETER IntermediateSqlServer
The SQL Server instance that contains the intermediate table the script watches for new accounts.

.PARAMETER IntermediateDatabase
The database on the intermediate SQL server that holds the new-accounts table.

.PARAMETER NewAccountsTable
The table (in the intermediate DB) to query for rows needing provisioning. The script expects the table to contain a
row per person and columns such as id, nameFirst, nameLast, nameMiddle, siteCode, jobType, empId, emailWork and tempPw.

.PARAMETER IntermediateCredential
PSCredential used to authenticate to the intermediate SQL instance when reading and updating the new-accounts table.

.PARAMETER TargetOrgUnit
The distinguished name (AD path) of the target OU where new user objects should be created if the script creates users.

.PARAMETER Organization
String used to populate the Company attribute on AD objects.

.PARAMETER Domain1
Primary domain suffix used to form the organizational email address (example: '@chicousd.org').

.PARAMETER Domain2
Secondary domain suffix used to form the GSuite address (example: '@chicousd.net').

.PARAMETER WhatIf
Switch that enables a dry-run mode; when supplied the script will display planned actions but will not perform
changes that modify AD, SQL, Exchange Online, or the file server. Many internal calls pass -WhatIf to cmdlets when this
switch is set.

.EXAMPLE
.
PS> .\Create-StaffAccount.ps1 -IntermediateSqlServer 'sql01' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (Get-Credential) -DomainControllers 'dc01','dc02' -ActiveDirectoryCredential (Get-Credential) -O365Credential (Get-Credential) -FileServerCredential (Get-Credential)

Runs in normal (non-WhatIf) mode and will attempt to provision any rows returned by the intermediate query.

.EXAMPLE
# Dry-run: preview changes without making modifications to AD, Exchange or file servers
PS> .\Create-StaffAccount.ps1 -IntermediateSqlServer 'sql01' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (Get-Credential) -DomainControllers 'dc01' -ActiveDirectoryCredential (Get-Credential) -O365Credential (Get-Credential) -FileServerCredential (Get-Credential) -WhatIf

.EXAMPLE
# Scheduled/CI invocation example (PowerShell scheduled task or Jenkins):
PS> powershell -NoProfile -ExecutionPolicy Bypass -File "C:\path\to\Create-StaffAccount.ps1" -IntermediateSqlServer 'sql01' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (Get-Credential) -DomainControllers 'dc01' -ActiveDirectoryCredential (Get-Credential) -O365Credential (Get-Credential) -FileServerCredential (Get-Credential)

.NOTES
- The script requires the following modules to be available: ExchangeOnlineManagement, dbatools, CommonScriptFunctions.
- External helper scripts in the `lib\` directory are used. Minimal reference docs for those functions are included
    below so an operator can understand their contract without opening each file.
- The script optionally invokes `gam.exe` to query or manage GSuite accounts. Ensure `bin\gam.exe` is valid and
    configured if GSuite actions are required.

.EXTERNAL FUNCTION REFERENCE
Below are short, comment-style contracts for external functions called by this script. These are intended as a quick
reference and are not substitutes for the implementations in `lib\`.

New-ADUserObject
 - Inputs (via pipeline): a PSCustomObject with properties: name, fn, ln, mi, new (original row), samid, empid, emailWork,
     pw1 (plain-text temporary password), targetOU, company, site (site metadata), new.siteCode, new.jobType and WhatIf uses.
 - Output: Outputs the same pipeline object with AD object created; may write verbose/info messages.
 - Side effects: Creates AD user (New-ADUser) and updates attributes such as proxyAddresses, targetAddress and optionally
     AccountExpirationDate and middleName.

New-StaffHomeDir
 - Parameters: ($cred, [string[]]$full) where $cred is PSCredential used to mount a remote share and $full is an array
     of groups/users to grant ACLs to.
 - Inputs (via pipeline): a PSCustomObject including samid and site.FileServer where home directories should be created.
 - Behavior: Creates a directory \<FileServer>\User\<samid>\Documents, sets ACLs with ICACLS and grants access to the
     supplied $full accounts and the user.

New-PassPhrase
 - No parameters. Returns a string: a generated passphrase used as a temporary password (13-16 chars) composed from
     dictionaries under `lib\`.

New-SamID
 - Parameters: -First <string> -Middle <string> -Last <string>
 - Returns: a candidate SamAccountName string generated from name parts. Respects AD uniqueness by checking for existing
     SAMs and proxyaddresses.

Format-Name
 - Parameters: <string>
 - Returns: a normalized, title-cased name string with limited allowed characters (letters, apostrophe and spaces).

#>
[cmdletbinding()]
param(
 [Alias('DCs')]
 [string[]]$DomainControllers,
 [Alias('ADCred')][System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [string[]]$DefaultStaffGroups,
 [Alias('MSCred')] [System.Management.Automation.PSCredential]$O365Credential,
 [string]$O365Domain,
 [string]$ExchangeServer,
 [Alias('ExchCred')] [System.Management.Automation.PSCredential]$ExchangeCredential,
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
  if ( $_.ad -and $_.ad.WhenCreated -lt ((Get-Date).AddDays(-180)) ) {
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
  $ErrorActionPreference = 'Continue'
  ($gUser = & $gam print users query "email: $($_.gSuite)" | ConvertFrom-Csv)*>$null
  $ErrorActionPreference = 'Stop'
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

function Connect-LocalExchangeServer {
 param (
  [string]$Server,
  [System.Management.Automation.PSCredential]$Credential
 )
 process {
  Write-Host ('{0}' -f $MyInvocation.MyCommand.Name) -F Green
  $sessionParams = @{
   ConfigurationName = 'Microsoft.Exchange'
   ConnectionUri     = "http://$Server/PowerShell/"
   Authentication    = 'Kerberos'
   Credential        = $Credential
   ErrorAction       = 'Stop'
  }
  $session = New-PSSession @sessionParams
  Import-PSSession $session -CommandName Get-RemoteMailbox, Enable-RemoteMailbox
 }
}

function Convert-FromSharedMailbox {
 process {
  $msgData = $MyInvocation.MyCommand.Name, $_.ad.EmployeeID, $_.ad.Mail
  $params = @{
   Filter = "UserPrincipalName -eq '{0}'" -f $_.ad.UserPrincipalName
  }
  if ((Get-Mailbox @params).IsShared -ne $true) { return $_ } # Shared already False
  Set-Mailbox -Identity $_.ad.UserPrincipalName -Type Regular -WhatIf:$WhatIf # Convert to Shared Mailbox
  if (!$WhatIf) { Start-Sleep -Seconds 10 } # Allow time for mailbox type change to propagate
  if ((Get-Mailbox @params).IsShared -eq $true) {
   Write-Warning ('{0},{1},{2}, Mailbox still in Shared state.' -f $msgData)
   if (!$WhatIf) { return } # Skip further processing
  }
  Write-Host ('{0},{1},{2},Mailbox converted from Shared to Regular' -f $msgData) -F Green
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

function Enable-ExchRemoteMailbox ($domain) {
 process {
  $remoteMailBox = Get-RemoteMailbox -Filter "PrimarySmtpAddress -eq '$($_.emailWork )'"
  if ($remoteMailBox) { return $_ } # Already enabled
  Write-Host ('{0}' -f $MyInvocation.MyCommand.Name)
  $params = @{
   Identity             = $_.name
   RemoteRoutingAddress = $_.samid + $domain
   WhatIf               = $WhatIf
  }
  Enable-RemoteMailbox @params
  $_
 }
}

function Enable-EmailRetention {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F DarkYellow
  $_.mailbox | Set-Mailbox -RetainDeletedItemsFor 30 -WhatIf:$WhatIf
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

# function New-HomeDir ($fsUser, $full) {
#  process {
#   # Begin Switch to 'New HDrive Location' AD group
#   $_ | New-StaffHomeDir $fsUser $full
#   $_
#  }
# }

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

function Set-AdExpirationDate {
 begin {
  $shortTermTypes = 'student teacher', 'coach', 'volunteer', 'student worker', 'intern'
 }
 process {
  if ( ($_.empid -match '^\d{7,}$') -or ($shortTermTypes -match $_.new.jobType) ) {
   # ♥ If current month is greater than 6 (June), set AccountExpirationDate to after the end of the current school term. ♥
   $year = '{0:yyyy}' -f $(if ([int](Get-Date -f MM) -gt 6) { (Get-Date).AddYears(1) } else { Get-Date })
   $accountExpirationDate = Get-Date "July 30 $year"
   Write-Host ('{0},{1},Setting Account Expiration to: {2}' -f $MyInvocation.MyCommand.Name, $samid, $accountExpirationDate) -F DarkCyan
   if (!$WhatIf) { Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $AccountExpirationDate }
  }
  $_
 }
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
# ===================================================== main =====================================================
Import-Module -Name ExchangeOnlineManagement -Cmdlet Connect-ExchangeOnline, Get-EXOMailBox, Disconnect-ExchangeOnline, Set-Mailbox
Import-Module -Name dbatools -Cmdlet Set-DbatoolsConfig, Invoke-DbaQuery, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module -Name CommonScriptFunctions -Cmdlet Show-TestRun, New-SqlOperation, Clear-SessionData, New-RandomPassword

Show-BlockInfo Main
if ($WhatIf) { Show-TestRun }

# Imported Functions
. .\lib\New-ADUserObject.ps1
. .\lib\New-PassPhrase.ps1
# . .\lib\New-StaffHomeDir.ps1

$gam = 'C:\GAM7\gam.exe'

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
  Connect-LocalExchangeServer -Server $ExchangeServer -Credential $ExchangeCredential
  Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdlets -Cred $ActiveDirectoryCredential
 }

 $accountObjs = $newAccounts |
  Format-Object |
   Add-EmpId |
    Add-ADData |
     Add-ADName |
      Add-ADSamId |
       Add-Info
 $accountObjs |
  Update-IntDBEmpId $intSQLInstance $NewAccountsTable |
   Add-EmailAddresses $Domain1 $Domain2 |
    Add-AccountStatus |
     Add-SiteData $lookupTable |
      New-UserADObj |
       Set-AdExpirationDate |
        Update-ADGroups |
         Enable-ExchRemoteMailbox $O365Domain |
          Confirm-OrgEmail |
           Confirm-GSuite |
            Update-ADPW |
             Convert-FromSharedMailbox |
              Enable-EmailForwarding |
               Enable-EmailRetention |
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