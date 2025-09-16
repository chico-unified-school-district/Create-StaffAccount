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
  $_.info = $_.empId, $_.samid, ($_.fn + ' ' + $_.ln) -join ','
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

function Add-EmpId {
 process {
  $_.empId = if ($_.new.empid) { $_.new.empID } else { Get-Random -Min 1000000 -Max 10000000 }
  $_
 }
}

function Add-GSuiteAddress ($gsuiteDomain) {
 process {
  $_.gsuite = $_.samId + $gsuiteDomain
  $_
 }
}

function Add-O365Address ($o365Domain) {
 process {
  $_.emailWork = $_.samid + $o365Domain
  $_
 }
}

function Add-SiteData {
 begin {
  $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json
 }
 process {
  # Skip blank or null site codes
  $sc = $_.new.siteCode
  $sd = $_.new.siteDescr
  $_.site = $lookupTable.Where({ [int]$_.SiteCode -eq [int]$sc })
  if (!$_.site) { $_.site = $lookupTable.Where({ $_.SiteDescr -eq $sd }) }
  if (!$_.site) { Write-Host ('{0},{1},{2},No Site match.' -f $MyInvocation.MyCommand.Name, $_.empId, $sc) -f Magenta }
  $_
 }
}

function Complete-Processing {
 process {
  if (!$_.new) { return }
  Write-Verbose ($MyInvocation.MyCommand.Name, $_ | Out-String)
  if (!$_.gsuiteReady -or !$_.emailWorkReady) { return }
  $msg = $MyInvocation.MyCommand.Name, $_.info, (Get-Date -Format G), ('=' * (20 - $str.length))
  Write-Host ('{0},[{1}],{2} <{3}' -f $msg) -F Cyan
 }
}

function Confirm-GSuite ($dbParams, $table) {
 begin {
  $gam = '.\bin\gam.exe'
  $updateSql = "UPDATE $table SET gsuite = @gmail WHERE id = @id"
 }
 process {
  if (!$_.new) { return $_ }
  if ($_.new.gsuite -eq $_.gsuite) { $_.gsuiteReady = $true ; return $_ } # If gsuite already in db then it was synced successfully.

  Write-Verbose ('{0},[{1}],Getting Gsuite User' -f $MyInvocation.MyCommand.Name, $_.gsuite)
  ($guser = & $gam print users query "email: $($_.gsuite)" | ConvertFrom-Csv)*>$null
  Write-Verbose ($guser | Out-String )
  if ($guser.PrimaryEmail -ne $_.gsuite) {
   Write-Verbose ('{0},[{1}],Gsuite User NOT Found' -f $MyInvocation.MyCommand.Name, $_.gsuite)
   return
  }
  Write-Host ('{0},[{1}],Gsuite User Found' -f $MyInvocation.MyCommand.Name, $_.gsuite) -F Blue
  $_.gsuiteReady = $true
  # Update the intDB once the gsuite account is synced to the cloud
  $sqlVars = "gmail=$($_.gsuite)", "id=$($_.new.id)"
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
  if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
  $_
 }
}

function Confirm-OrgEmail ($dbParams, $table, $exchCred) {
 begin {
  $updateSql = "UPDATE $table SET emailWork = @mail WHERE id = @id"
 }
 process {
  if (!$_.ad) { return $_ }
  Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.emailWork)
  $upn = $_.ad.UserPrincipalName
  $params = @{
   filter      = "UserPrincipalName -eq '$upn'"
   ErrorAction = 'SilentlyContinue'
  }
  $mailBox = Get-EXOMailbox @params
  # Stop processing until mailbox exists in the cloud
  if ($mailBox.UserPrincipalName -ne $_.ad.UserPrincipalName) { return }
  Write-Host ('{0},[{1}],Mailbox found!' -f $MyInvocation.MyCommand.Name, $_.emailWork) -F Blue
  $_.emailWorkReady = $true
  $sqlVars = "mail=$($_.emailWork)", "id=$($_.new.id)"
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
  <# Once the intDB has the emailWork entered no more subsequent runs will occur.
      An associated Laserfiche Workflow will then handle the next steps #>
  if (!$WhatIf -and $_.gsuiteReady) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
  $_
 }
}

function Format-UserObject {
 begin {
  . .\lib\New-PassPhrase.ps1
 }
 process {
  [PSCustomObject]@{
   ad             = $null
   new            = $_
   site           = $null
   empId          = $_.empId
   fn             = $fn
   ln             = $ln
   mi             = $mn
   name           = $newName
   samid          = $samId
   emailWork      = $null
   emailWorkReady = $null
   gsuite         = $null
   gsuiteReady    = $null
   company        = $Organization
   pw1            = New-PassPhrase
   pw2            = New-PassPhrase
   targetOU       = $TargetOrgUnit
   info           = $null
   status         = $null
  }
 }
}

function New-HomeDir ($dbParams, $table, $cred, $full) {
 begin {
  . .\lib\New-StaffHomeDir.ps1
 }
 process {
  $_ | New-StaffHomeDir $cred $full
  $_
 }
}

function New-UserADObj ($dbParams, $table) {
 begin {
  . .\lib\New-ADUserObject.ps1
  $updateSql = "UPDATE $table SET samAccountName = @samid, empId = @empId WHERE id = @id;"
 }
 process {
  if ($_.ad) { return $_ }
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F Green
  $_ | New-ADUserObject
  $sqlVars = "samid=$($_.samid)", "empId=$($_.empId)" , "id=$($_.new.id)"
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
  if (!$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
  $_
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

function Update-ADGroups ($dbParams, $table) {
 begin {
  $A5 = 'Office365_A5_Faculty' # Microsoft 365 License for admin and office staff
  $A1 = 'Office365_A1_Faculty' # Microsoft 365 License for general staff
 }
 process {
  if (!$_.ad) { return $_ }
  if ($_.ad.memberof) { return $_ } # if groups present then, likely already ran
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.samid)
  $licenseGroup = if ($_.BargUnitId -eq 'CUMA') { $A5 } else { $A1 }
  # Add user to various groups
  $groups = $DefaultStaffGroups + $licenseGroup
  if ( $_.site.Groups ) { $groups += $_.site.Groups.Split(',') }

  $msg = $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, ($groups -join ',')
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F Blue

  if ( -not$WhatIf ) { Add-ADPrincipalGroupMembership -Identity $_.ad.ObjectGUID -MemberOf $groups }
  $_
 }
}

function Update-ADPW ($dbParams, $table) {
 begin {
  $updateSql = "UPDATE $table SET tempPw = @pw WHERE id = @id"
 }
 process {
  Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info )
  if (!$_.gsuiteReady) { return $_ }
  if (($_.status -eq 'Account Already Exists')) { $_.pw2 = $_.status } else {
   Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info ) -F DarkGreen
   $securePw = ConvertTo-SecureString -String $_.pw2 -AsPlainText -Force
   # Updating the password activates the GSuite account
   Set-ADAccountPassword -Identity $_.ad.ObjectGUID -NewPassword $securePw -Confirm:$false -WhatIf:$WhatIf
  }
  $sqlVars = "pw=$($_.pw2)", "id=$($_.new.id)"
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
  if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
  $_
 }
}

function Update-EmpEmailWork ($dbParams, $table) {
 begin {
  $updateSql = "UPDATE $table SET EmailWork = @mail WHERE EmpID = @empId"
 }
 process {
  if (!$_.emailWorkReady) { return $_ }
  $sqlVars = "mail=$($_.emailWork)", "empId=$($_.empId)"
  Write-Host ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ',') ) -F Cyan
  if (!$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
  $_
 }
}

function Update-IntDB ($dbParams, $table) {
 begin { $updateSql = "UPDATE $table SET sourceSystem = @sys, dts = CURRENT_TIMESTAMP WHERE id = @id;" }
 process {
  $sqlVars = "sys=$ENV:COMPUTERNAME", "id=$($_.new.id)"
  if ($WhatIf) { return $_ }
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
  New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars
  $_
 }
}

Import-Module -Name ExchangeOnlineManagement -Cmdlet 'Connect-ExchangeOnline', 'Get-EXOMailBox', 'Disconnect-ExchangeOnline'
Import-Module -Name dbatools -Cmdlet 'Set-DbatoolsConfig', 'Invoke-DbaQuery'
Import-Module -Name CommonScriptFunctions

Show-BlockInfo Main
# Imported Functions
. .\lib\New-StaffHomeDir.ps1

if ($WhatIf) { Show-TestRun }
Disconnect-ExchangeOnline -Confirm:$false

$empBParams = @{
 Server     = $EmployeeServer
 Database   = $EmployeeDatabase
 Credential = $EmployeeCredential
}

$intDBparams = @{
 Server     = $IntermediateSqlServer
 Database   = $IntermediateDatabase
 Credential = $IntermediateCredential
}

$cmdlets = 'Get-ADUser', 'New-ADuser',
'Set-ADUser', 'Add-ADPrincipalGroupMembership' , 'Set-ADAccountPassword'

$newAccountSql = 'SELECT * FROM {0} WHERE emailWork IS NULL' -f $NewAccountsTable

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date '5:00pm' }
$delay = if ($WhatIf) { 0 } else { 180 }

do {
 $objs = New-SqlOperation @intDBparams -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv
 if ($objs) {
  Connect-ExchangeOnline -Credential $O365Credential -ShowBanner:$false
  Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdlets -Cred $ActiveDirectoryCredential
 }

 $objs |
  Format-UserObject |
   Add-EmpId |
    Add-ADData |
     Add-ADName |
      Add-ADSamId |
       Add-Info |
        Add-O365Address $Domain1 |
         Add-GSuiteAddress $Domain2 |
          Add-AccountStatus |
           Add-SiteData |
            New-UserADObj $intDBparams $NewAccountsTable |
             Add-ADData |
              Update-ADGroups $intDBparams $NewAccountsTable |
               New-HomeDir $intDBparams $NewAccountsTable $FileServerCredential $FullAccess |
                Confirm-GSuite $intDBparams $NewAccountsTable |
                 Update-ADPW $intDBparams $NewAccountsTable |
                  Confirm-OrgEmail $intDBparams $NewAccountsTable $O365Credential |
                   Update-EmpEmailWork $empBParams $EmployeeTable |
                    Update-IntDB $intDBparams $NewAccountsTable |
                     Complete-Processing

 Clear-SessionData
 Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
 Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay)))
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )
if (!$WhatIf) { Remove-TmpEXOs }
if ($WhatIf) { Show-TestRun }