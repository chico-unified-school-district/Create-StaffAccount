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
  [string[]]$DefaultStaffGroups,
  [Alias('MSCred')]
  [System.Management.Automation.PSCredential]$O365Credential,
  # [Alias('MgToken')]
  # [System.Management.Automation.PSCredential]$MGGraphToken,
  [Alias('FSCred')]
  [System.Management.Automation.PSCredential]$FileServerCredential,
  [string[]]$FullAccess,
  [string]$EmployeeServer,
  [string]$EmployeeDatabase,
  [string]$EmployeeTable,
  [System.Management.Automation.PSCredential]$EmployeeCredential,
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
  [string]$Domain1,
  [string]$Domain2,
  [Alias('wi')]
  [switch]$WhatIf
)
function Complete-Processing {
  begin { $i = 0 }
  process {
    if ($_) { $i++ }
    Write-Host ('{0},[{1}] <{2}' -f $MyInvocation.MyCommand.Name, $_.info, ('=' * (20 - $str.length))) -F Cyan
  }
  end {
    if ($i -eq 0 ) { return }
    Write-Verbose ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, $i)
  }
}

function Confirm-NetEmail ($dbParams, $table) {
  begin {
    $gam = '.\bin\gam.exe'
    $updateSql = "UPDATE $table SET gsuite = @gmail WHERE id = @id"
  }
  process {
    if ($_.db_gsuite -eq $_.gsuite) { $_ ; return } # Skip following logic if true

    Write-Verbose ('{0},[{1}],Getting Gsuite User' -f $MyInvocation.MyCommand.Name, $_.gsuite)
    $guser = $null
    ($guser = & $gam print users query "email: $($_.gsuite)" | ConvertFrom-CSV)*>$null
    Write-Verbose ($guser | Out-String )
    if ($guser.PrimaryEmail -ne $_.gsuite) {
      Write-Verbose ('{0},[{1}],Gsuite User NOT Found' -f $MyInvocation.MyCommand.Name, $_.gsuite)
      return
    }
    Write-Host ('{0},[{1}],Gsuite User Found' -f $MyInvocation.MyCommand.Name, $_.gsuite) -F Blue

    # Update the intDB once the gsuite account is synced to the cloud
    $sqlVars = "gmail=$($_.gsuite)", "id=$($_.id)"
    Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
    $_
  }
}

function Confirm-OrgEmail ($dbParams, $table, $exchCred) {
  begin {
    $updateSql = "UPDATE $table SET emailWork = @mail WHERE id = @id"
    function New-ExchangeConnection ([int]$attempts, $myCred) {
      process {
        $exchStatus = $null
        try { Connect-ExchangeOnline -Credential $myCred -ShowBanner:$false -ErrorAction SilentlyContinue -ErrorVariable exchStatus }
        catch {
          if ($exchStatus -and ($attempts -gt 0)) {
            $attempts--
            Write-Host ('{0},Error. Trying again in 30 seconds...,{1} Attempts remaining.' -f $MyInvocation.MyCommand.Name, $attempts) -F Yellow
            if (!$WhatIf) { Start-Sleep 30 }
            New-ExchangeConnection $attempts $myCred
          }
          else {
            Write-Host ('{0},Failed.' -f $MyInvocation.MyCommand.Name) -F Red
            Exit
          }
        }
      }
    }
    if (!$WhatIf) { New-ExchangeConnection 10 $exchCred }
  }
  process {
    Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.emailWork)
    $mailBox = Get-EXOMailbox -Identity $_.emailWork -ErrorAction SilentlyContinue
    # Update the intDB once the Outlook account is synced to the cloud
    if (-not$mailBox) { return }

    Write-Host ('{0},[{1}],Mailbox found!' -f $MyInvocation.MyCommand.Name, $_.emailWork) -F Blue

    $sqlVars = "mail=$($_.emailWork)", "id=$($_.id)"
    Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
    <# Once the intDB has the emailWork entered no more subsequent runs will occur.
      An associated Laserfiche Workflow will then handle the next steps #>
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
    $_
  }
  end {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
  }
}

function Get-CreationTime {
  begin {
  }
  process {
    $filter = "EmployeeID -eq '{0}'" -f $_.empId
    $adObj = Get-ADUser -Filter $filter -Properties WhenCreated
    $msg = $MyInvocation.MyCommand.Name, $_.info , (((Get-Date) - $adObj.WhenCreated).totalseconds / 60)
    Write-Host ('{0},[{1}],Creation Time (mins): {2:n2}' -f $msg) -F Cyan
    $_
  }
}

function New-UserADObj ($dbParams, $table) {
  begin {
    . .\lib\New-ADUserObject.ps1
    $updateSql = "UPDATE $table SET samAccountName = @samid, empId = @empId WHERE id = @id;"
  }
  process {
    # Skip user new creation if AD Obj already present
    if ($_.adObj) { $_ ; return }

    Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F Green

    $newObj = $_ | New-ADUserObject

    $filter = "EmployeeID -eq '{0}'" -f $_.empId
    $adObj = Get-ADUser -Filter $filter -Properties *

    if ($adObj) { $_.adObj = $adObj }

    if ( !$newObj -or !$adObj) { return }

    $sqlVars = "samid=$($_.samid)", "empId=$($_.empId)" , "id=$($_.id)"
    Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }

    $_
  }
}

function New-HomeDir ($dbParams, $table, $cred, $full, $readWrite) {
  begin {
    . .\lib\New-StaffHomeDir.ps1
  }
  process {
    # Skip home folder creation if AD Obj already present
    if ($_.adObj) { return $_ }
    Write-Host ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $_.samId, $_.fileServer) -F Green
    $_ | New-StaffHomeDir $cred $full $readWrite
    # No need to pass obj down pipe if home dir is new as cloud syncs usually have not occurred yet.
  }
}

function Set-EmpId {
  process {
    $name = $_.nameLast + ', ' + $_.nameFirst
    Write-Verbose ('{0},EmpId: [{1}],{2}' -f $MyInvocation.MyCommand.Name, $_.empId, $name)

    $adObj = if ($_.empId -match '\d') {
      Write-Verbose ('{0},EmpId: [{1}],EmpId Pulled from intDB. Checking AD' -f $MyInvocation.MyCommand.Name, $_.empId)
      $filter = "EmployeeID -eq '{0}'" -f $_.empId
      Get-ADUser -Filter $filter -Properties *
    }

    $empId = if ($adObj) { $adObj.EmployeeID }
    elseif ($_.empId -match '\d' -and $_.empId -ne 0) { $_.empId }
    else { Get-Random -Min 1000000 -Max 10000000 }

    $_.empId = $empId
    $_ | Add-Member -MemberType NoteProperty -Name adObj -Value $adObj
    $_
  }
}

function Format-UserObject {
  begin {
    . .\lib\Format-Name.ps1
    . .\lib\New-Name.ps1
    . .\lib\New-PassPhrase.ps1
    . .\lib\New-SamID.ps1
  }
  process {
    # Write-Verbose ($_ | Out-string)
    $fn, $ln, $mn = (Format-Name $_.nameFirst), (Format-Name $_.nameLast), (Format-Name $_.nameMiddle)
    $newName = if ($_.adObj) { $_.adObj.name } else { New-Name -F $fn -M $mn -L $ln }
    $samId = if ($_.adObj) { $_.adObj.samAccountName } else { New-SamID -F $fn -M $mn -L $ln }
    $siteData = $_ | Set-Site
    $psObj = [PSCustomObject]@{
      id         = $_.id
      adObj      = $_.adObj
      db_gsuite  = $_.gsuite
      fn         = $fn
      ln         = $ln
      mi         = $mn
      name       = $newName
      samid      = $samId
      empId      = $_.empId
      emailWork  = $samid + $Domain1
      gsuite     = $samid + $Domain2
      siteDescr  = $siteData.SiteDescr
      siteCode   = $_.siteCode
      company    = $Organization
      jobType    = $_.jobType
      pw1        = New-PassPhrase
      pw2        = New-PassPhrase
      groups     = $siteData.Groups
      fileServer = $siteData.FileServer
      targetOU   = $TargetOrgUnit
      info       = $_.empId, $samid, ($fn + ' ' + $ln) -join ','
    }
    Write-Verbose ($psObj | Out-String)
    $psObj
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

function Set-Site {
  begin {
    $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json
  }
  process {
    # Skip blank or null site codes
    if (  $_.siteCode -lt 0 ) { return }
    $sc = $_.siteCode
    $sd = $_.siteDescr
    $siteData = $lookupTable | Where-Object { [int]$_.siteCode -eq [int]$sc }
    $siteData = $lookupTable | Where-Object { $_.siteDescr -eq $sd }
    if (-not$siteData) { return (Write-Host ('{0},{1},{2},No Site match.' -f $MyInvocation.MyCommand.Name, $_.empId, $sc) -f Magenta) }
    Write-Verbose ('{0},{1},{2},Site match: {3}' -f $MyInvocation.MyCommand.Name, $_.empId, $sc, $siteData.SiteDescr)
    $siteData
  }
}

function Update-Groups ($dbParams, $table) {
  begin {
    $A5 = 'Office365_A5_Faculty' # Microsoft 365 License for admin and office staff
    $A1 = 'Office365_A1_Faculty' # Microsoft 365 License for general staff
  }
  process {
    # If group memberships present then this likely already ran.
    if ($_.adObj.memberof) { $_ ; return }
    $azureGroup = if ($_.BargUnitId -eq 'CUMA') { $A5 } else { $A1 }
    # Add user to various groups
    $groups = $DefaultStaffGroups + $azureGroup
    if ( $_.groups ) { $groups += $_.groups.Split(",") }

    $msg = $MyInvocation.MyCommand.Name, $_.samid, ($groups -join ',')
    Write-Host ('{0},[{1}],[{2}]' -f $msg) -F Blue

    if ( -not$WhatIf ) { Add-ADPrincipalGroupMembership -Identity $_.samid -MemberOf $groups }
    $_
  }
}

function Update-EmpEmailWork ($dbParams, $table) {
  begin {
    $testSql = "SELECT empId FROM $table WHERE empID = @id"
    $updateSql = "UPDATE $table SET EmailWork = @mail WHERE EmpID = @id"
  }
  process {
    $sqlVars = "id=$($_.empId)"
    $result = New-SqlOperation @dbParams -Query $testSql -Parameters $sqlVars
    if (-not$result) {
      Write-Verbose ('{0},[{1}],EmpId not found in Database' -f $MyInvocation.MyCommand.Name, $_.info)
      return $_
    }
    $sqlVars = "mail=$($_.emailWork)", "id=$($_.empId)"
    Write-Host ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ',') ) -F Magenta
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
    $_
  }
}

function Update-IntDB ($dbParams, $table) {
  begin { $updateSql = "UPDATE $table SET sourceSystem = @sys, dts = CURRENT_TIMESTAMP WHERE id = @id;" }
  process {
    $sqlVars = "sys=$ENV:COMPUTERNAME", "id=$($_.id)"
    Write-Host ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
    $_
  }
}

function Update-ADPW ($dbParams, $table) {
  begin {
    $updateSql = "UPDATE $table SET tempPw = @pw WHERE id = @id"
  }
  process {
    Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info )
    $obj = Get-ADUser -Identity $_.samid -Properties WhenCreated, PasswordLastSet
    <# PasswordLastSet and WhenCreated differ by less than a second on new accounts.
    5 second threshold used for safety #>
    if ( ($obj.PasswordLastSet - $obj.WhenCreated).TotalMilliseconds -gt 5000 ) { $_ ; return }
    $securePw = ConvertTo-SecureString -String $_.pw2 -AsPlainText -Force
    if (-not$WhatIf) {
      Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info ) -F DarkGreen
      <# Once Gsuite account is synced then the password reset is picked up
      via the Gsuite service running on the asscociated Domain Controller
      and this activates the gsuite account #>
      Set-ADAccountPassword -Identity $_.samid -NewPassword $securePw -Confirm:$false
    }
    $sqlVars = "pw=$($_.pw2)", "id=$($_.id)"
    Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $updateSql, ($sqlVars -join ','))
    # Add pw2 to intDB to allow for
    if (-not$WhatIf) { New-SqlOperation @dbParams -Query $updateSql -Parameters $sqlVars }
    $_
  }
}

Import-Module -Name ExchangeOnlineManagement -Cmdlet 'Connect-ExchangeOnline', 'Get-EXOMailBox'
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

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date "5:00pm" }
$delay = if ($WhatIf ) { 0 } else { 180 }

do {
  $objs = New-SqlOperation @intDBparams -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv
  if ($objs) {
    $dc = Select-DomainController $DomainControllers
    New-ADsession -DC $dc -cmdlets $cmdlets -Cred $ActiveDirectoryCredential
  }

  $objs | Set-EmpId | Format-UserObject |
    New-UserADObj $intDBparams $NewAccountsTable |
      Update-Groups $intDBparams $NewAccountsTable |
        New-HomeDir $intDBparams $NewAccountsTable $FileServerCredential $FullAccess |
          Confirm-NetEmail $intDBparams $NewAccountsTable |
            Update-ADPW $intDBparams $NewAccountsTable |
              Confirm-OrgEmail $intDBparams $NewAccountsTable $O365Credential |
                Update-EmpEmailWork $empBParams $EmployeeTable |
                  Update-IntDB $intDBparams $NewAccountsTable | Get-CreationTime | Complete-Processing

  Write-Verbose "Pausing for $delay seconds before next run..."
  Clear-SessionData
  Start-Sleep $delay
} Until ( $WhatIf -or ((Get-Date) -ge $stopTime) )
if (!$WhatIf) { Remove-TmpEXOs }
if ($WhatIf) { Show-TestRun }