function Create-StaffHomeDir {
	[cmdletbinding()]
	param (
  [Parameter(Mandatory = $True)]
  [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 1 })]
  [string]$FileServer,
  [Parameter(Mandatory = $True)]
  [Alias('FSCred')]
  [System.Management.Automation.PSCredential]$ServerCredential,
  [Parameter(Mandatory = $True)]
  [string]$samid,
  [switch]$WhatIf
	)

 if ($WhatIf) { "[TEST],[CREATE-HOMEDIR],\\$FileServer\User\$samid" }
 else {
  $originalPath = Get-Location

  Write-Verbose "Adding PSDrive"
  New-PSDrive -name share -Root \\$FileServer\User -PSProvider FileSystem -Credential $ServerCredential | Out-Null

  Set-Location -Path share:

  $homePath = ".\$samid"
  $docsPath = ".\$samid\Documents"

  if (!(Test-Path -Path $docsPath)) {
   Write-Verbose "Creating HomeDir for $samid on $FileServer."
   New-Item -Path $docsPath -ItemType Directory -Confirm:$false | Out-Null
   # Remove Inheritance and add users and groups
   ICACLS $homePath /inheritance:r /grant "Chico\CreateHomeDir:(OI)(CI)(F)" "BUILTIN\Administrators:(OI)(CI)(F)" | Out-Null
   Start-Sleep 5 # A delay is needed to ensure objects can be mapped to ACLs properly
   ICACLS $homePath /grant "SYSTEM:(OI)(CI)(F)" "chico\veritas:(OI)(CI)(M)" "Chico\Domain Admins:(OI)(CI)(F)" | Out-Null
   ICACLS $homePath /grant "Chico\IS-All:(OI)(CI)(M)" | Out-Null
   ICACLS $homePath /grant "${samid}:(OI)(CI)(RX)" | Out-Null
   ICACLS $docsPath /grant "${samid}:(OI)(CI)(M)" | Out-Null
   $regthis = "($($samid):\(OI\)\(CI\)\(M\))"
   if ( (ICACLS $docsPath) -match $regthis ) {
    "HomeDir Created,ACLs correct,\\$FileServer\User\$samid"
   }
   else {
    "ERROR,ACLs not correct for \\$FileServer\User\$samid\Documents"
   }
  }
  else { "\\$FileServer\User\$samid exists" }

  Set-Location $originalPath

  Write-Verbose "Removing PSDrive"
  Remove-PSDrive -Name share -Confirm:$false -Force | Out-Null
 }
}