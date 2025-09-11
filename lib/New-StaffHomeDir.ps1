function New-StaffHomeDir ($cred, [string[]]$full) {
 process {
  $samid = $_.samid
  $fileServer = $_.site.FileServer
  if ( $fileServer -match '[A-Za-z]' ) {
   if (-not(Test-Connection -ComputerName $fileServer -Count 2)) {
    Write-Host ('{0},Server not found [{1}]' -f $MyInvocation.MyCommand.Name, $fileServer)
    return
   }
   Write-Verbose ('{0},\\{1}\User\{2}' -f $MyInvocation.MyCommand.Name, $fileServer, $samid)
   if ($WhatIf) { return }
   $originalPath = Get-Location

   Write-Verbose 'Adding PSDrive'
   New-PSDrive -Name share -Root \\$fileServer\User -PSProvider FileSystem -Credential $cred | Out-Null

   Set-Location -Path share:

   $homePath = ".\$samid"
   $docsPath = ".\$samid\Documents"

   if ( -not(Test-Path -Path $docsPath) ) {
    Write-Host "Creating HomeDir for $samid on $fileServer."
    New-Item -Path $docsPath -ItemType Directory -Confirm:$false | Out-Null

    # Remove Inheritance and add users and groups
    ICACLS $homePath /inheritance:d /grant 'Chico\CreateHomeDir:(OI)(CI)(F)' | Out-Null
    Start-Sleep 5 # A delay is needed to ensure objects can be mapped to ACLs properly
    ICACLS $homePath /grant 'SYSTEM:(OI)(CI)(F)' | Out-Null
    ICACLS $homePath /remove 'BUILTIN\Administrators' | Out-Null
    ICACLS $homePath /remove 'BUILTIN\BUILTIN' | Out-Null
    ICACLS $homePath /remove 'CHICO\Administrator' | Out-Null

    foreach ($item in $full) {
     Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $homePath, $item) -F Blue
     ICACLS $homePath /grant "${$item}:(OI)(CI)(F)" | Out-Null
    }

    ICACLS $homePath /grant "${samid}:(OI)(CI)(RX)" | Out-Null
    ICACLS $docsPath /grant "${samid}:(OI)(CI)(M)" | Out-Null

    $regthis = "($($samid):\(OI\)\(CI\)\(M\))"
    if ( (ICACLS $docsPath) -match $regthis ) {
     Write-Verbose "HomeDir Created,ACLs correct,\\$fileServer\User\$samid"
    }
    else {
     Write-Error "ERROR,ACLs not correct for \\$fileServer\User\$samid\Documents"
    }
   }

   Set-Location $originalPath

   Write-Verbose 'Removing PSDrive'
   Remove-PSDrive -Name share -Confirm:$false -Force | Out-Null
  }
 }
}