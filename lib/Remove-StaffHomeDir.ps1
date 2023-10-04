function Remove-StaffHomeDir ($samid, $fileServer, $cred) {
  iF ( $fileServer -match "[A-Za-z]" ) {
    if (-not(Test-Connection -ComputerName $fileServer -Count 2)) {
      Write-Verbose ('{0},Server not found [{1}]' -f $MyInvocation.MyCommand.Name, $fileServer)
      return
    }
    Write-Verbose ('{0},\\{1}\User\{2}' -f $MyInvocation.MyCommand.Name, $fileServer, $samid)
    # if ($WhatIf) { return }
    $originalPath = Get-Location

    Write-Verbose "Adding PSDrive"
    New-PSDrive -name share -Root \\$fileServer\User -PSProvider FileSystem -Credential $cred | Out-Null

    Set-Location -Path share:
    $docsPath = ".\$samid\Documents"

    if ( Test-Path -Path $docsPath ) {
      Write-Host "Removing HomeDir for $samid on $fileServer."
      Remove-Item -Path .\$samid -Confirm:$true -Force:$false
    }
    else {
      Write-Host "Aint no HomeDir for $samid on $fileServer."
    }

    Set-Location $originalPath

    Write-Verbose "Removing PSDrive"
    Remove-PSDrive -Name share -Confirm:$false -Force | Out-Null
  }
}