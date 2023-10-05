# Remmove-Chars.ps1

function Remove-Chars {
 begin {
  function Remove-Chars { process { [regex]::Replace($_, "[^A-Za-z]", "") } }
 }
 process {
  Write-Host ('{0}' -f $MyInvocation.MyCommand.Name)
  [regex]::Replace($_, "[^A-Za-z]", "")
 }
}
