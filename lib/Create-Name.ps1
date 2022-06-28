function Create-Name {
 [CmdletBinding()]
 param(
  [Parameter(Mandatory = $true)]
  [string]$First,
  [Parameter(Mandatory = $false)]
  [string]$Middle = $null,
  [Parameter(Mandatory = $true)]
  [string]$Last
 )
 begin {
  # $First + $Middle + $Last
  function Confirm-FreeName {
   process {
    if (-not(Get-ADUser -LDAPFilter "(name=$_)")) { $_ ; return }
   }
  }
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
  $fn = Format-FirstLetter $First
  $mn = Format-FirstLetter $Middle
  $ln = Format-FirstLetter $Last
  $list = @(
   $fn, $ln -join ' '
   if ($mn.Length -ge 1) { $fn, ($mn.Substring(0, 1) + '.'), $ln -join ' ' }
   if ($mn.Length -gt 1) { $fn, $mn, $ln -join ' ' }
   $fn, $ln, 2 -join ' '
   $fn, $ln, 3 -join ' '
  )
  Write-Verbose ( $list | Out-String )
  foreach ($n in $list) {
   # pick the first valid, free name from the list array
   if ($n | Confirm-FreeName) {
    $n
    return
   }
  }
 }
}
