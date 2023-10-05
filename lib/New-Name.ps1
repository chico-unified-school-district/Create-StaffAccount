function New-Name ($first, $middle, $last) {
 begin {
  function Confirm-FreeName { process { if (-not(Get-ADUser -LDAPFilter "(name=$_)")) { $_ ; return } } }
  $TextInfo = (Get-Culture).TextInfo
 }
 process {
  $fn = $TextInfo.ToTitleCase($first)
  $mn = $TextInfo.ToTitleCase($middle)
  $ln = $TextInfo.ToTitleCase($last)
  $list = @(
   $fn, $ln -join ' '
   if ($mn.Length -ge 1) { $fn, ($mn.Substring(0, 1) + '.'), $ln -join ' ' }
   if ($mn.Length -gt 1) { $fn, $mn, $ln -join ' ' }
   $fn, $ln, 2 -join ' '
   $fn, $ln, 3 -join ' '
  )
  # Write-Verbose ( $list | Out-String )
  foreach ($n in $list) {
   # pick the first valid, free name from the list array
   if ($n | Confirm-FreeName) {
    $n
    return
   }
  }
 }
}
