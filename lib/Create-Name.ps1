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
  function nameFree {
   process {
    if (-not(Get-ADUser -LDAPFilter "(name=$_)")) { $_ ; return }
   }
  }
 }
 process {
  $list = @(
   $First, $Last -join ' '
   if ($Middle.Length -eq 1) { $First, $Middle, $Last -join ' ' }
   if ($Middle.Length -gt 1) { $First, $Middle.Substring(0, 1), $Last -join ' ' }
   if ($Middle.Length -gt 1) { $First, $Middle.Substring, $Last -join ' ' }
   $First, $Last, 2 -join ' '
   $First, $Last, 3 -join ' '
  )
  # $list
  foreach ($n in $list) {
   # pick the first valid, free name from the list array
   if ($n | namefree) {
    $n
    return
   }
  }
 }
}
