<#
.SYNOPSIS
Generate a unique SamAccountName candidate from name parts.

.DESCRIPTION
New-SamID accepts first, middle and last name parts and produces a candidate SAM account name. It tests candidate
values against AD for existing SamAccountName and proxyaddresses and returns the first available candidate. The
returned value is truncated to 20 characters to meet AD limits.

.PARAMETER First
Given (first) name.

.PARAMETER Middle
Middle name or initial (optional).

.PARAMETER Last
Family (last) name.

.OUTPUTS
String - a SAM account name candidate (max 20 characters).
#>
function New-SamID {
 [cmdletbinding()]
 param(
  [Parameter(Position = 0, Mandatory = $true)]
  [string]$First,
  [Parameter(Position = 1, Mandatory = $false)]
  [string]$Middle,
  [Parameter(Position = 2, Mandatory = $true)]
  [string]$Last
 )

 function Format-FirstLetter ($str) {
  if ($str.length -gt 1) { $str.substring(0, 1).ToUpper() + $str.substring(1) }
 }

 function makeNameObj ($f, $m, $l) {
  New-Object psobject -Property @{
   f = removeNonLetters $f
   m = removeNonLetters $m
   l = removeNonLetters $l
  }
 }
 function outputFreeSam ($sam) {
  if (
   -not( Get-ADUser -LDAPFilter "(samAccountName=$sam)" ) -and
   -not( Get-ADUser -LDAPFilter "(proxyaddresses=smtp:$sam@*)" )
  ) { $sam }
 }
 function removeNonLetters ( $str ) { if ($str) { $str -replace '[^a-zA-Z]' } }
 function testSams ($nameObj) {
  $possibleNames = @(
   $nameObj.f.Substring(0, 1) + $nameObj.l
   $( if ($nameObj.m) { $nameObj.f.Substring(0, 1) + $nameObj.m.SubString(0, 1) + $nameObj.l } )
   $nameObj.f + '.' + $nameObj.l
   $nameObj.f.Substring(0, 1) + $nameObj.l + '1'
   $nameObj.f.Substring(0, 1) + $nameObj.l + '2'
   $nameObj.f.Substring(0, 1) + $nameObj.l + '3'
  )
  if ($Middle) {
  }
  foreach ($name in $possibleNames) {
   if (outputFreeSam $name) { $name; return }
  }
 }
 function truncateSam {
  process {
   # Limit samid to 20 chars per Microsoft's specification
   if ($_.length -gt 20) { $_.substring(0, 20) } else { $_ }
  }
 }
 # ===================================================================
 Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $First, $Middle, $Last)
 # process
 $nameObj = makeNameObj -f (Format-FirstLetter $First) -m (Format-FirstLetter $Middle) -l (Format-FirstLetter $Last)
 testSams $nameObj | truncateSam
}