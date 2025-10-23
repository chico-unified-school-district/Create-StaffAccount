<#
.SYNOPSIS
Generates a secure, human-readable passphrase for temporary user passwords.

.DESCRIPTION
New-PassPhrase builds a passphrase by randomly selecting words from local dictionary files (*.dic) found in the
working directory. The resulting phrase is title-cased, includes a digit and a symbol, and is between 13 and 16
characters in length.

.OUTPUTS
String - a generated passphrase suitable for short-term use.

.NOTES
This function expects dictionary files in the working directory (for example `colors.dic` and `fruits.dic`).
#>
function New-PassPhrase {
	[cmdletbinding()]
	param ()
	Get-ChildItem -Path *.dic -Depth 1 | ForEach-Object { New-Variable -Name $_.name -Value $_.pspath }
 $TextInfo = (Get-Culture).TextInfo
 $colors = Get-Variable -Name *colors*dic | Select-Object Value | ForEach-Object { Get-Content -Path $_.Value }
 $fruits = Get-Variable -Name *fruits*dic | Select-Object Value | ForEach-Object { Get-Content -Path $_.Value }
 [string]$num = Get-Random -Min 2 -Max 9
 [string]$sym = '@', '%', '&', '#', '!' | Get-Random
 do {
  $p1 = $colors[(Get-Random -Min 0 -Max $colors.count)]
  $p2 = $fruits[(Get-Random -Min 0 -Max $fruits.count)]
  $phrase = $TextInfo.ToTitleCase("$p1") + $num + $TextInfo.ToTitleCase("$p2") + $sym
 } until ( ($phrase.Length -ge 13) -and ($phrase.Length -le 16))
 $phrase
}

# http://www.ashley-bovan.co.uk/words/partsofspeech.html