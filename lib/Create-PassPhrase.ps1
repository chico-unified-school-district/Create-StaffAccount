function Create-PassPhrase {
	[cmdletbinding()] 
	param ()
	Get-ChildItem -Path *.dic -Depth 1 | Foreach { New-Variable -Name $_.name -Value $_.pspath }
 $TextInfo = (Get-Culture).TextInfo
 $colors = Get-Variable -Name *colors*dic | Select-Object Value | foreach { Get-Content -path $_.Value }
 $fruits = Get-Variable -Name *fruits*dic | Select-Object Value | foreach { Get-Content -path $_.Value }
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