# Format-Name.ps1

function fName ($str) {
 if (-not$str) { return }
 $TextInfo = (Get-Culture).TextInfo
 $s1 = $TextInfo.ToTitleCase($str)
 $s2 = $s1 -replace '`', "'"
 [regex]::Replace($s2, "[^A-Za-z\'\s]", "")
}
