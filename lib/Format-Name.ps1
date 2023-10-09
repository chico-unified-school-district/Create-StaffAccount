# Format-Name.ps1

function Format-Name ($str) {
 if (-not$str) { return }
 $TextInfo = (Get-Culture).TextInfo
 $s1 = $str -replace '`', "'"
 $s2 = [regex]::Replace($s1, "[^A-Za-z\'\s]", "")
 $s3 = $TextInfo.ToTitleCase($s2)
 $s4 = [regex]::Replace($s3, '(?<='').', { param($m) "$m".ToUpper() })
 $s5 = [regex]::Replace($s4, '\s{2,}', ' ')
 $s5.Trim()
}
