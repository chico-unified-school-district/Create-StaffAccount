<#
.SYNOPSIS
Normalize and title-case name strings for AD account creation.

.DESCRIPTION
Format-Name removes disallowed characters (leaving letters, apostrophes and spaces), collapses multiple spaces,
and returns a culture-aware title-cased string. It preserves capitalization following an apostrophe.

.PARAMETER str
The raw input name string to normalize.

.OUTPUTS
String - title-cased normalized name suitable for use in AD attributes and display names.
#>
function Format-Name ($str) {
 if (-not$str) { return }
 $TextInfo = (Get-Culture).TextInfo
 $s1 = $str -replace '`', "'"
 $s2 = [regex]::Replace($s1, "[^A-Za-z\'\s]", '')
 $s3 = $TextInfo.ToTitleCase($s2)
 $s4 = [regex]::Replace($s3, '(?<='').', { param($m) "$m".ToUpper() })
 $s5 = [regex]::Replace($s4, '\s{2,}', ' ')
 $s5.Trim()
}
