Describe 'Format-Name function' {
 # Dot-source the helper (tests located in tests/ folder relative to repo root)
 $libPath = Join-Path $PSScriptRoot '..\lib\Format-Name.ps1'
 . $libPath

 It 'returns $null for null or empty input' {
  $result = Format-Name $null
  Should BeNullOrEmpty $result
 }

 It 'title-cases and preserves capitalization after apostrophe' {
  $result = Format-Name "d'angelo"
  $result | Should Be "D'Angelo"
 }

 It 'collapses multiple spaces to a single space' {
  $result = Format-Name 'john   doe'
  $result | Should Be 'John Doe'
 }

 It 'removes non-letters except apostrophes and spaces' {
  $result = Format-Name 'ann-marie'
  $result | Should Be 'Annmarie'
 }
}
