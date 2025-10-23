Describe 'New-SamID function' {
 $libPath = Join-Path $PSScriptRoot '..\lib\New-SamID.ps1'
 . $libPath

 Context 'basic generation' {
  It 'generates a samid using first initial + last name' {
   # Mock AD lookups to simulate no conflicts
   Mock -CommandName Get-ADUser -MockWith { $null }

   $sam = New-SamID -First 'John' -Middle 'Q' -Last 'Public'
   $sam | Should Match '^(JPublic|John\.Public|JPublic[0-9]*)'
  }

  It 'truncates long samids to 20 characters' {
   Mock -CommandName Get-ADUser -MockWith { $null }
   $longFirst = 'Verylongfirstname'
   $longLast = 'Verylonglastnamewithlots'
   $sam = New-SamID -First $longFirst -Last $longLast
   ($sam.Length -le 20) | Should Be $true
  }
 }

 Context 'conflict avoidance' {
  It 'skips existing samAccountName and returns next available' {
   # Simulate Get-ADUser returning a non-null result for first call then null for subsequent calls
   $script:call = 0
   Mock -CommandName Get-ADUser -MockWith {
    $script:call++ | Out-Null
    if ($script:call -eq 1) { return @{ SamAccountName = 'JSmith' } } else { return $null }
   }

   $sam = New-SamID -First 'John' -Last 'Smith'
   $sam | Should Not Be 'JSmith'
  }
 }
}
