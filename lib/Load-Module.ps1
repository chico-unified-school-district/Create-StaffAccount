function Load-Module {
 begin {
  $myModules = Get-Module
 }
 process {
  if ($myModules.name -contains $_) { return }
  if (-not(Get-Module -Name $_ -ListAvailable)) {
   Install-Module -Name $_ -Scope CurrentUser -AllowClobber -Confirm:$false -Force
  }
  Import-Module -Name $_ -Force -ErrorAction Stop -Verbose:$false | Out-Null
  $curModule = Get-Module -Name $_
  if (-not$curModule) {
   Write-Host ('{0},{1},Module not loaded for some reason' -f $MyInvocation.MyCommand.Name, $_)
   return
  }
  Write-Host ('{0},{1},Module loaded' -f $MyInvocation.MyCommand.Name, $_)
 }
}