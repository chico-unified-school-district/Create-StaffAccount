function Compare-Commit {

 Write-Host ('{0}' -f $MyInvocation.MyCommand.Name)
 $hash = git rev-parse HEAD


}