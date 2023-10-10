
[cmdletbinding()]
param (
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [int[]]$EmployeeId,
 [string[]]$Username,
 [Alias('wi')]
 [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
$user = Get-ADUser -filter "EmployeeID -eq '$EmployeeId'" -Properties *
if ($user){

 $user | Select-Object -Property SamAccountName, HomePage, departmentNumber

 $gam = '.\bin\gam.exe'
 Write-Host ('{0},{1},Removing Gsuite' -f $MyInvocation.MyCommand.Name, $user.Homepage)
 & $gam delete user $user.HomePage

 Write-Host ('{0},{1},Removing AD' -f $MyInvocation.MyCommand.Name, $user.SamAccountName)
 $user | Remove-ADObject

 $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json
 $sc = $user.departmentNumber | Out-String
 $siteData = $lookupTable | Where-Object { [int]$_.SiteCode -eq [int]$sc }
 $msg = $MyInvocation.MyCommand.Name, $user.SamAccountName, $siteData.FileServer
 Write-Host ('{0},{1},{2},Removing Home Directory' -f $msg)
 . .\lib\Remove-StaffHomeDir.ps1
 Remove-StaffHomeDir -samid $user.SamAccountName -fileserver $siteData.FileServer -cred $FileServerCredential
}