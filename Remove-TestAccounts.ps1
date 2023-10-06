
[cmdletbinding()]
param (
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [int[]]$EmployeeId,
 [string[]]$Username,
 [Alias('wi')]
 [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
($user = Get-ADUser -filter "EmployeeID -eq '$EmployeeId'" -Properties SamAccountName, HomePage, departmentNumber)
$user | Select-Object -Property SamAccountName, HomePage, departmentNumber
$gam = '.\bin\gam.exe'
Write-Host ('{0},{1},Removing Gsuite' -f $MyInvocation.MyCommand.Name, $user.Homepage)
& $gam delete user $user.HomePage
Write-Host ('{0},{1},Removing AD' -f $MyInvocation.MyCommand.Name, $user.SamAccountName)
$user | Remove-ADObject
$lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json
$user.deptartmentNumber
$siteData = $lookupTable | Where-Object { [int]$_.SiteCode -eq [int]$user.deptartmentNumber }
Write-Host ('{0},{1},{2},Removing Home Directory' -f $MyInvocation.MyCommand.Name, $user.SamAccountName, $siteData.FileServer)
. .\lib\Remove-StaffHomeDir.ps1
Remove-StaffHomeDir -samid $user.SamAccountName -fileserver $siteData.FileServer -cred $FileServerCredential

# function Get-UserObj {
#   begin {
#     $properties = 'EmployeeID', 'HomePage', 'departmentNumber'
#   }
#   process {
#     if ($null -eq $_) { return }
#     Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_)
#     $filter = "EmployeeID -eq '{0}' -or SamAccountName -eq '{1}'" -f $_, $_
#     Write-Verbose $filter
#     $adParams = @{
#       Filter     = $filter
#       Properties = $properties
#     }
#     $adObj = Get-ADUser @adParams
#     if (-not$adObj) {
#       Write-Host ('{0},Object not found with samid or empid of: {1}' -f $MyInvocation.MyCommand.Name, $_) -F Magenta
#       return
#     }
#     Write-Verbose ( $adObj | Out-String )
#     $adObj
#   }
# }

# function Remove-ADObj {
#   begin {
#   }
#   process {
#     Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.SamAccountName)
#     $_ | Remove-ADObject -Confirm:$true -WhatIf:$WhatIf
#     $_
#   }
# }

# function Remove-GSuite ($dbParams, $table) {
#   begin {
#     $gam = '.\bin\gam.exe'
#   }
#   process {
#     if ($_.HomePage -notlike '*@chicousd.net') { $_ ; return }
#     $gsuite = $_.HomePage
#     Write-Verbose ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $gsuite)

#     ($guser = & $gam print users query "email: $gsuite" | ConvertFrom-CSV)*>$null
#     Write-Verbose ($guser | Out-String )
#     if ($guser.PrimaryEmail -ne $gsuite) {
#       Write-Verbose ('{0},Gsuite User NOT Found: [{1}]' -f $MyInvocation.MyCommand.Name, $gsuite)
#       $_
#       return
#     }

#     Write-Host ('{0},Removing Gsuite User: [{1}]' -f $MyInvocation.MyCommand.Name, $gsuite)
#     $answer = Read-Host 'Remove Gsuite? Y | N'
#     if ($answer -match 'y') {
#       $myFilter = "EmployeeID -eq '{0}'" -f $_.EmployeeID
#       $adObj = Get-ADUser -Filter $myFilter
#       if ($adObj) {
#         Write-Host ('{0},{1},AD Object not poofed. Why? WHY!??!' -f $MyInvocation.MyCommand.Name, $_.SamAccountName)
#         $_
#         return
#       }
#       & $gam delete user $gsuite
#       $_
#       return
#     }
#     Write-Host "You did't choose the yes" -f Red
#     $_
#   }
# }

# function Remove-HomeDir {
#   begin {
#     . .\lib\Remove-StaffHomeDir.ps1
#     $lookupTable = Get-Content -Path .\json\site-lookup-table.json | ConvertFrom-Json
#   }
#   process {
#     $dept = if ($_.departmentNumber) { $_.departmentNumber }
#     else {
#       $answer = Read-Host 'Enter a dept num if Ye dare!'
#       if (-not$answer) { return }
#       $answer
#     }
#     Write-Verbose $dept
#     $siteData = $lookupTable | Where-Object { [int]$_.SiteCode -eq [int]$dept }
#     Write-Verbose ($siteData | Out-String)
#     Remove-StaffHomeDir -samid $_.SamAccountName -file $siteData.FileServer -cred $FileServerCredential
#   }
# }

# # ==================== Main =====================
# # Imported Functions
# # $dc = Select-DomainController $DomainControllers
# # $cmdlets = 'Get-ADUser', 'Remove-ADObject'
# # New-ADsession -DC $dc -cmdlets $cmdlets -Cred $ActiveDirectoryCredential

# if (-not$EmployeeId -and -not$Username) { exit }

# $EmployeeId, $Username | Get-UserObj |
# Remove-ADObj | Remove-GSuite | Remove-HomeDir