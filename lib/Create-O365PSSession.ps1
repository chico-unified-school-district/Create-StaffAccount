function Create-O365PSSession {
 [cmdletbinding()]
 param(
  [System.Management.Automation.PSCredential]$Credential
 )

 'Creating Office365 session'
 $o365SessionParams = @{
  ConfigurationName = 'Microsoft.Exchange'
  ConnectionUri     = 'https://outlook.office365.com/powershell-liveid/'
  Credential        = $Credential
  Authentication    = 'Basic'
  AllowRedirection  = $True
 }
 $O365Session = New-PSSession @o365SessionParams -ErrorAction Stop
 $o365CmdLets = 'Get-Mailbox', 'Get-User', 'Set-MailboxRegionalConfiguration', 'Get-MailboxRegionalConfiguration'
 Import-PSSession -Session $O365Session -CommandName $o365CmdLets -ErrorAction Stop
}