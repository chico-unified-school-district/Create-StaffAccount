# Create-StaffAccount

Monitors an independent database for new entries and prepares accounts when needed.

## Dependencies

- PowerShell modules: ExchangeOnlineManagement, dbatools, CommonScriptFunctions
- Local `lib\` helper scripts must be present: `New-ADUserObject.ps1`, `New-StaffHomeDir.ps1`, `New-PassPhrase.ps1`, `New-SamID.ps1`, `Format-Name.ps1`
- `bin\gam.exe` (optional) for GSuite checks / creation

## Example scheduled task / Jenkins invocation

Run from PowerShell (no profile) in a scheduled task or CI job. Ensure credentials are available to the job (use secure credential store):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "G:\My Drive\CUSD\Scripts\GitHub\Create-StaffAccount\Create-StaffAccount.ps1" `
 -IntermediateSqlServer 'sql01' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' `
 -IntermediateCredential (Get-Credential) -DomainControllers 'dc01' -ActiveDirectoryCredential (Get-Credential) `
 -O365Credential (Get-Credential) -FileServerCredential (Get-Credential)
```

For a dry-run preview use the `-WhatIf` switch.

```

```
