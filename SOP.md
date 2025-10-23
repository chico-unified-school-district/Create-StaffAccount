# SOP: Create-StaffAccount (Account Provisioning)

Version: 1.0
Date: 2025-10-23
Author: Operations

## Purpose

This Standard Operating Procedure (SOP) explains how to run, test, schedule, and troubleshoot the
Create-StaffAccount process. The process monitors an intermediate SQL table for new employee rows and
performs provisioning tasks: creating AD users, home directories, and cloud mailbox / forwarding setup.

## Scope

Applies to IT staff and automation engineers responsible for onboarding staff accounts for the organization
using the scripts in this repository.

## Roles & Responsibilities

- Operator: run manual or scheduled jobs, review logs and test runs, escalate issues.
- AD Admin: approves organizational OU, AD permissions, and group membership changes.
- File Server Admin: ensures file servers are available and share permissions are correct.
- Cloud Admin: ensures ExchangeOnline and GSuite connectivity and licensing.

## Prerequisites

- PowerShell on the host that runs the script (Windows PowerShell or PowerShell Core). The repository was tested with
  Windows PowerShell that includes Pester v3 in this environment; adjust tests if using Pester v5.
- Required PowerShell modules installed and accessible: ExchangeOnlineManagement, dbatools, CommonScriptFunctions.
- The following files must be present in the repo root: `Create-StaffAccount.ps1`, `lib\New-ADUserObject.ps1`,
  `lib\New-StaffHomeDir.ps1`, `lib\New-PassPhrase.ps1`, `lib\New-SamID.ps1`, `lib\Format-Name.ps1`.
- `bin\gam.exe` if GSuite checks / operations are needed (optional).
- Service accounts / PSCredentials with least privilege for AD, SQL, file server and Exchange Online operations.

## Security & Credentials

- Do NOT store plain-text passwords in the repository. Use secure stores for credentials (Windows Credential Manager,
  Azure Key Vault, Jenkins credentials store, or similar).
- The script accepts PSCredential parameters for AD, SQL and file server. When scheduling, either use a system
  account with appropriate privileges or configure the job to retrieve PSCredential securely.
- Limit access to the repository and scripts to authorized staff only.

## Quick references

- Main script: `Create-StaffAccount.ps1`
- Intermediate SQL table: configurable via `-NewAccountsTable` parameter
- Lookup table for sites: `json\site-lookup-table.json`

## How it works (high level)

1. Connects to the intermediate SQL instance and queries the configured `NewAccountsTable` for rows where EmailWork is NULL.
2. For each returned row it builds a processing object, formats names, generates/validates EmployeeID and SAM ID,
   and checks for existing AD / mailbox / GSuite accounts.
3. Where required, creates AD user (via `lib\New-ADUserObject.ps1`), creates home directory (via
   `lib\New-StaffHomeDir.ps1`), sets groups, configures mailbox forwarding/retention, and writes back to the
   intermediate and employee DBs.

## Running the process — Manual

1. Open a PowerShell session on a machine with network access to AD, file servers, and SQL servers.
2. From the repository root, prepare credentials:

```powershell
$adCred = Get-Credential -Message 'AD service account'
$o365Cred = Get-Credential -Message 'O365 service account'
$fsCred = Get-Credential -Message 'File server account'
$intCred = Get-Credential -Message 'Intermediate SQL account'
$empCred = Get-Credential -Message 'Employee SQL account'
```

3. Run a dry-run (WhatIf) first to preview actions:

```powershell
.\Create-StaffAccount.ps1 -IntermediateSqlServer 'int-sql.example.local' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential $intCred -DomainControllers 'dc1','dc2' -ActiveDirectoryCredential $adCred -O365Credential $o365Cred -FileServerCredential $fsCred -EmployeeServer 'emp-sql.example.local' -EmployeeDatabase 'hrdb' -EmployeeCredential $empCred -WhatIf
```

4. Review output for any warnings or errors. If output looks correct, run without `-WhatIf`.

## Running the process — Scheduled (Windows Task Scheduler)

1. Create a scheduled task that runs as a service account with the required privileges.
2. Use an action similar to the following (update paths and parameters to your environment):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "G:\My Drive\CUSD\Scripts\GitHub\Create-StaffAccount\Create-StaffAccount.ps1" -IntermediateSqlServer 'int-sql.example.local' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (Get-Credential) -DomainControllers 'dc1' -ActiveDirectoryCredential (Get-Credential) -O365Credential (Get-Credential) -FileServerCredential (Get-Credential)
```

Tip: Use credential stores in Task Scheduler or use a wrapper script to securely fetch credentials from a vault.

## Running the process — CI / Jenkins

- Use the Jenkins credentials store and pass credentials as secure parameters or inject them into the build agent.
- Example Jenkins pipeline step (pseudo):

```groovy
withCredentials([usernamePassword(credentialsId: 'adCred', usernameVariable: 'AD_USER', passwordVariable: 'AD_PW')]) {
  sh "powershell -NoProfile -ExecutionPolicy Bypass -File Create-StaffAccount.ps1 -IntermediateSqlServer 'int-sql' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (New-Object System.Management.Automation.PSCredential($AD_USER,(ConvertTo-SecureString $AD_PW -AsPlainText -Force))) ..."
}
```

## Testing — Pester

- Unit tests are in the `tests\` folder. They cover `Format-Name` and `New-SamID`.
- To run tests locally:

```powershell
# from repo root
Invoke-Pester -Path .\tests\ -OutputFile Detailed -OutputFormat LegacyNUnitXml
```

Notes:

- The tests were written for the Pester version installed on the current host (Pester v3 in the test environment).
- If you run Pester v5, some assertions and mock behaviors may need updates (the repo already contains both v3 and
  v5-aware patterns in tests and the mocking approach).

## Troubleshooting

- Script cannot connect to SQL: confirm network connectivity, instance name, credentials, and that the SQL Server
  allows remote connections. Test with `Connect-DbaInstance` manually.
- AD cmdlet errors: ensure `ActiveDirectory.psd1` cmdlets are available on the host and that the AD credential has
  rights to create users and set attributes.
- Home directory creation fails: verify file server name in site lookup table, ensure the file server is reachable
  (Test-Connection), and that the file server account has rights to create folders and set ACLs.
- Exchange/EXO errors: ensure `ExchangeOnlineManagement` module is installed and the O365 credential can connect.
- GSuite checks using `gam.exe` fail: ensure `bin\gam.exe` is present and configured properly with an admin account.
- Common test issues: Pester version mismatch causes `Should` operator errors. Use the Pester version installed in the
  execution environment or adjust tests accordingly.

## Rollback & Safety

- The script updates the intermediate database and, in some cases, the employee database. Before running in production,
  ensure backups and point-in-time recovery are available for these databases.
- If an accidental run modifies AD or SQL in error:
  - Disable affected user accounts immediately in AD (Set-ADUser -Enabled $false).
  - Restore database values from backups or manually update the intermediate table to correct fields such as empId, emailWork, gSuite, or tempPw.
  - Remove mailbox forwarding if incorrectly set.

## Logging & Monitoring

- The script writes verbose and host output during runs. For scheduled runs, capture stdout/stderr into a rotating log
  file and surface important errors to the monitoring/alerting system.
- Consider adding structured logging to emit JSON lines for easier ingestion.

## Change Management

- Any change to this process (script edits, lookup table changes, schedule changes) must be reviewed and tested in a
  staging environment before being promoted to production.

## Contact / Escalation

- Primary: AD Team (ad-team@example.local)
- Secondary: Platform Automation (automation@example.local)

## Appendix — Useful Commands

- Dry-run (WhatIf):

```powershell
.\Create-StaffAccount.ps1 -WhatIf -IntermediateSqlServer 'int-sql' -IntermediateDatabase 'intdb' -NewAccountsTable 'dbo.NewEmployees' -IntermediateCredential (Get-Credential) -DomainControllers 'dc1' -ActiveDirectoryCredential (Get-Credential) -O365Credential (Get-Credential) -FileServerCredential (Get-Credential)
```

- Run tests:

```powershell
Invoke-Pester -Path .\tests\ -OutputFile Detailed -OutputFormat LegacyNUnitXml
```

---

Revision History

- 1.0 — 2025-10-23 — Initial SOP created and added to repository.
