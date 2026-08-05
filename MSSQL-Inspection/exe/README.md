# Executables

## Script

- generate-remediation.ps1: generate prioritized remediation plan from an inspection JSON
- complete-inspection/run-inspection.ps1: merged complete inspection runner (MVP + full audit)
- complete-inspection/run-batch-inspection.ps1: merged complete inspection batch runner

## Quick start

1. Open PowerShell in this folder.
2. Run with Windows auth:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -Database master
3. Run with SQL auth:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -Database master -Username sql_user -Password "<PASSWORD>"
4. Run with explicit TDS for legacy targets (ODBC mode):
   - .\\complete-inspection\\run-inspection.ps1 -Server legacy-db.example.com -Database master -Username sql_user -Password "<PASSWORD>" -ConnectionProvider Odbc -TdsVersion 7.1
5. Run with default merged configuration:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com
6. Choose report formats:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -ReportFormats "json,html"
7. Choose report language:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -ReportLanguage zh-CN
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -ReportLanguage en-US
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -ReportLanguage zh
8. Optional JSON status localization:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -ReportLanguage zh -LocalizeJsonStatus

## Output

- Output root comes from `..\\skill\\config.json` (`output.defaultRootDir`)
- When no explicit folder parameter is provided, script prompts ONE working folder selection (Windows dialog first, then CLI fallback)
- The selected working folder is shared by report output and `servers.config`
- Report language comes from `..\\skill\\config.json` (`report.language`) and can be overridden by `-ReportLanguage`
- JSON status localization comes from `..\\skill\\config.json` (`report.localizeJsonStatus`) and can be overridden by `-LocalizeJsonStatus`
- Single inspection output: `<root>\\<date>\\complete-inspection-<timestamp>.*`
- Batch output: `<root>\\<date>\\complete-inspection-batch-<timestamp>\\...`
- Remediation output: `<root>\\<date>\\remediation-<timestamp>.*`

## Batch run

1. Edit instances.example.json and add multiple servers.
2. Run:
   - .\\complete-inspection\\run-batch-inspection.ps1 -InstancesPath .\\complete-inspection\\instances.example.json
3. Review batch output:
   - <root>\\<date>\\complete-inspection-batch-<timestamp>\\complete-inspection-batch-summary.json
   - <root>\\<date>\\complete-inspection-batch-<timestamp>\\complete-inspection-batch-summary.html
   - Includes merged score (combined/mvp/audit), status, and probe error overview

## Remediation-only run

- Generate from a specific inspection JSON:
   - .\\generate-remediation.ps1 -InspectionJsonPath "<path-to-inspection-json>"
- Auto pick latest inspection JSON under configured output path:
   - .\\generate-remediation.ps1
- Generate Chinese remediation report:
   - .\\generate-remediation.ps1 -ReportLanguage zh
- Localize status labels in remediation output by config or parameter:
   - .\\generate-remediation.ps1 -ReportLanguage zh -LocalizeJsonStatus

## Automation (non-interactive)

- Disable prompt/dialog for scheduled tasks or CI:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -NoFolderPrompt
   - .\\generate-remediation.ps1 -InspectionJsonPath "<path-to-inspection-json>" -NoFolderPrompt
- Override output path explicitly (no prompt):
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -OutputDir "D:\\download\\SQL\\manual"
   - .\\complete-inspection\\run-batch-inspection.ps1 -InstancesPath .\\complete-inspection\\instances.example.json -OutputRoot "D:\\download\\SQL\\manual"
   - .\\generate-remediation.ps1 -InspectionJsonPath "<path-to-inspection-json>" -OutputDir "D:\\download\\SQL\\manual"

## Complete Inspection (Merged)

- Run merged complete inspection for one instance:
   - .\\complete-inspection\\run-inspection.ps1 -Server db-host.example.com -Database master
- Run merged complete inspection in batch mode:
   - .\\complete-inspection\\run-batch-inspection.ps1 -InstancesPath .\\complete-inspection\\instances.example.json
- Check merged module details:
   - .\\complete-inspection\\README.md

### Server Login Config Item

- Real server config is external and stored in selected working folder:
   - <selected-working-folder>\\servers.config
- If missing, script auto-generates sample `servers.config` and asks user to fill actual server list before rerun.
- Source repository keeps sample only:
   - .\\complete-inspection\\config\\servers.sample.tsv
- Config file format has two sections:
   - `[global]` with `language` and `reportFormats`
   - `[servers]` tab-separated columns: `ip`, `port`, `username`, `password`

## Report formats

- Supported formats: json, html
- Single-run: use -ReportFormats
- Per-instance output: set reportFormats in instances.example.json

## Key checks in current version

- PLE and Buffer Cache Hit Ratio
- Blocking and longest running request
- Deadlocks from system_health (24h)
- Last full backup age and unprotected DB count
- Log usage and data file volume usage
- Failed jobs and disabled jobs
- AG unhealthy replicas (if AG enabled)
- Sysadmin membership count and detail
- Weak-policy SQL logins
- New server logins and database users in last 90 days

## Notes

- The merged workflow includes both lightweight metrics and extended full-audit checks.
- Heavy checks should still be scheduled off-peak for busy production systems.
- Some probes require VIEW SERVER STATE and access to msdb history tables.
- SqlClient mode uses native protocol negotiation and does not expose an explicit TDS setting.
- For legacy instances or protocol edge cases, use Odbc mode and set tdsVersion per instance.
- Suggested TDS hints:
   - SQL Server 2008/2008 R2: 7.1
   - SQL Server 2012/2014: 7.3
   - SQL Server 2016+: 7.4 (or keep empty and let driver negotiate)
