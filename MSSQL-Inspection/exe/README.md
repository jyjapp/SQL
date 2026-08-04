# Executables

## Script

- run-inspection.ps1: single-instance MVP inspector
- run-batch-inspection.ps1: multi-instance batch runner with summary report

## Quick start

1. Open PowerShell in this folder.
2. Run with Windows auth:
   - .\\run-inspection.ps1 -Server 10.0.0.10 -Database master
3. Run with SQL auth:
   - .\\run-inspection.ps1 -Server 10.0.0.10 -Database master -Username sa -Password "your_password"
4. Run with explicit TDS for legacy targets (ODBC mode):
   - .\\run-inspection.ps1 -Server 10.0.0.55 -Database master -Username sa -Password "your_password" -ConnectionProvider Odbc -TdsVersion 7.1
5. Run with automatic fallback (recommended for mixed-version estates):
   - .\\run-inspection.ps1 -Server 10.0.0.55 -Database master -Username sa -Password "your_password" -ConnectionProvider SqlClient -EnableFallback $true -FallbackTdsVersions "7.4,7.3,7.2,7.1,7.0"
6. Choose report formats:
   - .\\run-inspection.ps1 -Server 10.0.0.10 -ReportFormats "json,html,md"
   - .\\run-inspection.ps1 -Server 10.0.0.10 -ReportFormats "json,pdf"

## Output

- JSON file in .\\output
- HTML file in .\\output

## Batch run

1. Edit instances.example.json and add multiple servers.
2. Run:
   - .\\run-batch-inspection.ps1 -InstancesPath .\\instances.example.json
   - .\\run-batch-inspection.ps1 -InstancesPath .\\instances.example.json -SummaryFormats "json,html,md"
3. Review batch output:
   - .\\output\\batch-<timestamp>\\batch-summary.json
   - .\\output\\batch-<timestamp>\\batch-summary.html
   - Includes risk ranking, top-risk list, and suggested remediation actions

## Report formats

- Supported formats: json, html, md, docx, pdf
- Single-run: use -ReportFormats
- Batch summary: use -SummaryFormats
- Per-instance output: set reportFormats in instances.example.json
- docx/pdf export requires Microsoft Word on Windows (COM automation)

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

- This MVP focuses on lightweight online checks.
- Heavy checks (CHECKDB and deep index analysis) should be scheduled off-peak.
- Some probes require VIEW SERVER STATE and access to msdb history tables.
- SqlClient mode uses native protocol negotiation and does not expose an explicit TDS setting.
- For legacy instances or protocol edge cases, use Odbc mode and set tdsVersion per instance.
- If initial connection fails, fallback mode retries ODBC with ordered TDS candidates until one succeeds.
- Suggested TDS hints:
   - SQL Server 2008/2008 R2: 7.1
   - SQL Server 2012/2014: 7.3
   - SQL Server 2016+: 7.4 (or keep empty and let driver negotiate)
