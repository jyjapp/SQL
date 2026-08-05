# MSSQL Inspection

SQL-Server-Health-Sentinel MVP for SQL Server 2016+.

This module focuses on systematic inspection across six dimensions:

1. Instance and service health
2. Resource and performance bottlenecks
3. Data safety and integrity
4. Security and permission audit
5. Automation and job monitoring
6. High availability

## Folder layout

- skill/: design, probe definitions, thresholds
- exe/: runnable scripts and runtime examples

## Output Directory Configuration

Inspection output location is controlled by `skill/config.json`:

```json
{
	"output": {
		"defaultRootDir": "D:\\download\\SQL",
		"promptForFolder": true,
		"rememberSelectedRootAsDefault": true,
		"useDateSubfolder": true,
		"dateFolderFormat": "yyyyMMdd"
	},
	"report": {
		"language": "en-US",
		"localizeJsonStatus": false
	}
}
```

- `defaultRootDir`: default output root directory
- `promptForFolder`: prompt for one working folder at runtime when folder parameters are not explicitly provided
- `rememberSelectedRootAsDefault`: persist your selected folder as next run default
- `useDateSubfolder`: create date-based subfolder automatically
- `dateFolderFormat`: date folder naming format used by `Get-Date -Format`
- `report.language`: report language, supports `en-US` and `zh-CN`
- `report.localizeJsonStatus`: when true, JSON status fields are localized by report language

The same `report` settings are used by complete inspection reports and remediation reports.

When running `exe/complete-inspection/run-inspection.ps1` without explicit folder parameters, the script first tries Windows folder picker and falls back to command-line input. Press Enter keeps the configured default.

The selected folder is shared by:

- report output root (with date subfolder)
- server config file `servers.config`

If `servers.config` does not exist in the selected folder, script generates a sample and prompts the user to fill actual server inventory before rerun.

Server config format is split into two sections:

- `[global]`: `language` and `reportFormats`
- `[servers]`: tab-separated columns `ip`, `port`, `username`, `password`

All execution scripts now print the resolved output directory/root at runtime for easier verification.

## Current MVP coverage

- Version and uptime
- Memory and PLE basics
- Database/log file capacity summary
- Failed SQL Agent jobs in last 24 hours
- Last backup time summary
- Blocking snapshot

## Complete Merged Inspection

The merged and complete inspection skill is available in `exe/complete-inspection/`.
It combines the original MVP inspection pipeline and the extended full-audit feature set into one unified execution and reporting flow.

## Next milestones

- Add full index fragmentation and CHECKDB evidence chain
- Add AG/FCI deep checks
- Add export to Excel and richer HTML cards
- Add scheduler and notification integration
