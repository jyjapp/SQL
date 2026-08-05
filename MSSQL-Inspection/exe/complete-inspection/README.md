# SQL Server Complete Inspection

This module is the single merged inspection implementation.
It combines the original MVP metric pipeline and the extended full-audit checks.

## Scripts

- run-inspection.ps1: single-instance complete inspection
- run-batch-inspection.ps1: batch complete inspection
- instances.example.json: batch input template

## Module Layout

- checks/: feature-based SQL checks (15 groups)
- config/score-rules.json: scoring configuration
- config/servers.sample.tsv: sample only (do not store real credentials in repository)

## Dependencies

- Threshold rules: ../../skill/rules/thresholds.json
- Output folder configuration: ../../skill/config.json
- Report language configuration: ../../skill/config.json (report.language)

## Report Language

Set report language in `../../skill/config.json`:

```json
"report": {
	"language": "en-US",
	"localizeJsonStatus": false
}
```

Supported values:

- `en-US` (English)
- `zh-CN` (Chinese)
- aliases also supported: `en`, `english`, `zh`, `cn`, `zh-hans`

`localizeJsonStatus`:

- `false`: JSON keeps status values as `Good/Warning/Critical`
- `true`: JSON status values follow report language (for Chinese: `良好/警告/严重`)

You can also override at runtime:

```powershell
powershell -File .\run-inspection.ps1 -WorkspaceDir "D:\download\SQL" -ReportLanguage zh-CN
powershell -File .\run-batch-inspection.ps1 -WorkspaceDir "D:\download\SQL" -ReportLanguage en-US
powershell -File .\run-inspection.ps1 -WorkspaceDir "D:\download\SQL" -ReportLanguage zh -LocalizeJsonStatus
```

## Server Configuration Item

At runtime, scripts prompt user to select ONE working folder.
This folder is used for both:

- report output root (date subfolder under this root)
- server config file (`servers.config` in this same root)

If `servers.config` is missing in the selected folder, script auto-generates a sample and exits with a message:

- "Sample server list generated: <selected-folder>\\servers.config"
- "Please fill server config with actual server list, then rerun this skill."

Source code keeps sample only:

- `config/servers.sample.tsv`

`servers.config` format (two sections):

```text
[global]
language	zh-CN
reportFormats	json,html

[servers]
ip	port	username	password
db-host.example.com	1433	sql_user	<PASSWORD>
db-host-2.example.com	1433		
```

- `[global]`: global defaults for language and report formats.
- `[servers]`: tab-separated rows with columns `ip`, `port`, `username`, `password`.

## Single Instance

```powershell
powershell -File .\run-inspection.ps1 -Server db-host.example.com -Database master
```

SQL auth example:

```powershell
powershell -File .\run-inspection.ps1 -Server db-host.example.com -Database master -Username sql_user -Password "<PASSWORD>"
```

Use config item directly:

```powershell
powershell -File .\run-inspection.ps1
```

Specify one shared working folder explicitly:

```powershell
powershell -File .\run-inspection.ps1 -WorkspaceDir "D:\download\SQL"
```

## Batch

```powershell
powershell -File .\run-batch-inspection.ps1 -InstancesPath .\instances.example.json
```

Use batch config item directly:

```powershell
powershell -File .\run-batch-inspection.ps1
```

Specify external server-config directory explicitly:

```powershell
powershell -File .\run-batch-inspection.ps1 -WorkspaceDir "D:\download\SQL"
```

## Outputs

- Default root folder: configured by skill/config.json
- Date folder: automatically created (yyyyMMdd)
- Single report files: complete-inspection-<timestamp>.json/.html
- Batch summary files: complete-inspection-batch-summary.json/.html
