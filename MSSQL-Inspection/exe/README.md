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

## Output

- JSON file in .\\output
- HTML file in .\\output

## Batch run

1. Edit instances.example.json and add multiple servers.
2. Run:
   - .\\run-batch-inspection.ps1 -InstancesPath .\\instances.example.json
3. Review batch output:
   - .\\output\\batch-<timestamp>\\batch-summary.json
   - .\\output\\batch-<timestamp>\\batch-summary.html

## Key checks in current version

- PLE and Buffer Cache Hit Ratio
- Blocking and longest running request
- Deadlocks from system_health (24h)
- Last full backup age and unprotected DB count
- Log usage and data file volume usage
- Failed jobs and disabled jobs
- AG unhealthy replicas (if AG enabled)

## Notes

- This MVP focuses on lightweight online checks.
- Heavy checks (CHECKDB and deep index analysis) should be scheduled off-peak.
- Some probes require VIEW SERVER STATE and access to msdb history tables.
