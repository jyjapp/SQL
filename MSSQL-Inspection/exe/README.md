# Executables

## Script

- run-inspection.ps1: single-instance MVP inspector

## Quick start

1. Open PowerShell in this folder.
2. Run with Windows auth:
   - .\\run-inspection.ps1 -Server 10.0.0.10 -Database master
3. Run with SQL auth:
   - .\\run-inspection.ps1 -Server 10.0.0.10 -Database master -Username sa -Password "your_password"

## Output

- JSON file in .\\output
- HTML file in .\\output

## Notes

- This MVP focuses on lightweight online checks.
- Heavy checks (CHECKDB and deep index analysis) should be scheduled off-peak.
