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

## Current MVP coverage

- Version and uptime
- Memory and PLE basics
- Database/log file capacity summary
- Failed SQL Agent jobs in last 24 hours
- Last backup time summary
- Blocking snapshot

## Next milestones

- Add full index fragmentation and CHECKDB evidence chain
- Add AG/FCI deep checks
- Add export to Excel and richer HTML cards
- Add scheduler and notification integration
