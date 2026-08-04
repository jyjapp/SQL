# SQL-Server-Health-Sentinel Skill Spec

## 1. Goal

Build a lightweight, pluggable SQL Server inspection skill that collects health signals with low production impact and produces structured outputs for operations teams.

## 2. Architecture

- Collector layer:
  - T-SQL probes based on DMVs and system tables
  - Optional OS/service probes via PowerShell
- Processor layer:
  - Rule engine compares metrics against configurable thresholds
  - Status labels: Good, Warning, Critical
- Reporter layer:
  - JSON output for platform integration
  - HTML report for human review

## 3. Modules

### Module A: Connection and privilege

- Support Windows auth and SQL auth
- Batch mode from instance configuration file
- Query timeout and fail-fast controls

### Module B: T-SQL probe library

- Probes stored in skill/sql/probes.sql
- Read-first strategy; avoid high-impact operations in peak hours
- Use short-running DMV queries for online checks

### Module C: Rules engine

- Thresholds in skill/rules/thresholds.json
- Per-metric mapping to Good/Warning/Critical
- Customizable to business SLA

### Module D: Report renderer

- Raw results persisted as JSON
- Summary HTML with status cards

## 4. Data model

Each metric output should include:

- metricKey
- category
- value
- unit
- threshold
- status
- collectedAt
- sourceProbe

## 5. Risk controls

- Schedule heavy checks (CHECKDB, full fragmentation scans) in off-peak windows
- Enforce query timeout to avoid global task stalls
- Use least privilege service account whenever possible
- Grant only required permissions such as VIEW SERVER STATE for most probes

## 6. Roadmap

- Phase 1 (MVP): single instance, JSON + base HTML report
- Phase 2: batch mode, security deep checks, Excel summary
- Phase 3: REST API service, scheduling, alert integrations
