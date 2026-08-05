WITH backup_rollup AS (
  SELECT
    b.database_name,
    MAX(CASE WHEN b.type = 'D' THEN b.backup_finish_date END) AS last_full_backup,
    MAX(CASE WHEN b.type = 'I' THEN b.backup_finish_date END) AS last_diff_backup,
    MAX(CASE WHEN b.type = 'L' THEN b.backup_finish_date END) AS last_log_backup
  FROM msdb.dbo.backupset AS b WITH (NOLOCK)
  GROUP BY b.database_name
)
SELECT
  d.name AS database_name,
  d.recovery_model_desc,
  br.last_full_backup,
  br.last_diff_backup,
  br.last_log_backup,
  CASE WHEN br.last_full_backup IS NULL THEN NULL ELSE DATEDIFF(HOUR, br.last_full_backup, GETDATE()) END AS full_backup_age_hours,
  CASE WHEN br.last_diff_backup IS NULL THEN NULL ELSE DATEDIFF(HOUR, br.last_diff_backup, GETDATE()) END AS diff_backup_age_hours,
  CASE WHEN br.last_log_backup IS NULL THEN NULL ELSE DATEDIFF(HOUR, br.last_log_backup, GETDATE()) END AS log_backup_age_hours,
  CASE
    WHEN br.last_full_backup IS NULL OR DATEDIFF(HOUR, br.last_full_backup, GETDATE()) > 168 THEN 'critical'
    WHEN DATEDIFF(HOUR, br.last_full_backup, GETDATE()) > 48 THEN 'warning'
    ELSE 'ok'
  END AS full_backup_status,
  CASE
    WHEN br.last_diff_backup IS NULL THEN 'warning'
    WHEN DATEDIFF(HOUR, br.last_diff_backup, GETDATE()) > 72 THEN 'critical'
    WHEN DATEDIFF(HOUR, br.last_diff_backup, GETDATE()) > 24 THEN 'warning'
    ELSE 'ok'
  END AS diff_backup_status,
  CASE
    WHEN d.recovery_model_desc IN ('SIMPLE') THEN 'not_applicable'
    WHEN br.last_log_backup IS NULL OR DATEDIFF(HOUR, br.last_log_backup, GETDATE()) > 24 THEN 'critical'
    WHEN DATEDIFF(HOUR, br.last_log_backup, GETDATE()) > 4 THEN 'warning'
    ELSE 'ok'
  END AS log_backup_status
FROM sys.databases AS d WITH (NOLOCK)
LEFT JOIN backup_rollup AS br ON br.database_name = d.name
WHERE d.source_database_id IS NULL
  AND d.database_id > 4
ORDER BY d.name
OPTION (RECOMPILE);
