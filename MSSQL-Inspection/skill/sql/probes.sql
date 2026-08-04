/* SQL-Server-Health-Sentinel probe library (MVP) */

/* 1) Instance basics */
SELECT
  @@SERVERNAME AS server_name,
  @@VERSION AS product_version,
  sqlserver_start_time
FROM sys.dm_os_sys_info
OPTION (RECOMPILE);

/* 2) Memory and cache */
SELECT
  MAX(CASE WHEN counter_name = 'Page life expectancy' THEN cntr_value END) AS ple_seconds,
  MAX(CASE WHEN counter_name = 'Buffer cache hit ratio' THEN cntr_value END) AS buffer_cache_hit_ratio
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name IN ('Page life expectancy', 'Buffer cache hit ratio')
OPTION (RECOMPILE);

/* 3) Database and log usage summary */
SELECT
  DB_NAME(database_id) AS database_name,
  type_desc,
  CAST(size * 8.0 / 1024 AS DECIMAL(18,2)) AS size_mb
FROM sys.master_files WITH (NOLOCK)
OPTION (RECOMPILE);

SELECT
  DB_NAME(database_id) AS database_name,
  total_log_size_in_bytes / 1024.0 / 1024.0 AS total_log_mb,
  used_log_space_in_bytes / 1024.0 / 1024.0 AS used_log_mb,
  used_log_space_in_percent
FROM sys.dm_db_log_space_usage
OPTION (RECOMPILE);

/* 4) Blocking snapshot */
SELECT
  COUNT(1) AS blocking_sessions
FROM sys.dm_exec_requests WITH (NOLOCK)
WHERE blocking_session_id <> 0
OPTION (RECOMPILE);

/* 5) Failed SQL Agent jobs in last 24h */
SELECT
  COUNT(1) AS failed_jobs_24h
FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
INNER JOIN msdb.dbo.sysjobs j WITH (NOLOCK) ON h.job_id = j.job_id
WHERE h.step_id = 0
  AND h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
OPTION (RECOMPILE);

/* 6) Last full backup age */
SELECT
  MAX(backup_finish_date) AS last_full_backup_time
FROM msdb.dbo.backupset WITH (NOLOCK)
WHERE type = 'D'
OPTION (RECOMPILE);
