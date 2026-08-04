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

/* 4.1) Long-running request (current max elapsed time) */
SELECT
  ISNULL(MAX(total_elapsed_time / 1000.0), 0) AS longest_running_query_seconds
FROM sys.dm_exec_requests WITH (NOLOCK)
WHERE session_id <> @@SPID
  AND status <> 'background'
OPTION (RECOMPILE);

/* 4.2) Deadlock events in last 24h from system_health ring buffer */
;WITH rb AS (
  SELECT CAST(st.target_data AS XML) AS x
  FROM sys.dm_xe_session_targets st
  INNER JOIN sys.dm_xe_sessions s ON s.address = st.event_session_address
  WHERE s.name = 'system_health'
    AND st.target_name = 'ring_buffer'
)
SELECT
  COUNT(1) AS deadlocks_24h
FROM rb
CROSS APPLY x.nodes('//event[@name="xml_deadlock_report"]') AS d(e)
WHERE d.e.value('@timestamp', 'datetime2') >= DATEADD(HOUR, -24, SYSUTCDATETIME())
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

/* 5.1) Disabled SQL Agent jobs */
SELECT
  COUNT(1) AS disabled_jobs_count
FROM msdb.dbo.sysjobs WITH (NOLOCK)
WHERE enabled = 0
OPTION (RECOMPILE);

/* 6) Last full backup age */
SELECT
  MAX(backup_finish_date) AS last_full_backup_time
FROM msdb.dbo.backupset WITH (NOLOCK)
WHERE type = 'D'
OPTION (RECOMPILE);

/* 6.1) User DBs without recent full backup */
DECLARE @cutoff DATETIME = DATEADD(HOUR, -72, GETDATE());
SELECT
  COUNT(1) AS databases_without_recent_full_backup
FROM sys.databases d WITH (NOLOCK)
LEFT JOIN (
  SELECT database_name, MAX(backup_finish_date) AS last_full_backup
  FROM msdb.dbo.backupset WITH (NOLOCK)
  WHERE type = 'D'
  GROUP BY database_name
) b ON b.database_name = d.name
WHERE d.database_id > 4
  AND d.state_desc = 'ONLINE'
  AND (b.last_full_backup IS NULL OR b.last_full_backup < @cutoff)
OPTION (RECOMPILE);

/* 7) Data file volume usage */
SELECT
  MAX(
    CASE WHEN vs.total_bytes > 0
      THEN (1.0 - (vs.available_bytes * 1.0 / vs.total_bytes)) * 100
      ELSE NULL END
  ) AS data_disk_used_percent
FROM sys.master_files mf WITH (NOLOCK)
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs
WHERE mf.type_desc = 'ROWS'
OPTION (RECOMPILE);

/* 8) AG unhealthy replica count on local primary */
SELECT
  COUNT(1) AS ag_unhealthy_replicas
FROM sys.dm_hadr_availability_replica_states WITH (NOLOCK)
WHERE is_local = 1
  AND role_desc = 'PRIMARY'
  AND synchronization_health_desc <> 'HEALTHY'
OPTION (RECOMPILE);

/* 9) Sysadmin principal count */
SELECT
  COUNT(1) AS sysadmin_login_count
FROM sys.server_role_members srm WITH (NOLOCK)
INNER JOIN sys.server_principals r WITH (NOLOCK) ON r.principal_id = srm.role_principal_id
WHERE r.name = 'sysadmin'
OPTION (RECOMPILE);

/* 10) Weak-policy SQL login count */
SELECT
  COUNT(1) AS weak_policy_logins_count
FROM sys.sql_logins WITH (NOLOCK)
WHERE is_disabled = 0
  AND (is_policy_checked = 0 OR is_expiration_checked = 0)
OPTION (RECOMPILE);

/* 11) New server logins in last 90 days */
SELECT
  COUNT(1) AS new_logins_90d_count
FROM sys.server_principals WITH (NOLOCK)
WHERE type IN ('S', 'U', 'G')
  AND create_date >= DATEADD(DAY, -90, GETDATE())
  AND name NOT LIKE '##%'
OPTION (RECOMPILE);

/* 12) New database users in last 90 days */
CREATE TABLE #recent_users (
  database_name SYSNAME,
  user_name SYSNAME,
  type_desc NVARCHAR(60),
  create_date DATETIME
);

DECLARE @db SYSNAME;
DECLARE @sql NVARCHAR(MAX);
DECLARE db_cur CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases WITH (NOLOCK)
WHERE state_desc = 'ONLINE'
  AND database_id > 4;

OPEN db_cur;
FETCH NEXT FROM db_cur INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
  BEGIN TRY
    INSERT INTO #recent_users(database_name, user_name, type_desc, create_date)
    SELECT N''' + REPLACE(@db, '''', '''''') + ''', name, type_desc, create_date
    FROM ' + QUOTENAME(@db) + '.sys.database_principals WITH (NOLOCK)
    WHERE type IN (''S'', ''U'', ''G'')
      AND principal_id > 4
      AND create_date >= DATEADD(DAY, -90, GETDATE());
  END TRY
  BEGIN CATCH
  END CATCH;';

  EXEC sys.sp_executesql @sql;
  FETCH NEXT FROM db_cur INTO @db;
END

CLOSE db_cur;
DEALLOCATE db_cur;

SELECT COUNT(1) AS new_db_users_90d_count
FROM #recent_users
OPTION (RECOMPILE);
