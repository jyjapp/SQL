SELECT
  r.session_id,
  s.login_name,
  s.host_name,
  DB_NAME(r.database_id) AS database_name,
  r.status,
  r.command,
  r.wait_type,
  r.wait_time / 1000.0 AS wait_seconds,
  r.cpu_time / 1000.0 AS cpu_seconds,
  r.total_elapsed_time / 1000.0 AS elapsed_seconds,
  LEFT(REPLACE(REPLACE(txt.text, CHAR(13), ' '), CHAR(10), ' '), 400) AS query_text
FROM sys.dm_exec_requests AS r WITH (NOLOCK)
INNER JOIN sys.dm_exec_sessions AS s WITH (NOLOCK) ON s.session_id = r.session_id
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) AS txt
WHERE r.session_id <> @@SPID
  AND r.total_elapsed_time >= 300000
ORDER BY r.total_elapsed_time DESC
OPTION (RECOMPILE);
