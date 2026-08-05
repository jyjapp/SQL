SELECT
  COUNT(1) AS total_sessions,
  SUM(CASE WHEN s.status = 'running' THEN 1 ELSE 0 END) AS running_sessions,
  SUM(CASE WHEN s.status = 'sleeping' THEN 1 ELSE 0 END) AS sleeping_sessions,
  SUM(CASE WHEN r.blocking_session_id <> 0 THEN 1 ELSE 0 END) AS blocked_requests,
  SUM(CASE WHEN s.is_user_process = 1 THEN 1 ELSE 0 END) AS user_sessions
FROM sys.dm_exec_sessions AS s WITH (NOLOCK)
LEFT JOIN sys.dm_exec_requests AS r WITH (NOLOCK) ON r.session_id = s.session_id
WHERE s.session_id <> @@SPID
OPTION (RECOMPILE);
