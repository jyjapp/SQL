SELECT
  j.name AS job_name,
  msdb.dbo.agent_datetime(h.run_date, h.run_time) AS run_time,
  h.run_status,
  h.retries_attempted,
  h.message,
  CASE
    WHEN h.run_duration <= 0 THEN 0
    ELSE ((h.run_duration / 10000) * 3600)
      + (((h.run_duration % 10000) / 100) * 60)
      + (h.run_duration % 100)
  END AS duration_seconds
FROM msdb.dbo.sysjobhistory AS h WITH (NOLOCK)
INNER JOIN msdb.dbo.sysjobs AS j WITH (NOLOCK) ON j.job_id = h.job_id
WHERE h.step_id = 0
  AND h.run_status IN (0, 2, 3)
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
ORDER BY run_time DESC
OPTION (RECOMPILE);
