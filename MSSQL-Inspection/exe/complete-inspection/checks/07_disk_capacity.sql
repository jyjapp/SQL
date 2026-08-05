SELECT
  vs.volume_mount_point,
  MAX(vs.logical_volume_name) AS logical_volume_name,
  MAX(vs.file_system_type) AS file_system_type,
  MAX(vs.total_bytes) / 1024.0 / 1024.0 / 1024.0 AS total_gb,
  MAX(vs.available_bytes) / 1024.0 / 1024.0 / 1024.0 AS free_gb,
  CASE
    WHEN MAX(vs.total_bytes) > 0
      THEN (MAX(vs.available_bytes) * 100.0 / MAX(vs.total_bytes))
    ELSE NULL
  END AS free_percent
FROM sys.master_files AS mf WITH (NOLOCK)
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) AS vs
GROUP BY vs.volume_mount_point
ORDER BY free_percent ASC
OPTION (RECOMPILE);
