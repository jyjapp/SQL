SELECT
  d.name AS database_name,
  d.database_id,
  d.state_desc,
  d.recovery_model_desc,
  d.user_access_desc,
  d.compatibility_level,
  d.create_date,
  CASE WHEN d.source_database_id IS NULL THEN 0 ELSE 1 END AS is_snapshot,
  CASE
    WHEN d.database_id <= 4 THEN 'system'
    WHEN d.name = 'tempdb' THEN 'tempdb'
    ELSE 'user'
  END AS database_class
FROM sys.databases AS d WITH (NOLOCK)
WHERE d.source_database_id IS NULL
  AND (d.database_id > 4 OR d.name = 'tempdb')
ORDER BY d.name
OPTION (RECOMPILE);
