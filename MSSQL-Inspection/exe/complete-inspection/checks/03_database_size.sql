SELECT
  DB_NAME(mf.database_id) AS database_name,
  SUM(CASE WHEN mf.type_desc = 'ROWS' THEN mf.size END) * 8.0 / 1024 AS data_size_mb,
  SUM(CASE WHEN mf.type_desc = 'LOG' THEN mf.size END) * 8.0 / 1024 AS log_size_mb,
  SUM(mf.size) * 8.0 / 1024 AS total_size_mb
FROM sys.master_files AS mf WITH (NOLOCK)
INNER JOIN sys.databases AS d WITH (NOLOCK) ON d.database_id = mf.database_id
WHERE d.source_database_id IS NULL
  AND (d.database_id > 4 OR d.name = 'tempdb')
GROUP BY DB_NAME(mf.database_id)
ORDER BY total_size_mb DESC
OPTION (RECOMPILE);
