SELECT TOP (30)
  DB_NAME(mid.database_id) AS database_name,
  OBJECT_SCHEMA_NAME(mid.object_id, mid.database_id) AS schema_name,
  OBJECT_NAME(mid.object_id, mid.database_id) AS table_name,
  migs.user_seeks,
  migs.user_scans,
  migs.avg_total_user_cost,
  migs.avg_user_impact,
  CAST((migs.user_seeks + migs.user_scans) * migs.avg_total_user_cost * (migs.avg_user_impact / 100.0) AS DECIMAL(18,2)) AS improvement_score,
  mid.equality_columns,
  mid.inequality_columns,
  mid.included_columns
FROM sys.dm_db_missing_index_group_stats AS migs WITH (NOLOCK)
INNER JOIN sys.dm_db_missing_index_groups AS mig WITH (NOLOCK)
  ON migs.group_handle = mig.index_group_handle
INNER JOIN sys.dm_db_missing_index_details AS mid WITH (NOLOCK)
  ON mig.index_handle = mid.index_handle
WHERE mid.database_id > 4
ORDER BY improvement_score DESC
OPTION (RECOMPILE);
