CREATE TABLE #frag (
  database_name SYSNAME,
  schema_name SYSNAME,
  table_name SYSNAME,
  index_name SYSNAME,
  page_count BIGINT,
  avg_fragmentation_percent DECIMAL(9,2)
);

DECLARE @db SYSNAME;
DECLARE @sql NVARCHAR(MAX);
DECLARE db_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT TOP (3) d.name
FROM sys.databases d WITH (NOLOCK)
LEFT JOIN (
  SELECT database_id, SUM(size) AS total_pages
  FROM sys.master_files WITH (NOLOCK)
  GROUP BY database_id
) mf ON mf.database_id = d.database_id
WHERE d.database_id > 4
  AND d.state_desc = 'ONLINE'
  AND d.source_database_id IS NULL
ORDER BY ISNULL(mf.total_pages, 0) DESC, d.name ASC;

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
  BEGIN TRY
    ;WITH idx_candidates AS (
      SELECT TOP (30)
        ps.object_id,
        ps.index_id,
        SUM(ps.used_page_count) AS used_page_count
      FROM ' + QUOTENAME(@db) + '.sys.dm_db_partition_stats ps WITH (NOLOCK)
      INNER JOIN ' + QUOTENAME(@db) + '.sys.indexes i0 WITH (NOLOCK) ON i0.object_id = ps.object_id AND i0.index_id = ps.index_id
      INNER JOIN ' + QUOTENAME(@db) + '.sys.tables t0 WITH (NOLOCK) ON t0.object_id = i0.object_id
      WHERE i0.index_id > 0
      GROUP BY ps.object_id, ps.index_id
      HAVING SUM(ps.used_page_count) > 1000
      ORDER BY SUM(ps.used_page_count) DESC
    )
    INSERT INTO #frag(database_name, schema_name, table_name, index_name, page_count, avg_fragmentation_percent)
    SELECT
      N''' + REPLACE(@db, '''', '''''') + ''',
      sc.name,
      t.name,
      i.name,
      ips.page_count,
      CAST(ips.avg_fragmentation_in_percent AS DECIMAL(9,2))
    FROM idx_candidates c
    CROSS APPLY ' + QUOTENAME(@db) + '.sys.dm_db_index_physical_stats(DB_ID(N''' + REPLACE(@db, '''', '''''') + '''), c.object_id, c.index_id, NULL, ''LIMITED'') ips
    INNER JOIN ' + QUOTENAME(@db) + '.sys.indexes i ON i.object_id = ips.object_id AND i.index_id = ips.index_id
    INNER JOIN ' + QUOTENAME(@db) + '.sys.tables t ON t.object_id = i.object_id
    INNER JOIN ' + QUOTENAME(@db) + '.sys.schemas sc ON sc.schema_id = t.schema_id
    WHERE ips.avg_fragmentation_in_percent > 10
      AND i.index_id > 0;
  END TRY
  BEGIN CATCH
  END CATCH;';

  EXEC sys.sp_executesql @sql;
  FETCH NEXT FROM db_cursor INTO @db;
END

CLOSE db_cursor;
DEALLOCATE db_cursor;

SELECT TOP (200)
  database_name,
  schema_name,
  table_name,
  index_name,
  page_count,
  avg_fragmentation_percent
FROM #frag
ORDER BY avg_fragmentation_percent DESC, page_count DESC
OPTION (RECOMPILE);
