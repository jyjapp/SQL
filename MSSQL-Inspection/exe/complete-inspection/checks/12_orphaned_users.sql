CREATE TABLE #orphaned_users (
  database_name SYSNAME,
  user_name SYSNAME,
  user_type_desc NVARCHAR(60),
  default_schema_name SYSNAME,
  create_date DATETIME
);

DECLARE @db SYSNAME;
DECLARE @sql NVARCHAR(MAX);
DECLARE db_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases WITH (NOLOCK)
WHERE database_id > 4
  AND state_desc = 'ONLINE'
  AND source_database_id IS NULL;

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
  BEGIN TRY
    INSERT INTO #orphaned_users(database_name, user_name, user_type_desc, default_schema_name, create_date)
    SELECT
      N''' + REPLACE(@db, '''', '''''') + ''',
      dp.name,
      dp.type_desc,
      dp.default_schema_name,
      dp.create_date
    FROM ' + QUOTENAME(@db) + '.sys.database_principals AS dp WITH (NOLOCK)
    LEFT JOIN master.sys.server_principals AS sp WITH (NOLOCK) ON dp.sid = sp.sid
    WHERE dp.authentication_type <> 0
      AND dp.principal_id > 4
      AND dp.sid IS NOT NULL
      AND dp.type IN (''S'', ''U'', ''G'')
      AND sp.sid IS NULL
      AND dp.name NOT IN (''guest'', ''INFORMATION_SCHEMA'', ''sys'');
  END TRY
  BEGIN CATCH
  END CATCH;';

  EXEC sys.sp_executesql @sql;
  FETCH NEXT FROM db_cursor INTO @db;
END

CLOSE db_cursor;
DEALLOCATE db_cursor;

SELECT
  database_name,
  user_name,
  user_type_desc,
  default_schema_name,
  create_date
FROM #orphaned_users
ORDER BY database_name, user_name
OPTION (RECOMPILE);
