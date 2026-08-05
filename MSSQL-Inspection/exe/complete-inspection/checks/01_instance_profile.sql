SELECT
  @@SERVERNAME AS server_name,
  CAST(SERVERPROPERTY('MachineName') AS NVARCHAR(128)) AS machine_name,
  CAST(SERVERPROPERTY('Edition') AS NVARCHAR(128)) AS edition,
  CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS product_version,
  CAST(SERVERPROPERTY('ProductLevel') AS NVARCHAR(128)) AS product_level,
  CAST(SERVERPROPERTY('EngineEdition') AS INT) AS engine_edition,
  osi.sqlserver_start_time,
  DATEDIFF(MINUTE, osi.sqlserver_start_time, GETDATE()) AS uptime_minutes
FROM sys.dm_os_sys_info AS osi
OPTION (RECOMPILE);
