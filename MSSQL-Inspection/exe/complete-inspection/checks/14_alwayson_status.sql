IF SERVERPROPERTY('IsHadrEnabled') = 1
BEGIN
  SELECT
    ag.name AS ag_name,
    ar.replica_server_name,
    ars.role_desc,
    ars.operational_state_desc,
    ars.connected_state_desc,
    ars.synchronization_health_desc,
    drs.database_id,
    DB_NAME(drs.database_id) AS database_name,
    drs.synchronization_state_desc,
    drs.is_suspended
  FROM sys.availability_groups AS ag WITH (NOLOCK)
  INNER JOIN sys.availability_replicas AS ar WITH (NOLOCK) ON ar.group_id = ag.group_id
  INNER JOIN sys.dm_hadr_availability_replica_states AS ars WITH (NOLOCK) ON ars.replica_id = ar.replica_id
  LEFT JOIN sys.dm_hadr_database_replica_states AS drs WITH (NOLOCK)
    ON drs.replica_id = ars.replica_id
   AND drs.group_id = ag.group_id
  ORDER BY ag.name, ar.replica_server_name, database_name
  OPTION (RECOMPILE);
END
ELSE
BEGIN
  SELECT
    CAST(NULL AS NVARCHAR(128)) AS ag_name,
    CAST(NULL AS NVARCHAR(128)) AS replica_server_name,
    CAST('NOT_ENABLED' AS NVARCHAR(32)) AS role_desc,
    CAST(NULL AS NVARCHAR(32)) AS operational_state_desc,
    CAST(NULL AS NVARCHAR(32)) AS connected_state_desc,
    CAST(NULL AS NVARCHAR(32)) AS synchronization_health_desc,
    CAST(NULL AS INT) AS database_id,
    CAST(NULL AS NVARCHAR(128)) AS database_name,
    CAST(NULL AS NVARCHAR(32)) AS synchronization_state_desc,
    CAST(NULL AS BIT) AS is_suspended;
END
