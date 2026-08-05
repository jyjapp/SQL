SELECT
  sp.name,
  sp.type_desc,
  sp.is_disabled,
  sp.default_database_name,
  sp.create_date,
  CASE WHEN srp.role_principal_id IS NULL THEN 0 ELSE 1 END AS is_sysadmin
FROM sys.server_principals AS sp WITH (NOLOCK)
LEFT JOIN sys.server_role_members AS srp WITH (NOLOCK)
  ON srp.member_principal_id = sp.principal_id
LEFT JOIN sys.server_principals AS rolep WITH (NOLOCK)
  ON rolep.principal_id = srp.role_principal_id
 AND rolep.name = 'sysadmin'
WHERE sp.type IN ('S', 'U', 'G')
  AND sp.name NOT LIKE '##%'
ORDER BY sp.name
OPTION (RECOMPILE);
