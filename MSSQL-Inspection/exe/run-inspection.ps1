param(
  [string]$Server = "127.0.0.1",
  [string]$Database = "master",
  [string]$Username = "",
  [string]$Password = "",
  [int]$Port = 1433,
  [int]$TimeoutSeconds = 15,
  [ValidateSet("SqlClient", "Odbc")]
  [string]$ConnectionProvider = "SqlClient",
  [string]$TdsVersion = "",
  [string]$OdbcDriver = "ODBC Driver 18 for SQL Server",
  [bool]$EnableFallback = $true,
  [string]$FallbackTdsVersions = "7.4,7.3,7.2,7.1,7.0",
  [string]$ReportFormats = "json,html",
  [string]$ThresholdPath = "..\\skill\\rules\\thresholds.json",
  [string]$OutputDir = ".\\output"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-ReportFormats {
  param([string]$ReportFormats)

  $allowed = @("json", "html", "md", "docx", "pdf")
  $normalized = New-Object System.Collections.Generic.List[string]
  foreach ($token in ($ReportFormats -split ',')) {
    $fmt = $token.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($fmt)) {
      continue
    }
    if ($allowed -contains $fmt -and -not $normalized.Contains($fmt)) {
      $normalized.Add($fmt)
    }
  }

  if ($normalized.Count -eq 0) {
    $normalized.Add("json")
    $normalized.Add("html")
  }

  return $normalized
}

function Convert-HtmlToOffice {
  param(
    [string]$HtmlPath,
    [string]$DocxPath,
    [string]$PdfPath,
    [bool]$ExportDocx,
    [bool]$ExportPdf
  )

  $result = [ordered]@{
    docxOk = $false
    pdfOk = $false
    warning = $null
  }

  if (-not $ExportDocx -and -not $ExportPdf) {
    return $result
  }

  $word = $null
  $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open((Resolve-Path $HtmlPath).Path)

    if ($ExportDocx) {
      $doc.SaveAs((Resolve-Path (Split-Path -Parent $DocxPath)).Path + "\\" + (Split-Path -Leaf $DocxPath), 16)
      $result.docxOk = $true
    }
    if ($ExportPdf) {
      $doc.SaveAs((Resolve-Path (Split-Path -Parent $PdfPath)).Path + "\\" + (Split-Path -Leaf $PdfPath), 17)
      $result.pdfOk = $true
    }
  }
  catch {
    $result.warning = "docx/pdf export skipped: " + $_.Exception.Message
  }
  finally {
    if ($null -ne $doc) {
      try { $doc.Close() } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
    }
    if ($null -ne $word) {
      try { $word.Quit() } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  }

  return $result
}

function Convert-ToNumericOrNull {
  param([object]$Value)

  if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) {
    return $null
  }

  $parsed = 0.0
  if ([double]::TryParse($Value.ToString(), [ref]$parsed)) {
    return [double]$parsed
  }

  return $null
}

function Get-ConnectionString {
  param(
    [string]$Server,
    [string]$Database,
    [string]$Username,
    [string]$Password,
    [int]$Port,
    [string]$ConnectionProvider,
    [string]$TdsVersion,
    [string]$OdbcDriver
  )

  if ($ConnectionProvider -eq "Odbc") {
    $authPart = ""
    if ([string]::IsNullOrWhiteSpace($Username)) {
      $authPart = "Trusted_Connection=Yes;"
    }
    else {
      $authPart = "UID=$Username;PWD=$Password;"
    }

    $tdsPart = ""
    if (-not [string]::IsNullOrWhiteSpace($TdsVersion)) {
      $tdsPart = "TDS_Version=$TdsVersion;"
    }

    return "Driver={$OdbcDriver};Server=$Server;Port=$Port;Database=$Database;$authPart$tdsPartEncrypt=No;TrustServerCertificate=Yes;"
  }

  if ([string]::IsNullOrWhiteSpace($Username)) {
    return "Server=$Server,$Port;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
  }

  return "Server=$Server,$Port;Database=$Database;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
}

function New-DbConnection {
  param(
    [string]$ConnectionString,
    [string]$ConnectionProvider
  )

  if ($ConnectionProvider -eq "Odbc") {
    return (New-Object System.Data.Odbc.OdbcConnection($ConnectionString))
  }

  return (New-Object System.Data.SqlClient.SqlConnection($ConnectionString))
}

function Get-TdsFallbackList {
  param(
    [string]$FallbackTdsVersions,
    [string]$PreferredTdsVersion
  )

  $list = New-Object System.Collections.Generic.List[string]

  if (-not [string]::IsNullOrWhiteSpace($PreferredTdsVersion)) {
    $list.Add($PreferredTdsVersion.Trim())
  }

  foreach ($item in ($FallbackTdsVersions -split ',')) {
    $tds = $item.Trim()
    if (-not [string]::IsNullOrWhiteSpace($tds) -and -not $list.Contains($tds)) {
      $list.Add($tds)
    }
  }

  return $list
}

function Get-ConnectionAttempts {
  param(
    [string]$Server,
    [string]$Database,
    [string]$Username,
    [string]$Password,
    [int]$Port,
    [string]$ConnectionProvider,
    [string]$TdsVersion,
    [string]$OdbcDriver,
    [bool]$EnableFallback,
    [string]$FallbackTdsVersions
  )

  $attempts = New-Object System.Collections.Generic.List[object]
  $tdsCandidates = Get-TdsFallbackList -FallbackTdsVersions $FallbackTdsVersions -PreferredTdsVersion $TdsVersion

  if ($ConnectionProvider -eq "SqlClient") {
    $attempts.Add([ordered]@{
      provider = "SqlClient"
      tdsVersion = $null
      connectionString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -ConnectionProvider "SqlClient" -TdsVersion "" -OdbcDriver $OdbcDriver
      label = "SqlClient(auto-negotiation)"
    })

    if ($EnableFallback) {
      foreach ($tds in $tdsCandidates) {
        $attempts.Add([ordered]@{
          provider = "Odbc"
          tdsVersion = $tds
          connectionString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -ConnectionProvider "Odbc" -TdsVersion $tds -OdbcDriver $OdbcDriver
          label = "Odbc(TDS=$tds)"
        })
      }
    }

    return $attempts
  }

  if ($tdsCandidates.Count -eq 0) {
    $tdsCandidates.Add("")
  }

  $firstTds = $tdsCandidates[0]
  $attempts.Add([ordered]@{
    provider = "Odbc"
    tdsVersion = if ([string]::IsNullOrWhiteSpace($firstTds)) { $null } else { $firstTds }
    connectionString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -ConnectionProvider "Odbc" -TdsVersion $firstTds -OdbcDriver $OdbcDriver
    label = if ([string]::IsNullOrWhiteSpace($firstTds)) { "Odbc(auto-negotiation)" } else { "Odbc(TDS=$firstTds)" }
  })

  if ($EnableFallback) {
    foreach ($tds in $tdsCandidates | Select-Object -Skip 1) {
      $attempts.Add([ordered]@{
        provider = "Odbc"
        tdsVersion = $tds
        connectionString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -ConnectionProvider "Odbc" -TdsVersion $tds -OdbcDriver $OdbcDriver
        label = "Odbc(TDS=$tds)"
      })
    }
  }

  return $attempts
}

function Invoke-Scalar {
  param(
    [System.Data.IDbConnection]$Connection,
    [string]$Sql,
    [int]$TimeoutSeconds
  )

  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  $value = $cmd.ExecuteScalar()
  return (Convert-ToNumericOrNull -Value $value)
}

function Invoke-DataTable {
  param(
    [System.Data.IDbConnection]$Connection,
    [string]$Sql,
    [int]$TimeoutSeconds
  )

  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  $table = New-Object System.Data.DataTable
  $reader = $cmd.ExecuteReader()
  try {
    $table.Load($reader)
  }
  finally {
    $reader.Close()
  }
  return $table
}

function Get-Status {
  param(
    [object]$Value,
    [hashtable]$Rule
  )

  $numericValue = Convert-ToNumericOrNull -Value $Value

  if ($null -eq $numericValue) {
    return "Warning"
  }

  if ($Rule.ContainsKey("criticalBelow") -and $numericValue -lt [double]$Rule["criticalBelow"]) { return "Critical" }
  if ($Rule.ContainsKey("warningBelow") -and $numericValue -lt [double]$Rule["warningBelow"]) { return "Warning" }
  if ($Rule.ContainsKey("criticalAbove") -and $numericValue -gt [double]$Rule["criticalAbove"]) { return "Critical" }
  if ($Rule.ContainsKey("warningAbove") -and $numericValue -gt [double]$Rule["warningAbove"]) { return "Warning" }

  return "Good"
}

function Escape-Html {
  param([string]$Text)

  if ($null -eq $Text) { return "" }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-StatusBadgeClass {
  param([string]$Status)

  switch ($Status) {
    "Critical" { return "status-critical" }
    "Warning" { return "status-warning" }
    default { return "status-good" }
  }
}

function New-Metric {
  param(
    [string]$MetricKey,
    [string]$Category,
    [object]$Value,
    [string]$Unit,
    [hashtable]$Rule,
    [string]$SourceProbe
  )

  $status = Get-Status -Value $Value -Rule $Rule
  return [ordered]@{
    metricKey = $MetricKey
    category = $Category
    value = $Value
    unit = $Unit
    threshold = $Rule
    status = $status
    collectedAt = (Get-Date).ToString("s")
    sourceProbe = $SourceProbe
  }
}

function Invoke-SafeScalarProbe {
  param(
    [System.Data.IDbConnection]$Connection,
    [string]$ProbeName,
    [string]$Sql,
    [int]$TimeoutSeconds,
    [System.Collections.Generic.List[object]]$Errors
  )

  try {
    return Invoke-Scalar -Connection $Connection -Sql $Sql -TimeoutSeconds $TimeoutSeconds
  }
  catch {
    $Errors.Add([ordered]@{
      probe = $ProbeName
      message = $_.Exception.Message
      at = (Get-Date).ToString("s")
    })
    return $null
  }
}

function Invoke-SafeTableProbe {
  param(
    [System.Data.IDbConnection]$Connection,
    [string]$ProbeName,
    [string]$Sql,
    [int]$TimeoutSeconds,
    [System.Collections.Generic.List[object]]$Errors
  )

  try {
    return Invoke-DataTable -Connection $Connection -Sql $Sql -TimeoutSeconds $TimeoutSeconds
  }
  catch {
    $Errors.Add([ordered]@{
      probe = $ProbeName
      message = $_.Exception.Message
      at = (Get-Date).ToString("s")
    })
    $empty = New-Object System.Data.DataTable
    return $empty
  }
}

if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$thresholdFullPath = Join-Path $PSScriptRoot $ThresholdPath
$thresholdJson = Get-Content -Raw -Path $thresholdFullPath | ConvertFrom-Json
$ruleMap = @{}
$thresholdJson.rules.PSObject.Properties | ForEach-Object {
  $name = $_.Name
  $ruleMap[$name] = @{}
  $_.Value.PSObject.Properties | ForEach-Object { $ruleMap[$name][$_.Name] = $_.Value }
}

$attemptTrace = New-Object System.Collections.Generic.List[object]
$attempts = Get-ConnectionAttempts -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -ConnectionProvider $ConnectionProvider -TdsVersion $TdsVersion -OdbcDriver $OdbcDriver -EnableFallback $EnableFallback -FallbackTdsVersions $FallbackTdsVersions

$conn = $null
$activeProvider = $null
$activeTdsVersion = $null

foreach ($attempt in $attempts) {
  try {
    $conn = New-DbConnection -ConnectionString $attempt.connectionString -ConnectionProvider $attempt.provider
    $conn.Open()
    $activeProvider = $attempt.provider
    $activeTdsVersion = $attempt.tdsVersion
    $attemptTrace.Add([ordered]@{ label = $attempt.label; success = $true; message = "connected" })
    break
  }
  catch {
    $attemptTrace.Add([ordered]@{ label = $attempt.label; success = $false; message = $_.Exception.Message })
    if ($null -ne $conn) {
      try { $conn.Dispose() } catch {}
      $conn = $null
    }
  }
}

if ($null -eq $conn) {
  $traceText = ($attemptTrace | ForEach-Object { "$($_.label): $($_.message)" }) -join " | "
  throw "Unable to connect with all attempts. $traceText"
}

try {
  $metrics = New-Object System.Collections.Generic.List[object]
  $probeErrors = New-Object System.Collections.Generic.List[object]

  $ple = Invoke-SafeScalarProbe -Connection $conn -ProbeName "ple_seconds" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT MAX(cntr_value)
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name = 'Page life expectancy'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "ple_seconds" -Category "resource_performance" -Value $ple -Unit "seconds" -Rule $ruleMap["ple_seconds"] -SourceProbe "dm_os_performance_counters"))

  $bchr = Invoke-SafeScalarProbe -Connection $conn -ProbeName "buffer_cache_hit_ratio" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT MAX(cntr_value)
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name = 'Buffer cache hit ratio'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "buffer_cache_hit_ratio" -Category "resource_performance" -Value $bchr -Unit "percent" -Rule $ruleMap["buffer_cache_hit_ratio"] -SourceProbe "dm_os_performance_counters"))

  $blocking = Invoke-SafeScalarProbe -Connection $conn -ProbeName "blocking_sessions" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM sys.dm_exec_requests WITH (NOLOCK)
WHERE blocking_session_id <> 0
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "blocking_sessions" -Category "concurrency" -Value $blocking -Unit "count" -Rule $ruleMap["blocking_sessions"] -SourceProbe "dm_exec_requests"))

  $longestRunningSeconds = Invoke-SafeScalarProbe -Connection $conn -ProbeName "longest_running_query_seconds" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT ISNULL(MAX(total_elapsed_time / 1000.0), 0)
FROM sys.dm_exec_requests WITH (NOLOCK)
WHERE session_id <> @@SPID
  AND status <> 'background'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "longest_running_query_seconds" -Category "concurrency" -Value $longestRunningSeconds -Unit "seconds" -Rule $ruleMap["longest_running_query_seconds"] -SourceProbe "dm_exec_requests"))

  $deadlocks24h = Invoke-SafeScalarProbe -Connection $conn -ProbeName "deadlocks_24h" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
;WITH rb AS (
  SELECT CAST(st.target_data AS XML) AS x
  FROM sys.dm_xe_session_targets st
  INNER JOIN sys.dm_xe_sessions s ON s.address = st.event_session_address
  WHERE s.name = 'system_health'
    AND st.target_name = 'ring_buffer'
)
SELECT COUNT(1)
FROM rb
CROSS APPLY x.nodes('//event[@name="xml_deadlock_report"]') AS d(e)
WHERE d.e.value('@timestamp', 'datetime2') >= DATEADD(HOUR, -24, SYSUTCDATETIME())
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "deadlocks_24h" -Category "concurrency" -Value $deadlocks24h -Unit "count" -Rule $ruleMap["deadlocks_24h"] -SourceProbe "system_health_xevent"))

  $failedJobs = Invoke-SafeScalarProbe -Connection $conn -ProbeName "failed_jobs_24h" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
WHERE h.step_id = 0
  AND h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "failed_jobs_24h" -Category "automation_jobs" -Value $failedJobs -Unit "count" -Rule $ruleMap["failed_jobs_24h"] -SourceProbe "msdb.sysjobhistory"))

  $disabledJobs = Invoke-SafeScalarProbe -Connection $conn -ProbeName "disabled_jobs_count" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM msdb.dbo.sysjobs WITH (NOLOCK)
WHERE enabled = 0
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "disabled_jobs_count" -Category "automation_jobs" -Value $disabledJobs -Unit "count" -Rule $ruleMap["disabled_jobs_count"] -SourceProbe "msdb.sysjobs"))

  $lastFullBackup = Invoke-SafeTableProbe -Connection $conn -ProbeName "last_full_backup_time" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT MAX(backup_finish_date) AS last_full_backup_time
FROM msdb.dbo.backupset WITH (NOLOCK)
WHERE type = 'D'
OPTION (RECOMPILE);
"@

  $backupAgeHours = $null
  if ($lastFullBackup.Rows.Count -gt 0 -and $lastFullBackup.Rows[0]["last_full_backup_time"] -ne [System.DBNull]::Value) {
    $lastTime = [datetime]$lastFullBackup.Rows[0]["last_full_backup_time"]
    $backupAgeHours = [Math]::Round(((Get-Date) - $lastTime).TotalHours, 2)
  }
  $metrics.Add((New-Metric -MetricKey "last_full_backup_age_hours" -Category "data_safety" -Value $backupAgeHours -Unit "hours" -Rule $ruleMap["last_full_backup_age_hours"] -SourceProbe "msdb.backupset"))

  $dbWithoutRecentBackup = Invoke-SafeScalarProbe -Connection $conn -ProbeName "databases_without_recent_full_backup" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
DECLARE @cutoff DATETIME = DATEADD(HOUR, -72, GETDATE());
SELECT COUNT(1)
FROM sys.databases d WITH (NOLOCK)
LEFT JOIN (
  SELECT database_name, MAX(backup_finish_date) AS last_full_backup
  FROM msdb.dbo.backupset WITH (NOLOCK)
  WHERE type = 'D'
  GROUP BY database_name
) b ON b.database_name = d.name
WHERE d.database_id > 4
  AND d.state_desc = 'ONLINE'
  AND (b.last_full_backup IS NULL OR b.last_full_backup < @cutoff)
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "databases_without_recent_full_backup" -Category "data_safety" -Value $dbWithoutRecentBackup -Unit "count" -Rule $ruleMap["databases_without_recent_full_backup"] -SourceProbe "sys.databases+msdb.backupset"))

  $logUsage = Invoke-SafeTableProbe -Connection $conn -ProbeName "avg_log_used_percent" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT AVG(used_log_space_in_percent * 1.0) AS avg_log_used_percent
FROM sys.dm_db_log_space_usage
OPTION (RECOMPILE);
"@
  $avgLog = $null
  if ($logUsage.Rows.Count -gt 0 -and $logUsage.Rows[0]["avg_log_used_percent"] -ne [System.DBNull]::Value) {
    $avgLog = [double]$logUsage.Rows[0]["avg_log_used_percent"]
  }
  $metrics.Add((New-Metric -MetricKey "log_used_percent" -Category "io_disk" -Value $avgLog -Unit "percent" -Rule $ruleMap["log_used_percent"] -SourceProbe "dm_db_log_space_usage"))

  $dataDiskUsedPercent = Invoke-SafeScalarProbe -Connection $conn -ProbeName "data_disk_used_percent" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT MAX(
  CASE WHEN vs.total_bytes > 0
    THEN (1.0 - (vs.available_bytes * 1.0 / vs.total_bytes)) * 100
    ELSE NULL END
)
FROM sys.master_files mf WITH (NOLOCK)
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs
WHERE mf.type_desc = 'ROWS'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "data_disk_used_percent" -Category "io_disk" -Value $dataDiskUsedPercent -Unit "percent" -Rule $ruleMap["data_disk_used_percent"] -SourceProbe "dm_os_volume_stats"))

  $agUnhealthyReplicas = Invoke-SafeScalarProbe -Connection $conn -ProbeName "ag_unhealthy_replicas" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM sys.dm_hadr_availability_replica_states WITH (NOLOCK)
WHERE is_local = 1
  AND role_desc = 'PRIMARY'
  AND synchronization_health_desc <> 'HEALTHY'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "ag_unhealthy_replicas" -Category "high_availability" -Value $agUnhealthyReplicas -Unit "count" -Rule $ruleMap["ag_unhealthy_replicas"] -SourceProbe "dm_hadr_availability_replica_states"))

  $sysadminLoginCount = Invoke-SafeScalarProbe -Connection $conn -ProbeName "sysadmin_login_count" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM sys.server_role_members srm WITH (NOLOCK)
INNER JOIN sys.server_principals r WITH (NOLOCK) ON r.principal_id = srm.role_principal_id
WHERE r.name = 'sysadmin'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "sysadmin_login_count" -Category "security_audit" -Value $sysadminLoginCount -Unit "count" -Rule $ruleMap["sysadmin_login_count"] -SourceProbe "sys.server_role_members"))

  $weakPolicyLogins = Invoke-SafeScalarProbe -Connection $conn -ProbeName "weak_policy_logins_count" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM sys.sql_logins WITH (NOLOCK)
WHERE is_disabled = 0
  AND (is_policy_checked = 0 OR is_expiration_checked = 0)
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "weak_policy_logins_count" -Category "security_audit" -Value $weakPolicyLogins -Unit "count" -Rule $ruleMap["weak_policy_logins_count"] -SourceProbe "sys.sql_logins"))

  $newLogins90d = Invoke-SafeScalarProbe -Connection $conn -ProbeName "new_logins_90d_count" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT COUNT(1)
FROM sys.server_principals WITH (NOLOCK)
WHERE type IN ('S', 'U', 'G')
  AND create_date >= DATEADD(DAY, -90, GETDATE())
  AND name NOT LIKE '##%'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "new_logins_90d_count" -Category "security_audit" -Value $newLogins90d -Unit "count" -Rule $ruleMap["new_logins_90d_count"] -SourceProbe "sys.server_principals"))

  $newDbUsers90d = Invoke-SafeScalarProbe -Connection $conn -ProbeName "new_db_users_90d_count" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
CREATE TABLE #recent_users (
  database_name SYSNAME,
  user_name SYSNAME,
  type_desc NVARCHAR(60),
  create_date DATETIME
);

DECLARE @db SYSNAME;
DECLARE @sql NVARCHAR(MAX);
DECLARE db_cur CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases WITH (NOLOCK)
WHERE state_desc = 'ONLINE'
  AND database_id > 4;

OPEN db_cur;
FETCH NEXT FROM db_cur INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
  BEGIN TRY
    INSERT INTO #recent_users(database_name, user_name, type_desc, create_date)
    SELECT N''' + REPLACE(@db, '''', '''''') + ''', name, type_desc, create_date
    FROM ' + QUOTENAME(@db) + '.sys.database_principals WITH (NOLOCK)
    WHERE type IN (''S'', ''U'', ''G'')
      AND principal_id > 4
      AND create_date >= DATEADD(DAY, -90, GETDATE());
  END TRY
  BEGIN CATCH
  END CATCH;';

  EXEC sys.sp_executesql @sql;
  FETCH NEXT FROM db_cur INTO @db;
END

CLOSE db_cur;
DEALLOCATE db_cur;

SELECT COUNT(1)
FROM #recent_users
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "new_db_users_90d_count" -Category "security_audit" -Value $newDbUsers90d -Unit "count" -Rule $ruleMap["new_db_users_90d_count"] -SourceProbe "sys.database_principals"))

  $failedJobsDetails = Invoke-SafeTableProbe -Connection $conn -ProbeName "failed_jobs_details_24h" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT TOP (20)
  j.name AS job_name,
  msdb.dbo.agent_datetime(h.run_date, h.run_time) AS run_time,
  h.message
FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
INNER JOIN msdb.dbo.sysjobs j WITH (NOLOCK) ON j.job_id = h.job_id
WHERE h.run_status = 0
  AND h.step_id = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
ORDER BY run_time DESC
OPTION (RECOMPILE);
"@

  $longRunningTop = Invoke-SafeTableProbe -Connection $conn -ProbeName "long_running_queries_top5" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT TOP (5)
  r.session_id,
  DB_NAME(r.database_id) AS database_name,
  r.status,
  r.wait_type,
  CAST(r.total_elapsed_time / 1000.0 AS DECIMAL(18,2)) AS elapsed_seconds,
  LEFT(REPLACE(REPLACE(t.text, CHAR(13), ' '), CHAR(10), ' '), 300) AS query_text
FROM sys.dm_exec_requests r WITH (NOLOCK)
CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.session_id <> @@SPID
ORDER BY r.total_elapsed_time DESC
OPTION (RECOMPILE);
"@

  $backupAgeByDb = Invoke-SafeTableProbe -Connection $conn -ProbeName "backup_age_by_db" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT
  d.name AS database_name,
  b.last_full_backup,
  CAST(CASE WHEN b.last_full_backup IS NULL THEN NULL ELSE DATEDIFF(MINUTE, b.last_full_backup, GETDATE()) / 60.0 END AS DECIMAL(18,2)) AS backup_age_hours
FROM sys.databases d WITH (NOLOCK)
LEFT JOIN (
  SELECT database_name, MAX(backup_finish_date) AS last_full_backup
  FROM msdb.dbo.backupset WITH (NOLOCK)
  WHERE type = 'D'
  GROUP BY database_name
) b ON b.database_name = d.name
WHERE d.database_id > 4
ORDER BY backup_age_hours DESC
OPTION (RECOMPILE);
"@

  $sysadminLoginDetails = Invoke-SafeTableProbe -Connection $conn -ProbeName "sysadmin_login_details" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT
  p.name,
  p.type_desc,
  p.is_disabled,
  p.create_date
FROM sys.server_role_members srm WITH (NOLOCK)
INNER JOIN sys.server_principals r WITH (NOLOCK) ON r.principal_id = srm.role_principal_id
INNER JOIN sys.server_principals p WITH (NOLOCK) ON p.principal_id = srm.member_principal_id
WHERE r.name = 'sysadmin'
ORDER BY p.name
OPTION (RECOMPILE);
"@

  $weakPolicyLoginDetails = Invoke-SafeTableProbe -Connection $conn -ProbeName "weak_policy_login_details" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT TOP (50)
  name,
  is_policy_checked,
  is_expiration_checked,
  create_date,
  modify_date
FROM sys.sql_logins WITH (NOLOCK)
WHERE is_disabled = 0
  AND (is_policy_checked = 0 OR is_expiration_checked = 0)
ORDER BY create_date DESC
OPTION (RECOMPILE);
"@

  $recentLoginsDetails = Invoke-SafeTableProbe -Connection $conn -ProbeName "recent_logins_90d_details" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
SELECT TOP (50)
  name,
  type_desc,
  create_date,
  default_database_name
FROM sys.server_principals WITH (NOLOCK)
WHERE type IN ('S', 'U', 'G')
  AND create_date >= DATEADD(DAY, -90, GETDATE())
  AND name NOT LIKE '##%'
ORDER BY create_date DESC
OPTION (RECOMPILE);
"@

  $recentDbUsersDetails = Invoke-SafeTableProbe -Connection $conn -ProbeName "recent_db_users_90d_details" -TimeoutSeconds $TimeoutSeconds -Errors $probeErrors -Sql @"
CREATE TABLE #recent_users (
  database_name SYSNAME,
  user_name SYSNAME,
  type_desc NVARCHAR(60),
  create_date DATETIME
);

DECLARE @db SYSNAME;
DECLARE @sql NVARCHAR(MAX);
DECLARE db_cur CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases WITH (NOLOCK)
WHERE state_desc = 'ONLINE'
  AND database_id > 4;

OPEN db_cur;
FETCH NEXT FROM db_cur INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
  BEGIN TRY
    INSERT INTO #recent_users(database_name, user_name, type_desc, create_date)
    SELECT N''' + REPLACE(@db, '''', '''''') + ''', name, type_desc, create_date
    FROM ' + QUOTENAME(@db) + '.sys.database_principals WITH (NOLOCK)
    WHERE type IN (''S'', ''U'', ''G'')
      AND principal_id > 4
      AND create_date >= DATEADD(DAY, -90, GETDATE());
  END TRY
  BEGIN CATCH
  END CATCH;';

  EXEC sys.sp_executesql @sql;
  FETCH NEXT FROM db_cur INTO @db;
END

CLOSE db_cur;
DEALLOCATE db_cur;

SELECT TOP (50)
  database_name,
  user_name,
  type_desc,
  create_date
FROM #recent_users
ORDER BY create_date DESC
OPTION (RECOMPILE);
"@

  $criticalCount = ($metrics | Where-Object { $_.status -eq "Critical" }).Count
  $warningCount = ($metrics | Where-Object { $_.status -eq "Warning" }).Count
  $goodCount = ($metrics | Where-Object { $_.status -eq "Good" }).Count
  $healthScore = [Math]::Max(0, 100 - ($criticalCount * 20 + $warningCount * 8))

  $overallStatus = "Good"
  if ($criticalCount -gt 0) {
    $overallStatus = "Critical"
  }
  elseif ($warningCount -gt 0) {
    $overallStatus = "Warning"
  }

  $failedJobsDetailRows = @()
  foreach ($row in $failedJobsDetails.Rows) {
    $failedJobsDetailRows += [ordered]@{
      job_name = [string]$row["job_name"]
      run_time = [string]$row["run_time"]
      message = [string]$row["message"]
    }
  }

  $longRunningDetailRows = @()
  foreach ($row in $longRunningTop.Rows) {
    $longRunningDetailRows += [ordered]@{
      session_id = [int]$row["session_id"]
      database_name = [string]$row["database_name"]
      status = [string]$row["status"]
      wait_type = [string]$row["wait_type"]
      elapsed_seconds = Convert-ToNumericOrNull -Value $row["elapsed_seconds"]
      query_text = [string]$row["query_text"]
    }
  }

  $backupAgeDetailRows = @()
  foreach ($row in $backupAgeByDb.Rows) {
    $backupAgeDetailRows += [ordered]@{
      database_name = [string]$row["database_name"]
      last_full_backup = if ($row["last_full_backup"] -eq [System.DBNull]::Value) { $null } else { [string]$row["last_full_backup"] }
      backup_age_hours = Convert-ToNumericOrNull -Value $row["backup_age_hours"]
    }
  }

  $sysadminDetailRows = @()
  foreach ($row in $sysadminLoginDetails.Rows) {
    $sysadminDetailRows += [ordered]@{
      name = [string]$row["name"]
      type_desc = [string]$row["type_desc"]
      is_disabled = [string]$row["is_disabled"]
      create_date = [string]$row["create_date"]
    }
  }

  $weakPolicyDetailRows = @()
  foreach ($row in $weakPolicyLoginDetails.Rows) {
    $weakPolicyDetailRows += [ordered]@{
      name = [string]$row["name"]
      is_policy_checked = [string]$row["is_policy_checked"]
      is_expiration_checked = [string]$row["is_expiration_checked"]
      create_date = [string]$row["create_date"]
      modify_date = [string]$row["modify_date"]
    }
  }

  $recentLoginsDetailRows = @()
  foreach ($row in $recentLoginsDetails.Rows) {
    $recentLoginsDetailRows += [ordered]@{
      name = [string]$row["name"]
      type_desc = [string]$row["type_desc"]
      create_date = [string]$row["create_date"]
      default_database_name = [string]$row["default_database_name"]
    }
  }

  $recentDbUsersDetailRows = @()
  foreach ($row in $recentDbUsersDetails.Rows) {
    $recentDbUsersDetailRows += [ordered]@{
      database_name = [string]$row["database_name"]
      user_name = [string]$row["user_name"]
      type_desc = [string]$row["type_desc"]
      create_date = [string]$row["create_date"]
    }
  }

  $result = [ordered]@{
    project = "SQL-Server-Health-Sentinel"
    server = $Server
    database = $Database
    collectedAt = (Get-Date).ToString("s")
    connection = [ordered]@{
      requestedProvider = $ConnectionProvider
      requestedTdsVersion = if ([string]::IsNullOrWhiteSpace($TdsVersion)) { $null } else { $TdsVersion }
      activeProvider = $activeProvider
      activeTdsVersion = $activeTdsVersion
      fallbackEnabled = $EnableFallback
      attempts = $attemptTrace
      port = $Port
    }
    summary = [ordered]@{
      overallStatus = $overallStatus
      healthScore = $healthScore
      critical = $criticalCount
      warning = $warningCount
      good = $goodCount
      probeErrors = $probeErrors.Count
    }
    metrics = $metrics
    details = [ordered]@{
      longRunningQueriesTop5 = $longRunningDetailRows
      failedJobs24h = $failedJobsDetailRows
      backupAgeByDatabase = $backupAgeDetailRows
      sysadminLogins = $sysadminDetailRows
      weakPolicyLogins = $weakPolicyDetailRows
      recentLogins90d = $recentLoginsDetailRows
      recentDbUsers90d = $recentDbUsersDetailRows
      probeErrors = $probeErrors
    }
  }

  $basePath = Join-Path $OutputDir ("inspection-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
  $jsonPath = "$basePath.json"
  $htmlPath = "$basePath.html"
  $mdPath = "$basePath.md"
  $docxPath = "$basePath.docx"
  $pdfPath = "$basePath.pdf"

  $selectedFormats = Normalize-ReportFormats -ReportFormats $ReportFormats
  $formatSet = @{}
  foreach ($fmt in $selectedFormats) { $formatSet[$fmt] = $true }

  $generatedFiles = New-Object System.Collections.Generic.List[string]
  $outputWarnings = New-Object System.Collections.Generic.List[string]

  $rows = $metrics | ForEach-Object {
    $statusClass = Get-StatusBadgeClass -Status $_.status
    "<tr><td>$(Escape-Html -Text $_.metricKey)</td><td>$(Escape-Html -Text $_.category)</td><td>$(Escape-Html -Text ([string]$_.value))</td><td>$(Escape-Html -Text $_.unit)</td><td><span class='badge $statusClass'>$(Escape-Html -Text $_.status)</span></td></tr>"
  }

  $topQueryRows = $longRunningDetailRows | ForEach-Object {
    "<tr><td>$($_.session_id)</td><td>$(Escape-Html -Text $_.database_name)</td><td>$($_.elapsed_seconds)</td><td>$(Escape-Html -Text $_.wait_type)</td><td>$(Escape-Html -Text $_.query_text)</td></tr>"
  }

  if ($topQueryRows.Count -eq 0) {
    $topQueryRows = @("<tr><td colspan='5'>No active long-running queries.</td></tr>")
  }

  $jobRows = $failedJobsDetailRows | ForEach-Object {
    "<tr><td>$(Escape-Html -Text $_.job_name)</td><td>$(Escape-Html -Text $_.run_time)</td><td>$(Escape-Html -Text $_.message)</td></tr>"
  }

  if ($jobRows.Count -eq 0) {
    $jobRows = @("<tr><td colspan='3'>No failed jobs in last 24h.</td></tr>")
  }
  
  $sysadminRows = $sysadminDetailRows | ForEach-Object {
    "<tr><td>$(Escape-Html -Text $_.name)</td><td>$(Escape-Html -Text $_.type_desc)</td><td>$(Escape-Html -Text $_.is_disabled)</td><td>$(Escape-Html -Text $_.create_date)</td></tr>"
  }
  if ($sysadminRows.Count -eq 0) {
    $sysadminRows = @("<tr><td colspan='4'>No sysadmin principals found.</td></tr>")
  }
  
  $weakPolicyRows = $weakPolicyDetailRows | ForEach-Object {
    "<tr><td>$(Escape-Html -Text $_.name)</td><td>$(Escape-Html -Text $_.is_policy_checked)</td><td>$(Escape-Html -Text $_.is_expiration_checked)</td><td>$(Escape-Html -Text $_.create_date)</td><td>$(Escape-Html -Text $_.modify_date)</td></tr>"
  }
  if ($weakPolicyRows.Count -eq 0) {
    $weakPolicyRows = @("<tr><td colspan='5'>No weak-policy SQL logins detected.</td></tr>")
  }
  
  $recentIdentityRows = $recentLoginsDetailRows | ForEach-Object {
    "<tr><td>Server Login</td><td>$(Escape-Html -Text $_.name)</td><td>$(Escape-Html -Text $_.type_desc)</td><td>$(Escape-Html -Text $_.create_date)</td><td>$(Escape-Html -Text $_.default_database_name)</td></tr>"
  }
  $recentDbIdentityRows = $recentDbUsersDetailRows | ForEach-Object {
    "<tr><td>DB User</td><td>$(Escape-Html -Text $_.user_name)</td><td>$(Escape-Html -Text $_.type_desc)</td><td>$(Escape-Html -Text $_.create_date)</td><td>$(Escape-Html -Text $_.database_name)</td></tr>"
  }
  $recentIdentityRows = @($recentIdentityRows) + @($recentDbIdentityRows)
  if ($recentIdentityRows.Count -eq 0) {
    $recentIdentityRows = @("<tr><td colspan='5'>No new logins/users in last 90 days.</td></tr>")
  }

  $overallBadgeClass = Get-StatusBadgeClass -Status $overallStatus

  $html = @"
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>SQL Server Inspection Report</title>
  <style>
    :root {
      --bg: #f4f7fb;
      --card-bg: #ffffff;
      --ink: #1f2937;
      --sub: #5b6472;
      --line: #d9e2ec;
      --good: #0f9d58;
      --warn: #d38b00;
      --crit: #c62828;
      --brand: #0b5cab;
    }
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: var(--bg); color: var(--ink); }
    h1 { margin: 0 0 6px 0; }
    h2 { margin: 20px 0 10px 0; }
    .meta { color: var(--sub); margin-bottom: 16px; }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 16px; }
    .card { background: var(--card-bg); border: 1px solid var(--line); border-radius: 10px; padding: 12px; }
    .card .k { color: var(--sub); font-size: 12px; }
    .card .v { font-size: 26px; font-weight: 700; margin-top: 5px; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; color: #fff; font-size: 12px; }
    .status-good { background: var(--good); }
    .status-warning { background: var(--warn); }
    .status-critical { background: var(--crit); }
    .panel { background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 12px; margin-bottom: 16px; }
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid var(--line); padding: 8px 10px; text-align: left; vertical-align: top; }
    th { background: #edf3fa; }
    .mono { font-family: Consolas, 'Courier New', monospace; }
  </style>
</head>
<body>
  <h1>SQL-Server-Health-Sentinel Report</h1>
  <div class=\"meta\">Server: $(Escape-Html -Text $Server) | Database: $(Escape-Html -Text $Database) | Collected: $(Get-Date -Format "s") | Overall: <span class='badge $overallBadgeClass'>$(Escape-Html -Text $overallStatus)</span></div>

  <div class=\"cards\">
    <div class=\"card\"><div class=\"k\">Health Score</div><div class=\"v\">$healthScore</div></div>
    <div class=\"card\"><div class=\"k\">Critical</div><div class=\"v\">$criticalCount</div></div>
    <div class=\"card\"><div class=\"k\">Warning</div><div class=\"v\">$warningCount</div></div>
    <div class=\"card\"><div class=\"k\">Good</div><div class=\"v\">$goodCount</div></div>
    <div class=\"card\"><div class=\"k\">Probe Errors</div><div class=\"v\">$($probeErrors.Count)</div></div>
  </div>

  <div class=\"panel\">
    <h2>Metric Summary</h2>
    <table>
      <thead>
        <tr><th>Metric</th><th>Category</th><th>Value</th><th>Unit</th><th>Status</th></tr>
      </thead>
      <tbody>
        $($rows -join "`n")
      </tbody>
    </table>
  </div>
  
  <div class="panel">
    <h2>Security Audit: Sysadmin Principals</h2>
    <table>
      <thead>
        <tr><th>Principal</th><th>Type</th><th>Disabled</th><th>Create Time</th></tr>
      </thead>
      <tbody>
        $($sysadminRows -join "`n")
      </tbody>
    </table>
  </div>
  
  <div class="panel">
    <h2>Security Audit: Weak Policy SQL Logins</h2>
    <table>
      <thead>
        <tr><th>Login</th><th>Password Policy</th><th>Expiration Policy</th><th>Create Time</th><th>Modify Time</th></tr>
      </thead>
      <tbody>
        $($weakPolicyRows -join "`n")
      </tbody>
    </table>
  </div>
  
  <div class="panel">
    <h2>Security Audit: New Identities in Last 90 Days</h2>
    <table>
      <thead>
        <tr><th>Scope</th><th>Name</th><th>Type</th><th>Create Time</th><th>Database</th></tr>
      </thead>
      <tbody>
        $($recentIdentityRows -join "`n")
      </tbody>
    </table>
  </div>

  <div class=\"panel\">
    <h2>Top Long-Running Queries</h2>
    <table>
      <thead>
        <tr><th>Session</th><th>Database</th><th>Elapsed(s)</th><th>Wait Type</th><th>Query Text</th></tr>
      </thead>
      <tbody>
        $($topQueryRows -join "`n")
      </tbody>
    </table>
  </div>

  <div class=\"panel\">
    <h2>Failed Jobs (24h)</h2>
    <table>
      <thead>
        <tr><th>Job</th><th>Run Time</th><th>Message</th></tr>
      </thead>
      <tbody>
        $($jobRows -join "`n")
      </tbody>
    </table>
  </div>

  <div class=\"meta mono\">Generated by SQL-Server-Health-Sentinel MVP.</div>
</body>
</html>
"@

  $metricsMdRows = $metrics | ForEach-Object {
    "| $($_.metricKey) | $($_.category) | $($_.value) | $($_.unit) | $($_.status) |"
  }
  $criticalMetricList = @($metrics | Where-Object { $_.status -eq "Critical" } | ForEach-Object { $_.metricKey })
  if ($criticalMetricList.Count -eq 0) {
    $criticalMetricList = @("None")
  }

  $md = @"
# SQL-Server-Health-Sentinel Report

- Server: $Server
- Database: $Database
- CollectedAt: $(Get-Date -Format "s")
- OverallStatus: $overallStatus
- HealthScore: $healthScore
- ActiveProvider: $activeProvider
- ActiveTdsVersion: $(if ([string]::IsNullOrWhiteSpace([string]$activeTdsVersion)) { "auto" } else { $activeTdsVersion })

## Summary

| Item | Value |
| --- | --- |
| Critical | $criticalCount |
| Warning | $warningCount |
| Good | $goodCount |
| ProbeErrors | $($probeErrors.Count) |

## Critical Metrics

$($criticalMetricList -join ", ")

## Metrics

| Metric | Category | Value | Unit | Status |
| --- | --- | --- | --- | --- |
$($metricsMdRows -join "`n")
"@

if ($formatSet.ContainsKey("json")) {
  $result | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $jsonPath
  $generatedFiles.Add($jsonPath)
}

$htmlNeeded = $formatSet.ContainsKey("html") -or $formatSet.ContainsKey("docx") -or $formatSet.ContainsKey("pdf")
if ($htmlNeeded) {
  Set-Content -Encoding UTF8 -Path $htmlPath -Value $html
  if ($formatSet.ContainsKey("html")) {
    $generatedFiles.Add($htmlPath)
  }
}

if ($formatSet.ContainsKey("md")) {
  Set-Content -Encoding UTF8 -Path $mdPath -Value $md
  $generatedFiles.Add($mdPath)
}

if ($formatSet.ContainsKey("docx") -or $formatSet.ContainsKey("pdf")) {
  $officeResult = Convert-HtmlToOffice -HtmlPath $htmlPath -DocxPath $docxPath -PdfPath $pdfPath -ExportDocx $formatSet.ContainsKey("docx") -ExportPdf $formatSet.ContainsKey("pdf")
  if ($officeResult.docxOk) { $generatedFiles.Add($docxPath) }
  if ($officeResult.pdfOk) { $generatedFiles.Add($pdfPath) }
  if ($officeResult.warning) { $outputWarnings.Add($officeResult.warning) }
}

Write-Output "Inspection complete."
Write-Output "Requested formats: $($selectedFormats -join ',')"
foreach ($f in $generatedFiles) {
  Write-Output "Generated: $f"
}
foreach ($w in $outputWarnings) {
  Write-Output "Warning: $w"
}
}
finally {
  if ($null -ne $conn) {
    try {
      if ($conn.State -ne [System.Data.ConnectionState]::Closed) {
        $conn.Close()
      }
    }
    catch {
    }
    try { $conn.Dispose() } catch {}
  }
}
