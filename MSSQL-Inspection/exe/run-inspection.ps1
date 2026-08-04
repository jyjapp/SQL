param(
  [string]$Server = "127.0.0.1",
  [string]$Database = "master",
  [string]$Username = "",
  [string]$Password = "",
  [int]$Port = 1433,
  [int]$TimeoutSeconds = 15,
  [string]$ThresholdPath = "..\\skill\\rules\\thresholds.json",
  [string]$OutputDir = ".\\output"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
    [int]$Port
  )

  if ([string]::IsNullOrWhiteSpace($Username)) {
    return "Server=$Server,$Port;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
  }

  return "Server=$Server,$Port;Database=$Database;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
}

function Invoke-Scalar {
  param(
    [System.Data.SqlClient.SqlConnection]$Connection,
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
    [System.Data.SqlClient.SqlConnection]$Connection,
    [string]$Sql,
    [int]$TimeoutSeconds
  )

  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
  $table = New-Object System.Data.DataTable
  [void]$adapter.Fill($table)
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
    [System.Data.SqlClient.SqlConnection]$Connection,
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
    [System.Data.SqlClient.SqlConnection]$Connection,
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

$connString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port
$conn = New-Object System.Data.SqlClient.SqlConnection($connString)
$conn.Open()

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

  $result = [ordered]@{
    project = "SQL-Server-Health-Sentinel"
    server = $Server
    database = $Database
    collectedAt = (Get-Date).ToString("s")
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
      probeErrors = $probeErrors
    }
  }

  $jsonPath = Join-Path $OutputDir ("inspection-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".json")
  $htmlPath = [System.IO.Path]::ChangeExtension($jsonPath, ".html")

  $result | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $jsonPath

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

  Set-Content -Encoding UTF8 -Path $htmlPath -Value $html

  Write-Output "Inspection complete."
  Write-Output "JSON: $jsonPath"
  Write-Output "HTML: $htmlPath"
}
finally {
  $conn.Close()
}
