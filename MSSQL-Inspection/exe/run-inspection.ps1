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
  if ($null -eq $value -or $value -eq [System.DBNull]::Value) {
    return $null
  }
  return [double]$value
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
    [double]$Value,
    [hashtable]$Rule
  )

  if ($null -eq $Value) {
    return "Warning"
  }

  if ($Rule.ContainsKey("criticalBelow") -and $Value -lt [double]$Rule["criticalBelow"]) { return "Critical" }
  if ($Rule.ContainsKey("warningBelow") -and $Value -lt [double]$Rule["warningBelow"]) { return "Warning" }
  if ($Rule.ContainsKey("criticalAbove") -and $Value -gt [double]$Rule["criticalAbove"]) { return "Critical" }
  if ($Rule.ContainsKey("warningAbove") -and $Value -gt [double]$Rule["warningAbove"]) { return "Warning" }

  return "Good"
}

function New-Metric {
  param(
    [string]$MetricKey,
    [string]$Category,
    [double]$Value,
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

  $ple = Invoke-Scalar -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT MAX(cntr_value)
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name = 'Page life expectancy'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "ple_seconds" -Category "resource_performance" -Value $ple -Unit "seconds" -Rule $ruleMap["ple_seconds"] -SourceProbe "dm_os_performance_counters"))

  $bchr = Invoke-Scalar -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT MAX(cntr_value)
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name = 'Buffer cache hit ratio'
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "buffer_cache_hit_ratio" -Category "resource_performance" -Value $bchr -Unit "percent" -Rule $ruleMap["buffer_cache_hit_ratio"] -SourceProbe "dm_os_performance_counters"))

  $blocking = Invoke-Scalar -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM sys.dm_exec_requests WITH (NOLOCK)
WHERE blocking_session_id <> 0
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "blocking_sessions" -Category "concurrency" -Value $blocking -Unit "count" -Rule $ruleMap["blocking_sessions"] -SourceProbe "dm_exec_requests"))

  $failedJobs = Invoke-Scalar -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
WHERE h.step_id = 0
  AND h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
OPTION (RECOMPILE);
"@
  $metrics.Add((New-Metric -MetricKey "failed_jobs_24h" -Category "automation_jobs" -Value $failedJobs -Unit "count" -Rule $ruleMap["failed_jobs_24h"] -SourceProbe "msdb.sysjobhistory"))

  $lastFullBackup = Invoke-DataTable -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
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

  $logUsage = Invoke-DataTable -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT AVG(used_log_space_in_percent * 1.0) AS avg_log_used_percent
FROM sys.dm_db_log_space_usage
OPTION (RECOMPILE);
"@
  $avgLog = $null
  if ($logUsage.Rows.Count -gt 0 -and $logUsage.Rows[0]["avg_log_used_percent"] -ne [System.DBNull]::Value) {
    $avgLog = [double]$logUsage.Rows[0]["avg_log_used_percent"]
  }
  $metrics.Add((New-Metric -MetricKey "log_used_percent" -Category "io_disk" -Value $avgLog -Unit "percent" -Rule $ruleMap["log_used_percent"] -SourceProbe "dm_db_log_space_usage"))

  $result = [ordered]@{
    project = "SQL-Server-Health-Sentinel"
    server = $Server
    database = $Database
    collectedAt = (Get-Date).ToString("s")
    metrics = $metrics
  }

  $jsonPath = Join-Path $OutputDir ("inspection-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".json")
  $htmlPath = [System.IO.Path]::ChangeExtension($jsonPath, ".html")

  $result | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $jsonPath

  $rows = $metrics | ForEach-Object {
    "<tr><td>$($_.metricKey)</td><td>$($_.category)</td><td>$($_.value)</td><td>$($_.unit)</td><td>$($_.status)</td></tr>"
  }

  $html = @"
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>SQL Server Inspection Report</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f7f9fc; }
    h1 { margin: 0 0 8px 0; }
    .meta { color: #555; margin-bottom: 16px; }
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid #d9e2ec; padding: 10px; text-align: left; }
    th { background: #eef4ff; }
  </style>
</head>
<body>
  <h1>SQL-Server-Health-Sentinel Report</h1>
  <div class=\"meta\">Server: $Server | Database: $Database | Collected: $(Get-Date -Format "s")</div>
  <table>
    <thead>
      <tr><th>Metric</th><th>Category</th><th>Value</th><th>Unit</th><th>Status</th></tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>
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
