param(
  [string]$InspectionJsonPath = "",
  [string]$OutputDir = ".\output"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($InspectionJsonPath)) {
  $latest = Get-ChildItem -Path $OutputDir -Filter "inspection-*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if ($null -eq $latest) {
    throw "No inspection JSON found in $OutputDir"
  }
  $InspectionJsonPath = $latest.FullName
}

if (-not (Test-Path $InspectionJsonPath)) {
  throw "Inspection JSON not found: $InspectionJsonPath"
}

$data = Get-Content -Raw -Path $InspectionJsonPath | ConvertFrom-Json
$metricsMap = @{}
foreach ($m in $data.metrics) {
  $metricsMap[[string]$m.metricKey] = $m
}

$critical = @($data.metrics | Where-Object { $_.status -eq "Critical" } | ForEach-Object { $_.metricKey })
$warning = @($data.metrics | Where-Object { $_.status -eq "Warning" } | ForEach-Object { $_.metricKey })
if ($critical.Count -eq 0) { $critical = @("none") }
if ($warning.Count -eq 0) { $warning = @("none") }

$highActions = New-Object System.Collections.Generic.List[string]
$midActions = New-Object System.Collections.Generic.List[string]
$lowActions = New-Object System.Collections.Generic.List[string]

if ($metricsMap.ContainsKey("databases_without_recent_full_backup") -and $metricsMap["databases_without_recent_full_backup"].status -ne "Good") {
  $highActions.Add("Backup coverage: immediately schedule full backups for databases without recent full backup and validate restore path with RESTORE VERIFYONLY.")
}
if ($metricsMap.ContainsKey("weak_policy_logins_count") -and $metricsMap["weak_policy_logins_count"].status -ne "Good") {
  $highActions.Add("Security policy: enable CHECK_POLICY and CHECK_EXPIRATION for SQL logins, rotate weak credentials, disable stale accounts.")
}
if ($metricsMap.ContainsKey("sysadmin_login_count") -and $metricsMap["sysadmin_login_count"].status -ne "Good") {
  $highActions.Add("Privilege minimization: review sysadmin membership and remove non-essential high-privilege principals.")
}
if ($metricsMap.ContainsKey("ag_unhealthy_replicas") -and $metricsMap["ag_unhealthy_replicas"].status -ne "Good") {
  $highActions.Add("HA reliability: investigate AG unhealthy replicas and validate failover readiness.")
}

if ($metricsMap.ContainsKey("longest_running_query_seconds") -and $metricsMap["longest_running_query_seconds"].status -ne "Good") {
  $midActions.Add("Query latency: identify top long-running requests and optimize indexing, execution plans, and transaction scope.")
}
if ($metricsMap.ContainsKey("blocking_sessions") -and $metricsMap["blocking_sessions"].status -ne "Good") {
  $midActions.Add("Blocking control: analyze blocking chain and wait types, tune lock order and reduce hot-spot contention.")
}
if ($metricsMap.ContainsKey("deadlocks_24h") -and $metricsMap["deadlocks_24h"].status -ne "Good") {
  $midActions.Add("Deadlock reduction: capture deadlock graphs and apply index/order fixes for conflicting access patterns.")
}
if ($metricsMap.ContainsKey("disabled_jobs_count") -and $metricsMap["disabled_jobs_count"].status -ne "Good") {
  $midActions.Add("Job governance: verify disabled SQL Agent jobs and re-enable required operational jobs.")
}
if ($metricsMap.ContainsKey("log_used_percent") -and $metricsMap["log_used_percent"].status -ne "Good") {
  $midActions.Add("Log capacity: monitor log growth, review backup frequency, and prevent log full incidents.")
}
if ($metricsMap.ContainsKey("data_disk_used_percent") -and $metricsMap["data_disk_used_percent"].status -ne "Good") {
  $midActions.Add("Disk planning: expand storage headroom and enforce fixed-size autogrowth settings.")
}

if ($metricsMap.ContainsKey("new_logins_90d_count") -and $metricsMap["new_logins_90d_count"].status -ne "Good") {
  $lowActions.Add("Identity audit: review newly created server logins in last 90 days and confirm ownership/need.")
}
if ($metricsMap.ContainsKey("new_db_users_90d_count") -and $metricsMap["new_db_users_90d_count"].status -ne "Good") {
  $lowActions.Add("Database user audit: review newly created database users and align role grants with least privilege.")
}

if ($highActions.Count -eq 0) { $highActions.Add("No immediate high-priority risk detected in current thresholds.") }
if ($midActions.Count -eq 0) { $midActions.Add("No medium-priority risk detected in current thresholds.") }
if ($lowActions.Count -eq 0) { $lowActions.Add("No low-priority optimization item detected in current thresholds.") }

if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outMd = Join-Path $OutputDir ("remediation-" + $stamp + ".md")
$outHtml = Join-Path $OutputDir ("remediation-" + $stamp + ".html")

$mdLines = @()
$mdLines += "# SQL Server Inspection Remediation Plan"
$mdLines += ""
$mdLines += ("- Source report: " + $InspectionJsonPath)
$mdLines += ("- Server: " + $data.server)
$mdLines += ("- Database: " + $data.database)
$mdLines += ("- Collected: " + $data.collectedAt)
$mdLines += ("- Overall: " + $data.summary.overallStatus)
$mdLines += ("- Health score: " + $data.summary.healthScore)
$mdLines += ("- Critical/Warning/Good: " + $data.summary.critical + "/" + $data.summary.warning + "/" + $data.summary.good)
$mdLines += ("- Probe errors: " + $data.summary.probeErrors)
$mdLines += ""
$mdLines += "## Critical Metrics"
$mdLines += ""
$mdLines += ($critical -join ", ")
$mdLines += ""
$mdLines += "## Warning Metrics"
$mdLines += ""
$mdLines += ($warning -join ", ")
$mdLines += ""
$mdLines += "## High Priority"
foreach ($a in $highActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += "## Medium Priority"
foreach ($a in $midActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += "## Low Priority"
foreach ($a in $lowActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += "## 7-Day Execution Suggestion"
$mdLines += "1. Day 1-2: close high-priority security and backup gaps."
$mdLines += "2. Day 3-4: tune long-running queries and blocking/deadlock hotspots."
$mdLines += "3. Day 5-7: complete identity audit and baseline review."

Set-Content -Path $outMd -Encoding UTF8 -Value ($mdLines -join "`r`n")

$html = "<!doctype html><html><head><meta charset='utf-8'><title>SQL Remediation Plan</title><style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f6f8fb;color:#1f2937}h1,h2{margin:12px 0}ul{margin-top:6px}.card{background:#fff;border:1px solid #d9e2ec;border-radius:8px;padding:12px;margin:10px 0}</style></head><body>"
$html += "<h1>SQL Server Remediation Plan</h1>"
$html += "<div class='card'><b>Source:</b> " + [System.Net.WebUtility]::HtmlEncode($InspectionJsonPath)
$html += "<br><b>Server:</b> " + [System.Net.WebUtility]::HtmlEncode([string]$data.server)
$html += "<br><b>Overall:</b> " + [System.Net.WebUtility]::HtmlEncode([string]$data.summary.overallStatus)
$html += "<br><b>Health Score:</b> " + [string]$data.summary.healthScore
$html += "<br><b>Critical/Warning/Good:</b> " + [string]$data.summary.critical + "/" + [string]$data.summary.warning + "/" + [string]$data.summary.good
$html += "</div>"
$html += "<h2>High Priority</h2><ul>"
foreach ($a in $highActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul><h2>Medium Priority</h2><ul>"
foreach ($a in $midActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul><h2>Low Priority</h2><ul>"
foreach ($a in $lowActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul></body></html>"
Set-Content -Path $outHtml -Encoding UTF8 -Value $html

Write-Output ("Remediation plan generated.")
Write-Output ("MD: " + (Resolve-Path $outMd).Path)
Write-Output ("HTML: " + (Resolve-Path $outHtml).Path)
