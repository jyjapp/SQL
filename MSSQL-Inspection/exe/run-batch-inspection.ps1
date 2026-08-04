param(
  [string]$InstancesPath = ".\\instances.example.json",
  [string]$OutputRoot = ".\\output",
  [string]$SummaryFormats = "json,html",
  [bool]$GenerateRemediationPerInstance = $true,
  [string]$ThresholdPath = "..\\skill\\rules\\thresholds.json"
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

  $result = [ordered]@{ docxOk = $false; pdfOk = $false; warning = $null }
  if (-not $ExportDocx -and -not $ExportPdf) { return $result }

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

if (-not (Test-Path $InstancesPath)) {
  throw "Instances config not found: $InstancesPath"
}

$instancesConfig = Get-Content -Raw -Path $InstancesPath | ConvertFrom-Json
if ($null -eq $instancesConfig.instances -or $instancesConfig.instances.Count -eq 0) {
  throw "No instances found in config."
}

$batchFolder = Join-Path $OutputRoot ("batch-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
New-Item -ItemType Directory -Force -Path $batchFolder | Out-Null

$singleScript = Join-Path $PSScriptRoot "run-inspection.ps1"
if (-not (Test-Path $singleScript)) {
  throw "Single runner not found: $singleScript"
}

$remediationScript = Join-Path $PSScriptRoot "generate-remediation.ps1"

function Get-MetricStatusMap {
  param([object[]]$Metrics)

  $map = @{}
  foreach ($m in $Metrics) {
    $map[[string]$m.metricKey] = [string]$m.status
  }
  return $map
}

function Get-Recommendations {
  param([object[]]$Metrics)

  $statusMap = Get-MetricStatusMap -Metrics $Metrics
  $actions = New-Object System.Collections.Generic.List[string]

  if ($statusMap["weak_policy_logins_count"] -eq "Critical" -or $statusMap["weak_policy_logins_count"] -eq "Warning") {
    $actions.Add("Enable password and expiration policy for SQL logins without policy enforcement.")
  }
  if ($statusMap["sysadmin_login_count"] -eq "Critical" -or $statusMap["sysadmin_login_count"] -eq "Warning") {
    $actions.Add("Review sysadmin membership and remove unnecessary high-privilege principals.")
  }
  if ($statusMap["databases_without_recent_full_backup"] -eq "Critical" -or $statusMap["last_full_backup_age_hours"] -eq "Critical") {
    $actions.Add("Fix backup coverage immediately and verify restore path for unprotected databases.")
  }
  if ($statusMap["deadlocks_24h"] -eq "Critical" -or $statusMap["blocking_sessions"] -eq "Critical") {
    $actions.Add("Investigate blocking chain and deadlock patterns, then tune indexes and transaction scope.")
  }
  if ($statusMap["data_disk_used_percent"] -eq "Critical" -or $statusMap["log_used_percent"] -eq "Critical") {
    $actions.Add("Expand storage or free space before auto-growth and log saturation impacts write workload.")
  }
  if ($statusMap["ag_unhealthy_replicas"] -eq "Critical") {
    $actions.Add("Check AG replica synchronization health and failover readiness.")
  }

  if ($actions.Count -eq 0) {
    $actions.Add("No critical actions. Continue baseline tracking and periodic review.")
  }

  return $actions
}

function Get-RiskIndex {
  param(
    [int]$Critical,
    [int]$Warning,
    [int]$ProbeErrors
  )

  return ($Critical * 25) + ($Warning * 8) + ($ProbeErrors * 12)
}

$summary = New-Object System.Collections.Generic.List[object]

foreach ($instance in $instancesConfig.instances) {
  $instanceName = if ([string]::IsNullOrWhiteSpace([string]$instance.name)) { [string]$instance.server } else { [string]$instance.name }
  $instanceOutput = Join-Path $batchFolder $instanceName
  New-Item -ItemType Directory -Force -Path $instanceOutput | Out-Null

  try {
    $invokeParams = @{
      Server = [string]$instance.server
      Database = if ([string]::IsNullOrWhiteSpace([string]$instance.database)) { "master" } else { [string]$instance.database }
      Port = if ($null -eq $instance.port) { 1433 } else { [int]$instance.port }
      TimeoutSeconds = if ($null -eq $instance.timeoutSeconds) { 15 } else { [int]$instance.timeoutSeconds }
      ConnectionProvider = if ([string]::IsNullOrWhiteSpace([string]$instance.connectionProvider)) { "SqlClient" } else { [string]$instance.connectionProvider }
      TdsVersion = if ([string]::IsNullOrWhiteSpace([string]$instance.tdsVersion)) { "" } else { [string]$instance.tdsVersion }
      OdbcDriver = if ([string]::IsNullOrWhiteSpace([string]$instance.odbcDriver)) { "ODBC Driver 18 for SQL Server" } else { [string]$instance.odbcDriver }
      EnableFallback = if ($null -eq $instance.enableFallback) { $true } else { [bool]$instance.enableFallback }
      FallbackTdsVersions = if ([string]::IsNullOrWhiteSpace([string]$instance.fallbackTdsVersions)) { "7.4,7.3,7.2,7.1,7.0" } else { [string]$instance.fallbackTdsVersions }
      ThresholdPath = $ThresholdPath
      OutputDir = $instanceOutput
    }

    $instanceFormats = if ([string]::IsNullOrWhiteSpace([string]$instance.reportFormats)) { "json,html" } else { [string]$instance.reportFormats }
    $normalizedInstanceFormats = Normalize-ReportFormats -ReportFormats $instanceFormats
    if (-not ($normalizedInstanceFormats -contains "json")) {
      $normalizedInstanceFormats.Add("json")
    }
    $invokeParams.ReportFormats = ($normalizedInstanceFormats -join ",")

    if (-not [string]::IsNullOrWhiteSpace([string]$instance.username)) {
      $invokeParams.Username = [string]$instance.username
      $invokeParams.Password = [string]$instance.password
    }

    & $singleScript @invokeParams | Out-Null

    $latestJson = Get-ChildItem -Path $instanceOutput -Filter "inspection-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1

    if ($null -eq $latestJson) {
      throw "No inspection JSON generated."
    }

    $result = Get-Content -Raw -Path $latestJson.FullName | ConvertFrom-Json
    $criticalMetrics = @($result.metrics | Where-Object { $_.status -eq "Critical" } | ForEach-Object { [string]$_.metricKey })
    $warningMetrics = @($result.metrics | Where-Object { $_.status -eq "Warning" } | ForEach-Object { [string]$_.metricKey })
    $actions = Get-Recommendations -Metrics $result.metrics
    $riskIndex = Get-RiskIndex -Critical ([int]$result.summary.critical) -Warning ([int]$result.summary.warning) -ProbeErrors ([int]$result.summary.probeErrors)

    $remediationMdPath = ""
    $remediationHtmlPath = ""
    if ($GenerateRemediationPerInstance -and (Test-Path $remediationScript)) {
      try {
        $remOutput = & $remediationScript -InspectionJsonPath $latestJson.FullName -OutputDir $instanceOutput
        foreach ($line in @($remOutput)) {
          $text = [string]$line
          if ($text.StartsWith("MD: ")) { $remediationMdPath = $text.Substring(4).Trim() }
          if ($text.StartsWith("HTML: ")) { $remediationHtmlPath = $text.Substring(6).Trim() }
        }
      }
      catch {
      }
    }

    $summary.Add([ordered]@{
      instance = $instanceName
      server = [string]$instance.server
      overallStatus = [string]$result.summary.overallStatus
      healthScore = [int]$result.summary.healthScore
      critical = [int]$result.summary.critical
      warning = [int]$result.summary.warning
      good = [int]$result.summary.good
      probeErrors = [int]$result.summary.probeErrors
      riskIndex = $riskIndex
      criticalMetrics = $criticalMetrics
      warningMetrics = $warningMetrics
      suggestedActions = $actions
      jsonPath = $latestJson.FullName
      remediationMdPath = $remediationMdPath
      remediationHtmlPath = $remediationHtmlPath
    })
  }
  catch {
    $summary.Add([ordered]@{
      instance = $instanceName
      server = [string]$instance.server
      overallStatus = "Critical"
      healthScore = 0
      critical = 0
      warning = 0
      good = 0
      probeErrors = 1
      riskIndex = 100
      criticalMetrics = @("connection_or_runtime_failure")
      warningMetrics = @()
      suggestedActions = @("Fix connectivity/authentication first, then rerun inspection.")
      jsonPath = ""
      remediationMdPath = ""
      remediationHtmlPath = ""
      error = $_.Exception.Message
    })
  }
}

$sortedSummary = @($summary | Sort-Object -Property @{ Expression = { $_.riskIndex }; Descending = $true }, @{ Expression = { $_.critical }; Descending = $true }, @{ Expression = { $_.warning }; Descending = $true })

$topRisks = @($sortedSummary | Select-Object -First 10 | ForEach-Object {
  [ordered]@{
    instance = $_.instance
    server = $_.server
    riskIndex = $_.riskIndex
    overallStatus = $_.overallStatus
    criticalMetrics = $_.criticalMetrics
    suggestedActions = $_.suggestedActions
  }
})

$batchSummary = [ordered]@{
  project = "SQL-Server-Health-Sentinel"
  collectedAt = (Get-Date).ToString("s")
  summaryPolicy = "sorted_by_risk_desc"
  topRisks = $topRisks
  instances = $sortedSummary
}

$summaryBasePath = Join-Path $batchFolder "batch-summary"
$summaryJsonPath = "$summaryBasePath.json"
$summaryHtmlPath = "$summaryBasePath.html"
$summaryMdPath = "$summaryBasePath.md"
$summaryDocxPath = "$summaryBasePath.docx"
$summaryPdfPath = "$summaryBasePath.pdf"

$summaryFormats = Normalize-ReportFormats -ReportFormats $SummaryFormats
$summaryFormatSet = @{}
foreach ($fmt in $summaryFormats) { $summaryFormatSet[$fmt] = $true }
$generatedSummaryFiles = New-Object System.Collections.Generic.List[string]
$summaryWarnings = New-Object System.Collections.Generic.List[string]

$rank = 0
$rows = $sortedSummary | ForEach-Object {
  $rank += 1
  $statusClass = "status-good"
  if ($_.overallStatus -eq "Critical") { $statusClass = "status-critical" }
  elseif ($_.overallStatus -eq "Warning") { $statusClass = "status-warning" }

  $errorText = ""
  if ($_.Contains("error")) { $errorText = [System.Net.WebUtility]::HtmlEncode([string]$_.error) }

  $criticalMetricText = [System.Net.WebUtility]::HtmlEncode((@($_.criticalMetrics) -join ", "))
  if ([string]::IsNullOrWhiteSpace($criticalMetricText)) {
    $criticalMetricText = "-"
  }

  $actionText = [System.Net.WebUtility]::HtmlEncode((@($_.suggestedActions) -join " | "))
  if ([string]::IsNullOrWhiteSpace($actionText)) {
    $actionText = "-"
  }

  $remediationText = "-"
  if (-not [string]::IsNullOrWhiteSpace([string]$_.remediationMdPath) -or -not [string]::IsNullOrWhiteSpace([string]$_.remediationHtmlPath)) {
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$_.remediationMdPath)) { $parts += ("MD: " + [System.Net.WebUtility]::HtmlEncode([string]$_.remediationMdPath)) }
    if (-not [string]::IsNullOrWhiteSpace([string]$_.remediationHtmlPath)) { $parts += ("HTML: " + [System.Net.WebUtility]::HtmlEncode([string]$_.remediationHtmlPath)) }
    $remediationText = ($parts -join "<br>")
  }

  "<tr><td>$rank</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.instance))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.server))</td><td><span class='badge $statusClass'>$([System.Net.WebUtility]::HtmlEncode([string]$_.overallStatus))</span></td><td>$($_.riskIndex)</td><td>$($_.healthScore)</td><td>$($_.critical)</td><td>$($_.warning)</td><td>$($_.good)</td><td>$($_.probeErrors)</td><td class='mono'>$criticalMetricText</td><td class='mono'>$actionText</td><td class='mono'>$remediationText</td><td class='mono'>$errorText</td></tr>"
}

$topRiskRows = $topRisks | ForEach-Object {
  $actionText = [System.Net.WebUtility]::HtmlEncode((@($_.suggestedActions) -join " | "))
  $criticalMetricText = [System.Net.WebUtility]::HtmlEncode((@($_.criticalMetrics) -join ", "))
  "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.instance))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.server))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.overallStatus))</td><td>$([int]$_.riskIndex)</td><td class='mono'>$criticalMetricText</td><td class='mono'>$actionText</td></tr>"
}

if ($topRiskRows.Count -eq 0) {
  $topRiskRows = @("<tr><td colspan='6'>No risk data.</td></tr>")
}

$html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>SQL Server Batch Inspection Summary</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f4f7fb; }
    h1 { margin: 0 0 8px 0; }
    .meta { color: #5b6472; margin-bottom: 16px; }
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid #d9e2ec; padding: 9px; text-align: left; vertical-align: top; }
    th { background: #edf3fa; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; color: #fff; font-size: 12px; }
    .status-good { background: #0f9d58; }
    .status-warning { background: #d38b00; }
    .status-critical { background: #c62828; }
    .mono { font-family: Consolas, 'Courier New', monospace; }
  </style>
</head>
<body>
  <h1>SQL-Server-Health-Sentinel Batch Summary</h1>
  <div class="meta">Collected: $(Get-Date -Format "s") | Instances: $($sortedSummary.Count) | Policy: Sorted by risk index descending</div>

  <h2>Top Risk Instances</h2>
  <table>
    <thead>
      <tr><th>Instance</th><th>Server</th><th>Status</th><th>Risk Index</th><th>Critical Metrics</th><th>Suggested Actions</th></tr>
    </thead>
    <tbody>
      $($topRiskRows -join "`n")
    </tbody>
  </table>

  <h2>Full Ranking</h2>
  <table>
    <thead>
      <tr><th>Rank</th><th>Instance</th><th>Server</th><th>Status</th><th>Risk</th><th>Score</th><th>Critical</th><th>Warning</th><th>Good</th><th>Probe Errors</th><th>Critical Metrics</th><th>Suggested Actions</th><th>Remediation Reports</th><th>Error</th></tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>
</body>
</html>
"@

$topRiskMdRows = $topRisks | ForEach-Object {
  "| $($_.instance) | $($_.server) | $($_.overallStatus) | $($_.riskIndex) | $((@($_.criticalMetrics) -join ', ')) | $((@($_.suggestedActions) -join ' / ')) |"
}
if ($topRiskMdRows.Count -eq 0) {
  $topRiskMdRows = @("| - | - | - | - | - | - |")
}

$rankingMdRows = $sortedSummary | ForEach-Object {
  "| $($_.instance) | $($_.server) | $($_.overallStatus) | $($_.riskIndex) | $($_.healthScore) | $($_.critical) | $($_.warning) | $($_.probeErrors) |"
}

$md = @"
# SQL-Server-Health-Sentinel Batch Summary

- CollectedAt: $(Get-Date -Format "s")
- Instances: $($sortedSummary.Count)
- Policy: sorted_by_risk_desc

## Top Risk Instances

| Instance | Server | Status | RiskIndex | CriticalMetrics | SuggestedActions |
| --- | --- | --- | --- | --- | --- |
$($topRiskMdRows -join "`n")

## Full Ranking

| Instance | Server | Status | Risk | Score | Critical | Warning | ProbeErrors |
| --- | --- | --- | --- | --- | --- | --- | --- |
$($rankingMdRows -join "`n")
"@

if ($summaryFormatSet.ContainsKey("json")) {
  $batchSummary | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $summaryJsonPath
  $generatedSummaryFiles.Add($summaryJsonPath)
}

$summaryHtmlNeeded = $summaryFormatSet.ContainsKey("html") -or $summaryFormatSet.ContainsKey("docx") -or $summaryFormatSet.ContainsKey("pdf")
if ($summaryHtmlNeeded) {
  Set-Content -Encoding UTF8 -Path $summaryHtmlPath -Value $html
  if ($summaryFormatSet.ContainsKey("html")) {
    $generatedSummaryFiles.Add($summaryHtmlPath)
  }
}

if ($summaryFormatSet.ContainsKey("md")) {
  Set-Content -Encoding UTF8 -Path $summaryMdPath -Value $md
  $generatedSummaryFiles.Add($summaryMdPath)
}

if ($summaryFormatSet.ContainsKey("docx") -or $summaryFormatSet.ContainsKey("pdf")) {
  $officeResult = Convert-HtmlToOffice -HtmlPath $summaryHtmlPath -DocxPath $summaryDocxPath -PdfPath $summaryPdfPath -ExportDocx $summaryFormatSet.ContainsKey("docx") -ExportPdf $summaryFormatSet.ContainsKey("pdf")
  if ($officeResult.docxOk) { $generatedSummaryFiles.Add($summaryDocxPath) }
  if ($officeResult.pdfOk) { $generatedSummaryFiles.Add($summaryPdfPath) }
  if ($officeResult.warning) { $summaryWarnings.Add($officeResult.warning) }
}

Write-Output "Batch inspection complete."
Write-Output "Requested summary formats: $($summaryFormats -join ',')"
foreach ($f in $generatedSummaryFiles) {
  Write-Output "Generated: $f"
}
foreach ($w in $summaryWarnings) {
  Write-Output "Warning: $w"
}
