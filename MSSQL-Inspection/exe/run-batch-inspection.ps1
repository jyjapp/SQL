param(
  [string]$InstancesPath = ".\\instances.example.json",
  [string]$OutputRoot = ".\\output",
  [string]$ThresholdPath = "..\\skill\\rules\\thresholds.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
      ThresholdPath = $ThresholdPath
      OutputDir = $instanceOutput
    }

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
    $summary.Add([ordered]@{
      instance = $instanceName
      server = [string]$instance.server
      overallStatus = [string]$result.summary.overallStatus
      healthScore = [int]$result.summary.healthScore
      critical = [int]$result.summary.critical
      warning = [int]$result.summary.warning
      good = [int]$result.summary.good
      probeErrors = [int]$result.summary.probeErrors
      jsonPath = $latestJson.FullName
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
      jsonPath = ""
      error = $_.Exception.Message
    })
  }
}

$batchSummary = [ordered]@{
  project = "SQL-Server-Health-Sentinel"
  collectedAt = (Get-Date).ToString("s")
  instances = $summary
}

$summaryJsonPath = Join-Path $batchFolder "batch-summary.json"
$summaryHtmlPath = Join-Path $batchFolder "batch-summary.html"

$batchSummary | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $summaryJsonPath

$rows = $summary | ForEach-Object {
  $statusClass = "status-good"
  if ($_.overallStatus -eq "Critical") { $statusClass = "status-critical" }
  elseif ($_.overallStatus -eq "Warning") { $statusClass = "status-warning" }

  $errorText = ""
  if ($_.Contains("error")) { $errorText = [System.Net.WebUtility]::HtmlEncode([string]$_.error) }

  "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.instance))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.server))</td><td><span class='badge $statusClass'>$([System.Net.WebUtility]::HtmlEncode([string]$_.overallStatus))</span></td><td>$($_.healthScore)</td><td>$($_.critical)</td><td>$($_.warning)</td><td>$($_.good)</td><td>$($_.probeErrors)</td><td class='mono'>$errorText</td></tr>"
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
  <div class="meta">Collected: $(Get-Date -Format "s") | Instances: $($summary.Count)</div>
  <table>
    <thead>
      <tr><th>Instance</th><th>Server</th><th>Status</th><th>Score</th><th>Critical</th><th>Warning</th><th>Good</th><th>Probe Errors</th><th>Error</th></tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>
</body>
</html>
"@

Set-Content -Encoding UTF8 -Path $summaryHtmlPath -Value $html

Write-Output "Batch inspection complete."
Write-Output "Summary JSON: $summaryJsonPath"
Write-Output "Summary HTML: $summaryHtmlPath"
