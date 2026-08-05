param(
  [string]$InspectionJsonPath = "",
  [string]$OutputDir = ".\output",
  [string]$SkillConfigPath = "..\skill\config.json",
  [string]$ReportLanguage = "",
  [switch]$LocalizeJsonStatus,
  [switch]$NoFolderPrompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SkillConfig {
  param([string]$SkillConfigPath)

  $fullPath = Join-Path $PSScriptRoot $SkillConfigPath
  if (-not (Test-Path $fullPath)) {
    return $null
  }

  try {
    return Get-Content -Raw -Path $fullPath | ConvertFrom-Json
  }
  catch {
    Write-Warning "Failed to parse skill config at $fullPath. Using script defaults."
    return $null
  }
}

function Resolve-ReportLanguage {
  param(
    [string]$InputLanguage,
    [string]$SkillConfigPath
  )

  if (-not [string]::IsNullOrWhiteSpace($InputLanguage)) {
    $v = $InputLanguage.Trim().ToLowerInvariant()
    if ($v -in @("zh", "zh-cn", "zh_hans", "zh-hans", "cn", "chinese")) { return "zh-CN" }
    if ($v -in @("en", "en-us", "english")) { return "en-US" }
    return "en-US"
  }

  $cfg = Get-SkillConfig -SkillConfigPath $SkillConfigPath
  if ($null -ne $cfg -and $null -ne $cfg.report -and -not [string]::IsNullOrWhiteSpace([string]$cfg.report.language)) {
    $v = ([string]$cfg.report.language).Trim().ToLowerInvariant()
    if ($v -in @("zh", "zh-cn", "zh_hans", "zh-hans", "cn", "chinese")) { return "zh-CN" }
    if ($v -in @("en", "en-us", "english")) { return "en-US" }
  }

  return "en-US"
}

function Resolve-LocalizeJsonStatus {
  param(
    [bool]$FlagFromParam,
    [bool]$FlagProvided,
    [string]$SkillConfigPath
  )

  if ($FlagProvided) {
    return $FlagFromParam
  }

  $cfg = Get-SkillConfig -SkillConfigPath $SkillConfigPath
  if ($null -ne $cfg -and $null -ne $cfg.report -and $null -ne $cfg.report.localizeJsonStatus) {
    return [bool]$cfg.report.localizeJsonStatus
  }

  return $false
}

function Decode-UnicodeEscapes {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
  return [System.Text.RegularExpressions.Regex]::Replace(
    $Text,
    "\\u([0-9a-fA-F]{4})",
    { param($m) [char]([Convert]::ToInt32($m.Groups[1].Value, 16)) }
  )
}

function Get-I18n {
  param([string]$Language)

  if ($Language -eq "zh-CN") {
    $map = @{
      title = "SQL Server \u5de1\u68c0\u4fee\u590d\u8ba1\u5212"
      sourceReport = "\u6e90\u62a5\u544a"
      server = "\u670d\u52a1\u5668"
      database = "\u6570\u636e\u5e93"
      collected = "\u91c7\u96c6\u65f6\u95f4"
      overall = "\u603b\u4f53\u72b6\u6001"
      healthScore = "\u5065\u5eb7\u8bc4\u5206"
      statusBreakdown = "\u4e25\u91cd/\u8b66\u544a/\u826f\u597d"
      probeErrors = "\u63a2\u9488\u9519\u8bef"
      criticalMetrics = "\u4e25\u91cd\u6307\u6807"
      warningMetrics = "\u8b66\u544a\u6307\u6807"
      highPriority = "\u9ad8\u4f18\u5148\u7ea7"
      mediumPriority = "\u4e2d\u4f18\u5148\u7ea7"
      lowPriority = "\u4f4e\u4f18\u5148\u7ea7"
      executionPlan = "7\u5929\u6267\u884c\u5efa\u8bae"
      noHigh = "\u5f53\u524d\u9608\u503c\u4e0b\u672a\u53d1\u73b0\u9ad8\u4f18\u5148\u7ea7\u98ce\u9669\u3002"
      noMid = "\u5f53\u524d\u9608\u503c\u4e0b\u672a\u53d1\u73b0\u4e2d\u4f18\u5148\u7ea7\u98ce\u9669\u3002"
      noLow = "\u5f53\u524d\u9608\u503c\u4e0b\u672a\u53d1\u73b0\u4f4e\u4f18\u5148\u7ea7\u4f18\u5316\u9879\u3002"
      day1 = "1. \u7b2c1-2\u5929\uff1a\u4f18\u5148\u5173\u95ed\u9ad8\u4f18\u5148\u7ea7\u5b89\u5168\u4e0e\u5907\u4efd\u98ce\u9669\u3002"
      day2 = "2. \u7b2c3-4\u5929\uff1a\u4f18\u5316\u957f\u8017\u65f6\u8bed\u53e5\u4e0e\u963b\u585e/\u6b7b\u9501\u70ed\u70b9\u3002"
      day3 = "3. \u7b2c5-7\u5929\uff1a\u5b8c\u6210\u8eab\u4efd\u6743\u9650\u68b3\u7406\u4e0e\u57fa\u7ebf\u590d\u6838\u3002"
      actBackup = "\u5907\u4efd\u8986\u76d6\uff1a\u7acb\u5373\u4e3a\u7f3a\u5c11\u8fd1\u671f\u5168\u91cf\u5907\u4efd\u7684\u6570\u636e\u5e93\u5b89\u6392\u5168\u91cf\u5907\u4efd\uff0c\u5e76\u4f7f\u7528 RESTORE VERIFYONLY \u9a8c\u8bc1\u53ef\u6062\u590d\u6027\u3002"
      actPolicy = "\u5b89\u5168\u7b56\u7565\uff1a\u4e3a SQL \u767b\u5f55\u542f\u7528 CHECK_POLICY \u548c CHECK_EXPIRATION\uff0c\u8f6e\u6362\u5f31\u53e3\u4ee4\uff0c\u7981\u7528\u95f2\u7f6e\u8d26\u53f7\u3002"
      actSysadmin = "\u6743\u9650\u6700\u5c0f\u5316\uff1a\u5ba1\u6838 sysadmin \u6210\u5458\uff0c\u79fb\u9664\u4e0d\u5fc5\u8981\u7684\u9ad8\u6743\u9650\u4e3b\u4f53\u3002"
      actAg = "\u9ad8\u53ef\u7528\u6027\uff1a\u6392\u67e5 AG \u526f\u672c\u5065\u5eb7\u72b6\u6001\uff0c\u5e76\u9a8c\u8bc1\u6545\u969c\u5207\u6362\u9884\u6848\u3002"
      actQuery = "\u67e5\u8be2\u5ef6\u8fdf\uff1a\u5b9a\u4f4d\u957f\u8017\u65f6\u8bf7\u6c42\uff0c\u4f18\u5316\u7d22\u5f15\u3001\u6267\u884c\u8ba1\u5212\u548c\u4e8b\u52a1\u8303\u56f4\u3002"
      actBlocking = "\u963b\u585e\u63a7\u5236\uff1a\u5206\u6790\u963b\u585e\u94fe\u4e0e\u7b49\u5f85\u7c7b\u578b\uff0c\u4f18\u5316\u52a0\u9501\u987a\u5e8f\u5e76\u964d\u4f4e\u70ed\u70b9\u7ade\u4e89\u3002"
      actDeadlock = "\u6b7b\u9501\u6cbb\u7406\uff1a\u6293\u53d6\u6b7b\u9501\u56fe\uff0c\u5bf9\u51b2\u7a81\u8bbf\u95ee\u6a21\u5f0f\u8fdb\u884c\u7d22\u5f15\u4e0e\u8bbf\u95ee\u987a\u5e8f\u4f18\u5316\u3002"
      actJobs = "\u4f5c\u4e1a\u6cbb\u7406\uff1a\u68c0\u67e5\u5df2\u7981\u7528 SQL Agent \u4f5c\u4e1a\uff0c\u91cd\u65b0\u542f\u7528\u5fc5\u8981\u7684\u8fd0\u7ef4\u4f5c\u4e1a\u3002"
      actLog = "\u65e5\u5fd7\u5bb9\u91cf\uff1a\u76d1\u63a7\u65e5\u5fd7\u589e\u957f\uff0c\u5ba1\u6838\u5907\u4efd\u9891\u7387\uff0c\u9632\u6b62\u65e5\u5fd7\u5199\u6ee1\u4e8b\u6545\u3002"
      actDisk = "\u5b58\u50a8\u89c4\u5212\uff1a\u6269\u5bb9\u5b58\u50a8\u4f59\u91cf\uff0c\u5e76\u4f7f\u7528\u56fa\u5b9a\u5927\u5c0f\u7684\u81ea\u52a8\u589e\u957f\u7b56\u7565\u3002"
      actLoginAudit = "\u8eab\u4efd\u5ba1\u8ba1\uff1a\u590d\u6838\u8fd190\u5929\u65b0\u5efa\u7684\u670d\u52a1\u5668\u767b\u5f55\uff0c\u786e\u8ba4\u6240\u6709\u8005\u53ca\u5fc5\u8981\u6027\u3002"
      actDbUserAudit = "\u5e93\u7528\u6237\u5ba1\u8ba1\uff1a\u590d\u6838\u8fd190\u5929\u65b0\u5efa\u5e93\u7528\u6237\uff0c\u5e76\u4f7f\u89d2\u8272\u6388\u6743\u7b26\u5408\u6700\u5c0f\u6743\u9650\u539f\u5219\u3002"
    }
    foreach ($key in @($map.Keys)) {
      $map[$key] = Decode-UnicodeEscapes -Text ([string]$map[$key])
    }
    return $map
  }

  return @{
    title = "SQL Server Inspection Remediation Plan"
    sourceReport = "Source report"
    server = "Server"
    database = "Database"
    collected = "Collected"
    overall = "Overall"
    healthScore = "Health score"
    statusBreakdown = "Critical/Warning/Good"
    probeErrors = "Probe errors"
    criticalMetrics = "Critical Metrics"
    warningMetrics = "Warning Metrics"
    highPriority = "High Priority"
    mediumPriority = "Medium Priority"
    lowPriority = "Low Priority"
    executionPlan = "7-Day Execution Suggestion"
    noHigh = "No immediate high-priority risk detected in current thresholds."
    noMid = "No medium-priority risk detected in current thresholds."
    noLow = "No low-priority optimization item detected in current thresholds."
    day1 = "1. Day 1-2: close high-priority security and backup gaps."
    day2 = "2. Day 3-4: tune long-running queries and blocking/deadlock hotspots."
    day3 = "3. Day 5-7: complete identity audit and baseline review."
    actBackup = "Backup coverage: immediately schedule full backups for databases without recent full backup and validate restore path with RESTORE VERIFYONLY."
    actPolicy = "Security policy: enable CHECK_POLICY and CHECK_EXPIRATION for SQL logins, rotate weak credentials, disable stale accounts."
    actSysadmin = "Privilege minimization: review sysadmin membership and remove non-essential high-privilege principals."
    actAg = "HA reliability: investigate AG unhealthy replicas and validate failover readiness."
    actQuery = "Query latency: identify top long-running requests and optimize indexing, execution plans, and transaction scope."
    actBlocking = "Blocking control: analyze blocking chain and wait types, tune lock order and reduce hot-spot contention."
    actDeadlock = "Deadlock reduction: capture deadlock graphs and apply index/order fixes for conflicting access patterns."
    actJobs = "Job governance: verify disabled SQL Agent jobs and re-enable required operational jobs."
    actLog = "Log capacity: monitor log growth, review backup frequency, and prevent log full incidents."
    actDisk = "Disk planning: expand storage headroom and enforce fixed-size autogrowth settings."
    actLoginAudit = "Identity audit: review newly created server logins in last 90 days and confirm ownership/need."
    actDbUserAudit = "Database user audit: review newly created database users and align role grants with least privilege."
  }
}

function Normalize-Status {
  param([string]$Status)
  if ([string]::IsNullOrWhiteSpace($Status)) { return "Unknown" }
  $v = $Status.Trim().ToLowerInvariant()
  if ($v -in @("good", "\u826f\u597d")) { return "Good" }
  if ($v -in @("warning", "\u8b66\u544a")) { return "Warning" }
  if ($v -in @("critical", "\u4e25\u91cd")) { return "Critical" }
  return $Status
}

function Localize-Status {
  param(
    [string]$Status,
    [string]$Language,
    [bool]$EnableLocalization
  )

  $normalized = Normalize-Status -Status $Status
  if (-not $EnableLocalization -or $Language -ne "zh-CN") {
    return $normalized
  }

  switch ($normalized) {
    "Good" { return (Decode-UnicodeEscapes -Text "\u826f\u597d") }
    "Warning" { return (Decode-UnicodeEscapes -Text "\u8b66\u544a") }
    "Critical" { return (Decode-UnicodeEscapes -Text "\u4e25\u91cd") }
    default { return $normalized }
  }
}

function Get-SkillConfigFullPath {
  param([string]$SkillConfigPath)
  return (Join-Path $PSScriptRoot $SkillConfigPath)
}

function Select-OutputRootInteractively {
  param([string]$DefaultRoot)

  if ([Environment]::UserInteractive) {
    try {
      Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
      $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
      $dialog.Description = "Select output root folder for remediation reports"
      $dialog.ShowNewFolderButton = $true
      if (Test-Path $DefaultRoot) {
        $dialog.SelectedPath = (Resolve-Path $DefaultRoot).Path
      }

      $dialogResult = $dialog.ShowDialog()
      if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
        return $dialog.SelectedPath
      }
    }
    catch {
    }
  }

  $inputRoot = Read-Host "Enter remediation output root folder (press Enter to use default: $DefaultRoot)"
  if (-not [string]::IsNullOrWhiteSpace($inputRoot)) {
    return $inputRoot.Trim()
  }

  return $DefaultRoot
}

function Save-SelectedRootAsDefault {
  param(
    [string]$SkillConfigPath,
    [string]$SelectedRoot
  )

  $fullPath = Get-SkillConfigFullPath -SkillConfigPath $SkillConfigPath
  if (-not (Test-Path $fullPath)) {
    return
  }

  try {
    $config = Get-Content -Raw -Path $fullPath | ConvertFrom-Json
    if ($null -eq $config.output) {
      $config | Add-Member -MemberType NoteProperty -Name output -Value ([pscustomobject]@{})
    }

    if ($config.output.PSObject.Properties.Name -contains "defaultRootDir") {
      $config.output.defaultRootDir = $SelectedRoot
    }
    else {
      $config.output | Add-Member -MemberType NoteProperty -Name defaultRootDir -Value $SelectedRoot
    }

    $config | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -Path $fullPath
  }
  catch {
    Write-Warning "Failed to persist selected output root to config: $($_.Exception.Message)"
  }
}

function Resolve-OutputDirectory {
  param(
    [string]$RawOutputDir,
    [bool]$OutputDirExplicitlySet,
    [string]$SkillConfigPath,
    [bool]$NoFolderPrompt
  )

  if ($OutputDirExplicitlySet) {
    return $RawOutputDir
  }

  $config = Get-SkillConfig -SkillConfigPath $SkillConfigPath
  $defaultRoot = ".\output"
  $promptForFolder = $false
  $rememberSelectedRootAsDefault = $false
  $useDateSubfolder = $false
  $dateFolderFormat = "yyyyMMdd"

  if ($null -ne $config -and $null -ne $config.output) {
    if (-not [string]::IsNullOrWhiteSpace([string]$config.output.defaultRootDir)) {
      $defaultRoot = [string]$config.output.defaultRootDir
    }
    if ($null -ne $config.output.promptForFolder) {
      $promptForFolder = [bool]$config.output.promptForFolder
    }
    if ($null -ne $config.output.rememberSelectedRootAsDefault) {
      $rememberSelectedRootAsDefault = [bool]$config.output.rememberSelectedRootAsDefault
    }
    if ($null -ne $config.output.useDateSubfolder) {
      $useDateSubfolder = [bool]$config.output.useDateSubfolder
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$config.output.dateFolderFormat)) {
      $dateFolderFormat = [string]$config.output.dateFolderFormat
    }
  }

  $selectedRoot = $defaultRoot
  if ($promptForFolder -and -not $NoFolderPrompt -and [Environment]::UserInteractive) {
    $selectedRoot = Select-OutputRootInteractively -DefaultRoot $defaultRoot
    if ($rememberSelectedRootAsDefault -and -not [string]::Equals($selectedRoot, $defaultRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      Save-SelectedRootAsDefault -SkillConfigPath $SkillConfigPath -SelectedRoot $selectedRoot
    }
  }

  if ($useDateSubfolder) {
    return Join-Path $selectedRoot (Get-Date -Format $dateFolderFormat)
  }

  return $selectedRoot
}

$resolvedOutputDir = Resolve-OutputDirectory -RawOutputDir $OutputDir -OutputDirExplicitlySet $PSBoundParameters.ContainsKey("OutputDir") -SkillConfigPath $SkillConfigPath -NoFolderPrompt:$NoFolderPrompt
Write-Output "Resolved output directory: $resolvedOutputDir"

if ([string]::IsNullOrWhiteSpace($InspectionJsonPath)) {
  $latest = Get-ChildItem -Path $resolvedOutputDir -Filter "complete-inspection-*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  if ($null -eq $latest) {
    $latest = Get-ChildItem -Path $resolvedOutputDir -Filter "inspection-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
  }

  if ($null -eq $latest) {
    $latest = Get-ChildItem -Path $resolvedOutputDir -Filter "full-audit-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
  }

  if ($null -eq $latest) {
    $latest = Get-ChildItem -Path $resolvedOutputDir -Recurse -Filter "complete-inspection-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
  }

  if ($null -eq $latest) {
    $latest = Get-ChildItem -Path $resolvedOutputDir -Recurse -Filter "inspection-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
  }

  if ($null -eq $latest) {
    $latest = Get-ChildItem -Path $resolvedOutputDir -Recurse -Filter "full-audit-*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  }
  if ($null -eq $latest) {
    throw "No inspection JSON found in $resolvedOutputDir"
  }
  $InspectionJsonPath = $latest.FullName
}

if (-not (Test-Path $InspectionJsonPath)) {
  throw "Inspection JSON not found: $InspectionJsonPath"
}

$data = Get-Content -Raw -Path $InspectionJsonPath | ConvertFrom-Json
$hasMetrics = ($null -ne $data.PSObject.Properties["metrics"] -and $null -ne $data.metrics)
if (-not $hasMetrics) {
  throw "Invalid inspection JSON for remediation: metrics section not found. Please provide a single-run report file like complete-inspection-*.json (not complete-inspection-batch-summary.json)."
}
$resolvedReportLanguage = Resolve-ReportLanguage -InputLanguage $ReportLanguage -SkillConfigPath $SkillConfigPath
$enableStatusLocalization = Resolve-LocalizeJsonStatus -FlagFromParam $LocalizeJsonStatus.IsPresent -FlagProvided $PSBoundParameters.ContainsKey("LocalizeJsonStatus") -SkillConfigPath $SkillConfigPath
$i18n = Get-I18n -Language $resolvedReportLanguage

$metricsMap = @{}
foreach ($m in $data.metrics) {
  $metricsMap[[string]$m.metricKey] = $m
}

$critical = @($data.metrics | Where-Object { (Normalize-Status -Status ([string]$_.status)) -eq "Critical" } | ForEach-Object { $_.metricKey })
$warning = @($data.metrics | Where-Object { (Normalize-Status -Status ([string]$_.status)) -eq "Warning" } | ForEach-Object { $_.metricKey })
if ($critical.Count -eq 0) { $critical = @("none") }
if ($warning.Count -eq 0) { $warning = @("none") }

$criticalCount = @($data.metrics | Where-Object { (Normalize-Status -Status ([string]$_.status)) -eq "Critical" }).Count
$warningCount = @($data.metrics | Where-Object { (Normalize-Status -Status ([string]$_.status)) -eq "Warning" }).Count
$goodCount = @($data.metrics | Where-Object { (Normalize-Status -Status ([string]$_.status)) -eq "Good" }).Count

$highActions = New-Object System.Collections.Generic.List[string]
$midActions = New-Object System.Collections.Generic.List[string]
$lowActions = New-Object System.Collections.Generic.List[string]

if ($metricsMap.ContainsKey("databases_without_recent_full_backup") -and (Normalize-Status -Status ([string]$metricsMap["databases_without_recent_full_backup"].status)) -ne "Good") {
  $highActions.Add([string]$i18n.actBackup)
}
if ($metricsMap.ContainsKey("weak_policy_logins_count") -and (Normalize-Status -Status ([string]$metricsMap["weak_policy_logins_count"].status)) -ne "Good") {
  $highActions.Add([string]$i18n.actPolicy)
}
if ($metricsMap.ContainsKey("sysadmin_login_count") -and (Normalize-Status -Status ([string]$metricsMap["sysadmin_login_count"].status)) -ne "Good") {
  $highActions.Add([string]$i18n.actSysadmin)
}
if ($metricsMap.ContainsKey("ag_unhealthy_replicas") -and (Normalize-Status -Status ([string]$metricsMap["ag_unhealthy_replicas"].status)) -ne "Good") {
  $highActions.Add([string]$i18n.actAg)
}

if ($metricsMap.ContainsKey("longest_running_query_seconds") -and (Normalize-Status -Status ([string]$metricsMap["longest_running_query_seconds"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actQuery)
}
if ($metricsMap.ContainsKey("blocking_sessions") -and (Normalize-Status -Status ([string]$metricsMap["blocking_sessions"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actBlocking)
}
if ($metricsMap.ContainsKey("deadlocks_24h") -and (Normalize-Status -Status ([string]$metricsMap["deadlocks_24h"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actDeadlock)
}
if ($metricsMap.ContainsKey("disabled_jobs_count") -and (Normalize-Status -Status ([string]$metricsMap["disabled_jobs_count"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actJobs)
}
if ($metricsMap.ContainsKey("log_used_percent") -and (Normalize-Status -Status ([string]$metricsMap["log_used_percent"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actLog)
}
if ($metricsMap.ContainsKey("data_disk_used_percent") -and (Normalize-Status -Status ([string]$metricsMap["data_disk_used_percent"].status)) -ne "Good") {
  $midActions.Add([string]$i18n.actDisk)
}

if ($metricsMap.ContainsKey("new_logins_90d_count") -and (Normalize-Status -Status ([string]$metricsMap["new_logins_90d_count"].status)) -ne "Good") {
  $lowActions.Add([string]$i18n.actLoginAudit)
}
if ($metricsMap.ContainsKey("new_db_users_90d_count") -and (Normalize-Status -Status ([string]$metricsMap["new_db_users_90d_count"].status)) -ne "Good") {
  $lowActions.Add([string]$i18n.actDbUserAudit)
}

if ($highActions.Count -eq 0) { $highActions.Add([string]$i18n.noHigh) }
if ($midActions.Count -eq 0) { $midActions.Add([string]$i18n.noMid) }
if ($lowActions.Count -eq 0) { $lowActions.Add([string]$i18n.noLow) }

if (-not (Test-Path $resolvedOutputDir)) {
  New-Item -ItemType Directory -Force -Path $resolvedOutputDir | Out-Null
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outMd = Join-Path $resolvedOutputDir ("remediation-" + $stamp + ".md")
$outHtml = Join-Path $resolvedOutputDir ("remediation-" + $stamp + ".html")

$overallStatusDisplay = Localize-Status -Status ([string]$data.summary.overallStatus) -Language $resolvedReportLanguage -EnableLocalization $enableStatusLocalization

$mdLines = @()
$mdLines += ("# " + [string]$i18n.title)
$mdLines += ""
$mdLines += ("- " + [string]$i18n.sourceReport + ": " + $InspectionJsonPath)
$mdLines += ("- " + [string]$i18n.server + ": " + $data.server)
$mdLines += ("- " + [string]$i18n.database + ": " + $data.database)
$mdLines += ("- " + [string]$i18n.collected + ": " + $data.collectedAt)
$mdLines += ("- " + [string]$i18n.overall + ": " + $overallStatusDisplay)
$mdLines += ("- " + [string]$i18n.healthScore + ": " + $data.summary.healthScore)
$mdLines += ("- " + [string]$i18n.statusBreakdown + ": " + $criticalCount + "/" + $warningCount + "/" + $goodCount)
$mdLines += ("- " + [string]$i18n.probeErrors + ": " + $data.summary.probeErrors)
$mdLines += ""
$mdLines += ("## " + [string]$i18n.criticalMetrics)
$mdLines += ""
$mdLines += ($critical -join ", ")
$mdLines += ""
$mdLines += ("## " + [string]$i18n.warningMetrics)
$mdLines += ""
$mdLines += ($warning -join ", ")
$mdLines += ""
$mdLines += ("## " + [string]$i18n.highPriority)
foreach ($a in $highActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += ("## " + [string]$i18n.mediumPriority)
foreach ($a in $midActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += ("## " + [string]$i18n.lowPriority)
foreach ($a in $lowActions) { $mdLines += ("- " + $a) }
$mdLines += ""
$mdLines += ("## " + [string]$i18n.executionPlan)
$mdLines += [string]$i18n.day1
$mdLines += [string]$i18n.day2
$mdLines += [string]$i18n.day3

Set-Content -Path $outMd -Encoding UTF8 -Value ($mdLines -join "`r`n")

$html = "<!doctype html><html><head><meta charset='utf-8'><title>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.title) + "</title><style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f6f8fb;color:#1f2937}h1,h2{margin:12px 0}ul{margin-top:6px}.card{background:#fff;border:1px solid #d9e2ec;border-radius:8px;padding:12px;margin:10px 0}</style></head><body>"
$html += "<h1>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.title) + "</h1>"
$html += "<div class='card'><b>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.sourceReport) + ":</b> " + [System.Net.WebUtility]::HtmlEncode($InspectionJsonPath)
$html += "<br><b>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.server) + ":</b> " + [System.Net.WebUtility]::HtmlEncode([string]$data.server)
$html += "<br><b>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.overall) + ":</b> " + [System.Net.WebUtility]::HtmlEncode([string]$overallStatusDisplay)
$html += "<br><b>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.healthScore) + ":</b> " + [string]$data.summary.healthScore
$html += "<br><b>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.statusBreakdown) + ":</b> " + [string]$criticalCount + "/" + [string]$warningCount + "/" + [string]$goodCount
$html += "</div>"
$html += "<h2>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.highPriority) + "</h2><ul>"
foreach ($a in $highActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul><h2>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.mediumPriority) + "</h2><ul>"
foreach ($a in $midActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul><h2>" + [System.Net.WebUtility]::HtmlEncode([string]$i18n.lowPriority) + "</h2><ul>"
foreach ($a in $lowActions) { $html += "<li>" + [System.Net.WebUtility]::HtmlEncode($a) + "</li>" }
$html += "</ul></body></html>"
Set-Content -Path $outHtml -Encoding UTF8 -Value $html

Write-Output ("Remediation plan generated.")
Write-Output ("MD: " + (Resolve-Path $outMd).Path)
Write-Output ("HTML: " + (Resolve-Path $outHtml).Path)
