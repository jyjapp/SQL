param(
  [string]$InstancesPath = "",
  [string]$WorkspaceDir = "",
  [string]$ServerConfigDir = "",
  [string]$ServerConfigFile = "servers.config",
  [string]$SkillConfigPath = "..\..\skill\config.json",
  [string]$OutputRoot = "D:\download\SQL",
  [string]$ReportFormats = "json,html",
  [string]$ReportLanguage = "",
  [switch]$LocalizeJsonStatus,
  [int]$TimeoutSeconds = 30,
  [switch]$NoFolderPrompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SkillConfig {
  param([string]$RelativePath)
  $p = Join-Path $PSScriptRoot $RelativePath
  if (-not (Test-Path $p)) { return $null }
  try { return (Get-Content -Raw -Path $p | ConvertFrom-Json) }
  catch { return $null }
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

  $cfg = Get-SkillConfig -RelativePath $SkillConfigPath
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

  $cfg = Get-SkillConfig -RelativePath $SkillConfigPath
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
      batchTitle = "\u5b8c\u6574\u5de1\u68c0\u6279\u91cf\u6c47\u603b"
      generatedAt = "\u751f\u6210\u65f6\u95f4"
      instance = "\u5b9e\u4f8b"
      server = "\u670d\u52a1\u5668"
      status = "\u72b6\u6001"
      combinedScore = "\u7efc\u5408\u8bc4\u5206"
      mvpScore = "MVP \u8bc4\u5206"
      auditScore = "\u5b8c\u6574\u5de1\u68c0\u8bc4\u5206"
      mvpCritical = "MVP \u4e25\u91cd\u9879"
      mvpWarning = "MVP \u8b66\u544a\u9879"
      probeErrors = "\u63a2\u9488\u9519\u8bef"
      error = "\u9519\u8bef"
    }
    foreach ($key in @($map.Keys)) {
      $map[$key] = Decode-UnicodeEscapes -Text ([string]$map[$key])
    }
    return $map
  }

  return @{
    batchTitle = "Complete Inspection Batch Summary"
    generatedAt = "Generated at"
    instance = "Instance"
    server = "Server"
    status = "Status"
    combinedScore = "Combined Score"
    mvpScore = "MVP Score"
    auditScore = "Audit Score"
    mvpCritical = "MVP Critical"
    mvpWarning = "MVP Warning"
    probeErrors = "Probe Errors"
    error = "Error"
  }
}

function Localize-Status {
  param(
    [string]$Status,
    [string]$Language
  )

  if ($Language -ne "zh-CN") { return $Status }
  switch ($Status) {
    "Good" { return (Decode-UnicodeEscapes -Text "\u826f\u597d") }
    "Warning" { return (Decode-UnicodeEscapes -Text "\u8b66\u544a") }
    "Critical" { return (Decode-UnicodeEscapes -Text "\u4e25\u91cd") }
    default { return $Status }
  }
}

function Parse-ServerConfigText {
  param([string]$Text)

  $section = ""
  $globalSection = [ordered]@{}
  $instances = New-Object System.Collections.Generic.List[object]

  foreach ($rawLine in ($Text -split "`r?`n")) {
    $line = $rawLine.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line.StartsWith("#") -or $line.StartsWith(";")) { continue }

    if ($line -match "^\[(.+)\]$") {
      $section = $Matches[1].Trim().ToLowerInvariant()
      continue
    }

    if ($section -eq "global") {
      $parts = $line -split "`t", 2
      if ($parts.Count -lt 2) { $parts = $line -split "=", 2 }
      if ($parts.Count -lt 2 -and $line -match "^(\S+)[\t ]+(.+)$") {
        $parts = @($Matches[1], $Matches[2])
      }
      if ($parts.Count -lt 2) { continue }
      $k = $parts[0].Trim().ToLowerInvariant()
      $v = $parts[1].Trim()
      if (-not [string]::IsNullOrWhiteSpace($k)) {
        $globalSection[$k] = $v
      }
      continue
    }

    if ($section -eq "servers") {
      $parts = [regex]::Split($line, "[\t ]+") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      if ($parts.Count -lt 2) { continue }

      $col0 = $parts[0].Trim().ToLowerInvariant()
      $col1 = $parts[1].Trim().ToLowerInvariant()
      if (($col0 -eq "ip" -or $col0 -eq "server") -and $col1 -eq "port") { continue }

      $ip = $parts[0].Trim()
      if ([string]::IsNullOrWhiteSpace($ip)) { continue }

      $port = 1433
      if ($parts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($parts[1])) {
        [void][int]::TryParse($parts[1].Trim(), [ref]$port)
      }
      $username = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }
      $password = if ($parts.Count -ge 4) { $parts[3].Trim() } else { "" }

      $instances.Add([pscustomobject]@{
        name = $ip
        server = $ip
        database = "master"
        port = $port
        username = $username
        password = $password
        connectionProvider = "SqlClient"
        tdsVersion = ""
        odbcDriver = "ODBC Driver 18 for SQL Server"
        reportFormats = ""
        timeoutSeconds = 30
      })
    }
  }

  $defaultSingle = $null
  if ($instances.Count -gt 0) {
    $defaultSingle = $instances[0]
  }

  $cfg = @{}
  $cfg['globalSettings'] = $globalSection
  $cfg['defaultSingle'] = $defaultSingle
  $cfg['instances'] = $instances.ToArray()
  return $cfg
}

function Get-ServerConfigFromPath {
  param([string]$FullPath)
  if (-not (Test-Path $FullPath)) { return $null }

  $raw = Get-Content -Raw -Path $FullPath
  try {
    $obj = $raw | ConvertFrom-Json
    if ($null -ne $obj) { return $obj }
  }
  catch {
  }

  return (Parse-ServerConfigText -Text $raw)
}

function Select-WorkingRootInteractively {
  param([string]$DefaultRoot)

  if ([Environment]::UserInteractive) {
    try {
      Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
      $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
      $dialog.Description = "Select working folder for reports and servers.config"
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

  $inputRoot = Read-Host "Enter working folder (press Enter to use default: $DefaultRoot)"
  if (-not [string]::IsNullOrWhiteSpace($inputRoot)) {
    return $inputRoot.Trim()
  }

  return $DefaultRoot
}

function Resolve-WorkingDirectory {
  param(
    [string]$WorkspaceDir,
    [string]$OutputRoot,
    [string]$ServerConfigDir,
    [string]$SkillConfigPath,
    [bool]$NoFolderPrompt,
    [bool]$WorkspaceDirExplicitlySet,
    [bool]$OutputRootExplicitlySet,
    [bool]$ServerConfigDirExplicitlySet
  )

  $skillCfg = Get-SkillConfig -RelativePath $SkillConfigPath
  $defaultRoot = "D:\download\SQL"
  if ($null -ne $skillCfg -and $null -ne $skillCfg.output -and -not [string]::IsNullOrWhiteSpace([string]$skillCfg.output.defaultRootDir)) {
    $defaultRoot = [string]$skillCfg.output.defaultRootDir
  }

  if ($WorkspaceDirExplicitlySet -and -not [string]::IsNullOrWhiteSpace($WorkspaceDir)) {
    return $WorkspaceDir
  }

  if ($OutputRootExplicitlySet -and -not [string]::IsNullOrWhiteSpace($OutputRoot)) {
    return $OutputRoot
  }

  $configRoot = $defaultRoot
  if ($ServerConfigDirExplicitlySet -and -not [string]::IsNullOrWhiteSpace($ServerConfigDir)) {
    $configRoot = $ServerConfigDir
  }
  elseif (-not $NoFolderPrompt -and [Environment]::UserInteractive) {
    $configRoot = Select-WorkingRootInteractively -DefaultRoot $defaultRoot
  }

  return $configRoot
}

function Ensure-ServerConfigFile {
  param(
    [string]$WorkingDirectory,
    [string]$ServerConfigFile
  )

  $fullPath = Join-Path $WorkingDirectory $ServerConfigFile
  if (Test-Path $fullPath) {
    return $fullPath
  }

  $samplePath = Join-Path $PSScriptRoot ".\config\servers.sample.tsv"
  if (Test-Path $samplePath) {
    Copy-Item -Path $samplePath -Destination $fullPath -Force
  }
  else {
    $sampleContent = @"
[global]
language`tzh-CN
reportFormats`tjson,html

[servers]
ip`tport`tusername`tpassword
db-host.example.com`t1433`t`t
"@
    Set-Content -Encoding UTF8 -Path $fullPath -Value $sampleContent
  }

  Write-Output "Sample server list generated: $fullPath"
  Write-Output "Please fill server config with actual server list, then rerun this skill."
  throw "server config sample created. Fill actual values and rerun."
}

$instances = @()
$enableJsonStatusLocalization = Resolve-LocalizeJsonStatus -FlagFromParam $LocalizeJsonStatus.IsPresent -FlagProvided $PSBoundParameters.ContainsKey("LocalizeJsonStatus") -SkillConfigPath $SkillConfigPath
$workingDirectory = Resolve-WorkingDirectory -WorkspaceDir $WorkspaceDir -OutputRoot $OutputRoot -ServerConfigDir $ServerConfigDir -SkillConfigPath $SkillConfigPath -NoFolderPrompt:$NoFolderPrompt -WorkspaceDirExplicitlySet $PSBoundParameters.ContainsKey("WorkspaceDir") -OutputRootExplicitlySet $PSBoundParameters.ContainsKey("OutputRoot") -ServerConfigDirExplicitlySet $PSBoundParameters.ContainsKey("ServerConfigDir")
if (-not (Test-Path $workingDirectory)) {
  New-Item -ItemType Directory -Force -Path $workingDirectory | Out-Null
}
Write-Output "Selected working directory: $workingDirectory"

if ($PSBoundParameters.ContainsKey("InstancesPath") -and -not [string]::IsNullOrWhiteSpace($InstancesPath) -and (Test-Path $InstancesPath)) {
  $instancesConfig = Get-Content -Raw -Path $InstancesPath | ConvertFrom-Json
  if ($null -ne $instancesConfig.instances -and $instancesConfig.instances.Count -gt 0) {
    $instances = @($instancesConfig.instances)
    Write-Output "Instances loaded from: $InstancesPath"
  }
}

if ($instances.Count -eq 0) {
  $serverConfigFullPath = Ensure-ServerConfigFile -WorkingDirectory $workingDirectory -ServerConfigFile $ServerConfigFile
  Write-Output "Server config file: $serverConfigFullPath"
  if (Test-Path $serverConfigFullPath) {
    $serverConfig = Get-ServerConfigFromPath -FullPath $serverConfigFullPath

    if ($null -ne $serverConfig -and $null -ne $serverConfig.globalSettings) {
      $g = $serverConfig.globalSettings
      if (-not $PSBoundParameters.ContainsKey("ReportLanguage") -and $g.Contains("language") -and -not [string]::IsNullOrWhiteSpace([string]$g["language"])) {
        $ReportLanguage = [string]$g["language"]
      }
      if (-not $PSBoundParameters.ContainsKey("ReportFormats") -and $g.Contains("reportformats") -and -not [string]::IsNullOrWhiteSpace([string]$g["reportformats"])) {
        $ReportFormats = [string]$g["reportformats"]
      }
    }

    if ($null -ne $serverConfig.instances -and $serverConfig.instances.Count -gt 0) {
      $instances = @($serverConfig.instances)
      Write-Output "Instances loaded from: $serverConfigFullPath"
    }
    elseif ($null -ne $serverConfig.defaultSingle -and -not [string]::IsNullOrWhiteSpace([string]$serverConfig.defaultSingle.server)) {
      $instances = @($serverConfig.defaultSingle)
      Write-Output "Using defaultSingle from: $serverConfigFullPath"
    }
  }
}

if ($instances.Count -eq 0) {
  throw "No instances found. Configure either -InstancesPath or external server config file."
}

$resolvedReportLanguage = Resolve-ReportLanguage -InputLanguage $ReportLanguage -SkillConfigPath $SkillConfigPath
$i18n = Get-I18n -Language $resolvedReportLanguage

$datedRoot = Join-Path $workingDirectory (Get-Date -Format "yyyyMMdd")
$batchFolder = Join-Path $datedRoot ("complete-inspection-batch-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
New-Item -ItemType Directory -Force -Path $batchFolder | Out-Null
Write-Output "Resolved output directory: $batchFolder"

$singleScript = Join-Path $PSScriptRoot "run-inspection.ps1"
if (-not (Test-Path $singleScript)) {
  throw "Runner not found: $singleScript"
}

$summary = New-Object System.Collections.Generic.List[object]

foreach ($instance in $instances) {
  $instanceName = if ([string]::IsNullOrWhiteSpace([string]$instance.name)) { [string]$instance.server } else { [string]$instance.name }
  $instanceOutput = Join-Path $batchFolder $instanceName
  New-Item -ItemType Directory -Force -Path $instanceOutput | Out-Null

  try {
    $invokeParams = @{
      Server = [string]$instance.server
      Database = if ([string]::IsNullOrWhiteSpace([string]$instance.database)) { "master" } else { [string]$instance.database }
      Port = if ($null -eq $instance.port) { 1433 } else { [int]$instance.port }
      TimeoutSeconds = if ($null -eq $instance.timeoutSeconds) { $TimeoutSeconds } else { [int]$instance.timeoutSeconds }
      ConnectionProvider = if ([string]::IsNullOrWhiteSpace([string]$instance.connectionProvider)) { "SqlClient" } else { [string]$instance.connectionProvider }
      TdsVersion = if ([string]::IsNullOrWhiteSpace([string]$instance.tdsVersion)) { "" } else { [string]$instance.tdsVersion }
      OdbcDriver = if ([string]::IsNullOrWhiteSpace([string]$instance.odbcDriver)) { "ODBC Driver 18 for SQL Server" } else { [string]$instance.odbcDriver }
      OutputDir = $instanceOutput
      ReportFormats = if ([string]::IsNullOrWhiteSpace([string]$instance.reportFormats)) { $ReportFormats } else { [string]$instance.reportFormats }
      ReportLanguage = $resolvedReportLanguage
      LocalizeJsonStatus = $enableJsonStatusLocalization
      NoFolderPrompt = $true
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$instance.username)) {
      $invokeParams.Username = [string]$instance.username
      $invokeParams.Password = [string]$instance.password
    }

    & $singleScript @invokeParams | Out-Null

    $latestJson = Get-ChildItem -Path $instanceOutput -Recurse -Filter "complete-inspection-*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1

    if ($null -eq $latestJson) {
      throw "No complete inspection JSON generated."
    }

    $result = Get-Content -Raw -Path $latestJson.FullName | ConvertFrom-Json

    $summary.Add([ordered]@{
      instance = $instanceName
      server = [string]$instance.server
      overallStatus = if ($enableJsonStatusLocalization) { Localize-Status -Status ([string]$result.summary.overallStatus) -Language $resolvedReportLanguage } else { [string]$result.summary.overallStatus }
      healthScore = [int]$result.summary.healthScore
      mvpScore = [int]$result.summary.mvpScore
      auditScore = [int]$result.summary.auditScore
      mvpCritical = [int]$result.summary.mvpCritical
      mvpWarning = [int]$result.summary.mvpWarning
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
      mvpScore = 0
      auditScore = 0
      mvpCritical = 0
      mvpWarning = 0
      probeErrors = 1
      jsonPath = ""
      error = $_.Exception.Message
    })
  }
}

$sorted = @($summary | Sort-Object -Property @{ Expression = { $_.healthScore }; Descending = $false }, @{ Expression = { $_.probeErrors }; Descending = $true })
$report = [ordered]@{
  project = "SQL-Server-Complete-Inspection"
  collectedAt = (Get-Date).ToString("s")
  instances = $sorted
}

$basePath = Join-Path $batchFolder "complete-inspection-batch-summary"
$jsonPath = "$basePath.json"
$htmlPath = "$basePath.html"

$report | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $jsonPath

$rows = $sorted | ForEach-Object {
  $errorText = ""
  if ($_.Contains("error")) { $errorText = [System.Net.WebUtility]::HtmlEncode([string]$_.error) }
  "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.instance))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.server))</td><td>$([System.Net.WebUtility]::HtmlEncode((Localize-Status -Status ([string]$_.overallStatus) -Language $resolvedReportLanguage)))</td><td>$($_.healthScore)</td><td>$($_.mvpScore)</td><td>$($_.auditScore)</td><td>$($_.mvpCritical)</td><td>$($_.mvpWarning)</td><td>$($_.probeErrors)</td><td>$errorText</td></tr>"
}

$html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.batchTitle))</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f5f7fb; color: #1f2937; }
    table { border-collapse: collapse; width: 100%; background: #fff; }
    th, td { border: 1px solid #d9e2ec; padding: 8px 10px; text-align: left; vertical-align: top; }
    th { background: #edf3fa; }
  </style>
</head>
<body>
  <h1>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.batchTitle))</h1>
  <p>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.generatedAt)): $(Get-Date -Format "s")</p>
  <table>
    <thead>
      <tr><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.instance))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.server))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.status))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.combinedScore))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.mvpScore))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.auditScore))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.mvpCritical))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.mvpWarning))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.probeErrors))</th><th>$([System.Net.WebUtility]::HtmlEncode([string]$i18n.error))</th></tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>
</body>
</html>
"@

Set-Content -Encoding UTF8 -Path $htmlPath -Value $html
Write-Output "Generated: $jsonPath"
Write-Output "Generated: $htmlPath"
