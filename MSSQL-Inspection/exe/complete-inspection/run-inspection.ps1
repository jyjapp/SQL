param(
  [string]$Server = "localhost",
  [string]$Database = "master",
  [string]$Username = "",
  [string]$Password = "",
  [int]$Port = 1433,
  [int]$TimeoutSeconds = 30,
  [ValidateSet("SqlClient", "Odbc")]
  [string]$ConnectionProvider = "SqlClient",
  [string]$TdsVersion = "",
  [string]$OdbcDriver = "ODBC Driver 18 for SQL Server",
  [string]$ReportFormats = "json,html",
  [string]$ReportLanguage = "",
  [switch]$LocalizeJsonStatus,
  [string]$ThresholdPath = "..\\..\\skill\\rules\\thresholds.json",
  [string]$ScoreConfigPath = ".\\config\\score-rules.json",
  [string]$AuditChecksRoot = ".\\checks",
  [string]$WorkspaceDir = "",
  [string]$ServerConfigDir = "",
  [string]$ServerConfigFile = "servers.config",
  [string]$ServerConfigKey = "defaultSingle",
  [string]$SkillConfigPath = "..\\..\\skill\\config.json",
  [string]$OutputDir = ".\\output",
  [switch]$NoFolderPrompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-ReportFormats {
  param([string]$Formats)
  $allowed = @("json", "html")
  $list = New-Object System.Collections.Generic.List[string]
  foreach ($token in ($Formats -split ',')) {
    $v = $token.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($v)) { continue }
    if ($allowed -contains $v -and -not $list.Contains($v)) {
      $list.Add($v)
    }
  }
  if ($list.Count -eq 0) {
    $list.Add("json")
    $list.Add("html")
  }
  return $list
}

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
      reportTitle = "SQL Server \u5b8c\u6574\u5de1\u68c0\u62a5\u544a"
      batchTitle = "\u5b8c\u6574\u5de1\u68c0\u6279\u91cf\u6c47\u603b"
      generatedAt = "\u751f\u6210\u65f6\u95f4"
      server = "\u670d\u52a1\u5668"
      collectedAt = "\u91c7\u96c6\u65f6\u95f4"
      overall = "\u603b\u4f53\u72b6\u6001"
      combinedScore = "\u7efc\u5408\u8bc4\u5206"
      mvpScore = "MVP \u8bc4\u5206"
      fullAuditScore = "\u5b8c\u6574\u5de1\u68c0\u8bc4\u5206"
      mvpCritical = "MVP \u4e25\u91cd\u9879"
      mvpWarning = "MVP \u8b66\u544a\u9879"
      probeErrors = "\u63a2\u9488\u9519\u8bef"
      mvpMetrics = "MVP \u6307\u6807"
      metric = "\u6307\u6807"
      category = "\u5206\u7c7b"
      value = "\u6570\u503c"
      unit = "\u5355\u4f4d"
      status = "\u72b6\u6001"
      source = "\u6765\u6e90"
      rows = "\u884c\u6570"
      noRows = "\u65e0\u6570\u636e\u8fd4\u56de\u3002"
      instance = "\u5b9e\u4f8b"
      error = "\u9519\u8bef"
    }
    foreach ($key in @($map.Keys)) {
      $map[$key] = Decode-UnicodeEscapes -Text ([string]$map[$key])
    }
    return $map
  }

  return @{
    reportTitle = "SQL Server Complete Inspection Report"
    batchTitle = "Complete Inspection Batch Summary"
    generatedAt = "Generated at"
    server = "Server"
    collectedAt = "Collected"
    overall = "Overall"
    combinedScore = "Combined Score"
    mvpScore = "MVP Score"
    fullAuditScore = "Full Audit Score"
    mvpCritical = "MVP Critical"
    mvpWarning = "MVP Warning"
    probeErrors = "Probe Errors"
    mvpMetrics = "MVP Metrics"
    metric = "Metric"
    category = "Category"
    value = "Value"
    unit = "Unit"
    status = "Status"
    source = "Source"
    rows = "Rows"
    noRows = "No rows returned."
    instance = "Instance"
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

function Get-JsonResultWithLocalizedStatus {
  param(
    [object]$Result,
    [string]$Language,
    [bool]$EnableLocalization
  )

  if (-not $EnableLocalization) {
    return $Result
  }

  $cloned = ($Result | ConvertTo-Json -Depth 12 | ConvertFrom-Json)
  if ($null -ne $cloned.summary -and -not [string]::IsNullOrWhiteSpace([string]$cloned.summary.overallStatus)) {
    $cloned.summary.overallStatus = Localize-Status -Status ([string]$cloned.summary.overallStatus) -Language $Language
  }

  if ($null -ne $cloned.metrics) {
    foreach ($m in $cloned.metrics) {
      if ($null -ne $m -and -not [string]::IsNullOrWhiteSpace([string]$m.status)) {
        $m.status = Localize-Status -Status ([string]$m.status) -Language $Language
      }
    }
  }

  return $cloned
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

function Select-OutputRootInteractively {
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
    [string]$RawOutputDir,
    [string]$ServerConfigDir,
    [bool]$WorkspaceDirExplicitlySet,
    [bool]$OutputDirExplicitlySet,
    [bool]$ServerConfigDirExplicitlySet,
    [string]$SkillConfigPath,
    [bool]$NoFolderPrompt
  )

  $config = Get-SkillConfig -RelativePath $SkillConfigPath
  $defaultRoot = "D:\\download\\SQL"
  $promptForFolder = $true

  if ($null -ne $config -and $null -ne $config.output) {
    if (-not [string]::IsNullOrWhiteSpace([string]$config.output.defaultRootDir)) {
      $defaultRoot = [string]$config.output.defaultRootDir
    }
    if ($null -ne $config.output.promptForFolder) {
      $promptForFolder = [bool]$config.output.promptForFolder
    }
  }

  if ($WorkspaceDirExplicitlySet -and -not [string]::IsNullOrWhiteSpace($WorkspaceDir)) {
    return $WorkspaceDir
  }

  if ($OutputDirExplicitlySet -and -not [string]::IsNullOrWhiteSpace($RawOutputDir)) {
    return $RawOutputDir
  }

  if ($ServerConfigDirExplicitlySet -and -not [string]::IsNullOrWhiteSpace($ServerConfigDir)) {
    return $ServerConfigDir
  }

  if ($promptForFolder -and -not $NoFolderPrompt -and [Environment]::UserInteractive) {
    return (Select-OutputRootInteractively -DefaultRoot $defaultRoot)
  }

  return $defaultRoot
}

function Resolve-OutputDirectory {
  param(
    [string]$WorkingDirectory,
    [string]$SkillConfigPath
  )

  $config = Get-SkillConfig -RelativePath $SkillConfigPath
  $useDateSubfolder = $true
  $dateFolderFormat = "yyyyMMdd"

  if ($null -ne $config -and $null -ne $config.output) {
    if ($null -ne $config.output.useDateSubfolder) {
      $useDateSubfolder = [bool]$config.output.useDateSubfolder
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$config.output.dateFolderFormat)) {
      $dateFolderFormat = [string]$config.output.dateFolderFormat
    }
  }

  if ($useDateSubfolder) { return Join-Path $WorkingDirectory (Get-Date -Format $dateFolderFormat) }
  return $WorkingDirectory
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

  $samplePath = Join-Path $PSScriptRoot ".\\config\\servers.sample.tsv"
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

function Get-ConnectionString {
  param(
    [string]$Server,
    [string]$Database,
    [string]$Username,
    [string]$Password,
    [int]$Port,
    [int]$TimeoutSeconds,
    [string]$ConnectionProvider,
    [string]$TdsVersion,
    [string]$OdbcDriver
  )

  if ($ConnectionProvider -eq "Odbc") {
    $authPart = if ([string]::IsNullOrWhiteSpace($Username)) { "Trusted_Connection=Yes;" } else { "UID=$Username;PWD=$Password;" }
    $tdsPart = if ([string]::IsNullOrWhiteSpace($TdsVersion)) { "" } else { "TDS_Version=$TdsVersion;" }
    return "Driver={$OdbcDriver};Server=$Server;Port=$Port;Database=$Database;$authPart${tdsPart}Encrypt=No;TrustServerCertificate=Yes;Connection Timeout=$TimeoutSeconds;"
  }

  if ([string]::IsNullOrWhiteSpace($Username)) {
    return "Server=$Server,$Port;Database=$Database;Integrated Security=True;TrustServerCertificate=True;Connection Timeout=$TimeoutSeconds;"
  }

  return "Server=$Server,$Port;Database=$Database;User ID=$Username;Password=$Password;TrustServerCertificate=True;Connection Timeout=$TimeoutSeconds;"
}

function New-DbConnection {
  param([string]$ConnectionString, [string]$ConnectionProvider)
  if ($ConnectionProvider -eq "Odbc") { return (New-Object System.Data.Odbc.OdbcConnection($ConnectionString)) }
  return (New-Object System.Data.SqlClient.SqlConnection($ConnectionString))
}

function Invoke-DataTable {
  param([System.Data.IDbConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  $table = New-Object System.Data.DataTable
  $reader = $cmd.ExecuteReader()
  try { $table.Load($reader) } finally { $reader.Close() }
  # Return as a single object; otherwise PowerShell enumerates DataTable rows.
  return ,$table
}

function Invoke-ScalarSafe {
  param([System.Data.IDbConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  try {
    $v = $cmd.ExecuteScalar()
    if ($v -eq [System.DBNull]::Value) { return $null }
    return $v
  }
  catch {
    return $null
  }
}

function Test-IsTimeoutException {
  param([System.Exception]$Exception)
  if ($null -eq $Exception) { return $false }
  $msg = [string]$Exception.Message
  if ([string]::IsNullOrWhiteSpace($msg)) { return $false }
  return ($msg -match "超时|timeout")
}

function Convert-DataTableToRows {
  param([System.Data.DataTable]$Table)
  $rows = @()
  if ($null -eq $Table) { return $rows }
  foreach ($r in $Table.Rows) {
    $o = [ordered]@{}
    foreach ($c in $Table.Columns) {
      $v = $r[$c.ColumnName]
      $o[$c.ColumnName] = if ($v -eq [System.DBNull]::Value) { $null } else { $v }
    }
    $rows += [pscustomobject]$o
  }
  return $rows
}

function Get-StatusByRule {
  param([object]$Value, [hashtable]$Rule)
  if ($null -eq $Value) { return "Warning" }
  $numeric = 0.0
  if (-not [double]::TryParse($Value.ToString(), [ref]$numeric)) { return "Warning" }

  if ($Rule.ContainsKey("criticalBelow") -and $numeric -lt [double]$Rule["criticalBelow"]) { return "Critical" }
  if ($Rule.ContainsKey("warningBelow") -and $numeric -lt [double]$Rule["warningBelow"]) { return "Warning" }
  if ($Rule.ContainsKey("criticalAbove") -and $numeric -gt [double]$Rule["criticalAbove"]) { return "Critical" }
  if ($Rule.ContainsKey("warningAbove") -and $numeric -gt [double]$Rule["warningAbove"]) { return "Warning" }
  return "Good"
}

function Add-Metric {
  param(
    [System.Collections.Generic.List[object]]$Target,
    [string]$MetricKey,
    [string]$Category,
    [object]$Value,
    [string]$Unit,
    [hashtable]$Rule,
    [string]$Source
  )

  $status = Get-StatusByRule -Value $Value -Rule $Rule
  $Target.Add([ordered]@{
    metricKey = $MetricKey
    category = $Category
    value = $Value
    unit = $Unit
    threshold = $Rule
    status = $status
    sourceProbe = $Source
    collectedAt = (Get-Date).ToString("s")
  })
}

function Escape-Html {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Convert-RowsToHtmlTable {
  param([object[]]$Rows, [int]$MaxRows = 120)
  if ($null -eq $Rows -or $Rows.Count -eq 0) {
    $noRowsText = "No rows returned."
    if ($null -ne $script:I18n -and $script:I18n.ContainsKey("noRows")) {
      $noRowsText = [string]$script:I18n["noRows"]
    }
    return "<p>$(Escape-Html -Text $noRowsText)</p>"
  }

  $display = @($Rows | Select-Object -First $MaxRows)
  $cols = @($display[0].PSObject.Properties.Name)
  $head = "<tr>" + (($cols | ForEach-Object { "<th>$(Escape-Html -Text $_)</th>" }) -join "") + "</tr>"
  $body = $display | ForEach-Object {
    $cells = foreach ($c in $cols) {
      $v = $_.$c
      "<td>$(Escape-Html -Text ([string]$v))</td>"
    }
    "<tr>$($cells -join '')</tr>"
  }
  return "<table><thead>$head</thead><tbody>$($body -join '')</tbody></table>"
}

$enableJsonStatusLocalization = Resolve-LocalizeJsonStatus -FlagFromParam $LocalizeJsonStatus.IsPresent -FlagProvided $PSBoundParameters.ContainsKey("LocalizeJsonStatus") -SkillConfigPath $SkillConfigPath
$resolvedReportLanguage = "en-US"
$script:I18n = Get-I18n -Language $resolvedReportLanguage

$workingDirectory = Resolve-WorkingDirectory -WorkspaceDir $WorkspaceDir -RawOutputDir $OutputDir -ServerConfigDir $ServerConfigDir -WorkspaceDirExplicitlySet $PSBoundParameters.ContainsKey("WorkspaceDir") -OutputDirExplicitlySet $PSBoundParameters.ContainsKey("OutputDir") -ServerConfigDirExplicitlySet $PSBoundParameters.ContainsKey("ServerConfigDir") -SkillConfigPath $SkillConfigPath -NoFolderPrompt:$NoFolderPrompt
if (-not (Test-Path $workingDirectory)) {
  New-Item -ItemType Directory -Force -Path $workingDirectory | Out-Null
}
Write-Output "Selected working directory: $workingDirectory"

$shouldLoadServerConfig = -not (
  $PSBoundParameters.ContainsKey("Server") -or
  $PSBoundParameters.ContainsKey("Database") -or
  $PSBoundParameters.ContainsKey("Port") -or
  $PSBoundParameters.ContainsKey("Username") -or
  $PSBoundParameters.ContainsKey("Password")
)

if (
  $PSBoundParameters.ContainsKey("ServerConfigDir") -or
  $PSBoundParameters.ContainsKey("ServerConfigFile") -or
  $PSBoundParameters.ContainsKey("ServerConfigKey")
) {
  $shouldLoadServerConfig = $true
}

$serverConfigFullPath = ""
if ($shouldLoadServerConfig) {
  $serverConfigFullPath = Ensure-ServerConfigFile -WorkingDirectory $workingDirectory -ServerConfigFile $ServerConfigFile
  Write-Output "Server config file: $serverConfigFullPath"
}

$resolvedOutputDir = Resolve-OutputDirectory -WorkingDirectory $workingDirectory -SkillConfigPath $SkillConfigPath
if (-not (Test-Path $resolvedOutputDir)) {
  New-Item -ItemType Directory -Force -Path $resolvedOutputDir | Out-Null
}
Write-Output "Resolved output directory: $resolvedOutputDir"
$serverConfig = $null
if ($shouldLoadServerConfig -and -not [string]::IsNullOrWhiteSpace($serverConfigFullPath)) {
  $serverConfig = Get-ServerConfigFromPath -FullPath $serverConfigFullPath
}
$singleConfig = $null
if ($null -ne $serverConfig -and -not [string]::IsNullOrWhiteSpace($ServerConfigKey)) {
  if ($serverConfig.PSObject.Properties.Name -contains $ServerConfigKey) {
    $singleConfig = $serverConfig.$ServerConfigKey
  }
  elseif ($null -ne $serverConfig.instances -and $serverConfig.instances.Count -gt 0) {
    if ($ServerConfigKey -eq "defaultSingle") {
      $singleConfig = $serverConfig.instances[0]
    }
    else {
      $singleConfig = @($serverConfig.instances | Where-Object { $_.server -eq $ServerConfigKey -or $_.name -eq $ServerConfigKey } | Select-Object -First 1)
    }
  }
}
elseif ($null -ne $serverConfig -and $null -ne $serverConfig.defaultSingle) {
  $singleConfig = $serverConfig.defaultSingle
}

if ($null -ne $serverConfig -and $null -ne $serverConfig.globalSettings) {
  $g = $serverConfig.globalSettings
  if (-not $PSBoundParameters.ContainsKey("ReportLanguage") -and $g.Contains("language") -and -not [string]::IsNullOrWhiteSpace([string]$g["language"])) {
    $ReportLanguage = [string]$g["language"]
  }
  if (-not $PSBoundParameters.ContainsKey("ReportFormats") -and $g.Contains("reportformats") -and -not [string]::IsNullOrWhiteSpace([string]$g["reportformats"])) {
    $ReportFormats = [string]$g["reportformats"]
  }
}

if ($null -ne $singleConfig) {
  if (-not $PSBoundParameters.ContainsKey("Server") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.server)) { $Server = [string]$singleConfig.server }
  if (-not $PSBoundParameters.ContainsKey("Database") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.database)) { $Database = [string]$singleConfig.database }
  if (-not $PSBoundParameters.ContainsKey("Port") -and $null -ne $singleConfig.port) { $Port = [int]$singleConfig.port }
  if (-not $PSBoundParameters.ContainsKey("Username") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.username)) { $Username = [string]$singleConfig.username }
  if (-not $PSBoundParameters.ContainsKey("Password") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.password)) { $Password = [string]$singleConfig.password }
  if (-not $PSBoundParameters.ContainsKey("ConnectionProvider") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.connectionProvider)) { $ConnectionProvider = [string]$singleConfig.connectionProvider }
  if (-not $PSBoundParameters.ContainsKey("TdsVersion") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.tdsVersion)) { $TdsVersion = [string]$singleConfig.tdsVersion }
  if (-not $PSBoundParameters.ContainsKey("OdbcDriver") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.odbcDriver)) { $OdbcDriver = [string]$singleConfig.odbcDriver }
  if (-not $PSBoundParameters.ContainsKey("ReportFormats") -and -not [string]::IsNullOrWhiteSpace([string]$singleConfig.reportFormats)) { $ReportFormats = [string]$singleConfig.reportFormats }
  if (-not $PSBoundParameters.ContainsKey("TimeoutSeconds") -and $null -ne $singleConfig.timeoutSeconds) { $TimeoutSeconds = [int]$singleConfig.timeoutSeconds }
}

$resolvedReportLanguage = Resolve-ReportLanguage -InputLanguage $ReportLanguage -SkillConfigPath $SkillConfigPath
$script:I18n = Get-I18n -Language $resolvedReportLanguage

$thresholdFullPath = Join-Path $PSScriptRoot $ThresholdPath
$scoreConfigFullPath = Join-Path $PSScriptRoot $ScoreConfigPath
$auditChecksPath = Join-Path $PSScriptRoot $AuditChecksRoot

if (-not (Test-Path $thresholdFullPath)) { throw "Threshold config not found: $thresholdFullPath" }
if (-not (Test-Path $scoreConfigFullPath)) { throw "Score config not found: $scoreConfigFullPath" }
if (-not (Test-Path $auditChecksPath)) { throw "Audit checks path not found: $auditChecksPath" }

$thresholdJson = Get-Content -Raw -Path $thresholdFullPath | ConvertFrom-Json
$scoreConfig = Get-Content -Raw -Path $scoreConfigFullPath | ConvertFrom-Json

$ruleMap = @{}
$thresholdJson.rules.PSObject.Properties | ForEach-Object {
  $name = $_.Name
  $ruleMap[$name] = @{}
  $_.Value.PSObject.Properties | ForEach-Object { $ruleMap[$name][$_.Name] = $_.Value }
}

$connectionString = Get-ConnectionString -Server $Server -Database $Database -Username $Username -Password $Password -Port $Port -TimeoutSeconds $TimeoutSeconds -ConnectionProvider $ConnectionProvider -TdsVersion $TdsVersion -OdbcDriver $OdbcDriver
$conn = $null
$probeErrors = New-Object System.Collections.Generic.List[object]

$auditRegistry = @(
  @{ key = "instanceProfile"; file = "01_instance_profile.sql"; title = "Instance Profile" },
  @{ key = "databaseInventory"; file = "02_database_inventory.sql"; title = "Database Inventory" },
  @{ key = "databaseSize"; file = "03_database_size.sql"; title = "Database Size" },
  @{ key = "activeSessions"; file = "04_active_sessions.sql"; title = "Active Sessions" },
  @{ key = "memoryUsage"; file = "05_memory_usage.sql"; title = "Memory Usage" },
  @{ key = "cpuUsage"; file = "06_cpu_usage.sql"; title = "CPU Usage" },
  @{ key = "diskCapacity"; file = "07_disk_capacity.sql"; title = "Disk Capacity" },
  @{ key = "backupStatus"; file = "08_backup_status.sql"; title = "Backup Status" },
  @{ key = "failedJobs24h"; file = "09_failed_jobs_24h.sql"; title = "Failed Jobs Last 24h" },
  @{ key = "fragmentedIndexes"; file = "10_fragmented_indexes.sql"; title = "Fragmented Indexes" },
  @{ key = "missingIndexesTop30"; file = "11_missing_indexes_top30.sql"; title = "Missing Index Suggestions Top 30" },
  @{ key = "orphanedUsers"; file = "12_orphaned_users.sql"; title = "Orphaned Users" },
  @{ key = "longRunningQueries5m"; file = "13_long_running_queries_5m.sql"; title = "Long Running Queries Over 5 Minutes" },
  @{ key = "alwaysOnStatus"; file = "14_alwayson_status.sql"; title = "AlwaysOn Status" },
  @{ key = "serverLogins"; file = "15_server_logins.sql"; title = "Server Logins" }
)

$mvpMetrics = New-Object System.Collections.Generic.List[object]
$auditSections = [ordered]@{}

try {
  $conn = New-DbConnection -ConnectionString $connectionString -ConnectionProvider $ConnectionProvider
  $conn.Open()

  $ple = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT MAX(cntr_value)
FROM sys.dm_os_performance_counters WITH (NOLOCK)
WHERE object_name LIKE '%Buffer Manager%'
  AND counter_name = 'Page life expectancy';
"@
  Add-Metric -Target $mvpMetrics -MetricKey "ple_seconds" -Category "resource_performance" -Value $ple -Unit "seconds" -Rule $ruleMap["ple_seconds"] -Source "dm_os_performance_counters"

  $bchr = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
;WITH c AS (
  SELECT
    SUM(CASE WHEN counter_name = 'Buffer cache hit ratio' THEN cntr_value ELSE 0 END) AS ratio_value,
    SUM(CASE WHEN counter_name = 'Buffer cache hit ratio base' THEN cntr_value ELSE 0 END) AS ratio_base
  FROM sys.dm_os_performance_counters WITH (NOLOCK)
  WHERE object_name LIKE '%Buffer Manager%'
    AND counter_name IN ('Buffer cache hit ratio', 'Buffer cache hit ratio base')
)
SELECT CASE WHEN ratio_base > 0 THEN CAST((ratio_value * 100.0) / ratio_base AS DECIMAL(10,2)) ELSE NULL END
FROM c;
"@
  Add-Metric -Target $mvpMetrics -MetricKey "buffer_cache_hit_ratio" -Category "resource_performance" -Value $bchr -Unit "percent" -Rule $ruleMap["buffer_cache_hit_ratio"] -Source "dm_os_performance_counters"

  $blocking = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql "SELECT COUNT(1) FROM sys.dm_exec_requests WITH (NOLOCK) WHERE blocking_session_id <> 0;"
  Add-Metric -Target $mvpMetrics -MetricKey "blocking_sessions" -Category "concurrency" -Value $blocking -Unit "count" -Rule $ruleMap["blocking_sessions"] -Source "dm_exec_requests"

  $longest = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql "SELECT ISNULL(MAX(total_elapsed_time / 1000.0), 0) FROM sys.dm_exec_requests WITH (NOLOCK) WHERE session_id <> @@SPID AND status <> 'background';"
  Add-Metric -Target $mvpMetrics -MetricKey "longest_running_query_seconds" -Category "concurrency" -Value $longest -Unit "seconds" -Rule $ruleMap["longest_running_query_seconds"] -Source "dm_exec_requests"

  $deadlocks = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
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
WHERE d.e.value('@timestamp', 'datetime2') >= DATEADD(HOUR, -24, SYSUTCDATETIME());
"@
  Add-Metric -Target $mvpMetrics -MetricKey "deadlocks_24h" -Category "concurrency" -Value $deadlocks -Unit "count" -Rule $ruleMap["deadlocks_24h"] -Source "system_health_xevent"

  $failedJobs = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
WHERE h.step_id = 0
  AND h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE());
"@
  Add-Metric -Target $mvpMetrics -MetricKey "failed_jobs_24h" -Category "automation_jobs" -Value $failedJobs -Unit "count" -Rule $ruleMap["failed_jobs_24h"] -Source "msdb.sysjobhistory"

  $disabledJobs = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql "SELECT COUNT(1) FROM msdb.dbo.sysjobs WITH (NOLOCK) WHERE enabled = 0;"
  Add-Metric -Target $mvpMetrics -MetricKey "disabled_jobs_count" -Category "automation_jobs" -Value $disabledJobs -Unit "count" -Rule $ruleMap["disabled_jobs_count"] -Source "msdb.sysjobs"

  $lastFullBackup = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql "SELECT MAX(backup_finish_date) FROM msdb.dbo.backupset WITH (NOLOCK) WHERE type = 'D';"
  $backupAgeHours = $null
  if ($null -ne $lastFullBackup) {
    $backupAgeHours = [Math]::Round(((Get-Date) - [datetime]$lastFullBackup).TotalHours, 2)
  }
  Add-Metric -Target $mvpMetrics -MetricKey "last_full_backup_age_hours" -Category "data_safety" -Value $backupAgeHours -Unit "hours" -Rule $ruleMap["last_full_backup_age_hours"] -Source "msdb.backupset"

  $unprotected = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
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
  AND (b.last_full_backup IS NULL OR b.last_full_backup < @cutoff);
"@
  Add-Metric -Target $mvpMetrics -MetricKey "databases_without_recent_full_backup" -Category "data_safety" -Value $unprotected -Unit "count" -Rule $ruleMap["databases_without_recent_full_backup"] -Source "sys.databases+msdb.backupset"

  $avgLogUsed = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql "SELECT AVG(used_log_space_in_percent * 1.0) FROM sys.dm_db_log_space_usage;"
  Add-Metric -Target $mvpMetrics -MetricKey "log_used_percent" -Category "io_disk" -Value $avgLogUsed -Unit "percent" -Rule $ruleMap["log_used_percent"] -Source "dm_db_log_space_usage"

  $diskUsed = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT MAX(
  CASE WHEN vs.total_bytes > 0
    THEN (1.0 - (vs.available_bytes * 1.0 / vs.total_bytes)) * 100
    ELSE NULL END
)
FROM sys.master_files mf WITH (NOLOCK)
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs
WHERE mf.type_desc = 'ROWS';
"@
  Add-Metric -Target $mvpMetrics -MetricKey "data_disk_used_percent" -Category "io_disk" -Value $diskUsed -Unit "percent" -Rule $ruleMap["data_disk_used_percent"] -Source "dm_os_volume_stats"

  $agUnhealthy = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM sys.dm_hadr_availability_replica_states WITH (NOLOCK)
WHERE is_local = 1
  AND role_desc = 'PRIMARY'
  AND synchronization_health_desc <> 'HEALTHY';
"@
  Add-Metric -Target $mvpMetrics -MetricKey "ag_unhealthy_replicas" -Category "high_availability" -Value $agUnhealthy -Unit "count" -Rule $ruleMap["ag_unhealthy_replicas"] -Source "dm_hadr_availability_replica_states"

  $sysadminCount = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM sys.server_role_members srm WITH (NOLOCK)
INNER JOIN sys.server_principals r WITH (NOLOCK) ON r.principal_id = srm.role_principal_id
WHERE r.name = 'sysadmin';
"@
  Add-Metric -Target $mvpMetrics -MetricKey "sysadmin_login_count" -Category "security_audit" -Value $sysadminCount -Unit "count" -Rule $ruleMap["sysadmin_login_count"] -Source "sys.server_role_members"

  $weakPolicy = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM sys.sql_logins WITH (NOLOCK)
WHERE is_disabled = 0
  AND (is_policy_checked = 0 OR is_expiration_checked = 0);
"@
  Add-Metric -Target $mvpMetrics -MetricKey "weak_policy_logins_count" -Category "security_audit" -Value $weakPolicy -Unit "count" -Rule $ruleMap["weak_policy_logins_count"] -Source "sys.sql_logins"

  $newLogins = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
SELECT COUNT(1)
FROM sys.server_principals WITH (NOLOCK)
WHERE type IN ('S', 'U', 'G')
  AND create_date >= DATEADD(DAY, -90, GETDATE())
  AND name NOT LIKE '##%';
"@
  Add-Metric -Target $mvpMetrics -MetricKey "new_logins_90d_count" -Category "security_audit" -Value $newLogins -Unit "count" -Rule $ruleMap["new_logins_90d_count"] -Source "sys.server_principals"

  $newDbUsers = Invoke-ScalarSafe -Connection $conn -TimeoutSeconds $TimeoutSeconds -Sql @"
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

SELECT COUNT(1) FROM #recent_users;
"@
  Add-Metric -Target $mvpMetrics -MetricKey "new_db_users_90d_count" -Category "security_audit" -Value $newDbUsers -Unit "count" -Rule $ruleMap["new_db_users_90d_count"] -Source "sys.database_principals"

  foreach ($c in $auditRegistry) {
    $checkFile = Join-Path $auditChecksPath $c.file
    try {
      $sql = Get-Content -Raw -Path $checkFile
      $checkTimeoutSeconds = $TimeoutSeconds
      if ($c.key -eq "fragmentedIndexes") {
        # fragmented index scan can be expensive on large instances; keep a higher floor timeout.
        $checkTimeoutSeconds = [Math]::Max($TimeoutSeconds, 300)
      }
      $table = Invoke-DataTable -Connection $conn -Sql $sql -TimeoutSeconds $checkTimeoutSeconds
      # Force array semantics; a single row can otherwise be unwrapped and lose .Count.
      $rows = @(Convert-DataTableToRows -Table $table)
      $auditSections[$c.key] = [ordered]@{
        title = $c.title
        file = $c.file
        rowCount = $rows.Count
        rows = $rows
      }
    }
    catch {
      $isFragSoftFail = ($c.key -eq "fragmentedIndexes")
      if (-not $isFragSoftFail) {
        $probeErrors.Add([ordered]@{ scope = "audit"; check = $c.key; file = $c.file; message = $_.Exception.Message; at = (Get-Date).ToString("s") })
      }
      $auditSections[$c.key] = [ordered]@{
        title = $c.title
        file = $c.file
        rowCount = 0
        rows = @()
        warning = $(if ($isFragSoftFail) { "fragmentedIndexes probe failed and was skipped for this run: $($_.Exception.Message)" } else { $null })
      }
    }
  }

  $mvpCritical = @($mvpMetrics | Where-Object { $_.status -eq "Critical" }).Count
  $mvpWarning = @($mvpMetrics | Where-Object { $_.status -eq "Warning" }).Count
  $mvpScore = [Math]::Max(0, 100 - ($mvpCritical * 20 + $mvpWarning * 8))

  $weights = $scoreConfig.weights
  $diskThresholds = $scoreConfig.diskThresholdPercent

  $diskRows = @($auditSections.diskCapacity.rows)
  $diskCriticalCount = @($diskRows | Where-Object { $null -ne $_.free_percent -and [double]$_.free_percent -lt [double]$diskThresholds.criticalBelow }).Count
  $diskWarningCount = @($diskRows | Where-Object { $null -ne $_.free_percent -and [double]$_.free_percent -lt [double]$diskThresholds.warningBelow -and [double]$_.free_percent -ge [double]$diskThresholds.criticalBelow }).Count

  $backupRows = @($auditSections.backupStatus.rows)
  $fullBackupCriticalCount = @($backupRows | Where-Object { $_.full_backup_status -eq "critical" }).Count
  $fullBackupWarningCount = @($backupRows | Where-Object { $_.full_backup_status -eq "warning" }).Count

  $failedJobsCount = @($auditSections.failedJobs24h.rows).Count
  $fragmentedIndexCount = @($auditSections.fragmentedIndexes.rows).Count
  $missingIndexCount = @($auditSections.missingIndexesTop30.rows).Count
  $orphanedUsersCount = @($auditSections.orphanedUsers.rows).Count
  $longRunningCount = @($auditSections.longRunningQueries5m.rows).Count
  $alwaysOnIssueCount = @($auditSections.alwaysOnStatus.rows | Where-Object { $_.role_desc -ne "NOT_ENABLED" -and $_.synchronization_health_desc -ne $null -and $_.synchronization_health_desc -ne "HEALTHY" }).Count

  $auditPenalty = 0
  $auditPenalty += $diskCriticalCount * [int]$weights.diskCriticalPerVolume
  $auditPenalty += $diskWarningCount * [int]$weights.diskWarningPerVolume
  $auditPenalty += $fullBackupCriticalCount * [int]$weights.fullBackupCriticalPerDatabase
  $auditPenalty += $fullBackupWarningCount * [int]$weights.fullBackupWarningPerDatabase
  $auditPenalty += [Math]::Min($fragmentedIndexCount, [int]$weights.highFragmentedIndexMax) * [int]$weights.highFragmentedIndexPerItem
  $auditPenalty += [Math]::Min($missingIndexCount, [int]$weights.missingIndexMax) * [int]$weights.missingIndexPerItem
  if ($failedJobsCount -gt 0) { $auditPenalty += [int]$weights.failedJobsPenalty }
  if ($orphanedUsersCount -gt 0) { $auditPenalty += [int]$weights.orphanedUsersPenalty }
  if ($alwaysOnIssueCount -gt 0) { $auditPenalty += [int]$weights.alwaysOnIssuePenalty }
  if ($longRunningCount -gt 0) { $auditPenalty += [int]$weights.longRunningQueryPenalty }

  $auditScore = [Math]::Max(0, 100 - $auditPenalty)
  $combinedScore = [Math]::Round(($mvpScore * 0.45 + $auditScore * 0.55), 0)
  if ($combinedScore -lt 0) { $combinedScore = 0 }

  $overallStatus = "Good"
  if ($combinedScore -lt 70 -or $mvpCritical -gt 0) {
    $overallStatus = "Critical"
  }
  elseif ($combinedScore -lt 85 -or $mvpWarning -gt 0) {
    $overallStatus = "Warning"
  }

  $result = [ordered]@{
    project = "SQL-Server-Complete-Inspection"
    collectedAt = (Get-Date).ToString("s")
    server = $Server
    database = $Database
    summary = [ordered]@{
      overallStatus = $overallStatus
      healthScore = $combinedScore
      mvpScore = $mvpScore
      auditScore = $auditScore
      mvpCritical = $mvpCritical
      mvpWarning = $mvpWarning
      probeErrors = $probeErrors.Count
    }
    metrics = $mvpMetrics
    fullAudit = [ordered]@{
      penalty = $auditPenalty
      indicators = [ordered]@{
        diskCriticalVolumes = $diskCriticalCount
        diskWarningVolumes = $diskWarningCount
        fullBackupCriticalDatabases = $fullBackupCriticalCount
        fullBackupWarningDatabases = $fullBackupWarningCount
        failedJobs24h = $failedJobsCount
        highFragmentedIndexes = $fragmentedIndexCount
        missingIndexSuggestions = $missingIndexCount
        orphanedUsers = $orphanedUsersCount
        longRunningQueriesOver5m = $longRunningCount
        alwaysOnIssues = $alwaysOnIssueCount
      }
      sections = $auditSections
    }
    probeErrors = $probeErrors
  }

  $basePath = Join-Path $resolvedOutputDir ("complete-inspection-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
  $jsonPath = "$basePath.json"
  $htmlPath = "$basePath.html"

  $formats = Normalize-ReportFormats -Formats $ReportFormats

  if ($formats -contains "json") {
    $jsonResult = Get-JsonResultWithLocalizedStatus -Result $result -Language $resolvedReportLanguage -EnableLocalization $enableJsonStatusLocalization
    $jsonResult | ConvertTo-Json -Depth 9 | Set-Content -Encoding UTF8 -Path $jsonPath
    Write-Output "Generated: $jsonPath"
  }

  if ($formats -contains "html") {
    $metricRows = $mvpMetrics | ForEach-Object {
      "<tr><td>$(Escape-Html -Text $_.metricKey)</td><td>$(Escape-Html -Text $_.category)</td><td>$(Escape-Html -Text ([string]$_.value))</td><td>$(Escape-Html -Text $_.unit)</td><td>$(Escape-Html -Text (Localize-Status -Status $_.status -Language $resolvedReportLanguage))</td></tr>"
    }

    $auditSectionHtml = foreach ($entry in $auditSections.GetEnumerator()) {
      $title = [string]$entry.Value.title
      $file = [string]$entry.Value.file
      $rows = @($entry.Value.rows)
      "<div class='panel'><h2>$(Escape-Html -Text $title)</h2><div class='meta'>$(Escape-Html -Text ([string]$script:I18n.source)): $(Escape-Html -Text $file) | $(Escape-Html -Text ([string]$script:I18n.rows)): $($rows.Count)</div>$(Convert-RowsToHtmlTable -Rows $rows)</div>"
    }

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>$(Escape-Html -Text ([string]$script:I18n.reportTitle))</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f5f7fb; color: #1f2937; }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 16px; }
    .card { background: #fff; border: 1px solid #d9e2ec; border-radius: 10px; padding: 10px; }
    .k { color: #5b6472; font-size: 12px; }
    .v { font-size: 24px; font-weight: 700; }
    .panel { background: #fff; border: 1px solid #d9e2ec; border-radius: 10px; padding: 12px; margin-bottom: 14px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #d9e2ec; padding: 7px 9px; text-align: left; vertical-align: top; }
    th { background: #edf3fa; }
    .meta { color: #5b6472; margin: 6px 0 10px 0; }
  </style>
</head>
<body>
  <h1>$(Escape-Html -Text ([string]$script:I18n.reportTitle))</h1>
  <div class='meta'>$(Escape-Html -Text ([string]$script:I18n.server)): $(Escape-Html -Text $Server) | $(Escape-Html -Text ([string]$script:I18n.collectedAt)): $(Get-Date -Format "s") | $(Escape-Html -Text ([string]$script:I18n.overall)): $(Escape-Html -Text (Localize-Status -Status $overallStatus -Language $resolvedReportLanguage))</div>

  <div class='cards'>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.combinedScore))</div><div class='v'>$combinedScore</div></div>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.mvpScore))</div><div class='v'>$mvpScore</div></div>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.fullAuditScore))</div><div class='v'>$auditScore</div></div>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.mvpCritical))</div><div class='v'>$mvpCritical</div></div>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.mvpWarning))</div><div class='v'>$mvpWarning</div></div>
    <div class='card'><div class='k'>$(Escape-Html -Text ([string]$script:I18n.probeErrors))</div><div class='v'>$($probeErrors.Count)</div></div>
  </div>

  <div class='panel'>
    <h2>$(Escape-Html -Text ([string]$script:I18n.mvpMetrics))</h2>
    <table>
      <thead><tr><th>$(Escape-Html -Text ([string]$script:I18n.metric))</th><th>$(Escape-Html -Text ([string]$script:I18n.category))</th><th>$(Escape-Html -Text ([string]$script:I18n.value))</th><th>$(Escape-Html -Text ([string]$script:I18n.unit))</th><th>$(Escape-Html -Text ([string]$script:I18n.status))</th></tr></thead>
      <tbody>
        $($metricRows -join "`n")
      </tbody>
    </table>
  </div>

  $($auditSectionHtml -join "`n")
</body>
</html>
"@

    Set-Content -Encoding UTF8 -Path $htmlPath -Value $html
    Write-Output "Generated: $htmlPath"
  }
}
finally {
  if ($null -ne $conn) {
    try {
      if ($conn.State -ne [System.Data.ConnectionState]::Closed) { $conn.Close() }
    }
    catch {}
    try { $conn.Dispose() } catch {}
  }
}
