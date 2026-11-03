#requires -Version 5.1
<#
.SYNOPSIS
    Shows which SQL Server and database an iNovah instance's site is using, and verifies it.

.DESCRIPTION
    1. Locates the client's application folder(s) by matching the client name against
       IIS sites/applications (if IIS is installed) and folder names under the search root(s).
    2. Reads ONLY web.config and user.config in those folders. Any other file (backups such as
       web.config.bak, 'web - Copy.config', old .json/.config files) is ignored.
    3. Extracts connection strings from <connectionStrings>, <appSettings>, configSource files,
       encrypted sections (decrypted when run as admin on the app server), and JSON settings.
       If a user.config sits in the same folder as a web.config, its values supersede the
       web.config values with the same name; the output shows which file each value came from.
    4. Connects to each SQL Server found and checks that the database exists, its state,
       and whether it can actually be opened.

.PARAMETER Instance
    The iNovah instance to look up, e.g. FAY71T or FAY71D. (-ClientName also works.)
    You can also pass a URL here instead; it's detected automatically.

.PARAMETER Url
    The site URL, e.g. sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/ (http:// or https:// optional).
    The instance is taken from the URL path (iNovah2-FAY71T -> FAY71T) and the result is checked
    against each site's WebApplicationURL. If no folder is named for that instance, every
    web.config/user.config under the search roots is checked for a matching WebApplicationURL.

.PARAMETER SearchRoot
    One or more folders to search for client folders. Defaults to C:\ and E:\ (missing drives are skipped).

.PARAMETER Depth
    How many folder levels below each search root to look for client folders. Default 3.

.PARAMETER UrlKey
    The appSettings key that holds the site's URL. Default 'WebApplicationURL'.
    Like the connection strings, a value in user.config supersedes web.config.

.PARAMETER SkipIIS
    Don't look up IIS sites/applications.

.PARAMETER SkipVerify
    Only read the config files; don't connect to SQL Server.

.PARAMETER Credential
    SQL login to use for verification instead of the credentials in the config file.

.PARAMETER Detailed
    Show every connection string found in every config file (the old full output)
    instead of the one-block-per-database summary.

.PARAMETER PrimaryKey
    Name of the connection string iNovah uses for its own database. Default 'Data.ConnectionString'.
    If no entry has that name, the entry whose Initial Catalog matches the instance is used.

.PARAMETER All
    Also show other connection strings (integrations such as PeopleSoft/CCM, older settings).

.PARAMETER ShowPassword
    Show passwords from connection strings instead of masking them.

.EXAMPLE
    .\Get-ClientDatabase.ps1
    Prompts for an iNovah instance or URL, then shows the results. Offers to look up another.

.EXAMPLE
    .\Get-ClientDatabase.ps1 FAY71T

.EXAMPLE
    .\Get-ClientDatabase.ps1 sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/

.EXAMPLE
    .\Get-ClientDatabase.ps1 -Instance FAY71T | Select-Object 'SQL Instance (Data Source)', 'Database (Initial Catalog)', WebApplicationURL, Status

.EXAMPLE
    .\Get-ClientDatabase.ps1 FAY71D -Credential (Get-Credential) | Export-Csv .\FAY71D-db.csv -NoTypeInformation
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [Alias('ClientName')]
    [string]$Instance,

    # Site URL, e.g. sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/
    [string]$Url,

    [string[]]$SearchRoot = @('C:\', 'E:\'),

    # Top-level folders skipped on each search root (keeps C:\ searches fast)
    [string[]]$ExcludeFolders = @('Windows', '$Recycle.Bin', 'System Volume Information',
                                  'Recovery', 'PerfLogs', 'Users', 'ProgramData'),

    [int]$Depth = 3,

    # appSettings key holding the site's URL
    [string]$UrlKey = 'WebApplicationURL',

    [switch]$SkipIIS,
    [switch]$SkipVerify,
    [pscredential]$Credential,
    [int]$TimeoutSeconds = 5,
    [switch]$ShowPassword,

    # Show every connection string found, per config file, instead of the summary
    [switch]$Detailed,

    # The connection string iNovah itself uses. Others (integrations, old settings) are hidden unless -All.
    [string]$PrimaryKey = 'Data.ConnectionString',

    # Show every database connection found, not just the iNovah one
    [switch]$All
)

#region ---------- Helpers: parsing ----------

function Test-LooksLikeConnString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)(data source|server|address)\s*=') -and
           ($Value -match '(?i)(initial catalog|database)\s*=')
}

function Get-CsValue {
    param($Builder, [string[]]$Keys)
    foreach ($k in $Keys) {
        if ($Builder.ContainsKey($k)) { return [string]$Builder[$k] }
    }
    return $null
}

function ConvertFrom-ConnString {
    param([string]$ConnString)

    # Entity Framework style: metadata=...;provider connection string="..."
    if ($ConnString -match '(?i)provider connection string\s*=\s*"([^"]+)"') {
        $ConnString = $Matches[1]
    }

    # Generic builder accepts any keyword, so newer/unknown options won't break parsing
    $b = New-Object System.Data.Common.DbConnectionStringBuilder
    $parsedOk = $true
    try { $b.set_ConnectionString($ConnString) } catch { $parsedOk = $false }

    if (-not $parsedOk) {
        # Tolerate hand-edited strings (e.g. Data Source='SBOXSQLPE\FAY71 with no closing quote)
        $b = New-Object System.Data.Common.DbConnectionStringBuilder
        foreach ($part in $ConnString -split ';') {
            $kv = $part -split '=', 2
            if ($kv.Count -eq 2 -and $kv[0].Trim()) {
                try { $b[$kv[0].Trim()] = $kv[1].Trim().Trim([char[]]@("'", '"')) } catch { }
            }
        }
    }

    $integrated = Get-CsValue $b @('Integrated Security', 'Trusted_Connection')

    [pscustomobject]@{
        Server     = (Get-CsValue $b @('Data Source', 'Server', 'Address', 'Addr', 'Network Address'))
        Database   = (Get-CsValue $b @('Initial Catalog', 'Database'))
        UserId     = (Get-CsValue $b @('User ID', 'UID', 'User'))
        Password   = (Get-CsValue $b @('Password', 'PWD'))
        Integrated = ($integrated -match '^(?i)(true|yes|sspi)$')
        Provider   = (Get-CsValue $b @('Provider'))
    }
}

function Split-DataSource {
    param([string]$DataSource)
    # 'tcp:SBOXSQLPE\FAY71,1433' -> Host SBOXSQLPE, Instance FAY71, Port 1433
    $result = [ordered]@{ Host = $null; Instance = $null; Port = $null }
    if (-not $DataSource) { return [pscustomobject]$result }
    $ds = $DataSource.Trim().Trim([char[]]@("'", '"')) -replace '^(?i)(tcp|np|lpc):', ''
    if ($ds -match '^(.*?),\s*(\d+)$') { $ds = $Matches[1]; $result.Port = $Matches[2] }
    $parts = $ds -split '\\', 2
    $result.Host = $parts[0]
    if ($parts.Count -gt 1) { $result.Instance = $parts[1] } else { $result.Instance = 'MSSQLSERVER (default)' }
    [pscustomobject]$result
}

function Get-ProtectedConnStrings {
    param([string]$Path)
    # Encrypted <connectionStrings> (aspnet_regiis -pe). Decryption only works on the
    # machine holding the key, usually requiring an elevated session.
    try {
        Add-Type -AssemblyName System.Configuration
        $map = New-Object System.Configuration.ExeConfigurationFileMap
        $map.ExeConfigFilename = $Path
        $cfg = [System.Configuration.ConfigurationManager]::OpenMappedExeConfiguration(
            $map, [System.Configuration.ConfigurationUserLevel]::None)

        foreach ($c in $cfg.ConnectionStrings.ConnectionStrings) {
            # Skip entries inherited from machine.config (e.g. LocalSqlServer)
            if ($c.ElementInformation.Source -and $c.ElementInformation.Source -ne $Path) { continue }
            [pscustomobject]@{
                Name    = $c.Name
                Raw     = $c.ConnectionString
                Section = 'connectionStrings (encrypted - decrypted OK)'
                File    = $Path
            }
        }
    }
    catch {
        Write-Warning "Encrypted connectionStrings in '$Path' could not be decrypted. Run as admin on the app server. ($($_.Exception.GetBaseException().Message))"
        [pscustomobject]@{
            Name    = '(encrypted)'
            Raw     = $null
            Section = 'connectionStrings (encrypted - could not decrypt)'
            File    = $Path
        }
    }
}

function Get-XmlConnStrings {
    param([string]$Path)

    $results = @()
    try {
        [xml]$xml = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not parse XML '$Path': $($_.Exception.Message)"
        return
    }

    # local-name() so it still works if the config root has an xmlns attribute
    foreach ($cs in $xml.SelectNodes('//*[local-name()="connectionStrings"]')) {

        $configSource = $cs.GetAttribute('configSource')
        if ($configSource) {
            # Only web.config / user.config are read, so the referenced file is not opened
            Write-Warning "'$Path' moves its connectionStrings to '$configSource', which is not read (only web.config/user.config are used)."
            continue
        }

        if ($cs.GetAttribute('configProtectionProvider')) {
            $results += Get-ProtectedConnStrings -Path $Path
            continue
        }

        foreach ($add in $cs.SelectNodes('*[local-name()="add"]')) {
            $results += [pscustomobject]@{
                Name    = $add.GetAttribute('name')
                Raw     = $add.GetAttribute('connectionString')
                Section = 'connectionStrings'
                File    = $Path
            }
        }
    }

    # Some apps keep the connection string in appSettings instead
    foreach ($add in $xml.SelectNodes('//*[local-name()="appSettings"]/*[local-name()="add"]')) {
        $value = $add.GetAttribute('value')
        if (Test-LooksLikeConnString $value) {
            $results += [pscustomobject]@{
                Name    = $add.GetAttribute('key')
                Raw     = $value
                Section = 'appSettings'
                File    = $Path
            }
        }
    }

    $results
}

function Get-AppSettingValue {
    param([string]$Path, [string]$Key)
    if (-not $Path) { return $null }
    try { [xml]$xml = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop } catch { return $null }
    foreach ($add in $xml.SelectNodes('//*[local-name()="appSettings"]/*[local-name()="add"]')) {
        if ($add.GetAttribute('key') -eq $Key) { return $add.GetAttribute('value') }
    }
    return $null
}

function Search-JsonNode {
    param($Node, [string]$Prefix, [string]$Path)

    if ($Node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $Node.PSObject.Properties) {
            $key = if ($Prefix) { "${Prefix}:$($p.Name)" } else { $p.Name }
            $v = $p.Value
            if ($v -is [string]) {
                $inCsSection = $Prefix -match '(?i)(^|:)ConnectionStrings$'
                if (($inCsSection -and $v) -or (Test-LooksLikeConnString $v)) {
                    [pscustomobject]@{ Name = $key; Raw = $v; Section = 'json'; File = $Path }
                }
            }
            elseif ($v -is [System.Management.Automation.PSCustomObject] -or $v -is [array]) {
                Search-JsonNode -Node $v -Prefix $key -Path $Path
            }
        }
    }
    elseif ($Node -is [array]) {
        $i = 0
        foreach ($n in $Node) {
            Search-JsonNode -Node $n -Prefix "${Prefix}:$i" -Path $Path
            $i++
        }
    }
}

function Get-JsonConnStrings {
    param([string]$Path)

    $raw = Get-Content -LiteralPath $Path -Raw
    # Strip full-line // comments and /* */ blocks (appsettings allows them, ConvertFrom-Json in 5.1 doesn't)
    $clean = $raw -replace '(?m)^\s*//.*$', '' -replace '(?s)/\*.*?\*/', ''

    $obj = $null
    try { $obj = $clean | ConvertFrom-Json -ErrorAction Stop } catch { }

    if ($obj) {
        Search-JsonNode -Node $obj -Prefix '' -Path $Path
    }
    else {
        # Fallback: regex scan for "key": "value" pairs that look like connection strings
        Write-Verbose "JSON parse failed for '$Path', using regex fallback."
        foreach ($m in [regex]::Matches($raw, '"([^"]+)"\s*:\s*"((?:[^"\\]|\\.)*)"')) {
            $value = $m.Groups[2].Value -replace '\\\\', '\'
            if (Test-LooksLikeConnString $value) {
                [pscustomobject]@{ Name = $m.Groups[1].Value; Raw = $value; Section = 'json (regex)'; File = $Path }
            }
        }
    }
}

#endregion

#region ---------- Helper: verification ----------

$verifyCache = @{}

function Test-SqlDatabase {
    param($Parsed)

    $r = [ordered]@{
        Reachable      = $false
        DatabaseExists = $null
        DbState        = $null
        CanOpenDb      = $null
        SqlServerName  = $null
        SqlVersion     = $null
        AvailabilityGroup = $null
        ListenerName   = $null
        ReplicaRole    = $null
        PrimaryReplica = $null
        VerifyError    = $null
    }

    $sb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $sb['Data Source']      = $Parsed.Server
    $sb['Initial Catalog']  = 'master'
    $sb['Connect Timeout']  = $TimeoutSeconds
    $sb['Application Name'] = 'Get-ClientDatabase'

    if ($Credential) {
        $sb['User ID']  = $Credential.UserName
        $sb['Password'] = $Credential.GetNetworkCredential().Password
    }
    elseif ($Parsed.UserId -and -not $Parsed.Integrated) {
        $sb['User ID']  = $Parsed.UserId
        $sb['Password'] = $Parsed.Password
    }
    else {
        $sb['Integrated Security'] = $true
    }

    $conn = New-Object System.Data.SqlClient.SqlConnection $sb.ConnectionString
    try {
        $conn.Open()
        $r.Reachable = $true

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
SELECT CAST(SERVERPROPERTY('ServerName')     AS nvarchar(256)),
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64)),
       (SELECT state_desc FROM sys.databases WHERE name = @db)
"@
        $dbParam = if ($Parsed.Database) { $Parsed.Database } else { [DBNull]::Value }
        [void]$cmd.Parameters.AddWithValue('@db', $dbParam)

        $state = [DBNull]::Value
        $rd = $cmd.ExecuteReader()
        if ($rd.Read()) {
            $r.SqlServerName = $rd.GetValue(0)
            $r.SqlVersion    = $rd.GetValue(1)
            $state           = $rd.GetValue(2)
        }
        $rd.Close()

        if (-not $Parsed.Database) {
            $r.VerifyError = 'No database named in connection string (login default DB will be used).'
        }
        elseif ($state -is [DBNull]) {
            $r.DatabaseExists = $false
            $r.VerifyError = 'Database not found in sys.databases (or login lacks VIEW ANY DATABASE).'
        }
        else {
            $r.DatabaseExists = $true
            $r.DbState = $state
            try {
                $conn.ChangeDatabase($Parsed.Database)
                $r.CanOpenDb = $true
            }
            catch {
                $r.CanOpenDb = $false
                $r.VerifyError = $_.Exception.GetBaseException().Message
            }

            # Availability Group / listener details (skipped quietly if not AG or no permission)
            try {
                $conn.ChangeDatabase('master')
                $ag = $conn.CreateCommand()
                $ag.CommandText = @"
SELECT ag.name,
       (SELECT STUFF((SELECT ', ' + l.dns_name FROM sys.availability_group_listeners l
                      WHERE l.group_id = ag.group_id FOR XML PATH('')), 1, 2, '')),
       ars.role_desc,
       gs.primary_replica
FROM sys.databases d
JOIN sys.availability_replicas ar ON ar.replica_id = d.replica_id
JOIN sys.availability_groups  ag ON ag.group_id  = ar.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states ars ON ars.replica_id = ar.replica_id
LEFT JOIN sys.dm_hadr_availability_group_states   gs  ON gs.group_id   = ag.group_id
WHERE d.name = @db
"@
                [void]$ag.Parameters.AddWithValue('@db', $Parsed.Database)
                $rd = $ag.ExecuteReader()
                if ($rd.Read()) {
                    $r.AvailabilityGroup = $rd.GetValue(0)
                    $r.ListenerName      = $rd.GetValue(1)
                    $r.ReplicaRole       = $rd.GetValue(2)
                    $r.PrimaryReplica    = $rd.GetValue(3)
                }
                $rd.Close()
                foreach ($k in 'AvailabilityGroup','ListenerName','ReplicaRole','PrimaryReplica') {
                    if ($r[$k] -is [DBNull]) { $r[$k] = $null }
                }
            }
            catch {
                Write-Verbose "AG lookup skipped for $($Parsed.Server): $($_.Exception.GetBaseException().Message)"
            }
        }
    }
    catch {
        $r.VerifyError = $_.Exception.GetBaseException().Message
    }
    finally {
        $conn.Dispose()
    }

    # Some logins (Azure SQL, contained users) can't open master but can open their own database.
    if (-not $r.Reachable -and $Parsed.Database -and $r.VerifyError -match '(?i)login failed|cannot open database "master"|not able to access the database "master"') {
        $sb['Initial Catalog'] = $Parsed.Database
        $conn2 = New-Object System.Data.SqlClient.SqlConnection $sb.ConnectionString
        try {
            $conn2.Open()
            $c = $conn2.CreateCommand()
            $c.CommandText = "SELECT CAST(SERVERPROPERTY('ServerName') AS nvarchar(256)), CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64)), CAST(DATABASEPROPERTYEX(DB_NAME(), 'Status') AS nvarchar(60))"
            $rd = $c.ExecuteReader()
            if ($rd.Read()) { $r.SqlServerName = $rd.GetValue(0); $r.SqlVersion = $rd.GetValue(1); $r.DbState = $rd.GetValue(2) }
            $rd.Close()
            $r.Reachable = $true; $r.DatabaseExists = $true; $r.CanOpenDb = $true; $r.VerifyError = $null
        }
        catch {
            $r.VerifyError = $_.Exception.GetBaseException().Message
        }
        finally { $conn2.Dispose() }
    }

    [pscustomobject]$r
}

#endregion

#region ---------- Input: instance name or URL ----------

function ConvertTo-NormalUrl {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim() -replace '^(?i)https?://', ''
    return $v.TrimEnd('/').ToLowerInvariant()
}

function Split-InovahUrl {
    # 'https://SANDBOXE1.sandbox.inovahcloud.com/iNovah2-FAY71T/'
    #   -> Server 'sandboxe1', Domain 'sandbox.inovahcloud.com', Path 'inovah2-fay71t'
    # '.sandbox.inovahcloud.com/iNovah2-FAY71T' (leading dot) -> no server, Domain 'sandbox.inovahcloud.com'
    param([string]$Value)
    $n = ConvertTo-NormalUrl $Value
    if (-not $n) { return $null }
    $noServer = $n.StartsWith('.')
    $parts  = $n.TrimStart('.') -split '/', 2
    $hostName = $parts[0]
    $path   = if ($parts.Count -gt 1) { $parts[1].Trim('/') } else { '' }
    $labels = $hostName -split '\.'
    $server = $null; $domain = $hostName
    if (-not $noServer -and $labels.Count -ge 3) {
        $server = $labels[0]
        $domain = ($labels[1..($labels.Count - 1)]) -join '.'
    }
    [pscustomobject]@{ Host = $hostName; Server = $server; Domain = $domain; Path = $path }
}

function Get-UrlMatch {
    # Compares a config's WebApplicationURL with the requested URL, ignoring the server name
    # (the first part of the host name, e.g. SANDBOXE1).
    param([string]$ConfigUrl)
    $c = Split-InovahUrl $ConfigUrl
    if (-not $c) { return [pscustomobject]@{ Rank = 0; Label = 'No WebApplicationURL' } }
    if ($c.Path -ne $reqUrl.Path) { return [pscustomobject]@{ Rank = 0; Label = 'No' } }

    $domainOk = ($c.Domain -eq $reqUrl.Domain) -or ($c.Host -eq $reqUrl.Host) -or
                $c.Host.EndsWith(".$($reqUrl.Host)")      # requested 'sandbox.inovahcloud.com/...' without server
    if ($domainOk) {
        $srv = if ($c.Server) { " (server in config: $($c.Server.ToUpperInvariant()))" } else { '' }
        return [pscustomobject]@{ Rank = 2; Label = "Match$srv" }
    }
    return [pscustomobject]@{ Rank = 1; Label = "Path only - domain differs ($($c.Host))" }
}

function Get-UrlPath {
    param([string]$NormalUrl)
    if (-not $NormalUrl) { return $null }
    $parts = $NormalUrl -split '/', 2
    if ($parts.Count -gt 1) { return $parts[1] } else { return '' }
}

# ---------- Interactive mode: nothing passed on the command line ----------
if (-not $Instance -and -not $Url) {
    Write-Host ""
    Write-Host "iNovah instance lookup" -ForegroundColor Cyan
    Write-Host "  Enter an iNovah instance, e.g.  FAY71T"
    Write-Host "  or the site URL, e.g.           .sandbox.inovahcloud.com/iNovah2-FAY71T/"
    Write-Host "  (Q to quit)"
    while ($true) {
        $answer = ([string](Read-Host "`nInstance or URL")).Trim()
        if ($answer -match '^(?i)q(uit)?$') { break }
        if (-not $answer) { Write-Host "Please enter an instance or URL." -ForegroundColor Yellow; continue }

        # Run the lookup (same script, same switches) for what was entered
        & $PSCommandPath -Instance $answer @PSBoundParameters | Out-String | Write-Host

        $again = Read-Host "Look up another? (Y/N)"
        if ($again -notmatch '^(?i)y') { break }
    }
    return
}

# A URL passed in the Instance position (instance names never contain '.' or '/')
if (-not $Url -and $Instance -match '[./]') { $Url = $Instance; $Instance = $null }

$normUrl = ConvertTo-NormalUrl $Url
$reqUrl  = Split-InovahUrl $Url
$urlPath = if ($reqUrl) { $reqUrl.Path }

if ($Url -and -not $Instance) {
    $segments = @($urlPath -split '/' | Where-Object { $_ })
    $seg = $segments | Where-Object { $_ -match '^(?i)inovah\d*[-_]' } | Select-Object -First 1
    if (-not $seg -and $segments) { $seg = $segments[-1] }
    if ($seg) { $Instance = ($seg -replace '^(?i)inovah\d*[-_]', '').ToUpperInvariant() }
    if ($Instance) { Write-Host "URL '$Url' -> iNovah instance '$Instance'" -ForegroundColor Cyan }
}

#endregion

#region ---------- Step 1: find the client's app folders ----------

if ($Instance) { Write-Host "Searching for iNovah instance '$Instance'..." -ForegroundColor Cyan }

$locations = [ordered]@{}   # path -> how it was found

function Add-Location {
    param([string]$Path, [string]$Source)
    if (-not $Path) { return }
    $Path = [Environment]::ExpandEnvironmentVariables($Path).TrimEnd('\')
    if ((Test-Path -LiteralPath $Path) -and -not $locations.Contains($Path)) {
        $locations[$Path] = $Source
    }
}

if ($Instance -and -not $SkipIIS) {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        foreach ($site in Get-Website) {
            $sitePath = [Environment]::ExpandEnvironmentVariables($site.PhysicalPath)
            if ($site.Name -like "*$Instance*" -or $sitePath -like "*$Instance*") {
                Add-Location $sitePath "IIS site: $($site.Name)"
            }
            foreach ($app in Get-WebApplication -Site $site.Name) {
                $appPath = [Environment]::ExpandEnvironmentVariables($app.PhysicalPath)
                if ($app.Path -like "*$Instance*" -or $appPath -like "*$Instance*") {
                    Add-Location $appPath "IIS app: $($site.Name)$($app.Path)"
                }
            }
        }
    }
    catch {
        Write-Verbose "IIS lookup skipped: $($_.Exception.Message)"
    }
}

foreach ($root in $SearchRoot) {
    if (-not $Instance) { break }
    if (-not (Test-Path -LiteralPath $root)) {
        Write-Verbose "Search root not found, skipping: $root"
        continue
    }

    # Top-level folders, minus system folders
    $topDirs = Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { $ExcludeFolders -notcontains $_.Name }

    foreach ($dir in $topDirs) {
        if ($dir.Name -like "*$Instance*") { Add-Location $dir.FullName 'Folder name match' }
        Get-ChildItem -LiteralPath $dir.FullName -Directory -Recurse -Depth ($Depth - 1) -Filter "*$Instance*" -ErrorAction SilentlyContinue |
            ForEach-Object { Add-Location $_.FullName 'Folder name match' }
    }
}

# URL given but no folder named for it: look for a config whose WebApplicationURL matches
if ($normUrl -and $locations.Count -eq 0) {
    Write-Host "No folder named for '$Instance'; checking web.config/user.config files for WebApplicationURL '$Url'..." -ForegroundColor Cyan
    foreach ($root in $SearchRoot) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $topDirs = Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { $ExcludeFolders -notcontains $_.Name }
        foreach ($dir in $topDirs) {
            Get-ChildItem -LiteralPath $dir.FullName -File -Recurse -Depth ($Depth + 3) -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -eq 'web.config' -or $_.Name -eq 'user.config' } |
                ForEach-Object {
                    $cfgUrl = Get-AppSettingValue -Path $_.FullName -Key $UrlKey
                    if ($cfgUrl -and (Get-UrlMatch $cfgUrl).Rank -ge 1) {
                        Add-Location $_.DirectoryName "WebApplicationURL match ($($_.Name))"
                    }
                }
        }
    }
    # Instance name from a folder like 'iNovah 2 [FAY71T]', if we didn't get one from the URL
    if (-not $Instance -and $locations.Count) {
        $first = @($locations.Keys)[0]
        if ($first -match '\[([^\]]+)\]') { $Instance = $Matches[1] }
    }
}

if ($locations.Count -eq 0) {
    if ($normUrl) {
        Write-Warning "No site on this server has WebApplicationURL '$Url' (or a folder named for '$Instance'). The URL may be hosted on a different server."
        return
    }
    Write-Warning "No IIS sites or folders matching '$Instance' were found under: $($SearchRoot -join ', '). Try -SearchRoot with another path or a larger -Depth."
    return
}

Write-Verbose "Found $($locations.Count) location(s):"
$locations.GetEnumerator() | ForEach-Object { Write-Verbose "  $($_.Key)  [$($_.Value)]" }

#endregion

#region ---------- Step 2 & 3: read configs, apply user.config overrides, verify ----------

function Get-ConfigEntries {
    param([System.IO.FileInfo]$File)
    if ($File.Extension -eq '.json') { Get-JsonConnStrings -Path $File.FullName }
    else { Get-XmlConnStrings -Path $File.FullName }
}

function Get-EntrySummary {
    param($Entry)
    if (-not $Entry) { return $null }
    $p = if ($Entry.Raw) { ConvertFrom-ConnString $Entry.Raw }
    if ($p) { "Data Source=$($p.Server); Initial Catalog=$($p.Database)" } else { '(unreadable)' }
}

function New-ResultRow {
    param($Entry, [string]$FoundVia, [string]$ValueFrom, $OverriddenEntry,
          [string]$WebAppUrl, [string]$WebAppUrlFrom)

    $p  = if ($Entry.Raw) { ConvertFrom-ConnString $Entry.Raw } else { $null }
    $notes = @()

    # If user.config only sets some values, the rest still come from web.config
    if ($p -and $OverriddenEntry -and $OverriddenEntry.Raw) {
        $w = ConvertFrom-ConnString $OverriddenEntry.Raw
        if ($w) {
            $filled = @()
            if (-not $p.Server   -and $w.Server)   { $p.Server   = $w.Server;   $filled += 'Data Source' }
            if (-not $p.Database -and $w.Database) { $p.Database = $w.Database; $filled += 'Initial Catalog' }
            if (-not $p.UserId -and -not $p.Integrated) {
                if ($w.UserId)         { $p.UserId = $w.UserId; $p.Password = $w.Password; $filled += 'User ID/Password' }
                elseif ($w.Integrated) { $p.Integrated = $true; $filled += 'Integrated Security' }
            }
            if ($filled) { $notes += "Not set in user.config, taken from web.config: $($filled -join ', ')." }
        }
    }

    $ds = Split-DataSource $p.Server

    $auth = if (-not $p)            { $null }
            elseif ($p.Integrated)  { 'Windows (Integrated)' }
            elseif ($p.UserId)      { "SQL login: $($p.UserId)" }
            else                    { 'Unknown' }

    $pwd = if ($p -and $p.Password) { if ($ShowPassword) { $p.Password } else { '********' } }

    $dbMatches = if ($p -and $p.Database) { $p.Database -eq $Instance } else { $null }

    if ($dbMatches -eq $false -and $p.Database -notlike "$Instance*") {
        $notes += "Initial Catalog '$($p.Database)' does not match instance '$Instance'."
    }
    if ($p -and $p.Provider -and -not $p.Server) { $notes += "Non-SQL Server provider: $($p.Provider)" }

    # Same columns on every row so Format-Table / Export-Csv show everything
    $row = [ordered]@{
        Instance             = $Instance
        FoundVia             = $FoundVia
        Root                 = $loc
        Folder               = Split-Path -Parent $Entry.File
        ValueFrom            = $ValueFrom
        Section              = $Entry.Section
        Name                 = $Entry.Name
        DataSource           = $p.Server
        ServerHost           = $ds.Host
        NamedInstance        = $ds.Instance
        Database             = $p.Database
        MatchesInstance      = $dbMatches
        WebApplicationURL    = $WebAppUrl
        UrlFrom              = $WebAppUrlFrom
        UrlMatch             = $null
        Auth                 = $auth
        Password             = $pwd
        OverriddenWebConfig  = (Get-EntrySummary $OverriddenEntry)
        Reachable            = $null
        DatabaseExists       = $null
        DbState              = $null
        CanOpenDb            = $null
        SqlServerName        = $null
        SqlVersion           = $null
        AvailabilityGroup    = $null
        ListenerName         = $null
        ReplicaRole          = $null
        PrimaryReplica       = $null
        VerifyError          = $null
        Note                 = $null
        ConfigFile           = $Entry.File
    }

    if ($notes) { $row.Note = $notes -join ' ' }
    $obj = [pscustomobject]$row

    # Keep the parsed values aside (not on the output object, so passwords never leak into it)
    $parsedMap["$($Entry.File)|$($Entry.Name)"] = $p

    # In summary mode only the selected iNovah connection is verified, later
    if (-not $DeferVerify) { Add-Verification -Row $obj }
    $obj
}

function Add-Verification {
    param($Row)
    $p = $parsedMap["$($Row.ConfigFile)|$($Row.Name)"]
    if ($SkipVerify -or -not $p -or -not $p.Server) { return }

    $cacheKey = '{0}|{1}|{2}|{3}' -f $p.Server, $p.Database, $p.UserId, $p.Integrated
    if (-not $verifyCache.ContainsKey($cacheKey)) {
        Write-Verbose "Verifying $($p.Server) / $($p.Database)..."
        $verifyCache[$cacheKey] = Test-SqlDatabase -Parsed $p
    }
    foreach ($prop in $verifyCache[$cacheKey].PSObject.Properties) { $Row.$($prop.Name) = $prop.Value }

    if ($p.Integrated -and -not $Credential) {
        $extra = "Verified as $env:USERDOMAIN\$env:USERNAME, not the app pool identity."
        $Row.Note = (@($Row.Note, $extra) | Where-Object { $_ }) -join ' '
    }
}

$seenDirs    = @{}
$parsedMap   = @{}
$DeferVerify = -not $Detailed

$results = foreach ($loc in $locations.Keys) {
    $foundVia = $locations[$loc]

    $files = Get-ChildItem -LiteralPath $loc -Recurse -Depth 4 -File -ErrorAction SilentlyContinue |
        Where-Object {
            $name = $_.Name
            ($_.FullName -notmatch '\\(node_modules|obj|\.git)\\') -and
            ($name -eq 'web.config' -or $name -eq 'user.config')   # exact names only - no backups
        }

    # Work folder by folder so user.config can be paired with the web.config beside it
    foreach ($group in ($files | Group-Object DirectoryName)) {
        if ($seenDirs.ContainsKey($group.Name)) { continue }
        $seenDirs[$group.Name] = $true

        $web  = $group.Group | Where-Object { $_.Name -eq 'web.config' }  | Select-Object -First 1
        $user = $group.Group | Where-Object { $_.Name -eq 'user.config' } | Select-Object -First 1

        if (-not ($web -or $user)) { continue }

        # WebApplicationURL: user.config wins if it has a non-empty value
        $urlArgs = @{ WebAppUrl = $null; WebAppUrlFrom = $null }
        $userUrl = if ($user) { Get-AppSettingValue -Path $user.FullName -Key $UrlKey }
        $webUrl  = if ($web)  { Get-AppSettingValue -Path $web.FullName  -Key $UrlKey }
        if ($userUrl) {
            $urlArgs.WebAppUrl = $userUrl
            $urlArgs.WebAppUrlFrom = if ($webUrl -and $webUrl -ne $userUrl) { "user.config (web.config had: $webUrl)" } else { 'user.config' }
        }
        elseif ($webUrl) {
            $urlArgs.WebAppUrl = $webUrl
            $urlArgs.WebAppUrlFrom = 'web.config'
        }
        else {
            $urlArgs.WebAppUrlFrom = "'$UrlKey' not found"
        }

        $webEntries  = if ($web)  { @(Get-ConfigEntries $web) }  else { @() }
        $userEntries = if ($user) { @(Get-ConfigEntries $user) } else { @() }

        $webByName = @{}
        foreach ($w in $webEntries) { $webByName[[string]$w.Name] = $w }
        # A user.config entry only overrides if it actually sets a server or database.
        # Empty / placeholder entries fall through to web.config.
        $userEntries = @($userEntries | Where-Object {
            $up = if ($_.Raw) { ConvertFrom-ConnString $_.Raw }
            $keep = $up -and ($up.Server -or $up.Database)
            if (-not $keep -and $_.Name) { Write-Verbose "user.config entry '$($_.Name)' has no Data Source/Initial Catalog; using web.config." }
            $keep
        })

        $userByName = @{}
        foreach ($u in $userEntries) { $userByName[[string]$u.Name] = $u }

        # web.config values that user.config does NOT override
        foreach ($w in $webEntries) {
            if (-not $userByName.ContainsKey([string]$w.Name)) {
                $from = if (-not $user)             { 'web.config (no user.config)' }
                        elseif (-not $userEntries)  { 'web.config (user.config has no DB settings)' }
                        else                        { 'web.config (not overridden by user.config)' }
                New-ResultRow -Entry $w -FoundVia $foundVia -ValueFrom $from @urlArgs
            }
        }

        # user.config values win
        foreach ($u in $userEntries) {
            $o = $webByName[[string]$u.Name]
            $from = if ($o) { 'user.config (overrides web.config)' }
                    elseif ($web) { 'user.config (not in web.config)' }
                    else { 'user.config' }
            New-ResultRow -Entry $u -FoundVia $foundVia -ValueFrom $from -OverriddenEntry $o @urlArgs
        }
    }
}

if (-not $results) {
    Write-Warning "Found the instance folder(s) but no connection strings in any web.config or user.config."
    return
}

# ---------- If a URL was given, keep only sites whose WebApplicationURL matches it ----------
if ($normUrl) {
    $rank = @{}
    foreach ($r in $results) {
        $m = Get-UrlMatch $r.WebApplicationURL
        $r.UrlMatch = $m.Label
        $rank[$r.Folder] = [Math]::Max([int]$rank[$r.Folder], $m.Rank)
    }
    # Domain + path match (any server) wins; path-only is a fallback
    $matched = @($results | Where-Object { $rank[$_.Folder] -eq 2 })
    if (-not $matched) { $matched = @($results | Where-Object { $rank[$_.Folder] -eq 1 }) }
    if ($matched) {
        # Keep every connection string from the folder(s) that host that URL
        $folders = $matched.Folder | Sort-Object -Unique
        $results = @($results | Where-Object { $folders -contains $_.Folder })
    }
    else {
        Write-Warning "Found '$Instance', but no WebApplicationURL there matches '$Url'. Showing what was found."
    }
}

if ($Detailed) {
    $results
    return
}

# ---------- Pick the connection iNovah itself uses ----------
$dbRows = @($results | Where-Object { $_.DataSource -or $_.Database })
if (-not $dbRows) { $dbRows = @($results) }

$selected = @($dbRows | Where-Object { $_.Name -eq $PrimaryKey })
if (-not $selected) {
    $selected = @($dbRows | Where-Object { $_.Database -like "$Instance*" })
    if ($selected) { Write-Verbose "No '$PrimaryKey' entry found; using entries whose Initial Catalog matches '$Instance'." }
}
if (-not $selected) {
    Write-Warning "Could not tell which connection is iNovah's own (no '$PrimaryKey' entry and no Initial Catalog matching '$Instance'). Showing all."
    $selected = $dbRows
}

$hidden = @($dbRows | Where-Object { $selected -notcontains $_ })
if ($All) { $selected = $dbRows }
elseif ($hidden) {
    $hiddenNames = ($hidden.Name | Sort-Object -Unique) -join ', '
    Write-Host "($($hidden.Count) other connection string(s) not shown: $hiddenNames. Use -All to include them.)" -ForegroundColor DarkGray
}

foreach ($row in $selected) { Add-Verification -Row $row }
$dbRows = $selected

# ---------- Summary: one block per database ----------


$summary = $dbRows |
    Group-Object { '{0}|{1}' -f $_.DataSource, $_.Database } |
    Sort-Object Count -Descending |
    ForEach-Object {
        $g = $_.Group
        $f = $g[0]

        # URL from the same config folder(s) as this database; fall back to any URL found
        $urlRows = @($g | Where-Object WebApplicationURL)
        if (-not $urlRows) { $urlRows = @($results | Where-Object WebApplicationURL) }
        $urls    = @($urlRows.WebApplicationURL | Sort-Object -Unique)
        $urlFrom = if ($urlRows) { $urlRows[0].UrlFrom -replace '\s*\(web\.config had:.*\)$', '' } else { "'$UrlKey' not found" }

        $status = if ($SkipVerify)            { 'Not verified (-SkipVerify)' }
                  elseif ($f.CanOpenDb)       { "OK - database $($f.DbState), connected" }
                  elseif ($f.DatabaseExists)  { "Database $($f.DbState) but could not be opened" }
                  elseif ($f.Reachable)       { 'Server reachable, DATABASE NOT FOUND' }
                  elseif ($f.VerifyError -match '(?i)login failed') { 'LOGIN FAILED (server reachable)' }
                  elseif ($f.DataSource)      { 'SERVER UNREACHABLE' }
                  else                        { 'No server in config' }

        # Where this database is referenced, relative to the instance folder
        $usedBy = $g | ForEach-Object {
            $rel = if ($_.Root -and $_.Folder.StartsWith($_.Root)) { $_.Folder.Substring($_.Root.Length).TrimStart('\','/') } else { $_.Folder }
            if (-not $rel) { $rel = '(root)' }
            "$rel [$($_.Name)]"
        } | Sort-Object -Unique

        $warnings = @($g.Note) + @($g.VerifyError) | Where-Object { $_ } | Sort-Object -Unique

        $dsLabel = if ($f.ListenerName) { "$($f.DataSource)  (AG listener)" } else { $f.DataSource }

        $o = [pscustomobject][ordered]@{
            Instance          = $Instance
            WebApplicationURL = ($urls -join '; ')
            UrlSource         = $urlFrom
            'SQL Instance (Data Source)' = $dsLabel
            AvailabilityGroup = $f.AvailabilityGroup
            'Active SQL Instance'        = $f.SqlServerName
            ReplicaRole       = $f.ReplicaRole
            PrimaryReplica    = $f.PrimaryReplica
            'Database (Initial Catalog)' = $f.Database
            Status            = $status
            Login             = $f.Auth
            SqlVersion        = $f.SqlVersion
            SettingsFrom      = (($g.ValueFrom | Sort-Object -Unique) -join '; ')
            UsedBy            = ($usedBy -join "`n")
            Warnings          = ($warnings -join "`n")
        }
        if ($normUrl) {
            $o | Add-Member -NotePropertyName RequestedUrl -NotePropertyValue $Url
            $o | Add-Member -NotePropertyName UrlMatch     -NotePropertyValue ($g.UrlMatch | Where-Object { $_ } | Select-Object -First 1)
        }
        $o
    }

$summary
