# Get-ClientDatabase

A PowerShell script that answers one question quickly: **which SQL Server and database is this iNovah instance actually using, and is it working?**

Give it an iNovah instance name (e.g. `FAY71T`) or a site URL. It finds the instance's application folders on the server, reads the connection strings from the configuration, works out which values are really in effect, and then connects to SQL Server to confirm the database exists and can be opened.

## Requirements

- Windows PowerShell 5.1 or later, run on the iNovah **application server**.
- Network access to the SQL Server(s) named in the configuration.
- **Run as Administrator** if the connection strings are encrypted, or to look up IIS sites (needs the WebAdministration module).

## Quick start

Run the script with no parameters and it will prompt you:

```powershell
.\Get-ClientDatabase.ps1
```

```
iNovah instance lookup
  Enter an iNovah instance, e.g.  FAY71T
  or the site URL, e.g.           .sandbox.inovahcloud.com/iNovah2-FAY71T/
  (Q to quit)

Instance or URL: FAY71T
```

After each result it offers to look up another.

Or pass the instance or URL directly:

```powershell
.\Get-ClientDatabase.ps1 FAY71T
.\Get-ClientDatabase.ps1 sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/
```

> If PowerShell refuses to run the script because it was downloaded, run `Unblock-File .\Get-ClientDatabase.ps1` once, or start it with `powershell -ExecutionPolicy Bypass -File .\Get-ClientDatabase.ps1`.

## What you get

One block per database, for example:

```
Instance                   : FAY71T
WebApplicationURL          : https://sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/
UrlSource                  : user.config
SQL Instance (Data Source) : SBOXSQLPE\FAY71  (AG listener)
AvailabilityGroup          : AG-SANDBOX
Active SQL Instance        : SBOXSQL01\FAY71
ReplicaRole                : PRIMARY
Database (Initial Catalog) : FAY71T
Status                     : OK - database ONLINE, connected
Login                      : SQL login: inovahapp
SettingsFrom               : user.config (overrides web.config)
UsedBy                     : (root) [Data.ConnectionString]
Warnings                   :
```

**Status** tells you at a glance whether things are healthy:

| Status | Meaning |
|---|---|
| `OK - database ONLINE, connected` | Everything checks out. |
| `Database <state> but could not be opened` | The database exists but is offline, restoring, or the login has no access to it. See **Warnings**. |
| `Server reachable, DATABASE NOT FOUND` | The SQL Server answered, but no database has that name (or the login can't see it). Check the Initial Catalog. |
| `LOGIN FAILED (server reachable)` | Wrong username/password, or the login doesn't exist on that server. |
| `SERVER UNREACHABLE` | The server name is wrong, SQL Server is down, or a firewall is blocking the connection. |
| `Not verified (-SkipVerify)` | The configuration was read but no connection was attempted. |

**Warnings** flag things worth a second look — for example a database name that doesn't match the instance, settings that came partly from web.config and partly from user.config, or that the check ran as your Windows account rather than the app pool identity.

## How it works

1. **Finds the instance's folders** by matching the instance name against IIS site and application names/paths, and against folder names under the search roots (`C:\` and `E:\` by default, three levels deep; Windows, Users, ProgramData and other system folders are skipped).
2. **Reads only `web.config` and `user.config`.** Backups and copies such as `web.config.bak` or `web - Copy.config` are ignored, so old settings never confuse the result.
3. **Applies `user.config` overrides.** When a `user.config` sits beside a `web.config`, its connection strings and `WebApplicationURL` win. A `user.config` entry that doesn't set a server or database is ignored, and any value it leaves out is filled in from `web.config`. The output shows where each value came from.
4. **Picks iNovah's own connection** — the `Data.ConnectionString` entry, or failing that, the one whose database name matches the instance. Other connection strings (integrations such as PeopleSoft/CCM, older settings) are hidden unless you add `-All`.
5. **Verifies it** by connecting to SQL Server and checking that the database exists, its state, and that it can actually be opened. For Always On databases it also reports the availability group, listener, replica role and current primary.

### Looking up by URL

When you pass a URL, the instance name is taken from the path (`iNovah2-FAY71T` → `FAY71T`), and the result is checked against each site's `WebApplicationURL`. The server part of the host name (e.g. `SANDBOXE1`) is ignored when matching, so `.sandbox.inovahcloud.com/iNovah2-FAY71T/` works too. If no folder is named for the instance, every `web.config`/`user.config` under the search roots is checked for a matching URL — useful when folder names don't follow the usual pattern.

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Instance` (or `-ClientName`) | — | iNovah instance, e.g. `FAY71T`. A URL here is detected automatically. |
| `-Url` | — | Site URL, e.g. `sandboxe1.sandbox.inovahcloud.com/iNovah2-FAY71T/`. `http://`/`https://` optional. |
| `-SearchRoot` | `C:\`, `E:\` | Folders to search for instance folders. Missing drives are skipped. |
| `-ExcludeFolders` | system folders | Top-level folders to skip on each search root. |
| `-Depth` | `3` | How many folder levels below each root to search. |
| `-UrlKey` | `WebApplicationURL` | appSettings key that holds the site URL. |
| `-PrimaryKey` | `Data.ConnectionString` | Connection string name iNovah uses for its own database. |
| `-All` | off | Show every database connection, not just iNovah's. |
| `-Detailed` | off | Show every connection string from every config file, one row each. |
| `-SkipIIS` | off | Don't look up IIS sites/applications. |
| `-SkipVerify` | off | Read the config only; don't connect to SQL Server. |
| `-Credential` | — | SQL login to verify with instead of the one in the config. |
| `-TimeoutSeconds` | `5` | SQL connection timeout. |
| `-ShowPassword` | off | Show passwords instead of `********`. Use with care. |

## More examples

Pick just the key columns:

```powershell
.\Get-ClientDatabase.ps1 FAY71T | Select-Object 'SQL Instance (Data Source)', 'Database (Initial Catalog)', WebApplicationURL, Status
```

Verify with a different SQL login and save to CSV:

```powershell
.\Get-ClientDatabase.ps1 FAY71D -Credential (Get-Credential) | Export-Csv .\FAY71D-db.csv -NoTypeInformation
```

Search a different drive, deeper:

```powershell
.\Get-ClientDatabase.ps1 FAY71T -SearchRoot D:\ -Depth 5
```

See everything found, file by file:

```powershell
.\Get-ClientDatabase.ps1 FAY71T -Detailed | Format-List
```

## Notes and limitations

- **Windows authentication** connections are verified as *your* account, not the app pool identity, so a success here doesn't guarantee the app pool can connect. Use `-Credential` or test as the service account if in doubt.
- **Encrypted connection strings** can only be decrypted on the server that holds the key, from an elevated session. Otherwise the entry is shown as "could not decrypt".
- A `connectionStrings` section that points to another file via `configSource` is **not** read; the script warns when it sees one.
- "Database not found" can also mean the login lacks `VIEW ANY DATABASE` permission.
- Passwords are masked by default and are never written to the output objects unless `-ShowPassword` is used.
