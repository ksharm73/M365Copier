# OneDriveCopier v2.0 — README

Downloads files from **OneDrive or SharePoint** (via Microsoft Graph API) to a
network mapped drive.  Switch sources with **one line** in `appsettings.json`.
Designed for Autosys scheduling or manual execution — no OneDrive sync client required.

---

## Project Structure

```
OneDriveCopier/
├── src/
│   ├── OneDriveCopier.csproj   ← .NET 8 project + NuGet references
│   ├── appsettings.json        ← all config: auth, source, dest, options
│   ├── Program.cs              ← entry point, config loading, source factory
│   ├── ArgumentParser.cs       ← CLI --key value / --flag parser
│   ├── AppSettings.cs          ← strongly-typed config models
│   ├── GraphClientFactory.cs   ← creates authenticated GraphServiceClient
│   ├── IFileSource.cs          ← abstraction + RemoteFile + FolderNotFoundException
│   ├── OneDriveSource.cs       ← Graph API: user OneDrive implementation
│   ├── SharePointSource.cs     ← Graph API: SharePoint document library implementation
│   ├── FileCopier.cs           ← download engine (retry, integrity check)
│   ├── FileLogger.cs           ← dual console + timestamped file logger
│   └── ExitCodes.cs            ← standardised exit codes (0-6)
└── README.md
```

---

## Architecture

```
Program.cs
  │
  ├── loads appsettings.json  ──→  AppSettings (strongly typed)
  ├── creates GraphServiceClient (app-only / client credentials auth)
  │
  ├── SourceMode = "OneDrive"    ──→  OneDriveSource   ──┐
  └── SourceMode = "SharePoint"  ──→  SharePointSource ──┤  implements IFileSource
                                                          │
                                                    FileCopier
                                                          │
                                              IFileSource.ListFilesAsync()
                                              IFileSource.OpenReadAsync()
                                                          │
                                              writes to Z:\Reports\20260522\
```

**To switch sources: change `"SourceMode"` in `appsettings.json`. Nothing else.**

---

## Azure AD App Registration

Create an **App Registration** in Azure AD (Entra ID):

1. Azure Portal → App Registrations → New Registration
2. Note the **Tenant ID**, **Client ID**
3. Certificates & Secrets → New Client Secret → note the **Value**
4. API Permissions → Add:

| Permission      | Type        | Purpose                            |
|-----------------|-------------|------------------------------------|
| Files.Read.All  | Application | Read OneDrive and SharePoint files |
| Sites.Read.All  | Application | Read SharePoint sites (SP mode)    |

5. **Grant admin consent** for both permissions

> ⚠️ Application permissions (not Delegated) are required — the tool runs
> without a signed-in user.

---

## Configuration (`appsettings.json`)

```jsonc
{
  "AzureAd": {
    "TenantId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret": "your-secret-value"
  },

  // ← ONE LINE TO SWITCH SOURCE
  "SourceMode": "OneDrive",          // "OneDrive" or "SharePoint"

  "OneDrive": {
    "UserId":         "john.doe@contoso.com",
    "BaseFolderPath": "Reports"      // folder inside drive root, excluding date folder
  },

  "SharePoint": {
    "Hostname":       "contoso.sharepoint.com",
    "SitePath":       "/sites/FinanceTeam",
    "LibraryName":    "Documents",
    "BaseFolderPath": "Reports"
  },

  "Destination": {
    "RootPath": "Z:\\Reports"        // date folder appended automatically
  },

  "CopyOptions": {
    "Overwrite":          false,
    "FileFilter":         "*.xls,*.xlsx",
    "MaxRetries":         3,
    "RetryDelaySeconds":  1
  },

  "Logging": {
    "LogDirectory": ""               // blank = exe folder
  }
}
```

### Switching to SharePoint

```jsonc
// Before (OneDrive)
"SourceMode": "OneDrive"

// After (SharePoint)
"SourceMode": "SharePoint"
```

That's the only change needed. All SharePoint-specific values are in
the `SharePoint` block which is ignored when mode is OneDrive, and vice versa.

---

## Build

### Prerequisites
- .NET 8 SDK: https://dotnet.microsoft.com/download/dotnet/8.0

### Self-contained Windows EXE (recommended — no runtime needed on target)
```batch
cd src
dotnet publish -c Release -r win-x64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  -o ..\publish
```

Output: `publish\OneDriveCopier.exe`  (~65 MB)
**Copy both `OneDriveCopier.exe` and `appsettings.json` to the target machine.**

### Framework-dependent (smaller, needs .NET 8 runtime)
```batch
dotnet publish -c Release -r win-x64 --self-contained false -o ..\publish
```

---

## Usage

```
OneDriveCopier.exe --folder-name <date-folder> [options]
```

### Required

| Argument        | Description                             |
|-----------------|-----------------------------------------|
| `--folder-name` | Date folder to copy, e.g. `20260522`    |

### Optional

| Argument     | Description                                              |
|--------------|----------------------------------------------------------|
| `--dest`     | Override `Destination.RootPath` from appsettings.json   |
| `--log-dir`  | Override log file directory                              |
| `--help`     | Print usage                                              |

### Examples

```batch
REM Standard run (all config from appsettings.json)
OneDriveCopier.exe --folder-name 20260522

REM Override destination at runtime
OneDriveCopier.exe --folder-name 20260522 --dest "\\fileserver\share\Reports"

REM Custom log directory
OneDriveCopier.exe --folder-name 20260522 --log-dir "C:\Logs\OneDriveCopier"
```

---

## Destination Path Behaviour

The date folder is **mirrored** under the destination root:

```
SourceMode = OneDrive
  OneDrive: /Reports/20260522/Q1_Risk.xlsx
  Written:  Z:\Reports\20260522\Q1_Risk.xlsx

SourceMode = SharePoint
  SharePoint: /sites/FinanceTeam/Documents/Reports/20260522/Q1_Risk.xlsx
  Written:    Z:\Reports\20260522\Q1_Risk.xlsx
```

The destination path is identical regardless of source — the on-prem
process consuming the files does not need to change when you switch sources.

---

## Logging

Each run produces a unique timestamped log file:
```
OneDriveCopier_20260522_060015.log
```

Sample output:
```
══════════════════════════════════════════════════════════════════════
  OneDrive / SharePoint → Network Drive File Copier
  Run started : 2026-05-22 06:00:15
  Machine     : APP_SERVER_01
  User        : svc_autosys
══════════════════════════════════════════════════════════════════════
2026-05-22 06:00:15.012 [INFO ] Source mode  : OneDrive
2026-05-22 06:00:15.015 [INFO ] Date folder  : 20260522
2026-05-22 06:00:15.016 [INFO ] Destination  : Z:\Reports
2026-05-22 06:00:15.020 [INFO ] Listing remote files via Graph API...
2026-05-22 06:00:15.834 [INFO ] Files found  : 3
────────────────────────────────────────────────────────────────────
2026-05-22 06:00:16.101 [OK   ] COPY  Q1_Risk.xlsx    (1.4 MB)
2026-05-22 06:00:16.340 [OK   ] COPY  Q1_Sales.xlsx   (980.2 KB)
2026-05-22 06:00:16.512 [WARN ] SKIP  Q1_Nikko.xlsx   (already exists at destination)
────────────────────────────────────────────────────────────────────
2026-05-22 06:00:16.513 [INFO ] Total files found : 3
2026-05-22 06:00:16.513 [OK   ] Copied            : 2
2026-05-22 06:00:16.514 [WARN ] Skipped (exist)   : 1
```

---

## Exit Codes (Autosys `exit_code_range`)

| Code | Meaning                                |
|------|----------------------------------------|
| 0    | Success — all files copied             |
| 1    | Invalid arguments / config error       |
| 2    | Remote folder not found                |
| 3    | Destination not writable               |
| 4    | Partial failure — some files failed    |
| 5    | Total failure — no files copied        |
| 6    | Auth failure — check AzureAd config    |

---

## Autosys Job Definition

```
insert_job: ONEDRIVE_COPY_DAILY   job_type: CMD
command: C:\Tools\OneDriveCopier\OneDriveCopier.exe --folder-name %YYYYMMDD%
machine: APP_SERVER_01
owner: svc_autosys
date_conditions: 1
days_of_week: all
start_times: "06:00"
condition: success(UPSTREAM_SYNC_JOB)
on_failure: MAIL_ALERT_JOB
exit_code_range: 4-6
```

### Dynamic date in Autosys
Use a pre-job command or Autosys global variable to set `%YYYYMMDD%`
to today's date in `yyyyMMdd` format.

### PowerShell Scheduled Task
```powershell
$date   = Get-Date -Format "yyyyMMdd"
$exeDir = "C:\Tools\OneDriveCopier"

$action = New-ScheduledTaskAction `
    -Execute "$exeDir\OneDriveCopier.exe" `
    -Argument "--folder-name $date" `
    -WorkingDirectory $exeDir

$trigger = New-ScheduledTaskTrigger -Daily -At 6:00AM

Register-ScheduledTask `
    -TaskName   "OneDriveCopier_Daily" `
    -Action     $action `
    -Trigger    $trigger `
    -RunLevel   Highest `
    -User       "DOMAIN\svc_account" `
    -Password   "password"
```

---

## Extending with a New Source

To add a new source (e.g. Azure Blob Storage):

1. Create `AzureBlobSource.cs` implementing `IFileSource`
2. Add its settings class in `AppSettings.cs`
3. Add a section in `appsettings.json`
4. Add one `case` in `Program.cs` source factory switch:
   ```csharp
   "azureblob" => new AzureBlobSource(settings.AzureBlob),
   ```
5. Change `SourceMode` in `appsettings.json`

No other code changes needed.
