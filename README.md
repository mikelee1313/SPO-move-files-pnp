# SharePoint Large Library Move Script

A PowerShell script for moving files between SharePoint Online document libraries with PnP.PowerShell.

The script is designed for very large source libraries by processing one folder at a time, retrying on throttling, writing logs directly to disk, and saving a checkpoint so a run can be resumed after interruption.

## What It Does

- Connects to source and target SharePoint Online sites with app-only authentication
- Walks the source library recursively, folder by folder
- Moves files to the target library while preserving the file name
- Creates missing destination folders as needed
- Retries transient SharePoint throttling errors
- Streams the result log directly to CSV on disk
- Saves a checkpoint file so interrupted runs can resume

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- PnP.PowerShell installed
- An Azure AD app registration with certificate-based authentication
- Permission to read from the source site and write to the target site

Install the module if needed:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Files
- [pnpmove.ps1](pnpmove.ps1) - main migration script

## Configuration

Update the values near the top of [pnpmove.ps1](pnpmove.ps1):

- `AppId`
- `TenantId`
- `CertThumbprint`
- `SourceSiteUrl`
- `TargetSiteUrl`
- `SourceLibrary`
- `TargetLibrary`
- `SourceRootFolder`
- `TargetRootFolder`

Logging and resume settings:

- `LogPath` - CSV output file for move results
- `CheckpointPath` - JSON checkpoint file for resume state
- `ResumeFromCheckpoint` - set to `$true` to continue from the last checkpoint
- `CheckpointEvery` - how often to write progress messages

## How To Run

1. Open [pnpmove.ps1](pnpmove.ps1) in PowerShell.
2. Verify the configuration values are correct for your environment.
3. Run the script:

```powershell
.\pnpmove.ps1
```

## Resume Behavior

The script writes a checkpoint file while it runs. If the run stops unexpectedly, you can start it again and it will skip items already processed up to the last saved source file URL.

If you want a completely fresh run, set:

```powershell
$ResumeFromCheckpoint = $false
```

That clears the log and checkpoint files before the run starts.

## Output

By default the script writes:

- CSV log: `C:\Temp\PnP_FB_MoveFiles_Log.csv`
- Checkpoint: `C:\Temp\PnP_FB_MoveFiles_Checkpoint.json`

## Notes

- The script processes one folder at a time so it can handle very large libraries more safely than a flat library-wide enumeration.
- The move still depends on SharePoint Online performance and throttling behavior, so very large migrations can take a long time.
- Do not commit real credentials, tenant IDs, or certificate thumbprints if you plan to publish this repository publicly.

## License

Add your preferred license before publishing to GitHub.
