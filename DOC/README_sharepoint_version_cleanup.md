# SharePoint version cleanup (Python)

## What this does
- Recursively walks subfolders under a root folder
- Lists files in each folder
- Reads each file's version history
- Deletes versions older than `SP_DAYS_TO_KEEP`
- Keeps the newest version by default
- Supports dry-run first

## Recommended auth
For SharePoint Online REST app-only, Microsoft recommends Microsoft Entra app-only with a certificate.

## Setup
1. Create an Entra app registration
2. Add SharePoint application permission
   - Prefer `Sites.Selected` and grant the target site explicitly
3. Upload/use a certificate for the app
4. Export the private key as PEM and set `SP_PRIVATE_KEY_PATH`
5. Fill `.env` from `.env.example`

## Install
```bash
pip install msal requests
```

## Run dry-run
```bash
python sharepoint_version_cleanup.py
```

## Run actual delete
```bash
python sharepoint_version_cleanup.py --execute
```

## Notes
- The script uses SharePoint REST endpoints similar to your Power Automate flow.
- Folder names like `Forms` are skipped by default.
- Paths containing `%` or `#` may need a ResourcePath-based variant if they occur often.
