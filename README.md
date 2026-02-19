# OneDrive Migration Script

This PowerShell script copies files from OneDrive to a destination folder (default: Nextcloud). It hydrates cloud-only files when needed, verifies each copy by size, and writes a log to the Desktop. Only source files can be dehydrated, and only if you choose that option.

## Features
- Detects OneDrive accounts (Personal/Business) and lets you choose the source
- Finds common cloud destinations (e.g., Nextcloud, Dropbox, Google Drive) or uses a custom folder
- Hydrates files before copying (downloads cloud-only files)
- Copies via streaming and verifies file size
- Optional dehydrate of source files only
- Progress UI in the terminal and log file on the Desktop
- Resume support: skips files that already exist with the same size

## Requirements
- Windows
- PowerShell 5.1 or newer
- OneDrive client installed and signed in
- Destination cloud client installed and signed in (so files can upload)

## Usage
1. Save the script locally, e.g. `OneDriveMigrationScript.ps1`.
2. Open PowerShell.
3. Run the script:

```powershell
# Example: run from the script folder
.\OneDriveMigrationScript.ps1
```

4. Follow the prompts:
   - Confirm OneDrive account type
   - Choose the source
   - Select the source dehydrate mode
   - Choose or enter a destination folder (for "Custom" the exact folder is used)

## What the script does
- Lists all files under the OneDrive source (no folders)
- Hydrates cloud-only files when needed
- Copies into `OneDriveMigration` for detected cloud destinations
- If you choose "Custom", files are copied directly into the folder you enter (no extra subfolder is created)
- Verifies each file (size check)
- Optionally dehydrates source files, based on your selection

## Skipped files
Files are skipped when:
- they already exist in the destination **and** the file size matches
- they are in the skip list (e.g., `desktop.ini`, `Thumbs.db`)

## Logging
- Log file: `OneDrive_Migration.log` on the Desktop
- Errors are collected and printed at the end

## Notes and safety
- Destination files are never dehydrated. The cloud client must upload them first.
- The script is conservative: copies are verified and faulty destination files are removed.
- Dehydration marks files as cloud-only. If not supported, files stay local.
- Test with a small folder first.
