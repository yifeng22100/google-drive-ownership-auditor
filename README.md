# Google Drive Ownership Auditor

A Google Apps Script that walks every folder and file under a Google Drive
folder and records **who owns each item** in a Google Sheets report. Items are
flagged as owned by you, by another account in your organisation, by an
external domain, or by a shared drive. Shortcuts are followed to check who
owns their targets.

It's built for large trees, resuming automatically across Apps Script's
6-minute execution limit.

The source Drive is **read-only**: the script never moves, edits or shares
anything.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | The script |
| `appsscript.json` | Manifest: enables the Drive advanced service and declares OAuth scopes |

## Setup

1. Create a new project at [script.google.com](https://script.google.com).
2. Replace `Code.gs` with the one from this repo.
3. Show the manifest: **Project Settings → Show "appsscript.json" manifest
   file in editor**. Replace its contents with `appsscript.json` from this
   repo. Change `timeZone` to your own (for example `Asia/Kuala_Lumpur`); it's
   used for the report name.
4. Edit `CONFIG` at the top of `Code.gs`:

   ```js
   ROOT_FOLDER_ID: "YOUR_ROOT_FOLDER_ID",  // Folder to audit
   MY_EMAIL: "you@example.com",            // Your account
   INTERNAL_DOMAINS: ["example.com"],      // Your organisation's domains
   ```

   The folder ID is the last part of the folder's URL:
   `https://drive.google.com/drive/folders/<ROOT_FOLDER_ID>`.

   `startOwnershipAudit()` refuses to run while any placeholder is left.

## Usage

| Function | When to run it |
|---|---|
| `startOwnershipAudit()` | Once, to create the report and start scanning. Approve the permissions prompt. |
| `resumeOwnershipAudit()` | Runs automatically every minute until the scan finishes. Safe to run by hand. |
| `checkOwnershipAuditStatus()` | Any time; logs progress and the report URL. |
| `retryErrorFolders()` | After completion, if any folders failed. Re-queues them from the failed page. |
| `resetOwnershipAudit()` | Before starting a brand-new audit. Leaves the old report in Drive. |

`startOwnershipAudit()` installs a 1-minute trigger. Each run works for about
4.5 minutes and saves its progress; when nothing is left to scan, it marks the
report `COMPLETED`, applies a filter and colour coding, and removes the
trigger.

Most rows in **Executions** will finish in 1–2 seconds with
`Audit already running. Skipping.` That's expected: the trigger fires every
minute while a run lasts about 5, and a lock stops runs from overlapping.

## The report

**Ownership Audit**: one row per item.

| Column | Meaning |
|---|---|
| Path | Folder path from the root folder |
| Name, Type | `FILE`, `FOLDER`, `SHORTCUT` or `SCAN ERROR` |
| Ownership Status | See below |
| Owner Email, Owner Domain | First owner of the item |
| File ID, File URL | |
| Shortcut Target ID / Owner / Status | Filled in for shortcuts |
| Remarks | Errors and notes |

**Ownership Status values**

| Status | Meaning |
|---|---|
| `OWNED BY ME` | Owned by `MY_EMAIL` |
| `OTHER ACCOUNT` | Owned by another account in `INTERNAL_DOMAINS` |
| `EXTERNAL DOMAIN` | Owned outside your organisation. Usually what you're looking for. |
| `SHARED DRIVE` | Lives in a shared drive (no individual owner) |
| `UNKNOWN OWNER` | The API returned no owner |
| `SCAN INCOMPLETE` | A folder that couldn't be listed |

Also filter **Shortcut Target Status** for `EXTERNAL DOMAIN`: these are
internally owned shortcuts that point at externally owned files, and they
aren't counted in the main status column.

**Scan Queue** is the script's work list (one row per folder, with page
checkpoints). **Summary** shows counts and overall status.

## After the scan

### Check for duplicate rows

If a run is cut off while saving, a few rows can be written twice. Files that
sit in more than one folder also legitimately appear once per location, with
the same File ID but a different Path. Count only the true duplicates (same
Path **and** File ID):

```
=COUNTA('Ownership Audit'!G2:G) - COUNTUNIQUE(FILTER('Ownership Audit'!A2:A&"|"&'Ownership Audit'!G2:G, 'Ownership Audit'!G2:G<>""))
```

If it's above 0: select columns A–L on **Ownership Audit**, choose
**Data → Data cleanup → Remove duplicates**, tick **Data has header row**,
and tick only **Path** and **File ID**.

### List the folders one level below a given folder

```
={"Folder Name"; SORT(UNIQUE(FILTER('Ownership Audit'!B2:B, 'Ownership Audit'!C2:C="FOLDER", REGEXMATCH('Ownership Audit'!A2:A, "(^|/)FOLDER NAME/[^/]+$"))))}
```

Replace `FOLDER NAME` with the parent folder's name, escaping regex characters
(for example `2024\. Archive`). For two levels down, change `[^/]+$` to
`[^/]+/[^/]+$`.

### Dealing with externally owned items

Google Drive doesn't allow ownership to be transferred to another
organisation. The usual fix is **Make a copy** into a folder you own, then
remove the original from your tree.

## Tuning

| Setting | Default | Notes |
|---|---|---|
| `CONCURRENCY` | 25 | Parallel Drive API requests per wave. 40 caused heavy throttling in testing; lower to 10–15 if the timing log shows throttle sleep. |
| `MAX_RUNTIME_MS` | 280000 | Stop starting new work after this; leaves margin under the 6-minute limit. |
| `FLUSH_EVERY_MS` | 60000 | How often progress is written to the sheet. |
| `MAX_RETRIES` | 5 | Per folder per run, for HTTP 429/5xx. Then it's marked `ERROR`. |
| `CHECK_SHORTCUT_TARGETS` | true | Turn off to skip the extra lookup per shortcut. |

Each run logs two lines in **Executions**:

```
Run: 285s | Waves: 140 | Items: 12000 | Queue: 40000 | Pending folders: 25000 | Errors: 0
Timing: load 10s | list 200s | shortcuts 20s (300 lookups) | throttle sleep 0s (0 throttled) | flush 30s
```

`Pending folders` rises at first as new subfolders are discovered, then falls
to 0.

## Limits

- **Visibility**: only items your account can see are reported.
- **Trash**: trashed items are skipped.
- **Owner**: only the first owner is recorded.
- **Paths**: a folder name containing `/` makes its path ambiguous.
- **Quotas** (Google Workspace): about 100,000 URL Fetch calls and 6 hours of
  trigger run time per user per day. One call is used per folder page and per
  shortcut target. Very large trees may need to continue the next day; the
  scan resumes where it stopped.

## Version history

**3.1**: Adds the `Timing:` log line showing where each run spends its time.

**3.0**: Rewrite for speed, about 60× faster than 2.0 on a large tree.
- Folders are listed in parallel with `UrlFetchApp.fetchAll` instead of one
  at a time.
- Shortcut targets are looked up in parallel and cached.
- The Scan Queue is read once per run and kept in memory. 2.0 re-read the
  whole queue sheet for every folder, which got slower as the queue grew.
- Sheet writes are buffered and saved in bulk.
- No per-run item cap.
- Rate-limited requests (429/5xx) are retried with backoff.
- Adds `retryErrorFolders()`.
- Uses the same sheets and state as 2.0, so a running 2.0 audit can be
  continued by swapping the code.

**2.0**: Batch, resumable audit with a Scan Queue sheet and a 1-minute
resume trigger, using the Drive advanced service one folder at a time.

**1.0**: Initial version.
