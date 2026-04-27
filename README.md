# Invoice Generator

Google Apps Script that turns rows of a spreadsheet into PDF invoices, files them in
Drive, writes back the PDF URL/ID, and locks the row.

## Files

- `Code.gs` — the full Apps Script (paste the whole file into your Apps Script editor)

## Version history

### 1.2.0 (current)
Bug-fix release. Drop-in replacement for 1.1.0.

**Fixes the errors you were seeing:**
- `API call to drive files delete failed with error: Empty response`
- `Service error: Drive`

**Changes:**

1. **`Drive.Files.remove()` → `DriveApp.setTrashed(true)`**
   The advanced Drive Service v3 returns an empty body for delete, which Apps
   Script raises as an error even though the file is gone. Switched to
   `setTrashed(true)` which is recoverable and doesn't have this bug.

2. **Added `retryOnDriveError()` wrapper** around `makeCopy` and `convertPDF`.
   Transient `Service error: Drive` failures retry up to 3× with linear backoff
   (2s, 4s) before giving up. Configurable via `SETTINGS.retry`.

3. **Hardened cleanup.** `invoiceId` is now declared before the `try` block, so
   the catch-block's cleanup call won't throw a `ReferenceError` when
   `makeCopy` fails on the very first line. A new `safeTrash()` helper
   swallows cleanup errors so the original error is the one that surfaces.

4. **`convertPDF()` rewritten.** Removed the deprecated `Drive.Files.update(meta, id, blob)`
   signature, which is the other place "Empty response" was hitting. Now uses
   plain `DriveApp` calls. When an existing PDF ID is provided, the old PDF is
   trashed and a new one is created (old behaviour: in-place overwrite, but the
   v2 API for that no longer works reliably).

5. **Per-row error handling in `sendInvoice()`.** A failure on row 5 used to
   silently abort the loop. Now each row is wrapped, failures are collected,
   and the script keeps processing. A summary dialog at the end lists which
   rows failed.

6. **`createSystem()` typo fixes.** The old code referenced
   `SETTINGS.col.systemCreated`, `SETTINGS.col.count`, `SETTINGS.col.templateId`,
   and `SETTINGS.col.folderId` — none of which existed in the SETTINGS object,
   meaning the function silently wrote to undefined ranges. Replaced with the
   real keys (`SystemCreated`, `Count`, `Original_ID`, `Original_Folder_ID`).

7. **Editors moved to config.** The hardcoded `anis@sli-eg.com` and
   `george@sli-eg.com` are now in `SETTINGS.editors`, so changing them no
   longer requires hunting through `protectRow()`.

8. **Misc.** Tightened scoping, removed unreachable comment blocks, replaced
   the `value =` global leak in `createHyperlinkString` with a normal return.

### 1.1.0
Auto configuration.

### 1.0.0
Initial release.

## Tuning

If you still see occasional Drive errors on heavy batches, increase the retries:

```javascript
SETTINGS.retry = {
  maxAttempts: 5,
  baseDelayMs: 3000
};
```

## Quick sanity check after pasting

1. Reload the spreadsheet so the new `onOpen()` runs.
2. Run `testUI` from the Apps Script editor — should show "Hello! The UI is connected."
3. Run `sendInvoice` on a small batch (1–2 rows) and watch **Executions** in
   the Apps Script editor. Every step is logged.
