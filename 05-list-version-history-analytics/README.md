# 05 · List Version History Analytics

SharePoint keeps every change to a list item as a version, but you cannot report on it directly. These scripts flatten the version history of a job-tracking list into a CSV that is loaded into a Dataverse table and used for performance analytics (cycle time, status changes, who changed what and when).

| Script | What it does |
|---|---|
| `Backfill-ListVersionLog.ps1` | One-time full export of all versions for the items in scope (assigned this year **or** not finished). Uses a **checkpoint file** and batched flushes so it can resume after interruption, retries failed items, and stops safely if the output CSV goes missing |
| `Refresh-ListVersionLog.ps1` | Incremental run: uses a **watermark** (last run time) plus a ledger of exported `ItemID|VersionID` to append only new versions. The watermark only moves forward when the run succeeds |

## Techniques

- CAML query with `RecursiveAll` + paging
- `Get-PnPProperty -Property Versions` per item
- Field-map hashtable (internal name → output column)
- `HashSet` checkpoint/dedupe

Output goes to `data/` (git-ignored). Auth: certificate app-only.
