# Backend changes

Two things landed together here: the **new flat / EV count fields**, and a fix
for a **column-misalignment bug** found while wiring them in.

---

## The bug, first

`handleSubmission` built its master row as a fixed 34-value array. The sheet's
last two columns are `latitude` and `longitude`, but the array's last two values
were `society_fee_11kw` and `agreement_tenure`:

| Sheet column 33/34 | What `appendRow` actually wrote |
|---|---|
| `latitude` | `parseFloat(data.society_fee_11kw) \|\| 0` |
| `longitude` | `data.agreement_tenure` |

**Every society from KZ-RWA-185 onward** (43–44 rows) has junk in those two
columns as a result — a `0` in latitude, usually blank in longitude.
KZ-RWA-185 and 186 carry `37.0625, -95.677068`, which is Google's
"geocoding failed" fallback, not a real location.

`handleUpdate` was never affected: it resolves columns by header name via
`colMap`, so the edit path has always been correctly aligned.

**The fix:** `handleSubmission` now builds its row by header name too. Anything
not in its value map — `latitude`, `longitude` — is written blank rather than
clobbered, and any field with no matching column is logged to `Debug_Log`
instead of vanishing. Reordering or extending the sheet can no longer misalign it.

---

## One file, nothing to configure

Everything is in **`Code.gs`**. Paste it over your existing script — that is the
whole install. `SPREADSHEET_ID`, `PARENT_FOLDER_ID` and `NOTIFY_EMAIL` are your
own values, copied verbatim from the file you sent.

The migration, the coordinate cleanup and the CSV import sit in a fenced
**ONE-TIME MAINTENANCE** section at the bottom. Nothing in the web app calls
into it — verified — so once those runs are done, select from the banner to the
end of the file and delete it. The only value you ever edit is
`BULK_CSV_FILE_ID`, and only when you reach step 4.

## Deploy order

Order matters. **Do step 1 first** — it is what stops the corruption.

1. **Paste `Code.gs` over the existing script and re-deploy the web app.**
   Safe on its own: with no new columns yet, the five extra fields are simply
   logged to `Debug_Log` as unmapped. Latitude/longitude stop being overwritten
   immediately.
2. **Run `addMissingColumns()`** from the editor. Appends five columns.
3. **Run `previewCoordCleanup()`, read the log, then `runCoordCleanup()`.**
   Blanks the junk already in latitude/longitude so a geocoding pass can refill
   them. Genuine coordinates (validated against India's bounding box) are kept.
4. **Later, when your CSV is ready** — set `BULK_CSV_FILE_ID`, then
   `previewBulkImport()` and `runBulkImport()`.

Running step 2 before step 1 would not help: the old positional `appendRow`
would keep writing into columns 33/34 regardless.

To run any of these: open the Apps Script editor, pick the function from the
dropdown beside **Run**, press Run, and read the log. Every destructive step has
a `preview…` twin that writes nothing.

## The new columns

`addMissingColumns()` appends five headers, and is safe to re-run:

| Column | Why |
|---|---|
| `flat_count`, `ev_2w_count`, `ev_4w_count` | New — the fields this change adds |
| `society_fee_11kw` | Pre-existing gap. Collected by the form, no column existed |
| `agreement_tenure` | Pre-existing gap. Same |

They are appended at the end rather than slotted next to their logical
neighbours. Now that both read and write paths are header-driven you *can*
reorder them by hand safely — but there is no need to.

### What the two pre-existing gaps cost

- **`agreement_tenure`** is the real loss: free text sales fills in on Step 2,
  discarded on every submission. Not recoverable from the sheet — it has to be
  re-read off the signed agreements.
- **`society_fee_11kw`** has cost nothing. No society in the sheet has any 11 kW
  chargers recorded, so it has never been populated. Adding the column means the
  first 11 kW deployment does not lose its commercial terms.

### UI behaviour for the counts

| Field | Required |
|---|---|
| `flat_count` | Yes on **new** onboarding, optional on edit |
| `ev_2w_count` | Yes on **new** onboarding, optional on edit |
| `ev_4w_count` | Yes on **new** onboarding, optional on edit |

Blank ≠ zero. A genuine `0` is stored and displayed; a blank means "not yet
collected". `handleUpdate` skips blank counts rather than writing an empty
cell, so editing an older society cannot wipe a value the bulk import filled in.

---

## Bulk-load the older RWAs

1. Upload your CSV to Drive.
2. Put its file ID in `BULK_CSV_FILE_ID` at the top of `FlatAndEvCounts.gs`.
   (From `https://drive.google.com/file/d/<FILE_ID>/view`.)
3. Run **`previewBulkImport()`** — writes nothing, logs exactly what would change.
4. Read the log, then run **`runBulkImport()`**.

### What the CSV needs

One identifying column plus at least one count column. Header spelling is
flexible — `Society ID` / `id` / `rwa_id`, `No. of Flats` / `flats` / `units`,
`2W EV` / `Two Wheelers` / `no_of_2w`, and so on.

Matching is by `society_id` first, falling back to `society_name` (case, spacing
and punctuation insensitive). Rows whose name matches more than one society are
reported as ambiguous and skipped rather than guessed at — add a `society_id`
for those.

Blank / `NA` / `-` cells are left alone rather than overwritten with zero, so a
partial CSV is safe. The import is idempotent, and only ever writes the three
count columns.

---

## Also fixed: multi-email truncation (frontend)

`submit()` in `App.jsx` set `payload.rwa_email` to only the **first** address
from the comma-separated box. `handleUpdate` writes that value straight into the
`rwa_email` column — so editing any society with several dashboard logins would
have silently dropped all but the first.

12 societies were exposed, 16 addresses at risk (KZ-RWA-106 alone has 5).
Nothing was lost yet — the truncation only fires on edit. The frontend now sends
the full list.

`User_Access`, which is what actually grants dashboard login, was never affected:
it has always read `data.rwa_emails` and split the whole list.
