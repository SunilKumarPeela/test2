# Point v8.20.6 — Consistent Excel Date and Time

Excel date/time cells are now normalized during import to the stable,
locale-independent `YYYY-MM-DD HH:MM:SS` format. Conversion is restricted to
columns whose headers indicate dates or times, so Employee IDs and other
ordinary numeric values are never reinterpreted as dates.

## Previous v8.20.5 changes

Point now detects the actual table header in generated Excel reports that
contain logos, descriptions, timestamps, or blank rows above the data. It also
ignores formatting-only blank columns after the table, preventing misleading
`Unnamed Column` fields while preserving intentional blank columns inside it.

Visible cells are now explicitly invalidated after viewport data is loaded.
This fixes stale blue selection or duplicate coloring that could remain until
the user clicked the grid when cell text itself had not changed. Repainting is
still limited to the visible viewport, so the large-data performance gains are
preserved. Missing, duplicate, selection, criteria, and change-state brushes
now appear immediately and accurately.

## Previous v8.20.3 changes

Search preparation now resolves names only for cells the user actually edited.
Generated result cells such as Email, State, and Computer Name are no longer
rescanned as possible name inputs on every Search. The pending set is keyed by
logical cell position, deduplicated automatically, cleared with the workspace,
and retained when an ambiguous name needs user correction. Existing unique-name
conversion, ambiguity blocking, exact matching, audit events, and validation
remain unchanged.

## Previous v8.20.2 changes

Windows can send resize messages before Point has finished creating its grid.
The optimized resize path now runs only after all grid controls are fully
initialized, preventing the installed application from closing during startup.
The viewport performance improvements remain enabled after initialization.

## Previous v8.20.1 changes

Point's Win32 grid now uses viewport-scoped loading, committing, and repainting.
Scrolling, resizing, selection, and workspace refreshes update only the rows and
columns currently visible instead of touching thousands of hidden controls.
Visible edits are committed before horizontal navigation, and newly revealed
cells are loaded immediately, preserving data integrity. Redundant control-text
updates are skipped. Validation, exact-match behavior, unmatched-row protection,
transformation previews, audit events, and local-only security controls remain
unchanged.

## Previous v8.20.0 changes

Point now provides a live **Transform Lens**: after selecting a data cell and
typing a pattern, the status line previews the results of all three operations
before data is changed. The toolbar uses the compact `<<`, `✂`, and `>>`
symbols. Keyboard users can press `Ctrl+,` to cut the left side, `Ctrl+.` to
cut the right side, `Ctrl+Z` to undo normal cell editing, and `Ctrl+U` to undo
Point's most recently learned single-cell or field transformation. These
features are local, deterministic, previewable, auditable, and do not send
workspace data to an external service.

## Previous v8.19.2 changes

The workspace toolbar now includes a compact **text / delimiter** box between
**Find Next** and the `<<`, `Cut`, `>>` buttons. Select a target data cell,
type a delimiter or text such as `/`, `global/`, or `@company.com`, then choose
an action. `<<` keeps the text on the left, `Cut` removes the entered pattern,
and `>>` keeps the text on the right. Highlighting text directly inside a cell
continues to work when the pattern box is empty. Learned rules can still be
previewed and applied to selected rows or the full field from the context menu.

## Previous v8.19.1 changes

The transformation controls are three compact, persistent buttons beside the
pattern box: `<<` keeps the left side, `Cut` removes the entered or highlighted
text, and `>>` keeps the right side. They remain visible and available for easy
discovery; select a data cell before applying an action.

All earlier learned-transformation actions remain available. The floating
action and highlighted-text context menu now also provide **Skip Left — Keep
Right** and **Skip Right — Keep Left**. These rules treat the selected text as
a delimiter, so every row may contain different words on either side. For
example, selecting `/` can transform `global/sunil`, `local/grace`, and
`admin/john` into `sunil`, `grace`, and `john` with one field-level action.

Directional rules support any 1–32 character delimiter, preserve unmatched
rows, show before/after previews and affected counts, work with selected-row or
full-field scope, remain undoable, and use the fast per-field row index.

The cell context menu continues to provide Remove, Split into Part Columns,
Skip Left, and Skip Right as keyboard/right-click alternatives.

Learned transformations now use a lightweight per-field row index. Preview,
field-wide apply, selected-row apply, and undo operate only on populated rows
in the chosen field instead of scanning unrelated workspace cells. Sequential
imports build the index without an extra sort, while deletions invalidate and
rebuild only when necessary.

Point can now learn a cleanup rule directly from highlighted text inside a
workspace cell. Highlight an unwanted prefix, suffix, or middle token (for
example `global/` in `global/speela`) and click the small scissors action that
appears at the cell's top-right. Point removes the selected token from that
cell and remembers its position-aware rule.

Right-click another cell in the same field to preview the learned rule, apply
it to every populated row, apply it only to a selected row range, or undo the
last application. Preview reports affected and unmatched counts with examples;
unmatched values are never changed. Matching is case-insensitive, changes stay
local, and audit records store counts and token length rather than sensitive
selected text.

If a user pastes a delimited value such as `global/speela` into the heading row
and leaves the column below empty, Split Column now recognizes the mistake,
moves the pasted text into the first data row, and produces `Original Value`,
`Original Value Part 1`, and `Original Value Part 2` instead of empty output.

Point now includes a **Data Tools** menu for reversible, case-insensitive row
filters and safe delimiter-based column splitting. Select a cell in the target
column, then choose Equals, Contains, Starts With, Ends With, Blank, or
Non-Blank. Use Clear Filter to restore the original rows.

The Split Column menu supports `/`, `\\`, `|`, comma, semicolon, colon, space,
and a custom delimiter typed in the Find box. Point previews the first three
rows, keeps the original column, and appends clearly named `Part 1`, `Part 2`,
and additional part columns when needed. Operations remain local and are added
to the audit log.

Username, email, Employee ID, and other exact identity lookups now keep the
original matching row as the identity authority. Related reports may supply
missing requested fields, but they cannot replace that person's identity with
another Employee ID encountered through group or membership relationships.

Universal lookup now treats `First Last`, `Last, First`, and `Last First` as
the same person-name token set. A uniquely matched name can return the stored
Display Name and any related requested fields. Point refuses to guess when a
name matches more than one person.

Universal lookup now preserves an unmatched ID exactly as entered, colors its
entire result row red, and writes `NOT FOUND` in every related output cell.
This makes failed lookups unmistakable even deep inside very large production
workspaces.

Universal mode now tries the exact field where a value was entered before
falling back to cross-field discovery. Adding another date, ID, or similarly
shaped output column no longer broadens a valid existing lookup unexpectedly,
while values pasted under a different heading can still resolve universally.

Point's application-owned native XLSX reader now runs reliably even when the
installer carries a Windows Internet-zone marker. The script path must be a
regular non-reparse file inside the installed Point tree, stale empty Excel
caches are invalidated, and Imported Files shows `Indexing workbook...` while
an asynchronous refresh is still processing.

Point Fetcher now securely publishes every completed, validated download from
its staging area into the main Point Inbox and immediately notifies an open
Point workspace to refresh Imported Files. Only the schedule explicitly run by
the user is published; unrelated older files in Staging remain untouched.

Point now builds exact-value indexes with constant-time canonical-column
resolution instead of reverse-scanning the header map for every data cell.
Large result sets also reserve workspace storage once before population,
reducing hash-table reallocations and making import, refresh, lookup, and
related-field expansion substantially smoother on large files.

Point now excludes Excel owner/lock files whose names begin with `~$` from
uploads, archive scans, imported-file lists, and workspace refreshes. Removal
errors also identify locked files and tell the user to close Excel.

Point no longer hides a match in a non-unique field merely because another
imported file contains a unique column with the same heading. Unique-column
preference is applied only when the entered value exists in that unique
column. For example, `Shanice` in a 5,000-row `First Name` field is found even
when a separate 10-row workbook also contains a unique `First Name` field.

# Point v8.16.7 — Blank Excel Header Recovery

Point now imports worksheets whose first row contains one or more blank header
cells. A blank heading is assigned a stable name such as `Unnamed Column 1`
instead of rejecting the complete sheet. This supports common exported Excel
files where column A contains row numbers but cell A1 is blank, including the
official File Examples XLSX samples.

# Point v8.16.6 — Automatic Excel Cache Repair

Point now versions its Excel worksheet cache and automatically discards empty
results created by an older importer. After upgrading, clicking **Refresh**
forces browser-downloaded workbooks through the current native XLSX reader.
Headerless cache files are never reused as successful imports.

# Point v8.16.5 — Native XLSX Import

Point now reads standard `.xlsx` and `.xlsm` worksheets directly from their
validated Office ZIP/XML package. This fixes browser-downloaded files that open
in Excel Protected View but previously appeared as `0 rows | 0 fields` in
Point. The native reader does not launch Excel, execute formulas, run macros,
follow links, or modify the protected Inbox original. Excel automation remains
as a compatibility fallback, and legacy `.xls` files continue to use Excel.

# Point v8.16.4 — Browser-Downloaded Excel Import

Point now imports browser-downloaded workbooks without requiring the user to
manually remove Windows' Internet-zone marker. The protected Inbox original is
left untouched. For extraction, Point creates a short-lived byte-only copy in
its private Excel cache, disables macros, events, external-link updates, and
alerts, tries Excel normal-load mode first, and uses repair mode only as a
fallback. The temporary copy is deleted after import or failure.

# Point v8.16.3 — Live Inbox Update

After Point Browser Fetcher validates and publishes a file, it sends a local
Windows notification to a running Point workspace. Point automatically
refreshes Imported Files and the field index while preserving existing
workspace headings and entered rows. If another refresh is already running,
the live update waits and retries instead of interrupting it.

# Point v8.16.2 — Automatic Browser Handoff

Manual direct-download attempts that require JavaScript or an interactive
website are handed to Point Browser Fetcher automatically with the original
HTTPS URL. The browser window opens on the correct page; the user completes
any login, MFA, CAPTCHA, or download interaction. Background schedules never
launch interactive windows.

# Point v8.16.1 — Verified Excel Downloads

Both Point Fetcher applications now validate the requested file type rather
than a temporary `.download` filename. XLSX/XLSM files must contain the real
Office workbook package structure before they can enter Staging or Inbox.
HTML redirect pages, arbitrary ZIP files, encrypted ZIP entries, unsafe paths,
and renamed non-Excel content are rejected and deleted.

When a manual **Download Now** request returns a JavaScript/login HTML page,
Point Fetcher now opens Point Browser Fetcher automatically and passes the
original HTTPS URL. Scheduled background jobs remain non-interactive and do
not unexpectedly open browser windows.

# Point v8.16.0 — Secure Browser-Assisted Downloads

## Point Browser Fetcher

`PointBrowserFetcher.exe` handles websites that require JavaScript redirects,
interactive sign-in, MFA, or browser cookies. It uses the installed Microsoft
Edge WebView2 Evergreen Runtime and never asks Point to collect website
passwords.

- HTTPS navigation only.
- Browser password saving and general autofill disabled.
- Developer tools and default context menus disabled.
- Only CSV, XLS, XLSX, and XLSM downloads are intercepted.
- HTML, login pages, access-denied pages, binary CSVs, false file extensions,
  empty files, and files over 2 GiB are rejected and deleted.
- Validated downloads are copied into Point's Inbox only after completion.
- Partial and interrupted downloads are removed.
- The isolated browser profile is stored per Windows user under the Point
  installation data directory.

Direct URLs and unattended API schedules should continue to use
`PointFetcher.exe`. Browser downloads may require the user to complete MFA or
CAPTCHA and are intentionally interactive.

# Point v8.15.6 — Permanent Application Branding

## Permanent Point icon (v8.15.6)

- A multi-resolution 16–256 pixel `point.ico` is embedded into both
  `Point.exe` and `PointFetcher.exe`.
- The main workspace, Fetcher, Chart, and Field Synonym Manager window classes
  use the same large and title-bar icons.
- The Inno Setup installer uses the Point icon, and installed shortcuts inherit
  the embedded executable icons.
- Release and debug builds now stop with an error when `point.ico` is missing,
  preventing GitHub from silently publishing an iconless build.

# Point v8.15.5 — Open Imported Files

## Open imported file (v8.15.5)

- Right-click an imported file in **Input Files** and select **Open File** to
  open the managed Inbox copy with its normal Windows application, typically
  Microsoft Excel.
- **View Columns** remains available in the same context menu, and double-click
  continues to show the file's column summary.
- Point rejects missing files, directories, reparse points, and unsupported
  file types before asking Windows to open the file.

# Point v8.15.4 — Reliable Incremental Row Detection

## Reliable incremental lookup fix (v8.15.4)

- Search now detects new or changed first-column values by comparing the grid
  with the last populated result, even if Windows did not deliver an edit
  notification for that cell.
- Earlier completed rows remain unchanged and each newly detected row is
  resolved into its related fields.

# Point v8.15.3 — Incremental Universal Row Lookup

## Incremental row lookup fix (v8.15.3)

- Previously populated Universal rows remain unchanged.
- Typing or pasting a new lookup into the next row and clicking **Search**
  fills that row with its related values without removing earlier rows.
- Multiple newly pasted rows are resolved as one incremental batch.

# Point v8.15.2 — Universal Search Result Replacement

## Universal search fix (v8.15.2)

- Entering or pasting a new Universal lookup after results are displayed now
  replaces the previous result rows instead of appending below them.
- A pasted batch is still processed together, so multiple newly pasted IDs
  produce only that batch's matching profiles.

# Point v8.15.1 — Synonym Refresh Workspace Preservation

## Synonym refresh fix (v8.15.1)

- **Save and Refresh** in Field Synonym Manager now preserves all workspace
  headings and entered cell values while rebuilding the field index.
- Normal input-file refreshes still clear the grid so stale source results are
  not retained after files are added, removed, or replaced.

# Point v8.15.0 — Secure Point Fetcher Companion

## Point Fetcher (v8.15.0)

The installer now includes a separate `PointFetcher.exe`. It is the only Point
component with outbound network behavior; the main `Point.exe` remains a local
workspace. Fetcher supports:

- Public HTTPS downloads
- HTTP Basic username/password authentication over HTTPS
- Bearer-token APIs
- API keys with a configurable safe header name
- Recurring 1, 2, 3, 4, 6, 8, 12, or 24-hour downloads
- A manual **Download Now** test
- Windows Credential Manager storage for every password, token, and API key
- A 2 GiB safety limit and CSV/XLS/XLSX/XLSM content validation
- Temporary downloads followed by atomic replacement of the last good file

Fetcher writes validated files to `Fetcher\Staging`. Select those staging
files in Point's Auto Schedule tab to push them into Inbox. Point's Auto
Schedule includes an **Open Point Fetcher** button.

Only HTTPS is accepted. Interactive form-login, CAPTCHA, MFA, and browser-only
SSO sites require an official API or a dedicated connector and are not treated
as generic username/password downloads.

## Automatic Imports and Collapsible Archive (v8.14.0)

## Auto Schedule tab (v8.14.0)

The new **Auto Schedule** tab lets an authorized local user select one or more
CSV/Excel source files and schedule them every 1, 2, 3, 4, 6, 8, 12, or 24
hours. While Point is open, due files replace their corresponding Inbox copies
and are indexed in one refresh batch. Schedules are stored in a current-user
protected configuration file and restored on startup.

Schedules support multi-selection removal and **Run Selected Now** for testing.
Missing or unreadable source files do not overwrite the last good Inbox copy.
Automatic imports are recorded in the Point audit log.

## Collapsible Available Files block (v8.14.0)

The small arrow beside **Available Files** collapses or expands the archive.
When collapsed, Imported Files automatically receives the released vertical
space.

## Inbox-Synchronized Synonym Fields (v8.13.14)

## Live Inbox headings in Synonym Manager (v8.13.14)

Canonical-field and synonym suggestions now come from the real headers in all
currently indexed Inbox datasets, including every worksheet imported from an
Excel workbook. Saved canonical fields and synonyms remain available. The old
built-in list is used only when the Inbox index is empty.

Prefix matches are ranked before other substring matches. The manager also
blocks opening while Refresh is running so it cannot display a partially
updated set of headings.

## Folder and Drive File Archive (v8.13.13)

## Available Files archive (v8.13.13)

The Input Files tab now includes an **Available Files** block. Select any
folder or drive and Point scans it and its subfolders for `.csv`, `.xlsx`,
`.xls`, and `.xlsm` files on a background thread. Inaccessible folders are
skipped and the first 10,000 matching files are shown.

Each available file has an `↑` action at the right. Clicking it sends that file
through the existing secure Inbox import flow. Files already present in the
Inbox show a green check mark. Scanning never imports a file by itself.

## Reliable Field Sample Preview (v8.13.12)

## Reliable hover panel (v8.13.12)

The field sample preview now uses a dedicated non-activating Win32 panel
instead of the tracking-tooltip API. This works reliably over the owner-drawn
Input Files list and automatically positions the panel inside the current
monitor.

## Input Field Sample Preview (v8.13.11)

## Hover sample values (v8.13.11)

Hover over any visible field badge on the Input Files tab to see up to five
distinct, non-empty sample values from that exact file and field. Previews are
cached after the first hover, long values are shortened, and fields classified
as highly sensitive do not expose sample data in the tooltip.

Common-field blue highlighting and the `+N` remaining-column panel continue to
work unchanged.

## Responsive Workspace Search (v8.13.10)

## Responsive loading for searches (v8.13.10)

Search preparation, Universal lookup, and workspace result population now
service the Windows message queue at bounded intervals. The progress bar keeps
moving and the window remains paintable and movable during large searches.

## Loading bar for searches (v8.13.9)

Workspace searches now display the existing bottom progress bar while work is
running. Universal mode reports **Resolving N of Total lookup values** with a
determinate percentage, followed by **Populating workspace results**. Updates
are throttled to percentage changes to avoid slowing large pasted batches. The
bar is hidden after success and on every blocked/error path.

## Catalog enrichment through any field (v8.13.8)

Point can now traverse any equivalent field across files when at least one side
is unique and the existing value-overlap threshold passes. This supports Group
ID to Group Owner, Request Code to Approver, Department ID to Department Owner,
and similar catalog enrichment patterns. User-configured synonyms can also use
contextual fields. Many-to-many relationships remain blocked to prevent unsafe
cross-products.

## Stable side-panel rendering (v8.13.7)

The remaining-columns display now uses a titled popup container with a
dedicated scrollable list instead of drawing a bare popup list over the input
rows. Closing the panel explicitly repaints the owner-drawn file list, avoiding
the white gaps and partially erased field badges seen in v8.13.6.

## Click-to-view hidden columns (v8.13.6)

The **+N** badge is now clickable. Point opens a compact, scrollable popup
beside the badge containing only the remaining hidden columns. Common fields
are marked with a blue-dot symbol. The popup stays inside the current monitor
and closes when the user clicks elsewhere, changes tabs, refreshes the list,
or presses Escape.

## GitHub/MSVC build correction (v8.13.5)

The Input Files UI now performs synonym-aware canonical matching through its
own validated resolver instead of accessing an Engine-private helper. Badge
width calculation also uses explicit integer types for current MSVC releases.

## Common-field intelligence and file inspection (v8.13.4)

- Common fields appearing in at least two source files are highlighted blue.
- Unique fields are gray, and synonym mappings participate in matching.
- Each file shows its imported row, field, and worksheet counts.
- The Input Files tab includes live search by filename or field name.
- Double-clicking a file opens its complete column list and marks common fields.
- Common-field counts are cached during refresh to keep repaint and resizing fast.

## Imported-column preview (v8.13.3)

Each file in the Input Files tab now displays its imported column headings as
compact badges beside the filename. Badges that do not fit are summarized as
**+N**. Excel worksheet headings are combined under their source workbook,
duplicate headings are suppressed, and unreadable or not-yet-refreshed files
show **Headers unavailable**. The completed refresh bar also remains hidden
when returning to Workspace instead of appearing as a solid blue strip.

## Remove several input files at once (v8.13.2)

The Input Files list now supports standard Windows multi-selection. Use
**Ctrl+click** to choose separate files or **Shift+click** to choose a range,
then select **Remove Selected**. Point asks once, removes only the selected
copies from its local Inbox, reports any failures, and refreshes the workspace
once after the batch finishes.

## Safe removal from Input Files (v8.13.1)

The Input Files tab now includes **Remove Selected**. Point confirms the exact
selected Inbox filename, removes only Point's local copy, refreshes the list,
and rebuilds the active data indexes. The original source file outside Point
is never deleted. Removal is blocked while a background refresh is running.

## Drag-and-drop input workflow (v8.13)

The main window now provides **Workspace** and **Input Files** tabs. Users can
drag CSV/XLSX/XLS/XLSM reports into Point or select multiple files with
**Upload Files**. Point validates the extensions, asks before replacement,
copies accepted reports into its protected local Inbox, lists the available
files, and starts the existing background refresh automatically. Audit entries
record counts only, without source paths or filenames.

Point v8.13 preserves the v8.12 one-click per-user installer, static runtime,
version resources, upgrade-safe data folders, and installer signing support.

## Installable Windows software (v8.12)

Point can now be distributed as one `Point-v8.12-Setup.exe`. End users only
double-click Setup; they do not need Command Prompt, Visual Studio, source code,
or the Visual C++ runtime. Setup provides:

- a modern installation wizard;
- Start menu, optional desktop, Inbox, and Uninstall shortcuts;
- optional sample reports;
- launch-after-install;
- Windows Apps & Features uninstall support;
- in-place upgrades that preserve Inbox, Workspace, Exports, Logs, synonym
  mappings, saved views, and the existing security policy;
- a per-user install under Local AppData, avoiding administrator requirements;
- a statically linked release executable (`/MT`) with hardened linker flags;
- executable version metadata for Point 8.12.

To produce the distributable installer on the release/build computer, install
Visual Studio Build Tools 2022 and Inno Setup 6, then double-click
`build_installer.bat`. The finished file is created at:

```text
build-installer\Point-v8.12-Setup.exe
```

Only the person producing the release needs those build tools. Every recipient
uses the resulting Setup file directly. The default installer policy is
single-user and owner-isolated. Company administrators can replace it with the
included Windows-group-enforced production policy during managed deployment.

## Point v8.11 — Background Refresh and Data Intelligence

## Production refresh and diagnostics (v8.11)

- Refresh now runs on a worker thread, keeping the window responsive.
- The Refresh button becomes Cancel; cancellation preserves the active data.
- File-level progress reports the current workbook or CSV and whether Point is
  importing, indexing, or reusing it.
- Unchanged CSV datasets reuse their existing in-memory parsed indexes; changed
  files alone are reparsed before relationships are safely rebuilt.
- **Relationship and Conflict Diagnostics** reports join fields, confidence,
  one-to-one/one-to-many cardinality, and inconsistent identity profile values.
- **Explain Selected Result and Lineage** shows the search method, contributing
  reports, and the selected row's displayed values.

During refresh, query/export controls are temporarily disabled to prevent a
partially rebuilt dataset from being used. The previous active dataset remains
intact if refresh is cancelled or fails.

## Point v8.10.13 — Fast Refresh with Progress

## Cached refresh and loading bar (v8.10.13)

Refresh now reuses converted Excel worksheets when the workbook names, sizes,
and modified times have not changed. This avoids reopening Excel and rebuilding
the worksheet cache on every unchanged refresh. A progress bar and stage text
show source checking, Excel import/cache reuse, indexing, and workspace setup;
the Refresh button is temporarily disabled to prevent overlapping refreshes.

## Point v8.10.12 — Synonym Field Auto-Complete

## Field and synonym suggestions (v8.10.12)

The Field Synonym Manager now offers live, filtered suggestions while typing.
Yellow canonical-field cells suggest common Point identity and output fields;
synonym rows suggest the actual raw column headings found in imported reports.
Use the mouse, or press Down and Enter, to select a suggestion. Empty cells do
not open a distracting list, and the existing red not-found and blue duplicate
validation remains active.

This makes mappings such as `Employee ID` to `eid`, `employee_number`, or the
exact spelling used by a loaded company report faster and less error-prone.

## Point v8.10.11 — User-Grouped Duplicate Blues

## One blue shade per user (v8.10.11)

Duplicate colors now alternate by user group instead of by physical row. Every
one-to-many row belonging to the same employee receives the same shade. The
next distinct employee group receives the other shade, followed by light,
dark, light, and so on for later users.

Point chooses the grouping identity from the displayed headings in this order:
Employee ID, User ID, Username, Email, User Principal Name, Display Name, then
Full Name. If none is displayed, it falls back to the first requested output.
For example, two computers for `E000002` are both light blue, while two
computers for `E000008` are both dark blue.

## Light and dark duplicate rows (v8.10.10)

Duplicate Universal results now alternate between light blue and darker blue.
This makes adjacent one-to-many results and multiple inputs resolving to the
same employee easier to follow across wide grids. Both shades retain dark blue
text for readability. Missing values remain red, ordinary rows remain white,
and selected cells keep their existing selection color.

The row-color lookup also now uses the correct logical row directly while the
workspace is vertically scrolled, so duplicate and missing colors remain
attached to the correct records beyond the first visible page.

## Search again after changing headings (v8.10.9)

Point now stores Universal lookup inputs separately from displayed result rows.
After adding, removing, or reordering requested headings, Search safely reruns
the original inputs under the new schema. For example, a user can first look up
Employee ID, Email, and Computer Name, then add Ticket Status and Account
Status and select Search again without retyping the Employee ID.

Selecting Search again with unchanged headings also refreshes the same input
history. Newly typed or pasted first-column values are added to that history.
Displayed Employee IDs, emails, computer names, and other output cells are
never treated as lookup inputs, so the former “current rows are Universal
results” blocking message is no longer needed.

## Same employee through different inputs (v8.10.8)

Point now preserves one Universal result row for every entered lookup. If
`E000011` is already displayed and the next input is that employee's email,
both inputs resolve to `E000011` and both result rows remain visible. Every
repeated resolved profile is colored blue to show that multiple inputs reached
the same employee.

Repeated inputs in one pasted batch are also retained instead of being
silently deduplicated. This restores the original duplicate-visibility rule:
blue communicates duplication; red continues to identify values that were not
found.

## Removed-heading and duplicate-result correction (v8.10.7)

Additive Universal results are now preserved only when the requested headings
are unchanged. Removing the Email heading removes previously displayed email
values on the next lookup instead of retaining them beneath a blank heading.
Adding, removing, reordering, or renaming any requested heading starts a clean
result schema using only the pending lookup rows.

Accumulated Universal results are also deduplicated across searches. Entering
the same Employee ID repeatedly keeps one resolved employee profile rather
than producing repeated Employee ID rows with empty companion cells.

## Preserve prior Universal results (v8.10.6)

Entering another lookup value beneath existing Universal results no longer
clears the rows above it. Point tracks only the newly typed or pasted
first-column rows as pending inputs, searches those values, and merges their
resolved profiles into the existing result grid.

Typing into an existing result row replaces that specific row's result on the
next Search. Typing into the next blank row appends the new result. Multi-row
paste appends a batch, while Clear Results and mode changes still reset the
workspace explicitly. Existing red missing-result and blue one-to-many colors
are retained and reindexed when results are merged.

## Employee ID versus Manager Employee ID correction (v8.10.5)

Universal Lookup now gives priority to exact matches found in the requested
output headings. If `E000002` exists both as one person's Employee ID and as
many employees' Manager Employee ID, requesting Employee ID, Username, and
Email returns only the employee whose Employee ID is `E000002`.

Other field types are searched only when the value is not found in a requested
output field. Therefore cross-type operations continue to work, including
Computer Name to Username, Email to Display Name, and Asset Tag to Employee ID.
The regression test includes a direct report whose Manager Employee ID equals
the entered Employee ID and verifies that only the employee profile is returned.

## Single-value Universal lookup correction (v8.10.4)

Universal result rows can no longer be reused accidentally as lookup inputs.
After results are displayed, typing or pasting a new value into the first
column starts a fresh lookup batch and clears the prior result rows. Pressing
Search without entering a new value is blocked with a clear instruction
instead of looking up every Employee ID already displayed.

A regression test verifies that entering one Employee ID with Employee ID,
Username, and Email headings returns exactly one matching employee profile.

## Resize performance correction (v8.10.3)

Maximize, restore, and window resizing now update only the columns and rows
that are actually visible. Thousands of permanently hidden edit controls are
no longer repositioned or redrawn on every Windows resize notification.
Off-screen columns are created hidden and remain untouched until they enter
the viewport. Visible cells are repositioned without an immediate repaint for
each individual control, followed by one combined grid redraw. The responsive
32-row auto-fill behavior is preserved.

## Responsive Excel-style workspace (v8.10.2)

The main data grid now calculates how many rows fit in the current window and
uses the available vertical space automatically. Maximizing or resizing Point
can display up to 32 workspace rows instead of leaving the lower part of the
window blank. The vertical scrollbar, Page Up/Page Down, keyboard navigation,
selection, search focus, and row colors all use the actual visible row count.

When a small amount of height remains after fitting complete rows, Point
slightly adjusts their displayed height so the grid, horizontal scrollbar, and
status area align cleanly with the bottom of the window.

## MSVC build correction (v8.10.1)

The Win32 source now declares the edit-control text setter before the
scrollable synonym-grid loader uses it. This resolves Visual Studio C3861
errors for `set_control_text` at the grid-loading calls.

## Scrollable Field Synonym Manager (v8.10)

The Field Synonym Manager now provides a 64-field by 64-synonym logical grid.
Eight field columns and twelve synonym rows stay visible at once for readability,
while horizontal and vertical scrollbars provide access to the rest. The mouse
wheel scrolls synonym rows; Shift+mouse wheel scrolls fields. Enter continues
down a field's synonym rows and automatically brings off-screen cells into view.

Field numbers and row numbers update with the viewport, so Field 9 through Field
64 can be edited without opening another window. Saving validates and persists
the complete 64 by 64 grid, including cells outside the current view.

## Numeric Employee ID normalization (v8.9.2)

Point now treats leading zeros as formatting for numeric Employee ID values.
For matching and relationship discovery, `39929`, `039929`, and `00039929` are
equivalent. The original report value is preserved in results, so a stored
`00039929` is still displayed and exported as `00039929`.

This normalization is deliberately restricted to canonical Employee ID, User
ID, and generic ID fields, including approved synonyms mapped to those fields.
It is not applied to ages, counts, dates, phone numbers, postal codes, ticket
values, or other numeric business data. If both padded and unpadded forms occur
as separate rows in one identity column, Point treats the key as duplicated
instead of guessing which record is correct.

## Workspace result colors and synonym entry (v8.9.1)

Universal Lookup now preserves an unmatched lookup value in the result grid and
highlights its row red. When one input produces multiple related results, every
returned row for that input is highlighted blue. A single unique result keeps
the normal white result color.

The Field Synonym Manager is faster to complete:

- Enter moves to the next synonym row under the same canonical field.
- Shift+Enter moves upward.
- After the last row, Enter moves to the next field heading.
- A user can enter several synonyms in one cell separated by commas or
  semicolons.
- Configurations containing more than twelve synonyms remain editable by
  scrolling down; commas or semicolons may still place several names in a cell.

## Universal Lookup in Normal mode (v8.9)

Normal mode now treats the heading as the desired output field and each value
under the first heading as an unknown-source lookup value. Point searches that
exact value across every imported field, identifies all matching source field
types, follows only validated relationships, and returns the requested output.

Examples:

| Desired heading | Pasted value | Returned result |
| --- | --- | --- |
| Username | Computer Name | Usernames related to that computer |
| Display Name | Employee ID | Employee display name |
| Email | Username | Related email address |
| Employee ID | Asset Tag | Employee IDs related to that asset |

The pasted value is never silently assigned to a guessed source field. Results
from every valid matching path are combined and deduplicated. If a common value
matches more than 500 source rows, Point blocks the lookup as too broad and
asks for a more specific value. Values that exist but cannot reach the desired
field through a validated relationship return no rows rather than an unsafe
association.

## Full-name identity conversion (v8.8)

When a workspace column is headed **Employee ID**, **Username**, **Email**,
**User ID**, or **User Principal Name**, Point accepts a pasted full name and
converts it to the corresponding identity value. Both `Last, First` and
`First Last` ordering are supported, including multi-token names.

Resolution uses Display Name, Full Name, Name, or combined First Name/Last Name
columns. It can follow Point's validated cross-report relationships when the
name and requested identity are stored in different reports. Conversion occurs
immediately after paste and is also checked when Search is selected.

Point converts only when exactly one distinct target identity is found. If two
people share the same normalized name, conversion is blocked and the user is
asked for a unique Employee ID, Username, or Email. Names and converted values
are not written to the audit log; only the number of successful conversions is
recorded.

### Live field validation colors (v8.7.3)

The synonym grid validates every nonblank entry against the currently loaded
report headers. A field name that is not present is highlighted red. A
normalized name entered more than once anywhere in the mapping grid is
highlighted blue, with duplicate blue taking precedence over missing red.
Blank cells keep their normal grid color and valid canonical headings remain
yellow. Colors update immediately while the user types.

### MSVC build correction (v8.7.2)

The Win32 source now declares the UTF-8 conversion and edit-control text
helpers before the schema grid reader uses them. This resolves Visual Studio
C3861 errors for `control_text` and `narrow`.

## Field Synonym Manager (v8.7)

Point can now connect reports whose identity columns use different headings.
Open **Workspace → Field Synonym Manager**. The manager now uses an
Excel-style grid instead of a text configuration box:

- Enter canonical identity fields in the yellow heading row.
- Enter every approved synonym in the numbered rows underneath its field.
- Eight field columns and twelve synonym rows are visible at once; scroll to
  edit up to 64 fields and 64 synonym rows.

Example grid:

| Employee ID | Username | Computer Name |
| --- | --- | --- |
| EID | Login Name | Hostname |
| Associate Number | SAM Account Name | Device Name |
| Worker ID | User Login | Machine Name |

The protected configuration is stored internally in this equivalent form:

```text
Employee ID = EID, Associate Number, Worker ID
Username = Login Name, SAM Account Name, User Login
Computer Name = Hostname, Device Name, Machine Name
```

Mappings are applied before relationship discovery. Point preserves the
original source columns while presenting the canonical field name in searches
and suggestions. Saving a mapping immediately refreshes every imported report.

Integrity controls:

- Only strong entity identifiers can be canonical fields.
- Classification fields such as Department, State, Status, Location, and Risk
  Level cannot become automatic relationship keys.
- A synonym cannot belong to more than one canonical field.
- A mapping that makes two columns in the same report equivalent is rejected.
- Existing uniqueness and 50% overlap validation still applies before a
  relationship is created.
- Configuration is protected with Windows DPAPI for the signed-in user,
  retained across restarts, excluded from workspace cleanup, and audited.
- Matching row values alone never create a relationship.

The supported canonical identity fields are Employee ID, Username, Email,
Computer ID/Name, Device ID/Name, Asset ID/Tag, Serial Number, User ID, User
Principal Name, and generic ID.

## Pivot removal (v8.6)

Pivot mode has been completely removed from the application, including its
mode-cycle entry, query engine API and implementation, saved-view support,
generated grid state, tests, documentation, and audit labels. Existing saved
views that specify `pivot` are now rejected as invalid instead of silently
running under a different mode.

## Chart field refresh fix (v8.1.1)

After a Chart search, editing a generated chart heading now returns the grid to
Chart setup mode immediately. The newly entered category or filter field is
retained, prior exact-filter values are preserved, stale result rows and the
old chart window are cleared, and the next Search builds an accurate chart from
the newly selected field.

## Multi-field charts (v8.2)

Chart mode now groups by one, two, or three fields. Enter each grouping field
as a heading and leave the first data-row value beneath it blank. To apply an
exact filter, add another heading and enter its required value beneath it.

Example: headings `Location`, `Risk Level`, and `Device Type`, all with blank
values, produce counts for each three-field combination. Adding `Department`
with `Security` beneath it limits that same chart to Security records. Existing
single-field charts continue to work without changes. Clicking a chart segment
drills down with every grouping value and filter, not only the first field.

## Independent multi-chart dashboard (v8.3)

Selecting two to four blank-value chart fields now creates a separate graph for
each field in one dashboard. For example, `Device Type` and `State` produce a
Device Type distribution graph and a State distribution graph; four selected
fields produce four panels. Counts are calculated independently, so the charts
do not accidentally count field combinations. Each panel applies the same
exact filters, uses its own scale and top-category limit, and supports accurate
field-specific drill-down. A single selected field still opens the original
full-size dynamic chart.

## Multiple series inside each graph (v8.4)

Chart mode now supports a shared series dimension in addition to multiple
independent graphs. The value under a selected heading defines its role:

- Leave it blank to create a graph for that field.
- Enter `@series` to use that field as the colored series inside every graph.
- Enter any other value to apply it as an exact filter.

Example: leave `Device Type` blank and enter `@series` under `State`. Point
creates a Device Type graph containing separate colored bars for every State
value. Leave `Device Type`, `Department`, and `Risk Level` blank while marking
`State` as `@series` to create three graphs, each with the same State series.
Point supports one shared series field, one to four graph fields, and any number
of exact filters within the visible grid. Drill-down applies the graph category,
series value, and every exact filter together.

### Windows build correction (v8.4.1)

The Win32 source now includes the standard map header explicitly and converts
the Win32 `RECT.right` `LONG` value to `int` before dashboard width arithmetic.
This resolves MSVC C2275/C2062/C3481 and C2672/C2737 build errors reported by
Visual Studio 2022 Build Tools.

## Dynamic chart dashboard (v8.5)

The chart dashboard now provides an advanced interactive presentation:

- Auto view selects 100%-stacked visualization for multi-series results.
- The View button cycles through Auto, Bar, Column, Pie, and Stacked %.
- Stacked segments use stable series colors and show category totals.
- Panel titles identify every active series value.
- Hover details show graph field, category, series, exact count, and percentage.
- Top 5/10/20/All ranks complete categories by total and retains every series
  segment, preventing partial or misleading category results.
- Live search filters graph fields, categories, and series values together.
- One to four panels continue to resize automatically with the chart window.
- Clicking any grouped or stacked segment applies its exact graph, category,
  series, and filter conditions during record drill-down.

Point is a local, dependency-free Windows data workspace for connecting CSV
and Excel reports. It discovers fields, identifies shared keys, narrows records
with multiple exact conditions, and builds joined result tables without Excel
formulas or a master worksheet.

## Compliance enforcement in v8.0

Point v8.0 adds technical controls that support NIST CSF 2.0, PCI DSS 4.0.1,
and GDPR deployments:

- Windows local/domain group authorization is enforced before the UI opens.
- `Inbox`, `Workspace`, `Exports`, and `Logs` receive protected Windows ACLs.
- Saved workspace views are encrypted with Windows DPAPI and are bound to the
  signed-in Windows user.
- Exports block values that pass payment-card length and Luhn checks.
- Passwords, secrets, tokens, keys, national IDs, PANs, and verification-code
  fields are masked in exports.
- Audit records neutralize control characters and use a chained SHA-256 digest.
- Configurable retention removes expired exports, workspaces, and logs.
- Configuration is fail-closed: missing or invalid security policy prevents
  startup.

These controls make Point a compliance-supporting application. They do not
certify the organization using it. See `COMPLIANCE.md` for the control matrix,
deployment responsibilities, limitations, and evidence checklist.

Version 8.1 retains the exact classic Point v7 display title and existing grid
presentation. Compliance controls operate behind the interface and do not
change the established buttons, fonts, colors, spacing, scrolling, or modes.

## Excel-style grid controls in v7.0

- Mouse wheel and precision-touchpad deltas scroll vertically.
- `Shift + Mouse Wheel` and horizontal touchpad gestures scroll columns.
- Arrow keys navigate cells; Left/Right still move the text cursor while it
  is inside text and move cells when the cursor reaches an edge.
- `Tab` / `Shift+Tab` move right/left.
- `Enter` / `Shift+Enter` move down/up.
- `Ctrl+Arrow`, `Page Up`, `Page Down`, `Home`, `End`, `Ctrl+Home`, and
  `Ctrl+End` provide Excel-style jumps.
- Hold `Shift` with navigation keys to extend a cell range.
- Drag the right edge of a heading to resize that column.
- Double-click a heading edge to auto-fit it to visible values.
- Drag a row’s lower edge to adjust row height; double-click to reset it.
- `Ctrl+C`, `Ctrl+X`, `Ctrl+V`, `Ctrl+A`, `Ctrl+F`, and `Delete` work with
  cells or selected ranges. Multi-cell clipboard data remains compatible with
  Excel.
- Headings remain fixed while the virtual data rows scroll.

These behaviors still use only the visible row controls, preserving Point’s
2,000,000-row virtual grid.

Version 6.3.1 explicitly converts Windows `POINT` coordinates from `LONG` to
`int` before clamping column-width and row-height drag values. This fixes the
MSVC C2672 template-type ambiguity reported by Visual Studio 2022 Build Tools.

## Performance in v7.0

Point builds reusable hash indexes for exact values and safe relationship keys.
The first use of a column builds its index once; later searches reuse it.
Exact lookups and cross-report relationship hops therefore use direct indexed
row lookups instead of rescanning every row in a 100,000-row report.

Relationship discovery now rejects non-identity fields before performing
expensive uniqueness work. Fields such as Status, Department, Location, Risk,
and other classifications no longer trigger unnecessary full-column scans
during Refresh. Retained Change baselines copy report data without duplicating
the runtime index cache.

## Security model

- Runs as a normal desktop user and never requests elevation.
- Performs no network access, telemetry, scripting, macros, or shell execution.
- Opens source reports read-only and never modifies them.
- Imports `.csv`, `.xlsx`, `.xls`, and `.xlsm` files from the local `Inbox`.
- Rejects symbolic links/reparse-point-like filesystem entries.
- Limits file size, rows, columns, header size, and cell size.
- Parses into temporary objects and publishes only complete datasets.
- Uses bounded joins; it never performs an unrestricted cross join.
- Escapes spreadsheet formula prefixes when exporting CSV.
- Writes an append-only local activity log without recording imported values.

## Build on Windows

Requirements: Windows 10/11 and Visual Studio 2022 Build Tools with the
**Desktop development with C++** workload. Those tools are needed to compile
Point but are not runtime dependencies.

Open **x64 Native Tools Command Prompt for VS 2022**, change to this directory,
run the one-time group configuration from an Administrator prompt:

```bat
configure_compliance.bat
```

Sign out and sign in again, then build:

```bat
build.bat
```

The release build treats every compiler warning as an error and enables
Microsoft SDL checks, Control Flow Guard, ASLR, and DEP-compatible linking.
For development diagnostics, run:

```bat
build_debug.bat
```

The debug build additionally enables runtime checks, debug symbols, and
unoptimized execution.

## Stability and recovery in v7.0

- Saved views use the bounded `POINT_VIEW_V3` format.
- Column widths and row height persist with the workspace view.
- Every loaded dimension is range-checked before it changes the UI.
- Views are written to a temporary file and published with an atomic,
  write-through replacement, preventing a partial save from corrupting the
  last valid view.
- Older `POINT_VIEW_V1` and `POINT_VIEW_V2` files remain supported with safe
  default dimensions.
- The grid right-click menu exposes Copy, Cut, Paste, Clear, Select Used Grid,
  and AutoFit through the same validated operations as the keyboard commands.
- Windows coordinate conversions use explicit types, and automated contracts
  guard against recurrence of the MSVC C2672 error.

The standalone executable is created at:

```text
build\Point.exe
```

Copy `Point.exe` anywhere and start it. On first launch it creates:

```text
Inbox\
Workspace\
Exports\
Logs\
```

Copy `point-security.conf` beside `Point.exe`. Point deliberately refuses to
start without this policy. Production administrators should manage the file
through their normal software deployment and change-control process.

Place reports in `Inbox` and select **Refresh**. Every heading and data cell is
a plain editable box with no dropdown arrow. Typing in a heading displays field
suggestions. Typing in a data cell displays matching values for that field.

Point starts in **Normal mode**. Use the mode button to cycle through
**Normal → Narrow → Count → Compare → Analyze → Insights → Chart → Change →
Normal**. Switching modes
preserves user-entered headings where their meaning remains compatible, removes
generated headings, and clears grid values so results from one workflow cannot
be mistaken for another mode's inputs.

### Normal mode

Normal mode preserves the original Point workflow:

- the first heading is the lookup field;
- enter one or more lookup values down the first column;
- the remaining headings are fields to return;
- each lookup can produce one or multiple result rows;
- results replace the entered lookup rows in a packed result table.

For example, headings `Age`, `Username`, and `Computer Name`, with `30` entered
under `Age`, return every matching person and associated computer.

### Narrow mode

In Narrow mode, the heading controls name the requested fields and the yellow
first grid row is the criteria row:

- a filled criteria cell is an exact-match condition;
- every filled condition is combined with `AND`;
- a blank criteria cell means return that field without filtering it;
- matching records are written from the second grid row downward.

For example, headings `Age`, `Department`, `Account Enabled`, `Username`, and
`Computer Name` with criteria `30`, `Information Security`, `Yes`, blank, and
blank mean:

```text
Age = 30
AND Department = Information Security
AND Account Enabled = Yes
```

`Username` and `Computer Name` are returned as output fields. Conditions may
come from different reports when Point can connect those reports through a
bounded relationship. All conditions must apply to the same connected entity.
Non-unique matches expand into separate result rows.

Point v7.0 promotes only identity-safe shared fields such as `Employee ID`,
`Username`, `Email`, `Computer ID`, and `Computer Name` into automatic
cross-report relationships. Classification fields such as `Location ID`,
`Department ID`, `Group ID`, status, age, and state remain searchable but
cannot carry a query into unrelated records. At least one side of an automatic
relationship must also contain unique populated values. Identical displayed
rows reached through different internal paths are returned only once.

Normal and Narrow modes also support one bounded one-to-many relationship. For
example, requesting `Group Name` for one username expands that user's groups
into separate rows:

| Username | Email | Age | Group Name |
|---|---|---:|---|
| speela | sunil@company.com | 30 | All-Employees |
| speela | sunil@company.com | 30 | Security-Admins |
| speela | sunil@company.com | 30 | Security-Readers |

Point permits one such fan-out but blocks a second simultaneous fan-out. For
example, independently expanding both multiple computers and multiple roles
could create an accidental Cartesian product, so Point stops that query instead
of producing misleading combinations.

### Count mode

Count mode supports both a complete grouped distribution and an exact selected
count.

To count every value, type one or more headings and leave every yellow
criterion blank:

| Location Name | Count |
|---|---:|
| Atlanta | 10 |
| Austin | 10 |
| Denver | 10 |
| Fort Collins | 10 |
| Greeley | 10 |

To count only a selected value, enter an exact value in the yellow first row:

| Location Name | Count |
|---|---:|
| Greeley | 2 |

Point automatically adds the `Count` heading and writes the result below the
yellow selection row.

Multiple selected fields are combined with exact `AND` logic:

| Location | Department | Count |
|---|---|---:|
| Greeley | Information Security | 1 |

This means:

```text
Location = Greeley
AND Department = Information Security
```

The workspace provides 256 logical columns with horizontal scrolling. Count
mode supports up to 255 selected fields because one column is reserved for
`Count`. Normal, Narrow, Compare, Analyze, Insights, Chart, and Change
can use up to 256 selected fields where their mode semantics allow it. Count
interprets each yellow cell independently:

- a **filled** yellow cell is an exact filter;
- a **blank** yellow cell is a grouping field.

Leaving every yellow value blank groups by every selected field. Filling every
yellow value counts one exact selection. Mixing filled and blank values filters
first and then groups the remaining fields. For example, `Account Status =
Active`, `Country = USA`, and a blank `State` cell returns one Count row per
state for active U.S. accounts. Counts represent matching source rows, so a
people count is accurate when each row in the selected report represents one
person.

The selected fields may come from different reports. Point v7.0 finds a shared
safe identity such as `Employee ID`, `Username`, or `Computer ID`, aligns the
records by that identity, and counts each connected identity once per valid
combination. For example, `Department Name` from Employees and
`Encryption Status` from Computers can be counted together.

Count results automatically include an object-aware column immediately after
`Count`; no click is required. When a Count is between 1 and 50, Point selects
the identifier that best explains what was counted:

- Device-security fields use `Affected Devices` and prefer Device ID, Computer
  ID, Computer Name, Asset ID, Asset Tag, or Serial Number.
- User, account, MFA, password, group, role, and access fields use
  `Affected Users` and prefer Employee ID or Username.
- Ticket and incident fields use `Affected Tickets` and prefer Ticket ID,
  Incident ID, or Case ID.
- Other fields use `Related Objects` with the best available stable identity.

For example:

| Patch Status | Count | Affected Devices |
|---|---:|---|
| Missing Patches | 6 | Device ID: D1002; D1044; D1308; D1510; D1782; D1901 |

| Device Type | Vulnerability Count | MFA Status | Count | Affected Devices |
|---|---:|---|---:|---|
| Laptop | 11 | Enabled | 2 | Device ID: D1024; D1841 |

For performance, a Count above 50 displays `Count above 50 - narrow filters`
without running a related-data query. Add more exact yellow-cell filters to
narrow that group to 50 or less. Count mode supports up to 254 selected fields
because `Count` and the affected-object column reserve two of the 256 grid
columns.

Point builds affected-object data with one linear scan of the compatible
source report. It creates an in-memory group-to-identity map and fills every
eligible Count row from that map. It does not run a separate relationship
query for each Count row.

For cross-report Counts, Point v7.0 carries the exact affected object IDs
through the Count calculation itself. If Department comes from Employee Data
and Patch Status comes from Device Security, each grouped Count retains the
Device IDs, Computer IDs, or other authoritative identifiers that produced
that number. This removes the earlier `Related data unavailable for
cross-report group` result without adding a second relationship query.

When a Count field appears with conflicting values in several reports, Point
v7.0 can select a uniquely authoritative source by field domain. Encryption,
patching, CrowdStrike, vulnerability, antivirus, endpoint, operating-system,
and device-status fields prefer a Device Security, Endpoint, Computer, or
Security report. Identity/access, employee/HR, ticket/incident, and training
fields use corresponding source-name signals. If two reports remain equally
plausible, Point still blocks the query instead of silently choosing one.

If repeated source evidence makes an occurrence count different from the
number of unique object rows, the inline rows and CSV preserve the returned
objects. A grouped row containing a blank value must be narrowed to an
explicit value before drill-down so Point never presents a broader record set
as exact evidence. Inline detail headings are automatically removed before
the next Count search so they are never mistaken for additional filters.

If one field has several values for an employee, such as `Group Name`, Point
can expand that one field. If two selected fields independently have several
values, Point blocks the count because multiplying those collections would
create an unsafe Cartesian product.

For accuracy, Point also blocks a field that has conflicting meanings in
multiple imported reports. Use **Workspace > Next Source** to select the
intended report, or give the source columns specific names such as
`Training Status` and `Ticket Status`.

### Compare mode

Compare mode is a local cyber-access intelligence view. It compares two
identities across as many selected evidence fields as needed and all compatible reports,
without constructing a Cartesian product. It is designed for group
memberships, application roles, licenses, assigned systems, locations, tags,
or other repeated identity attributes.

Use two to 256 headings:

1. the identity field, such as `Username`;
2. one or more evidence fields, such as `Group Name`, `Role Name`,
   `License Name`, `Application Name`, and `Account Status`.

Enter User A and User B in the first two yellow rows under the identity field:

| Username | Group Name | License Name | Application Name |
|---|---|---|---|
| speela | | | |
| jsmith | | | |

Select **Search**. Point replaces the input grid with an evidence report:

| Field | Result Type | Value | Source Reports | Risk | Recommended Action |
|---|---|---|---|---|---|
| Overall | Similarity | 50% (3 common / 6 unique) | 4 report(s) | High | Review asymmetric privileged access immediately. |
| Group Name | Common | All-Employees | Groups.csv, Reviews.csv | Info | Shared value; retain only while required by both roles. |
| Group Name | Only speela | Security-Admins | Groups.csv | High | Remove or document approved privileged-access exception. |
| License Name | Only jsmith | Power BI Pro | Licenses.csv | Low | Confirm this difference is role-based. |

Matching is exact after case-insensitive field/value normalization. Duplicate
evidence is deduplicated even when the same membership appears in several
reports. Source lineage retains every report that supplied each value.
Similarity is the Jaccard percentage of common selected field/value pairs over
their union. Deterministic risk rules flag asymmetric administrator,
privileged, vault, root, cloud, VPN, and remote-access evidence. They provide
review guidance but do not automatically revoke access or replace an
authorized security review.

If a source is selected through **Workspace > Next Source**, Compare is limited
to that report. In Automatic source mode it safely unions all compatible
reports. Select **Search** again to restore the cached headings and two
identities, then rerun the analysis after changing a field or identity.

### Analyze mode

Analyze mode investigates duplicate and missing records using one, two, or
three headings that you explicitly designate as key fields. For example, enter
`Employee ID`, or enter the composite key `Username` plus `Group Name`, and
select **Search**.

The yellow first row is an optional exact-value filter. Leave it blank to
analyze every value, or enter a value to inspect only that selection. For
example:

| Group Name |
|---|
| Security-Readers |

This analyzes only `Security-Readers`; unrelated repeated groups such as
`All-Employees` are excluded. With multiple key headings, every filled filter
is combined using exact `AND` logic, while blank filter cells remain
unrestricted.

Point replaces the grid headings with a detailed review table:

| Issue Type | Key | Source | Parsed Row | Record Details | Occurrences |
|---|---|---|---:|---|---:|
| Repeated Key Records (Different Details) | Asset Tag=A-100 | Quality.csv | 2 | Asset Tag=A-100; Description=First | 2 |
| Repeated Key Records (Different Details) | Asset Tag=A-100 | Quality.csv | 3 | Asset Tag=A-100; Description=Duplicate | 2 |
| Missing Key | Asset Tag=(blank) | Quality.csv | 4 | Asset Tag=; Description=Missing asset tag | 1 |

Every affected source record is displayed, so the user can see exactly what is
duplicated and where it came from. Issue types are:

- **Missing Key**: at least one selected key component is blank.
- **Exact Duplicate Row**: the selected key repeats and the complete records
  are identical.
- **Repeated Key Records (Different Details)**: the selected key repeats but
  other record values differ. This neutral label is intentional: repetition
  can be a data problem for `Employee ID`, but can be legitimate for a shared
  field such as `Group Name`.
- **Clean**: no duplicate or missing selected keys were found.

`Parsed Row` is the row number in Point's parsed report: the heading is row 1,
so the first data record is row 2. For CSV records containing embedded
newlines, this may differ from a physical text-editor line number.

When the source selector is **Automatic**, Analyze checks every imported report
that contains all selected key fields and can expose duplicate keys spanning
separate files. Select a particular report with **Workspace > Next Source** to
limit the review to that source.

The selected headings define what “duplicate” means. Do not select `Age`,
`Location`, `Department`, or another naturally repeated business attribute
unless repeated values in that field are genuinely an error. Use Count mode
when the goal is simply to measure how often a normal value occurs.

### Point Insight Agent

Insights mode is an offline, deterministic mini analysis agent. It does not
send report contents to an AI service. Instead, it inspects imported data,
selects relevant statistical and data-quality checks, ranks findings, explains
the evidence, and recommends a next action.

Leave every heading blank to profile all fields in the selected source, or type
up to 256 headings to focus the analysis. Select **Search** to produce:

| Priority | Source / Field | Insight | Evidence | Score | Recommended Action |
|---|---|---|---|---:|---|
| High | Users.csv / Employee ID | Possible duplicate identifiers | 3 repeated populated rows | 63 | Run Analyze with this field as the key and review records. |
| Medium | Users.csv / Location | Highly concentrated value | Greeley represents 88% | 88 | Confirm this is expected and not a default. |
| Info | Users.csv / Age | Numeric profile | min=21; median=34; max=67 | 10 | Use the chart to review range and distribution. |

The agent performs:

- missing-value measurement and risk scoring;
- duplicate detection for identifier-like fields;
- repeated-value and cardinality profiling;
- strong relationship-key identification;
- dominant/default-value detection;
- numeric minimum, median, and maximum profiling;
- statistical outlier detection using the 1.5×IQR rule;
- source-aware evidence and recommended next actions;
- priority-first ranking: High, Medium, Low, then Info.

Use **Workspace > Next Source** before running Insights to focus on one report.
With Automatic selected, Point profiles every compatible report.

## Dynamic charts

Chart mode creates a complete category distribution rather than requiring the
user to request each count separately:

1. enter the category heading in the first column, such as `Department`;
2. optionally add filter headings in the remaining columns;
3. enter an exact value under every filter heading in the yellow row;
4. select **Search**.

For example, `Department` as the category plus `Location = Greeley` as a filter
counts every department in Greeley, populates the source table, and opens the
chart automatically. Multiple filters use exact `AND` logic.

After running Count, Analyze, Insights, or another query containing a numeric
result column, the user can also choose **Workspace > Open Dynamic Chart**
manually.

The native chart:

- automatically chooses the rightmost usable numeric result column;
- provides Auto, Bar, Column, and Pie views without rerunning the query;
- provides a live category-search box that redraws as the user types;
- switches between Top 5, Top 10, Top 20, and All;
- sorts interactively from high-to-low or low-to-high;
- uses the mouse wheel to move through Top-N detail levels;
- scales bars dynamically when the window is resized;
- shows the exact label and value in the title when the pointer hovers;
- supports clickable selection highlighting;
- replaces the main grid with the exact source records behind a clicked Chart
  category while retaining the original Chart setup for the next Search;
- displays percentages automatically in Pie view;
- automatically chooses Pie, Column, or Bar in Auto view based on cardinality;
- uses the active grid result, so filters and selected sources carry through;
- requires no browser, network connection, JavaScript, or chart dependency.

For risk charts, run Insights and open the chart to visualize finding scores.

### Change Intelligence

Change mode compares the current reports against a retained in-memory
baseline. Repeated Refresh clicks do not silently replace that baseline.

1. Import the original reports and select
   **Workspace > Set Change Baseline**.
2. Replace or update reports in `Inbox`.
3. Select **Refresh**.
4. Cycle to **Change** mode.
5. Enter one to three stable key headings, such as `Employee ID`.
6. Select **Search**.

The result shows `Added`, `Removed`, and `Modified` records with the key, source
report, exact parsed source row, changed fields, complete Before details, and
complete After details. Every changed result row is displayed with a light-red
background and dark-red text, while `Source / Parsed Row` identifies where to
inspect the record in the imported report.
If a selected key is repeated, Point returns `Ambiguous Key` instead of
guessing which records correspond.

For convenience, the first prior successful Refresh is retained automatically
if no explicit baseline has been set. Use **Set Change Baseline** only when the
current reports have been reviewed and should become the new accepted state.
Change snapshots are held in process memory and are cleared when Point exits,
preventing hidden long-term retention of report contents.

### Saved workspace views

Use **Workspace > Save Workspace View** to preserve the current mode, selected
source, headings, and first two input rows. Use **Load Workspace View** to
restore them. The bounded local configuration is stored at:

```text
Workspace\point-saved-view.txt
```

The saved view stores only the mode, source name, up to 256 headings, and first
two grid rows; it never stores the complete imported report. Point v7.0 also
loads the previous six-column `POINT_VIEW_V1` format.

### Risk watchlist

Use **Workspace > Run Risk Watchlist** to run the local Insight Agent and show
only High and Medium findings. Every item includes its source/field, evidence,
transparent numeric score, and recommended action. If no rule triggers, Point
returns an explicit `Watchlist clear` result.

## Find, sorting, and clearing

- Type in the **Find in results** box and select **Find Next**.
- `Ctrl+F` focuses the Find box.
- Double-click a populated heading to sort that result column.
- Double-click the same heading again to reverse the sort direction.
- Numeric columns use numeric ordering; other columns use normalized text.
- `Ctrl+L` or **Workspace > Clear Results** removes result rows while
  preserving mode-specific input rows and headings.

## Import summary and source selection

Use **Workspace > Import Summary** to display each imported report, its row
count, field count, and rejected-report information.

Count and Compare operations block ambiguous fields that occur in multiple
reports. Analyze intentionally checks all compatible reports in Automatic mode
to detect cross-report duplicate keys. Use **Workspace > Next Source** to cycle
through imported reports and explicitly select one source. Cycling past the
final report returns selection to `Automatic`. The selected source is displayed
in the status line.

The build script copies the included sample reports into a new `build\Inbox`
so the prototype works immediately after compilation.

## 2,000,000-row grid and clipboard

Point exposes 2,000,000 logical rows and 256 logical columns using a sparse
virtualized store. Only 15 rows and the columns currently fitting in the
window are visible. Use the vertical scrollbar or mouse wheel to move through
rows, and the horizontal scrollbar to reach additional fields.

- `Ctrl+C` copies selected text, or the entire active cell when no text is
  selected.
- **Copy All** or `Ctrl+Shift+C` copies every used result column and row,
  including headings. Narrow and Count modes exclude their criteria row;
  Compare mode copies its complete six-column evidence report after Search.
- Drag from one cell into another cell to select a rectangular range. The
  selected cells are highlighted; `Ctrl+C` copies exactly that rectangle,
  including blank cells inside it.
- `Ctrl+V` pastes into the active cell.
- Tab/newline clipboard data copied from Excel is distributed across Point
  columns and rows beginning at the active cell.
- Export writes populated result rows and excludes mode-specific input rows.

## CSV expectations

- The first non-empty record is the header.
- UTF-8 (with or without BOM) and plain ASCII are supported.
- Commas, quotes, CRLF, and embedded newlines inside quoted fields are handled.
- Duplicate or blank headers cause the report to be quarantined logically
  (rejected and logged); Point never moves the original file.
- All values are treated as data, never as commands or formulas.

Example:

```csv
Employee ID,Username,Display Name,Email
E1001,speela,Sunil Kumar Peela,sunil@company.com
```

## Prototype limits

| Limit | Value |
|---|---:|
| File size | 500 MiB |
| Rows per report | 2,000,000 |
| Columns per report | 2,000 |
| Header length | 256 bytes |
| Cell length | 1 MiB |
| Result rows | 2,000,000 Normal / 1,999,999 Narrow plus one criteria row |

## Excel workbook import

Point v7.0 accepts `.xlsx`, `.xls`, and `.xlsm` when Microsoft Excel is
installed. Excel is used only as a local read-only import bridge:

- macros are forced off before any workbook opens;
- events, alerts, and external-link updates are disabled;
- workbooks open read-only and are never modified;
- every worksheet is exported to Point's private UTF-8 cache;
- cached CSV sheets are passed through the same bounded parser as ordinary CSV;
- the cache is recreated during Refresh.

This importer adds no bundled third-party runtime library, but it does require a
locally installed copy of Microsoft Excel. Point does not yet contain its own
independent XLSX/XLSM ZIP/XML or XLS OLE/BIFF parser.
