# CHD-Maid Changelog

All notable changes compared to release 02 (header version **2026.3.22**, ~549 lines) are reflected in the current **`chd-maid.ps1`** (header version **2026.3.27**).

## Release 03 [2026.3.27] changes since Release 02 [2026.3.22]

### Archives and 7-Zip

- **`.zip` / `.7z` / `.rar`** under `-source` are detected and extracted with **7-Zip** into a child folder under `-source` (named from the archive; ` (2)`, ` (3)`, … if the name collides). Archives are not unpacked under `-dest`.
- **Nested archives** under an extract are expanded in turn; disc conversion runs depth-first through the tree.
- **`-SevenZipPath`** (aliases **`-Path7z`**, **`-zpath`**) optionally points to `7z.exe` or its install folder; otherwise the script resolves 7-Zip beside the script, `Program Files`, etc.
- **`Invoke-ProcessWithStatus`** drives **7-Zip** extract progress on the console (same pattern as `chdman`).
- After a successful archive pass (no failed disc jobs under that unpack tree), the **extracted folder is removed before** the scanner continues with the next source item. **`-yes`** also removes the **archive file**; **`-no`** keeps the archive on disk but still removes the unpack tree when safe.
- **End-of-run** pass retries any extract records that still need cleanup (idempotent if already removed earlier).

### Source scanning and job model

- Replaced a single flat **“all .cue/.gdi/.iso”** list with a recursive **`Invoke-ChdMaidSourceDirectory`** walk: subfolders first, then archives at that level, then disc files—so **nested layouts** and **archives before cues in the same folder** behave predictably.
- **`-source`** and **`-dest`** resolution are split: source must exist (reprompt until valid); destination may be created when missing; interactive **Enter** uses the current directory for defaults.
- **`inputOutcome`** tracks per-input **success / skipped / failed** for archive cleanup rules and logging.

### Deleting sources (`-yes`)

- **`.cue`**: delete referenced **`.bin`** files (and optional cue variants in parsing).
- **`.gdi`**: delete **referenced track/data files** parsed from the GDI (same idea as cue sidecars).
- Removes the **whole game directory** under `-source` when it is a strict subfolder of `-source`, no **`.cue`/`.gdi`/`.iso`** remain, and no **unprocessed archives** remain. 
- **Junk and leftovers** go away, not only the primary image file.
- **Empty parent** folders after deletes are still pruned where applicable.
- Archive cleanup logic above aligns **extract removal** with **`-yes`/`-no`** for whether the **archive file** is deleted.

### CHD output and failures

- If **verify** fails after **create**, the new CHD is **removed** before surfacing failure; the **catch** block also removes **partial/failed** output if the file still exists.
- **Skipped / success** paths avoid counting failed or removed CHDs in **run totals** for new CHD size and space-saved math.

### Logging

- Default log path is **`logs\log-chd-maid-YYYY-MM-dd-HHmmss.log`** (UTF-8).
- The **`logs`** directory is **created if missing**.
- **`-nolog`**: no log for a **fully successful** run; on first failure, an **error-oriented** log can be created with a short header.
- **`quickCmd`** / flags in the log reflect **`-nolog`** and **`-SevenZipPath`** when bound.
- Release 02 used **`chd-maid-log-*.log`** in the **current directory** only; current uses **`logs\log-chd-maid-*.log`** and supports **`-nolog`**.
- Completion summary lines are **appended** to the log when appropriate (including failures).

### Completion summary

- **Elapsed time**, **CHDs created / skipped / failed**, **archives extracted count** are displayed (when non-zero).
- **Original size (processed disc sets)**: **archive file on-disk size** per extract pass (not unpacked folder size), **plus** once-per-folder loose payload (non-archive files) and **root-level** disc footprints under `-source`.
- **New CHD size (this run only)** and **Space saved (new CHDs vs source)** using **attributed** source material for successful jobs; **percentage of attributed source saved** may appear on the New CHD line.
- **Human-readable sizes** via **`Format-DataSizeGbMb`** (1024-based GB/MB; **N2** is culture-aware for thousands separators).
- **Lines with a numeric value of `0`** are **omitted** from the printed summary (elapsed and header always shown).

### Help and branding

- **Show-Help** expanded: workflow summary, **archive** behavior, **logging**, **parameters**, examples with **`-Path7z`** and **`-nolog`**.
- Typo fix in help: **“official”** (GitHub).

### Robustness

- **`Write-Fail`**: allows empty binding and substitutes a generic message if the error text is blank (avoids parameter binding failure when resolving 7-Zip).
- **`args` / `-help`** handling compatible with **PowerShell 7.6+** and **`-File` invocation without extra arguments**.
- **Console spacing** between unrelated jobs (e.g. after removing one game folder and before extracting the next archive) without inserting extra blanks **inside** the same archive’s extract → convert flow.

### Misc

- **`Get-CueReferencedBins`** returns a **case-insensitive deduped** set of bin paths (implementation detail vs. release 02 list).
---