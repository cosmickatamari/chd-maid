# Changelog

All notable changes compared to **release 03** (header version **2026.3.27**, **1341** lines in `release 03/chd-maid.ps1`) are reflected in the current **`chd-maid.ps1`** at the repo root (header version **2026.4.3**, **1492** lines).

## Release 04 [2026.4.3] changes since Release 03 [2026.3.27]

### Parameters and interactive behavior

- **`-ignore`**: skip a disc job when the matching **`.chd`** already exists under **`-dest`** (no verify/repair). For archives, matching uses the **archive base name** (e.g. `Title.7z` → `Title.chd` under **`-dest`**).
- **`-ignore`** cannot be combined with **`-yes`**; the script exits with **`Write-Fail`**.
- If you omit both **`-yes`** and **`-ignore`**, an **[I] Ignore / [R] Rescan** prompt may appear **only when** at least one **`.chd`** already exists under **`-dest`** (no prompt when the destination has no CHDs yet). Enter selects **Rescan**.
- Startup summary prints **Ignore sources:** (`Yes (-ignore)` / `No (rescan)`) and widens the label column so **Source:**, **Destination:**, **Delete sources:**, and **Ignore sources:** align.
- **`quickCmd`** / echoed flags include **`-ignore`** when that mode is active.

### Console progress and layout

- Each disc job prints a **`===[ current/total ]===`** line; **`current`** and **`total`** use US **thousands separators** via **`Format-UsInt32`**.
- **`total`** is initialized from the pre-run “compatible file(s)” count and **may increase** while processing if an archive yields **more disc jobs** than that nominal count.
- **Blank line** before the next job when the new input is under a **different parent folder** than the previous job (restores separation between games). **No** extra blank between **multi-disc** inputs in the **same** folder. Still coordinates with the existing **archive / folder-removal** section separator.
- Nested **extract roots** are kept **sorted longest-path-first** (**`Update-ChdmaidExtractDirRootsOrder`**) so path-prefix tests stay consistent as more archives are added.

### Archives and extract lifecycle

- When **`-ignore`** applies, **archives** can be **skipped entirely** if the expected **`.chd`** is already in **`-dest`** (no extract). The run adjusts the **progress total** downward for that skipped archive; optional **full log** lines record the ignore.
- **`Remove-ArchiveExtractionIfReady`**: detects leftover **archives** or **disc** files under an extract tree with a **single recursive directory walk** and **early exit** flags, instead of building intermediate **`Where-Object`** arrays for “still archives” / “still discs”.

### Logging

- **Per-disc batching**: while **full logging** (or on failure), each job appends a **contiguous block** of lines **`Complete-ChdmaidPerFileLog`** after that job finishes, so **chdman** live progress is **not interleaved** with finished-job log lines in the file.
- **`Add-ChdmaidPerFileLogLine`** accepts **empty strings** (**`[AllowEmptyString()]`**) for spacer lines in those blocks.
- Help text clarifies **completion summary** vs. **log path** behavior, **`-nolog`**, and **per-job** log layout.

### Completion summary

- New / split counters (when non-zero): **CHDs recreated**, **CHDs skipped (-ignore)**, **Archives skipped (-ignore)**.
- **“CHDs skipped (valid)”** renamed to **“CHDs skipped (verified existing)”**.
- **`success`**, **`skipped`**, **`failed`**, **`archives extracted`**, and related summary numbers use **`Format-UsInt32`**.
- **Elapsed time**: supports **days** in the display when the run crosses **≥ 1** day; sub-day times use a **zero-padded** `hh:mm:ss`-style presentation.
- **“New CHD size (this run only)”** no longer appends the old **“% of attributed source size saved”** parenthetical.
- **“Space saved (new CHDs vs source)”** includes **“N.N% difference”** vs **full processed original-size basis**.
- **Net space** uses **`bytesOriginalHandled - bytesChdThisRun`** (same basis as **Original size (processed disc sets)** for successful **new CHD** bytes), instead of maintaining a separate **`bytesOriginalAttributedSuccess`** tally and **`Add-ChdmaidSuccessOriginalAttributionForNet`** / per-success hashtables (**`chdmaidExtractRootUsedForNetSave`**, **`chdmaidLooseFolderNetDone`**, **`chdmaidLooseRootNetDone`**).
- Console: **CHDs failed** line highlights the **numeric** portion in **red**.
- Summary may list **Log file:** when a log path exists (including **error-only** logs under **`-nolog`**).

### Source pre-scan

- **One** recursive pass over **`-source`** files: count **archives** and collect **`.cue`/`.gdi`/`.iso`** (replaces **three** separate **`Get-ChildItem -Filter`** passes for `*.cue`, `*.gdi`, `*.iso`).

### Help and branding

- Header shows **Version 2026.4.3**.
- Archives section notes **`-no`/`-yes`** default when neither switch is passed; example **`.chd`** name uses **ImageName** for consistency with **`-ignore`** docs.
- New help for **`-ignore`** (prompt equivalence, archive naming, **`-yes`** restriction) and an **example** command line with **`-ignore`**.

### Misc

- “No jobs completed” guard treats **ignored** disc and archive skips as **work performed** so a run that only **`-ignore`**-skips does not mis-report an empty pass when inputs existed.

---
