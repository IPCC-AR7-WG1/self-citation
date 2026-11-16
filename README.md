# Zotero Reference and Self-Citation Analysis

Analyse the references in multiple Microsoft Word chapters, calculate separate self-citation statistics for **Coordinating Lead Authors / Lead Authors (CLA/LA)** and **Contributing Authors (CA)**, and generate a debug report for a **selected author in each chapter**.

The script reads the Zotero citation metadata embedded in Word (`ZOTERO_ITEM`) and copies the rendered bibliography from the `ZOTERO_BIBL` field. It does **not** change the original Word documents.

For instance, you can use the compiled FOD in [Documents/03.Chapters/40.Draft preparation/02.FOD/03.FOD compiled/](https://smu.sharepoint.com/:f:/r/teams/ipccwg1-ar7/Shared%20Documents/03.%20Chapters/40.%20Draft%20preparation/02.%20FOD/03.%20FOD%20compiled?d=w79bb9de1bbcb4aaf8469126041f7846a&csf=1&web=1&e=EJ0Tdi)

The script reads the Zotero citation metadata embedded in Word (`ZOTERO_ITEM`) and copies the rendered bibliography from the `ZOTERO_BIBL` field. It does **not** change the original Word documents.

## Requirements

-   Windows with desktop Microsoft Word
-   Python 3
-   `pywin32` (`pip install pywin32`)
-   An original `.docx` containing Zotero `ZOTERO_ITEM` citation fields
    and a `ZOTERO_BIBL` bibliography field

The Zotero Word add-in is recommended. This tool uses Word COM
automation and **does not run on macOS or Linux**.

## Installation
Install the Python dependency:

```cmd
py -m pip install pywin32
```

If `py` is unavailable, try `python -m pip install pywin32`.

## Files and folder setup

Place the script and the optional IPCC author database together (a .json file containing the database will be created when running the script for the first time!):

```text
analysis-tools/
├── zotero_reference_analysis.py
└── ipcc_author_database_partial.json       # optional; partial, advisory database
```

Keep your existing verification cache **in the folder containing the Word documents**, not necessarily beside the Python script:

```text
chapter-documents/
├── Chapter1.docx
├── Chapter2.docx
├── ...
└── author_verification_cache.json          # existing decisions; optional initially
```

If the cache does not yet exist, the script starts with an empty cache and creates it when you make a completed verification decision. Back up an existing cache before large batch runs.

## Run

Open Command Prompt in the script folder and run:

```cmd
py zotero_reference_analysis_batch_ipcc_select_debug_author.py
```

The script will:

1. Open a folder picker. Choose the folder containing the original chapter `.docx` files.
2. List eligible source files. Previously generated `_Reference_Analysis` and debug files, and Word `~$` temporary files, are excluded.
3. Ask whether to **replace existing analysis reports**. If you answer **No** (the default), any chapter with an existing analysis report is skipped entirely, including verification and debug-author selection.
4. Load the **historical verification cache** from the selected document folder and, if present, the **partial IPCC database** from the script folder.
5. For each processed document, extract Zotero items and chapter authors; verify ambiguous author matches; calculate the statistics.
6. Ask you to choose a **debug author separately for that document**.
7. Save the main analysis report and the selected author's debug report beside the source file.
8. At the end, write a CSV verification report and display a success/skip/failure summary.

### Select the debug author

For each processed document, the script prompts:

```text
Choose the author group for the debug report:
  1 = CLA/LA
  2 = CA
Enter 1 or 2:
```

It then prints the numbered authors in the chosen group:

```text
CLA_LA authors:
  1. Author A
  2. Author B
  3. Author C
Enter author number (1-3):
```

Enter **1** for the first displayed author, **2** for the second, and so on. The script converts this human-readable numbering to Python's zero-based indexing internally. This selection is repeated for **each chapter**, because the author lists can differ.

## Author-name verification

Full first-name + surname matches are counted automatically (middle names are ignored). A potential match involving an initial, such as **C. Smith** versus **Christopher Smith**, triggers a verification step.

**The order of operations is:**

1. Check `author_verification_cache.json` for the specific publication, author position and target chapter author.
2. Reuse a previously **confirmed** or **rejected** decision without prompting.
3. If no completed decision exists, display the publication and ambiguous author pair and offer:

   | Choice | Action | Internet? |
   |---|---|---|
   | **1** | Confirm the author manually | No |
   | **2** | Reject the match manually | No |
   | **3** | Review a suggestion from the local IPCC database, then confirm, reject, search online or skip | No for the local lookup |
   | **4** | Verify via Crossref, then OpenAlex if needed | **Yes** |
   | **5** | Skip; leave unresolved for a later run | No |

4. Store completed decisions in the existing JSON cache. The script writes the complete updated JSON file atomically rather than appending raw text. **Unresolved/skipped matches are not cached as final decisions** and can be revisited in a later run.
5. Use confirmed matches in the calculations. Rejected and unresolved matches are not counted.

### Local IPCC database

`ipcc_author_database_partial.json` is **optional** and currently **incomplete**. The script looks for it **beside the Python script** and indexes its records by DOI. It considers the listed **CLA/LA** names as possible chapter citation authors; CA and Review Editor names are not automatically treated as chapter authors.

The database is **advisory**: an absent name is **not** proof of non-authorship. After reviewing it, you can still confirm or reject the match manually. The script never automatically rejects a match merely because the local IPCC database lacks that name.

### Online verification and confidentiality

The script makes an external request **only if you select online verification**. It uses Crossref first and OpenAlex as a fallback, checking publication identity and author position before accepting a result. If online metadata is insufficient, you can still decide manually or leave the pair unresolved.

Online requests send bibliographic identifiers such as **DOIs or titles** to external services. Do not enable online verification for material whose disclosure is restricted without checking applicable confidentiality requirements. The script does not upload the entire Word document or bibliography as part of this lookup.

## Outputs

For each processed `Chapter1.docx`, the script creates:

| Output | Contents |
|---|---|
| `Chapter1_Reference_Analysis.docx` | Extracted chapter author lists, Word-rendered Zotero bibliography, and separate CLA/LA and CA self-citation statistics |
| `Chapter1_Reference_Analysis_DEBUG_CLA_LA_Author_A.docx` | First-author and co-author references for the selected CLA/LA author; `CA` is used in the filename if you selected that group |

The selected document folder also contains:

| File | Purpose |
|---|---|
| `author_verification_cache.json` | Persistent completed manual/online verification decisions; reused in subsequent runs and across chapters |
| `author_verification_report.csv` | Ambiguous matches encountered in the current batch, their status, method and details; **rewritten on each run**, not a permanent cumulative audit log |

**Keep a backup** of the JSON cache and any CSV reports you need to preserve before rerunning.

## Counting rules

For **each group separately** (CLA/LA and CA):

- **Individual first-author count:** Unique cited publications where that chapter author appears first.
- **Individual co-author count:** Unique cited publications where that chapter author appears after the first position.
- **Unique group first-author count:** Distinct cited publications with a group member in first position.
- **Unique group co-author count:** Distinct cited publications with at least one group member in a later position.
- **Unique group total:** The **union** of those publications: each cited publication is counted **once**, even if multiple members of the group wrote it.
- **Group self-citation rate:** `100 × unique group total / all unique cited publications`.

A publication can contribute to **both** the first-author and co-author group categories (if different group members occupy those positions), but only **once** to the group total. The **Total sum** row sums individual author counts and may double-count papers; use **Unique group self-cited papers** and **Unique self-citation percentage** for group-level interpretation.

The script does **not** calculate a combined all-chapter-authors rate. Publications without a chapter-author match remain in the denominator.

The script deduplicates structured Zotero records for analysis using DOI/title/year and fallback identifiers. It does **not** rewrite or deduplicate the bibliography displayed in the Word report.

## Troubleshooting

**The script asks me to verify an author I verified previously.** Check that the correct `author_verification_cache.json` is in the **selected Word-document folder**. Cache keys also depend on publication identity, author position and chapter-author name, so changed metadata may lead to a new prompt.

**I cannot select a debug author.** If the main analysis report already exists and you answered **No** to replacement, the chapter is skipped. Rerun and choose **Yes** to regenerate it.

**The local IPCC database is not found.** Place `ipcc_author_database_partial.json` beside the `.py` script. The script can still work without it using the historical cache, manual verification and optional online verification.

**No usable Zotero metadata or bibliography field found.** Use the original document with live Zotero fields, not a plain-text or unlinked bibliography.

**An author was not matched.** Check the Zotero given/family names. Initial-only names require verification; different full first names and nicknames are not automatically considered equivalent.

**A document fails.** The batch processor reports the error and continues with other files. Review the failure summary at the end.

## Limitations

- Requires **Windows and Microsoft Word COM automation**; not tested on macOS or Linux.
- The local IPCC database is partial and should not be considered a complete official author registry.
- Online bibliographic services may return incomplete or incorrect IPCC chapter authorship metadata.
- Joint first authorship is not inferred; the first author in Zotero's ordered author list is treated as first author.
- Results depend on the quality of the source Zotero metadata and the accuracy of manual verification decisions.
- The script was reviewed for its implemented workflow; Word automation and live API behavior must be tested on your Windows installation.


