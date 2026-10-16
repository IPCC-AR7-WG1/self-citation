# Zotero Reference and Self-Citation Analysis

A Python tool for analysing **WGI AR chapter references and chapter-author
self-citations** in Microsoft Word documents containing live Zotero
citation fields.

It produces a Word report with the chapter author lists, the rendered
Zotero bibliography, and separate self-citation statistics for
**Coordinating Lead Authors / Lead Authors (CLA/LA)** and **Contributing
Authors (CA)**. It also creates a detailed debug report for a selected
author.

## Requirements

-   Windows with desktop Microsoft Word
-   Python 3
-   `pywin32` (`pip install pywin32`)
-   An original `.docx` containing Zotero `ZOTERO_ITEM` citation fields
    and a `ZOTERO_BIBL` bibliography field

The Zotero Word add-in is recommended. This tool uses Word COM
automation and **does not run on macOS or Linux**.

## Installation

Clone or download this repository, then install the dependency:

``` bash
pip install pywin32
```

## Usage

Run the script from a terminal or command prompt:

``` bash
python zotero_reference_analysis_v11_deduplicated_interactive_debug_fixed.py
```

1.  Choose the original Word `.docx` in the file picker.

2.  The script extracts CLA, LA and CA names, Zotero citation metadata,
    and the rendered bibliography.

3.  In the terminal, select the author group for the debug report:

    ``` text
    Choose the author group for the debug report:
      1 = CLA/LA
      2 = CA
    Enter 1 or 2:
    ```

4.  The script prints the numbered authors in that group. Enter the
    **displayed number**, starting at **1**. Internally, the script uses
    `number - 1` to index the Python list.

5.  The script saves the output documents beside the original `.docx`.

**Important:** Keep the terminal open, because the author-selection
prompts use `input()`.

If an analysis report already exists, the script may offer to open it
instead of rerunning. Choose the rerun/replace option to recalculate
results and generate a new debug report.

## Output files

  ----------------------------------------------------------------------------------------------
  Output                                                     Description
  ---------------------------------------------------------- -----------------------------------
  `<input>_Reference_Analysis.docx`                          Author lists, formatted Zotero
                                                             bibliography, and separate CLA/LA
                                                             and CA self-citation tables

  `<input>_Reference_Analysis_DEBUG_<group>_<author>.docx`   First-author and co-author cited
                                                             works for the selected author,
                                                             including the complete Zotero
                                                             author lists, title, year,
                                                             publication venue and DOI when
                                                             available
  ----------------------------------------------------------------------------------------------

The original source document is **not modified**. The main output
preserves the rendered bibliography as it appears in Word; it does not
remove duplicate entries from the source bibliography.

## How self-citations are counted

### Author matching

An individual chapter author is matched to Zotero author metadata using:

-   Full **first given name** and **family name**
-   Middle names ignored
-   Case, diacritics, selected punctuation and spacing normalized

**Limitations:** Nicknames and different full first names (for example,
`Jone Smith` versus `Jonathen C. Smith`) are not automatically
treated as equivalent. Initial-only names such as `J. Smith` are not
treated as confirmed matches. These may require manual review.

### Publication deduplication

Before calculating statistics, the script tries to count each cited
publication only once using:

1.  Normalized DOI, where available
2.  Normalized title and publication year, when compatible
3.  Zotero item identifiers as a fallback

Records with conflicting non-empty DOIs are not merged solely because
their titles and years match. Among merged records, the script prefers
the one with the longer author list.

Deduplication is applied to the **analysis dataset**, not to the
rendered bibliography in the output Word document.

### Individual-author statistics

For each CLA/LA and CA, the table shows:

-   **First-author articles:** distinct cited works where that chapter
    author is listed first in Zotero
-   **Co-author articles:** distinct cited works where that chapter
    author appears after the first position
-   **Sum:** first-author + co-author works for that individual

The row labelled **Total sum** aggregates individual-author counts. It
can count a single paper more than once when several chapter authors
contributed to it. **Do not interpret that row as the number of unique
self-cited publications.**

### Unique group-level statistics

The report also includes a **Unique group self-cited papers** row for
each group (CLA/LA and CA), calculated separately:

  -----------------------------------------------------------------------
  Measure                             Definition
  ----------------------------------- -----------------------------------
  First-author                        Number of distinct cited papers
                                      whose first author belongs to the
                                      group

  Co-author                           Number of distinct cited papers
                                      with at least one group member in a
                                      non-first-author position

  Total                               Number of distinct cited papers
                                      with **any** group member anywhere
                                      in the author list
  -----------------------------------------------------------------------

First-author and co-author counts are **inclusive** and can overlap. For
example, if a CLA is first author and an LA is co-author of the same
paper, that paper contributes 1 to each category but only **1 to
Total**.

The **Unique self-citation percentage** row divides each unique count by
the total number of unique cited publications:

``` text
Unique group self-citation rate =
    (unique papers involving at least one group member / all unique cited papers) × 100
```

Publications with **no** chapter authors still count in the denominator.

The script **does not calculate a combined all-chapter-authors rate**;
CLA/LA and CA are reported separately. Their percentages should not be
added, because a paper may involve authors from both groups.

The table's individual-author totals and the unique group totals measure
different things. The group-level **Total** in the unique row is the
appropriate numerator for the proportion of cited publications involving
a member of that group.

## Example

Suppose the chapter cites 10 unique papers. Four involve at least one
CLA/LA, two involve at least one CA, and one of those papers involves
both groups.

-   CLA/LA unique self-citation rate: `4 / 10 × 100 = 40%`
-   CA unique self-citation rate: `2 / 10 × 100 = 20%`

The groups overlap, so these percentages must not be summed to estimate
an overall chapter rate.

## Troubleshooting

**No Zotero items found:** Check that the original document contains
live Zotero `ZOTERO_ITEM` fields rather than plain-text citations.

**Bibliography field not found:** The document needs a Zotero-generated
`ZOTERO_BIBL` field.

**Unexpected author names:** The script expects recognizable
`Coordinating Lead Authors`, `Lead Authors` and `Contributing Authors`
headings in the source document. Inspect the extracted lists.

**Duplicate references in the bibliography:** The analysis deduplicates
citation metadata, but the bibliography is copied from Word as rendered.
Correct duplicates in Zotero/the source document if the displayed
bibliography itself must be cleaned.

**Missing matches for an author:** Check Zotero's given/family names and
name variants; initials and nicknames are not automatically resolved.

**Word errors:** Close conflicting Word dialogs and ensure the source
file is accessible. The script uses a temporary local copy to reduce
problems with synced folders.

## Notes and limitations

-   This is a metadata-based counting tool, not an authorship
    identity-verification system.
-   It does not detect equal first authorship; the first entry in
    Zotero's ordered author list is treated as first author.
-   Results depend on the completeness and accuracy of embedded Zotero
    metadata.
-   The script has been syntax-checked, but its Word automation needs
    testing in the target Windows environment.
-   No license is specified here; add a `LICENSE` file if you want to
    define reuse permissions.
