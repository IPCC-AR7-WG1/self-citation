"""
Create a normal .docx containing:
  - merged Coordinating Lead Authors + Lead Authors list
  - Contributing Authors list
  - rendered Zotero bibliography
  - separate self-citation tables for CLA/LA and CA

Self-citation statistics are calculated from structured author metadata embedded
in ZOTERO_ITEM fields. No Zotero citation-style change is required.
The original source document is never edited.

Requirements:
    pip install pywin32

Run on Windows with Microsoft Word and the Zotero Word add-in installed.
"""

import os
import re
import json
import unicodedata
import shutil
import tempfile
import time
import csv
import copy
from difflib import SequenceMatcher
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen
import tkinter as tk
from tkinter import filedialog, messagebox

import pythoncom
import win32com.client as win32


WD_FORMAT_DOCX = 16
WD_COLLAPSE_END = 0
WD_AUTOFIT_CONTENT = 1


def choose_source():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select the original Word document",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
    )
    root.destroy()
    return path


def clean_para_text(text):
    return text.replace("\r", "").replace("\x07", "").strip()


def normalize_label(text):
    return re.sub(r"\s+", " ", clean_para_text(text)).strip()


def strip_country_parentheses(text):
    """
    Remove trailing IPCC country/affiliation parentheses robustly, including
    cases such as:
        Yongqiang Yu (China).
        Jatin Kala (Mauritius, Australia),
        Name (Country);
    """
    text = (text or "")
    text = text.replace("\x0b", " ").replace("\r", " ").replace("\n", " ")
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()

    # Remove trailing punctuation AFTER a closing country parenthesis first.
    # Then remove the complete parenthetical group. Repeat if necessary.
    while True:
        old = text
        text = re.sub(r"(?<=\))\s*[.,;:]+\s*$", "", text).strip()
        text = re.sub(r"\s*\([^()]*\)\s*$", "", text).strip()
        if text == old:
            break

    return text.strip(" ,;:.")

def split_author_entries(text):
    """
    Split an IPCC-style author list only at commas OUTSIDE parentheses.
    Country lists inside parentheses remain intact.
    """
    text = clean_para_text(text)
    if not text:
        return []

    parts = []
    start = 0
    depth = 0

    for i, ch in enumerate(text):
        if ch == "(":
            depth += 1
        elif ch == ")" and depth > 0:
            depth -= 1
        elif ch == "," and depth == 0:
            part = text[start:i].strip()
            if part:
                parts.append(part)
            start = i + 1

    last = text[start:].strip()
    if last:
        parts.append(last)

    return parts


def author_display_name(full_name):
    """
    Convert e.g.
      Li Cheng -> Cheng, L.
      Maria Elena Rossi -> Rossi, M. E.
      Jean-Pierre Dupont -> Dupont, J.-P.

    Keeps particles that are part of the surname reasonably well by treating
    common lowercase particles immediately before the final token as surname.
    """
    name = strip_country_parentheses(full_name)
    name = re.sub(r"\s+", " ", name).strip(" ,;")
    if not name:
        return ""

    # Already "Surname, Given" form.
    if "," in name:
        surname, given = [x.strip() for x in name.split(",", 1)]
        given_tokens = given.split()
    else:
        tokens = name.split()
        if len(tokens) == 1:
            return tokens[0]

        particle_words = {
            "da", "de", "del", "della", "der", "di", "dos", "du",
            "la", "le", "van", "von", "den", "ten", "ter"
        }
        surname_start = len(tokens) - 1
        while surname_start - 1 >= 1 and tokens[surname_start - 1].lower() in particle_words:
            surname_start -= 1

        surname = " ".join(tokens[surname_start:])
        given_tokens = tokens[:surname_start]

    initials = []
    for token in given_tokens:
        token = token.strip()
        if not token:
            continue

        # Hyphenated given name -> J.-P.
        if "-" in token:
            bits = [b for b in token.split("-") if b]
            hy = "-".join((b[0].upper() + ".") for b in bits if b)
            if hy:
                initials.append(hy)
        else:
            # Preserve an existing initial, otherwise take first character.
            m = re.match(r"([A-Za-zÀ-ÖØ-öø-ÿ])", token)
            if m:
                initials.append(m.group(1).upper() + ".")

    return surname + (", " + " ".join(initials) if initials else "")


def find_author_sections(doc):
    """
    Extract complete CLA/LA/CA author blocks. Word line wrapping is normalized
    before splitting, so names are never cut at a visual/manual line break.
    """
    raw_text = doc.Content.Text[:120000]
    raw_text = raw_text.replace("\x0b", " ").replace("\x07", " ")
    raw_text = raw_text.replace("\r", "\n")
    raw_text = re.sub(r"[ \t]+", " ", raw_text)

    labels = [
        ("coordinating lead authors", "cla"),
        ("lead authors", "la"),
        ("contributing authors", "ca"),
    ]

    hits = []
    for label, key in labels:
        m = re.search(rf"(?im)^\s*{re.escape(label)}\s*:?\s*", raw_text)
        if m:
            hits.append((m.start(), m.end(), key))
    hits.sort()

    found = {"cla": [], "la": [], "ca": []}

    for i, (_, content_start, key) in enumerate(hits):
        if i + 1 < len(hits):
            content_end = hits[i + 1][0]
        else:
            tail = raw_text[content_start:]
            # The author search must end at Review Editors. This prevents
            # names/affiliations after the author block from being interpreted
            # as Contributing Authors.
            stop = re.search(
                r"(?im)^\s*(?:review editors|chapter|executive summary|summary|references|"
                r"frequently asked questions?|faq|box|figure|table)\s*:?",
                tail
            )
            content_end = content_start + (stop.start() if stop else len(tail))

        block = raw_text[content_start:content_end]
        block = re.sub(r"\s*\n+\s*", " ", block)
        block = re.sub(r"\s+", " ", block).strip(" ,;")

        entries = split_author_entries(block)
        cleaned, seen = [], set()
        for entry in entries:
            name = strip_country_parentheses(entry)

            # Defensive guard against malformed leftovers such as "(China)."
            # or punctuation-only fragments being treated as an author.
            if not name or re.fullmatch(r"[\W_]+", name):
                continue
            if re.fullmatch(r"\([^()]+\)[.,;:]*", name):
                continue

            marker = name.casefold()
            if marker not in seen:
                cleaned.append(name)
                seen.add(marker)
        found[key] = cleaned

    return found

def split_person_name(full_name):
    """Return first given name and family name; ignore middle given names."""
    name = strip_country_parentheses(full_name)
    name = re.sub(r"\s+", " ", name).strip(" ,;")
    if not name:
        return "", ""

    if "," in name:
        family, given = [x.strip() for x in name.split(",", 1)]
        return (given.split()[0] if given else ""), family

    tokens = name.split()
    if len(tokens) < 2:
        return "", tokens[0] if tokens else ""

    particles = {
        "da", "de", "del", "della", "der", "di", "dos", "du",
        "la", "le", "van", "von", "den", "ten", "ter"
    }
    surname_start = len(tokens) - 1
    while surname_start - 1 >= 1 and tokens[surname_start - 1].lower() in particles:
        surname_start -= 1

    return tokens[0], " ".join(tokens[surname_start:])


def normalize_name_piece(value):
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.casefold()
    return re.sub(r"[\s\-\u2010-\u2015'’.]+", "", value)


def zotero_first_given(given):
    given = re.sub(r"\s+", " ", (given or "").strip())
    return given.split()[0] if given else ""


def same_person(chapter_name, zotero_author):
    """
    Match full first name + family name. Ignore middle names.
    Initial-only Zotero given names are not treated as full-name matches.
    """
    first, family = split_person_name(chapter_name)
    z_first = zotero_first_given(zotero_author.get("given", ""))
    z_family = zotero_author.get("family", "")

    if not first or not family or not z_first or not z_family:
        return False
    if len(normalize_name_piece(z_first)) <= 1:
        return False

    return (normalize_name_piece(first) == normalize_name_piece(z_first)
            and normalize_name_piece(family) == normalize_name_piece(z_family))


def extract_json_object(field_code):
    a = field_code.find("{")
    b = field_code.rfind("}")
    if a < 0 or b <= a:
        return None
    try:
        return json.loads(field_code[a:b + 1])
    except json.JSONDecodeError:
        return None


def normalize_doi(value):
    """Normalize DOI values, including URL and doi: forms."""
    value = (value or "").strip().casefold()
    value = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", value)
    value = re.sub(r"^doi:\s*", "", value)
    return value.strip().rstrip(" .")


def normalized_title_year(item):
    title = unicodedata.normalize("NFKD", item.get("title") or "")
    title = "".join(c for c in title if not unicodedata.combining(c))
    title = re.sub(r"[^a-z0-9]+", " ", title.casefold()).strip()
    try:
        year = str(item["issued"]["date-parts"][0][0])
    except (KeyError, IndexError, TypeError):
        year = ""
    return (title, year) if title and year else None


def unique_item_key(citation_item, item_data):
    """Fallback identifier for works lacking DOI and title/year."""
    uris = citation_item.get("uris") or []
    if uris:
        return "uri:" + str(uris[0])
    for key in ("citation-key", "URL", "id"):
        value = item_data.get(key)
        if value:
            return key + ":" + str(value).casefold()
    if citation_item.get("id") is not None:
        return "item-id:" + str(citation_item["id"])
    return "fallback:" + repr(item_data)


def extract_unique_zotero_items(doc):
    """
    Deduplicate Zotero citations by DOI, or normalized title + year when
    DOI is absent. Different Zotero URIs for the same work count only once.
    Preserve the most complete author metadata among duplicates.
    """
    unique = {}
    doi_index = {}
    title_index = {}
    citation_fields = 0
    parse_failures = 0
    duplicate_occurrences = 0

    for field in doc.Fields:
        try:
            code = field.Code.Text or ""
        except Exception:
            continue
        if "ZOTERO_ITEM" not in code.upper():
            continue
        citation_fields += 1
        payload = extract_json_object(code)
        if not payload:
            parse_failures += 1
            continue

        for citation_item in payload.get("citationItems") or []:
            item = citation_item.get("itemData") or {}
            if not item:
                continue
            doi = normalize_doi(item.get("DOI"))
            title_year = normalized_title_year(item)
            key = None
            if doi and doi in doi_index:
                key = doi_index[doi]
            elif title_year and title_year in title_index:
                candidate = title_index[title_year]
                existing_doi = normalize_doi(unique[candidate].get("DOI"))
                # Never merge two records with different nonempty DOIs.
                if not doi or not existing_doi or doi == existing_doi:
                    key = candidate

            if key is None:
                key = ("doi:" + doi) if doi else (
                    "title-year:" + repr(title_year) if title_year
                    else unique_item_key(citation_item, item)
                )
                if key not in unique:
                    unique[key] = item
                else:
                    duplicate_occurrences += 1
            else:
                duplicate_occurrences += 1
                # Prefer metadata with a longer complete author list.
                if len(item.get("author") or []) > len(unique[key].get("author") or []):
                    unique[key] = item

            if doi:
                doi_index[doi] = key
            if title_year:
                title_index[title_year] = key

    print(f"  Repeated citation/item occurrences removed: {duplicate_occurrences}")
    return list(unique.values()), citation_fields, parse_failures


def find_zotero_bibliography_field(doc):
    for field in doc.Fields:
        try:
            code = field.Code.Text or ""
            if "ZOTERO_BIBL" in code.upper():
                return field
        except Exception:
            pass
    return None


def paragraph_texts_from_formatted_bibliography(field):
    texts = []
    for p in field.Result.Paragraphs:
        t = clean_para_text(p.Range.Text)
        if t:
            texts.append(t)
    return texts



# ---------------- Optional author identity verification ----------------
# Only option 3 below makes an external HTTP request. No document, chapter
# title, chapter author list, or bibliography is transmitted.

def work_key(item):
    doi = normalize_doi(item.get("DOI"))
    if doi:
        return "doi:" + doi
    title_year = normalized_title_year(item)
    if title_year:
        return "title-year:" + "|".join(title_year)
    return "item:" + str(item.get("id") or item.get("title") or "unknown")


def verification_key(item, position, chapter_author):
    first, family = split_person_name(chapter_author)
    return json.dumps([work_key(item), position,
                       normalize_name_piece(first), normalize_name_piece(family)],
                      ensure_ascii=False)


def ambiguous_initial_match(chapter_author, z_author):
    first, family = split_person_name(chapter_author)
    z_first = zotero_first_given(z_author.get("given", ""))
    if not first or not family or not z_first:
        return False
    a = normalize_name_piece(first)
    b = normalize_name_piece(z_first)
    return (normalize_name_piece(family) == normalize_name_piece(z_author.get("family", ""))
            and len(a) > 1 and len(b) == 1 and a.startswith(b))


def request_json(url):
    req = Request(url, headers={"User-Agent": "ZoteroAuthorVerifier/1.0 (research metadata lookup)",
                                "Accept": "application/json"})
    with urlopen(req, timeout=15) as response:
        return json.load(response)


def comparable_title(value):
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def publication_verified(item, remote_title, remote_doi, remote_year=None):
    local_doi = normalize_doi(item.get("DOI"))
    remote_doi = normalize_doi(remote_doi)
    if local_doi and remote_doi:
        return local_doi == remote_doi
    if local_doi and not remote_doi:
        return False
    local_title = comparable_title(item.get("title"))
    other_title = comparable_title(remote_title)
    if not local_title or not other_title:
        return False
    score = SequenceMatcher(None, local_title, other_title).ratio()
    local_year = item_year(item)
    return score >= 0.96 and (not local_year or not remote_year or str(local_year) == str(remote_year))


def remote_author_matches(chapter_author, remote_authors, position):
    if position >= len(remote_authors):
        return "unresolved", "Author position missing from online metadata"
    candidate = remote_authors[position]
    given = candidate.get("given", "")
    family = candidate.get("family", "")
    if same_person(chapter_author, candidate):
        return "confirmed", (given + " " + family).strip()
    if ambiguous_initial_match(chapter_author, candidate):
        return "unresolved", (given + " " + family).strip()
    # Only reject if full given name and surname are supplied, and the
    # original initial/surname remains compatible with the candidate.
    if len(normalize_name_piece(zotero_first_given(given))) > 1 and family:
        return "rejected", (given + " " + family).strip()
    return "unresolved", (given + " " + family).strip()


def crossref_lookup(item):
    doi = normalize_doi(item.get("DOI"))
    if doi:
        data = request_json("https://api.crossref.org/works/" + quote(doi, safe=""))["message"]
    else:
        title = item.get("title") or ""
        data = request_json("https://api.crossref.org/works?" + urlencode({"query.title": title, "rows": 3}))["message"]
        data = next((candidate for candidate in data.get("items", [])
                     if publication_verified(item, (candidate.get("title") or [""])[0],
                                             candidate.get("DOI"),
                                             ((candidate.get("published") or {}).get("date-parts") or [[None]])[0][0])), None)
        if data is None:
            return None
    title = (data.get("title") or [""])[0]
    if not publication_verified(item, title, data.get("DOI"),
                                ((data.get("published") or {}).get("date-parts") or [[None]])[0][0]):
        return None
    return [{"given": a.get("given", ""), "family": a.get("family", "")}
            for a in data.get("author", [])]


def openalex_lookup(item):
    doi = normalize_doi(item.get("DOI"))
    if doi:
        data = request_json("https://api.openalex.org/works/https://doi.org/" + quote(doi, safe=""))
    else:
        data = request_json("https://api.openalex.org/works?" + urlencode({
            "search": item.get("title", ""), "per-page": 5}))
        data = next((candidate for candidate in data.get("results", [])
                     if publication_verified(item, candidate.get("title"), candidate.get("doi"),
                                             candidate.get("publication_year"))), None)
        if data is None:
            return None
    if not publication_verified(item, data.get("title"), data.get("doi"), data.get("publication_year")):
        return None
    result = []
    for authorship in data.get("authorships", []):
        display = (authorship.get("author") or {}).get("display_name", "")
        bits = display.strip().split()
        result.append({"given": " ".join(bits[:-1]), "family": bits[-1] if bits else ""})
    return result


def online_verify(item, position, chapter_author):
    """Crossref first; OpenAlex fallback; never accept contradictory results."""
    attempts = []
    for name, lookup in (("Crossref", crossref_lookup), ("OpenAlex", openalex_lookup)):
        try:
            authors = lookup(item)
            if authors is None:
                attempts.append((name, "unresolved", "Publication not confidently identified"))
                continue
            status, name_found = remote_author_matches(chapter_author, authors, position)
            attempts.append((name, status, name_found))
            if status == "confirmed":
                # Do not query the second provider if the first confidently confirms.
                return status, name, name_found
            # A negative identification from one source should be reviewed
            # rather than silently treated as definitive.
        except Exception as exc:
            attempts.append((name, "unresolved", "Lookup failed: " + str(exc)[:150]))
    confirmed = [a for a in attempts if a[1] == "confirmed"]
    rejected = [a for a in attempts if a[1] == "rejected"]
    if confirmed and rejected:
        return "unresolved", "Crossref/OpenAlex conflict", str(attempts)
    if confirmed:
        return "confirmed", confirmed[0][0], confirmed[0][2]
    return "unresolved", "Crossref/OpenAlex", str(attempts)


class IPCCAuthorDatabase:
    """Read-only, DOI-indexed reference list; never substitutes for user approval."""
    def __init__(self, path):
        self.path = path
        self.by_doi = {}
        if not path or not os.path.isfile(path):
            return
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
        for entry in data.get("records", []):
            doi = normalize_doi(entry.get("doi"))
            if doi:
                self.by_doi[doi] = entry

    def suggest(self, item, chapter_author, original):
        entry = self.by_doi.get(normalize_doi(item.get("DOI")))
        if not entry:
            return None
        # The chapter's credited authors are CLA/LA. CA and RE must not be
        # silently treated as chapter citation authors.
        names = (entry.get("authors") or {}).get("CLA", []) + (entry.get("authors") or {}).get("LA", [])
        target_first, target_last = split_person_name(chapter_author)
        candidates = []
        for name in names:
            first, last = split_person_name(name)
            if (normalize_name_piece(last) == normalize_name_piece(target_last)
                and normalize_name_piece(first) == normalize_name_piece(target_first)):
                candidates.append(name)
        return {"entry": entry, "candidates": candidates}


class VerificationSession:
    def __init__(self, folder, ipcc_db=None):
        self.ipcc_db = ipcc_db
        self.path = os.path.join(folder, "author_verification_cache.json")
        self.report_path = os.path.join(folder, "author_verification_report.csv")
        try:
            with open(self.path, encoding="utf-8") as handle:
                self.cache = json.load(handle)
            if not isinstance(self.cache, dict):
                raise ValueError("Expected JSON object")
        except FileNotFoundError:
            self.cache = {}
        self.report = []
        self.seen = set()

    def save(self):
        temp = self.path + ".tmp"
        with open(temp, "w", encoding="utf-8") as handle:
            json.dump(self.cache, handle, indent=2, ensure_ascii=False)
        os.replace(temp, self.path)

    def decide(self, item, position, chapter_author, original):
        key = verification_key(item, position, chapter_author)
        record = self.cache.get(key)
        if record is None or record.get("status") == "unresolved":
            print("\n" + "=" * 65)
            print("AMBIGUOUS AUTHOR MATCH")
            print("Chapter author:", author_display_name(chapter_author))
            print("Zotero author:", full_zotero_author_name(original))
            print("Position:", position + 1, "(First author)" if position == 0 else "(Co-author)")
            print("Publication:", item.get("title", "[No title]"))
            print("Year:", item_year(item), "DOI:", item.get("DOI", "[No DOI]"))
            print("  1 = Confirm this is the chapter author")
            print("  2 = Reject this match")
            local = self.ipcc_db.suggest(item, chapter_author, original) if self.ipcc_db else None
            if local:
                entry = local["entry"]
                print("  Local IPCC record:", entry.get("report"), "chapter", entry.get("chapter"))
                print("  Source:", entry.get("source_url", ""))
                print("  Database status:", entry.get("verification_status", "not audited"))
                print("  Matching official CLA/LA name(s):", ", ".join(local["candidates"]) or "NONE")
                print("  Note: database is partial and not independently audited.")
            print("  3 = Review local IPCC suggestion (no internet)")
            print("  4 = Verify online (Crossref, then OpenAlex)")
            print("  5 = Skip (unresolved; not counted)")
            while True:
                choice = input("Choice [1/2/3/4/5, default=5]: ").strip() or "5"
                if choice in ("1", "2", "3", "4", "5"):
                    break
                print("Please enter 1, 2, 3, 4, or 5.")
            if choice == "1":
                record = {"status": "confirmed", "method": "manual", "detail": "User confirmed"}
            elif choice == "2":
                record = {"status": "rejected", "method": "manual", "detail": "User rejected"}
            elif choice == "3":
                # The local database is advisory only. An absent name is never
                # proof that the chapter author did not write the publication.
                if local and local["candidates"]:
                    print("IPCC source:", local["entry"].get("source_url", ""))
                    print("Matching name(s):", ", ".join(local["candidates"]))
                else:
                    print("No matching local IPCC CLA/LA name found; not proof of non-authorship.")
                    print("The local database may be incomplete or the author may have another role.")
                # Always allow a manual decision after reviewing the IPCC data,
                # including when no DOI record or name match exists.
                while True:
                    response = input(
                        "Confirm manually [1], reject manually [2], "
                        "verify online [3], or skip for later [Enter]: "
                    ).strip()
                    if response in ("", "1", "2", "3"):
                        break
                    print("Please enter 1, 2, 3, or press Enter to skip.")
                if response == "1":
                    record = {"status": "confirmed", "method": "manual-after-ipcc-db",
                              "detail": "User confirmed after reviewing local IPCC database"}
                elif response == "2":
                    record = {"status": "rejected", "method": "manual-after-ipcc-db",
                              "detail": "User rejected after reviewing local IPCC database"}
                elif response == "3":
                    status, source, detail = online_verify(item, position, chapter_author)
                    print("Online result:", status, "|", source, "|", detail)
                    if status == "unresolved":
                        override = input(
                            "Confirm manually [1], reject manually [2], "
                            "or leave unresolved [Enter]? "
                        ).strip()
                        if override == "1":
                            record = {"status": "confirmed", "method": "manual-after-online", "detail": detail}
                        elif override == "2":
                            record = {"status": "rejected", "method": "manual-after-online", "detail": detail}
                        else:
                            record = {"status": "unresolved", "method": source, "detail": detail}
                    else:
                        record = {"status": status, "method": source, "detail": detail}
                else:
                    record = {"status": "unresolved", "method": "skipped-after-ipcc-db",
                              "detail": "User deferred decision after reviewing IPCC database"}
            elif choice == "4":
                status, source, detail = online_verify(item, position, chapter_author)
                print("Online result:", status, "|", source, "|", detail)
                if status == "unresolved":
                    override = input("Confirm manually [1], reject [2], or leave unresolved [Enter]? ").strip()
                    if override == "1":
                        record = {"status": "confirmed", "method": "manual-after-online", "detail": detail}
                    elif override == "2":
                        record = {"status": "rejected", "method": "manual-after-online", "detail": detail}
                    else:
                        record = {"status": status, "method": source, "detail": detail}
                else:
                    record = {"status": status, "method": source, "detail": detail}
            else:
                record = {"status": "unresolved", "method": "skipped", "detail": "No decision"}
            # Cache only completed decisions. Skipped/unresolved pairs will be
            # prompted again on future runs.
            if record["status"] != "unresolved":
                self.cache[key] = record
                self.save()
        if key not in self.seen:
            self.seen.add(key)
            self.report.append({
                "publication_key": work_key(item), "title": item.get("title", ""),
                "doi": item.get("DOI", ""), "chapter_author": author_display_name(chapter_author),
                "zotero_author": full_zotero_author_name(original),
                "author_position": position + 1, "status": record["status"],
                "method": record["method"], "detail": record.get("detail", "")})
        return record["status"]

    def resolve_items(self, items, cla_la, ca):
        """Use copies for counting; never mutate the original Word metadata."""
        resolved = copy.deepcopy(items)
        chapter_authors = list(dict.fromkeys(cla_la + ca))
        ambiguous = []
        for item in resolved:
            for position, author in enumerate(item.get("author") or []):
                for chapter_author in chapter_authors:
                    if ambiguous_initial_match(chapter_author, author):
                        ambiguous.append((item, position, chapter_author, copy.deepcopy(author)))
        print(f"  Ambiguous author-publication pairs: {len(ambiguous)}")
        for item, position, chapter_author, original in ambiguous:
            if self.decide(item, position, chapter_author, original) == "confirmed":
                # Store identity separately; no name is changed in the source
                # Zotero metadata. A pair-specific marker avoids confusing
                # two chapter authors with the same initial and surname.
                item["author"][position].setdefault("_verified_chapter_names", []).append(
                    verification_key(item, position, chapter_author))
        return resolved

    def write_report(self):
        fields = ["publication_key", "title", "doi", "chapter_author", "zotero_author",
                  "author_position", "status", "method", "detail"]
        with open(self.report_path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerows(self.report)


def verified_person(chapter_name, zotero_author, item=None, position=None):
    if same_person(chapter_name, zotero_author):
        return True
    if item is None or position is None:
        return False
    key = verification_key(item, position, chapter_name)
    return key in zotero_author.get("_verified_chapter_names", [])


def count_structured_author_occurrences(chapter_authors, zotero_items):
    rows = []
    total_first = total_co = 0

    for chapter_author in chapter_authors:
        first_count = co_count = 0

        for item in zotero_items:
            authors = item.get("author") or []
            match_position = None
            for idx, z_author in enumerate(authors):
                if verified_person(chapter_author, z_author, item, idx):
                    match_position = idx
                    break

            if match_position == 0:
                first_count += 1
            elif match_position is not None:
                co_count += 1

        rows.append((author_display_name(chapter_author),
                     first_count, co_count, first_count + co_count))
        total_first += first_count
        total_co += co_count

    return rows, total_first, total_co

def count_unique_group_self_citations(chapter_authors, zotero_items):
    """Inclusive first/co categories; union for the total, one count per work."""
    first = co = union = 0
    for item in zotero_items:
        positions = [
            idx for idx, author in enumerate(item.get("author") or [])
            if any(verified_person(name, author, item, idx) for name in chapter_authors)
        ]
        if not positions:
            continue
        union += 1
        first += int(0 in positions)
        co += int(any(idx > 0 for idx in positions))
    return first, co, union


def choose_debug_author(cla_la, ca):
    """Command-line selection, with human numbering starting at 1."""
    print("\nChoose the author group for the debug report:")
    print("  1 = CLA/LA")
    print("  2 = CA")
    while True:
        choice = input("Enter 1 or 2: ").strip()
        if choice in ("1", "2"):
            break
        print("Please enter 1 or 2.")

    authors = cla_la if choice == "1" else ca
    group_name = "CLA_LA" if choice == "1" else "CA"
    if not authors:
        print(f"No authors found in {group_name}; please select the other group.")
        return choose_debug_author(cla_la, ca)

    print(f"\n{group_name} authors:")
    for number, author in enumerate(authors, start=1):
        print(f"  {number}. {author_display_name(author)}")

    while True:
        answer = input(f"Enter author number (1-{len(authors)}): ").strip()
        try:
            number = int(answer)
        except ValueError:
            print("Please enter a whole number.")
            continue
        if 1 <= number <= len(authors):
            return authors[number - 1], group_name
        print(f"Please enter a number from 1 to {len(authors)}.")


def item_year(item):
    try:
        parts = (item.get("issued") or {}).get("date-parts") or []
        if parts and parts[0]:
            return str(parts[0][0])
    except Exception:
        pass
    return ""


def full_zotero_author_name(author):
    given = (author.get("given") or "").strip()
    family = (author.get("family") or "").strip()
    literal = (author.get("literal") or "").strip()
    if literal:
        return literal
    return " ".join(x for x in (given, family) if x).strip()


def debug_reference_text(item):
    """
    Human-readable reference reconstructed from embedded Zotero metadata.
    Crucially, this always shows the FULL detected author list.
    """
    authors = item.get("author") or []
    author_text = "; ".join(full_zotero_author_name(a) for a in authors if full_zotero_author_name(a))
    year = item_year(item)
    title = (item.get("title") or "").strip()
    container = (item.get("container-title") or "").strip()
    doi = (item.get("DOI") or "").strip()

    parts = []
    if author_text:
        parts.append(author_text)
    if year:
        parts.append(f"({year})")
    if title:
        parts.append(title)
    if container:
        parts.append(container)
    if doi:
        parts.append(f"DOI: {doi}")
    return ". ".join(parts)


def matched_items_for_author(chapter_author, zotero_items):
    """Return unique references where chapter_author is first author or co-author."""
    first_items = []
    co_items = []

    for item in zotero_items:
        authors = item.get("author") or []
        matched_position = None
        for idx, z_author in enumerate(authors):
            if verified_person(chapter_author, z_author, item, idx):
                matched_position = idx
                break

        if matched_position == 0:
            first_items.append(item)
        elif matched_position is not None:
            co_items.append(item)

    return first_items, co_items


def write_first_author_debug_doc(word, chapter_author, zotero_items, output_path):
    """
    Create a separate DOCX for the FIRST CLA/LA author.
    It lists every detected first-author and co-authored reference using the
    complete author arrays from Zotero itemData, not the rendered bibliography.
    """
    first_items, co_items = matched_items_for_author(chapter_author, zotero_items)

    debug_doc = word.Documents.Add()
    try:
        sel = word.Selection
        sel.HomeKey(Unit=6)  # wdStory

        sel.Style = debug_doc.Styles("Title")
        sel.TypeText("Self-citation Debug Report")
        sel.TypeParagraph()

        sel.Style = debug_doc.Styles("Normal")
        sel.TypeText(f"Chapter author checked: {author_display_name(chapter_author)}")
        sel.TypeParagraph()
        sel.TypeText("Matching: full first name + surname, plus confirmed ambiguous matches; middle names ignored.")
        sel.TypeParagraph()
        sel.TypeText(
            f"Detected matches: {len(first_items)} first-author references; "
            f"{len(co_items)} co-authored references."
        )
        sel.TypeParagraph()
        sel.TypeParagraph()

        sel.Style = debug_doc.Styles("Heading 1")
        sel.TypeText(f"First-author references ({len(first_items)})")
        sel.TypeParagraph()
        sel.Style = debug_doc.Styles("Normal")
        if first_items:
            for i, item in enumerate(first_items, 1):
                sel.TypeText(f"{i}. {debug_reference_text(item)}")
                sel.TypeParagraph()
        else:
            sel.TypeText("None detected.")
            sel.TypeParagraph()

        sel.TypeParagraph()
        sel.Style = debug_doc.Styles("Heading 1")
        sel.TypeText(f"Co-authored references ({len(co_items)})")
        sel.TypeParagraph()
        sel.Style = debug_doc.Styles("Normal")
        if co_items:
            for i, item in enumerate(co_items, 1):
                sel.TypeText(f"{i}. {debug_reference_text(item)}")
                sel.TypeParagraph()
        else:
            sel.TypeText("None detected.")
            sel.TypeParagraph()

        debug_doc.SaveAs2(output_path, FileFormat=16)  # wdFormatDocumentDefault (.docx)
    finally:
        debug_doc.Close(SaveChanges=False)

    return len(first_items), len(co_items)


def add_bold_label_and_text(doc, label, text):
    r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertAfter(label)
    start = r.Start
    end = start + len(label)
    doc.Range(start, end).Bold = True
    r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertAfter("\r" + text + "\r\r")


def add_heading(doc, text):
    r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    start = r.Start
    r.InsertAfter(text + "\r")
    doc.Range(start, start + len(text)).Bold = True


def add_stats_table(doc, title, rows, total_first, total_co, total_refs, unique_counts):
    add_heading(doc, title)

    insert_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(insert_range, len(rows) + 5, 4)

    headers = ["Author", "First-author articles", "Co-author articles", "Sum"]
    for c, h in enumerate(headers, 1):
        table.Cell(1, c).Range.Text = h
    table.Rows(1).Range.Bold = True

    row_idx = 2
    for author, first, co, total in rows:
        table.Cell(row_idx, 1).Range.Text = author
        table.Cell(row_idx, 2).Range.Text = str(first)
        table.Cell(row_idx, 3).Range.Text = str(co)
        table.Cell(row_idx, 4).Range.Text = str(total)
        row_idx += 1

    total_all = total_first + total_co
    table.Cell(row_idx, 1).Range.Text = "Total sum"
    table.Cell(row_idx, 2).Range.Text = str(total_first)
    table.Cell(row_idx, 3).Range.Text = str(total_co)
    table.Cell(row_idx, 4).Range.Text = str(total_all)
    table.Rows(row_idx).Range.Bold = True
    row_idx += 1

    def pct(n):
        return f"{(100.0 * n / total_refs):.1f}%" if total_refs else "0.0%"

    table.Cell(row_idx, 1).Range.Text = "Unique self-citation percentage"
    table.Cell(row_idx, 2).Range.Text = pct(total_first)
    table.Cell(row_idx, 3).Range.Text = pct(total_co)
    table.Cell(row_idx, 4).Range.Text = pct(total_all)
    table.Rows(row_idx).Range.Bold = True
    row_idx += 1

    table.Cell(row_idx, 1).Range.Text = "Total references"
    table.Cell(row_idx, 2).Range.Text = str(total_refs)
    table.Rows(row_idx).Range.Bold = True

    unique_first, unique_co, unique_total = unique_counts
    unique_row = len(rows) + 5
    table.Cell(unique_row, 1).Range.Text = "Unique group self-cited papers"
    table.Cell(unique_row, 2).Range.Text = str(unique_first)
    table.Cell(unique_row, 3).Range.Text = str(unique_co)
    table.Cell(unique_row, 4).Range.Text = str(unique_total)
    table.Rows(unique_row).Range.Bold = True

    # Group-level percentages are based on unique papers, not summed author matches.
    percentage_row = len(rows) + 3
    table.Cell(percentage_row, 2).Range.Text = pct(unique_first)
    table.Cell(percentage_row, 3).Range.Text = pct(unique_co)
    table.Cell(percentage_row, 4).Range.Text = pct(unique_total)
    table.Cell(percentage_row, 1).Range.Text = "Unique self-citation percentage"

    table.AutoFitBehavior(WD_AUTOFIT_CONTENT)

    end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end_range.InsertParagraphAfter()
    end_range.InsertParagraphAfter()


def choose_folder():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Choose folder containing original Word files")
    root.destroy()
    return folder


def is_source_docx(path):
    name = os.path.basename(path)
    lower = name.casefold()
    return (
        lower.endswith(".docx")
        and not name.startswith("~$")
        and "_reference_analysis" not in lower
        and "_debug_" not in lower
    )


def process_document(word, source, verifier, overwrite=False):
    """Process one source; close Word documents even when this file fails."""
    source = os.path.abspath(source)
    base, _ = os.path.splitext(source)
    output = base + "_Reference_Analysis.docx"
    if os.path.exists(output) and not overwrite:
        print(f"SKIPPED (output already exists): {source}")
        return "skipped"

    temp_dir = tempfile.mkdtemp(prefix="zotero_batch_")
    temp_source = os.path.join(temp_dir, os.path.basename(source))
    source_doc = out_doc = debug_doc = None
    try:
        shutil.copy2(source, temp_source)
        source_doc = word.Documents.Open(
            temp_source, ReadOnly=True, AddToRecentFiles=False,
            ConfirmConversions=False, Visible=False,
            OpenAndRepair=False, NoEncodingDialog=True
        )

        groups = find_author_sections(source_doc)
        cla_la = groups["cla"] + groups["la"]
        ca = groups["ca"]
        if not cla_la:
            raise RuntimeError("No CLA/LA authors found.")
        zotero_items, citation_field_count, parse_failures = extract_unique_zotero_items(source_doc)
        if not zotero_items:
            raise RuntimeError("No usable ZOTERO_ITEM metadata found.")
        bibl_field = find_zotero_bibliography_field(source_doc)
        if bibl_field is None:
            raise RuntimeError("No ZOTERO_BIBL field found.")

        zotero_items = verifier.resolve_items(zotero_items, cla_la, ca)
        la_rows, la_first, la_co = count_structured_author_occurrences(cla_la, zotero_items)
        ca_rows, ca_first, ca_co = count_structured_author_occurrences(ca, zotero_items)
        total_refs = len(zotero_items)

        # Select a debug author separately for each document (1-based menu).
        print(f"\nSelect a debug author for {os.path.basename(source)}:")
        debug_author, debug_group = choose_debug_author(cla_la, ca)
        safe_name = re.sub(r"[^\w.-]+", "_", debug_author, flags=re.UNICODE).strip("_")
        debug_output = base + f"_Reference_Analysis_DEBUG_{debug_group}_{safe_name}.docx"
        local_debug = os.path.join(temp_dir, os.path.basename(debug_output))
        dbg_first, dbg_co = write_first_author_debug_doc(
            word, debug_author, zotero_items, local_debug
        )

        out_doc = word.Documents.Add()
        add_bold_label_and_text(
            out_doc, "Coordinating Lead Authors and Lead Authors:",
            "; ".join(author_display_name(n) for n in cla_la)
        )
        add_bold_label_and_text(
            out_doc, "Contributing Authors:",
            "; ".join(author_display_name(n) for n in ca)
        )
        add_heading(out_doc, "References")
        dest = out_doc.Range(out_doc.Content.End - 1, out_doc.Content.End - 1)
        dest.FormattedText = bibl_field.Result.FormattedText
        end_r = out_doc.Range(out_doc.Content.End - 1, out_doc.Content.End - 1)
        end_r.InsertParagraphAfter()
        end_r.InsertParagraphAfter()

        add_stats_table(
            out_doc,
            "Coordinating Lead Authors and Lead Authors – self-citation statistics",
            la_rows, la_first, la_co, total_refs,
            count_unique_group_self_citations(cla_la, zotero_items)
        )
        add_stats_table(
            out_doc,
            "Contributing Authors – self-citation statistics",
            ca_rows, ca_first, ca_co, total_refs,
            count_unique_group_self_citations(ca, zotero_items)
        )

        local_output = os.path.join(temp_dir, os.path.basename(output))
        out_doc.SaveAs2(local_output, FileFormat=WD_FORMAT_DOCX)
        out_doc.Close(SaveChanges=False)
        out_doc = None
        source_doc.Close(SaveChanges=False)
        source_doc = None

        # Copy both results only after successful processing.
        shutil.copy2(local_output, output)
        shutil.copy2(local_debug, debug_output)

        print(f"  Unique cited references: {total_refs}")
        print(f"  Zotero citation fields: {citation_field_count}; unparsed: {parse_failures}")
        print(f"  Debug author: {debug_author} ({dbg_first} first; {dbg_co} co)")
        print(f"  Saved: {output}")
        print(f"  Saved: {debug_output}")
        return "success"
    finally:
        if out_doc is not None:
            try:
                out_doc.Close(SaveChanges=False)
            except Exception:
                pass
        if source_doc is not None:
            try:
                source_doc.Close(SaveChanges=False)
            except Exception:
                pass
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    folder = choose_folder()
    if not folder:
        print("No folder selected.")
        return
    sources = sorted(
        (os.path.join(folder, name) for name in os.listdir(folder)
         if is_source_docx(os.path.join(folder, name))),
        key=lambda path: os.path.basename(path).casefold()
    )
    if not sources:
        print("No eligible source .docx files found in the selected folder.")
        return

    print(f"Found {len(sources)} original Word files in: {folder}")
    for i, path in enumerate(sources, 1):
        print(f"  {i}. {os.path.basename(path)}")

    answer = input(
        "\nIf an analysis file already exists, replace it? [y/N]: "
    ).strip().casefold()
    overwrite = answer in ("y", "yes")
    print("\nYou will select a CLA/LA or CA debug author separately for each document.")

    default_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ipcc_author_database_partial.json")
    print("\nLocal IPCC database:", default_db)
    print("Found." if os.path.isfile(default_db) else "Not found; continuing with manual/cache/online verification.")
    ipcc_db = IPCCAuthorDatabase(default_db)
    print(f"Loaded {len(ipcc_db.by_doi)} DOI-indexed IPCC chapter records.")
    verifier = VerificationSession(folder, ipcc_db=ipcc_db)
    pythoncom.CoInitialize()
    word = None
    errors = []
    successes = skipped = 0
    old_options = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = True
        word.DisplayAlerts = 0
        old_options = (
            word.Options.CheckSpellingAsYouType,
            word.Options.CheckGrammarAsYouType,
            word.Options.CheckGrammarWithSpelling
        )
        word.Options.CheckSpellingAsYouType = False
        word.Options.CheckGrammarAsYouType = False
        word.Options.CheckGrammarWithSpelling = False
        word.ScreenUpdating = False

        for i, source in enumerate(sources, 1):
            print(f"\n[{i}/{len(sources)}] {os.path.basename(source)}")
            try:
                result = process_document(word, source, verifier, overwrite=overwrite)
                if result == "success":
                    successes += 1
                else:
                    skipped += 1
            except Exception as exc:
                errors.append((os.path.basename(source), str(exc)))
                print(f"  FAILED: {exc}")
    finally:
        if word is not None:
            try:
                if old_options is not None:
                    (word.Options.CheckSpellingAsYouType,
                     word.Options.CheckGrammarAsYouType,
                     word.Options.CheckGrammarWithSpelling) = old_options
                word.ScreenUpdating = True
            except Exception:
                pass
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    verifier.write_report()
    print(f"\nVerification cache: {verifier.path}")
    print(f"Verification report: {verifier.report_path}")
    print(f"\nFINISHED: {successes} succeeded, {skipped} skipped, {len(errors)} failed.")
    if errors:
        print("Failures:")
        for filename, error in errors:
            print(f"  {filename}: {error}")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
