"""
===============================================================================
finalise_v7.py: FINAL CHECKS BEFORE THE TURN-LEVEL BUILD
===============================================================================
Runs after apply_date_corrections.py.
Reads  harrison_v7_CLEAN4.json  (and the other sheets of CLEAN4.xlsx, such as
       Rejected, which are carried over)
Writes harrison_v7_CLEAN6.json  and  harrison_v7_CLEAN6.xlsx
CLEAN4 is never changed. Nothing is deleted and no date is changed. (CLEAN5
was a one-off addition to the corpus collected for this study and is not part
of this sequence; see the README.)

1. RECORD ID
   Records are numbered 1, 2, 3 ... in the order CLEAN4 already has (Grammy
   year, then act, then publication date), so in the delivered file a
   record's ID is its Excel row number minus one. The number stays with the
   record if the sheet is re-sorted. Rerunning this script on the same CLEAN4
   gives the same numbers. Once turns have been built, never renumber: a
   record added later takes the next unused number.

2. TURN SET EXCLUSION (new column; blank means the record goes forward)
   a. Index and listing pages: tag, category and author pages, month
      archives, paginated section pages, search and profile pages, and back-
      issue catalogues. Each is a list of excerpts from several posts, often
      about other artists, with no single publication date.
   b. Records that never name the act or any of its members anywhere in the
      title or full text: wrong-entity matches such as stray-cat welfare pages
      filed under Stray Cats or country-music news filed under Big Country.
   The turn set is then: high date confidence AND a blank exclusion.

3. MENTION COUNTS
   "Act Name Mentions" and "Member Mentions" (full names, aliases and a few
   distinctive surnames such as Knopfler or Setzer) for every record. The
   pipeline needs them: in a record that never names the act, a
   member's name is often a namesake (Nu Shooz's John Smith, Take 6's David
   Thomas), so it should not count as band voice without a check. For acts
   whose name is an ordinary word (Asia, Big Country), the act count is a
   count of the word, not proof the page is about the band.

4. LANGUAGE
   A language code for every record (py3langid), so non-English records can
   be found and handled.

5. GARBLED CHARACTERS
   The quote detection depends on these repairs. Each repaired record is
   noted in Notes.
   a. Full garbling ("â€œ" for a curly quote, "Ã©" for é) is repaired with
      ftfy's encoding fix. Where ftfy turns a damaged closing quote into a
      dagger (†) or a damaged ellipsis into "⦠", those are set right.
   b. Worse damage leaves only a lone "â" where a quote, apostrophe or dash
      was: "âDripping water,â she shows us". In English records with three
      or more of these, the punctuation is rebuilt from its position: inside
      a word it is an apostrophe, before a word an opening quote, after
      punctuation a closing quote, between spaces a dash, and after a word a
      closing quote if one is open, otherwise an apostrophe.

A "Turn set exclusions" sheet lists every excluded record with the reason,
so the exclusions can be spot-checked before the build.
===============================================================================
"""

import json
import os
import re

import ftfy
import pandas as pd
import py3langid as langid
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

VERSION = "finalise_v7 (28 Sep 2026: numbered IDs, punctuation rebuilt)"
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN4.json")
IN_XLSX = os.path.join(DIR, "harrison_v7_CLEAN4.xlsx")
OUT_JSON = os.path.join(DIR, "harrison_v7_CLEAN6.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_CLEAN6.xlsx")
EXCEL_CELL_CAP = 32000
REJECTED_TEXT_CAP = 1500

MEMBERS = {
    "Rickie Lee Jones": ["Rickie Lee Jones"],
    "Dire Straits": ["Mark Knopfler", "David Knopfler", "John Illsley", "Pick Withers"],
    "The Knack": ["Doug Fieger", "Berton Averre", "Prescott Niles", "Bruce Gary", "Burton Averre"],
    "Christopher Cross": ["Christopher Cross", "Chris Cross", "Christopher Charles Geppert"],
    "Irene Cara": ["Irene Cara", "Irene Cara Escalera"],
    "Robbie Dupree": ["Robbie Dupree", "Robert Dupuis"],
    "Amy Holland": ["Amy Holland"],
    "The Pretenders": ["Chrissie Hynde", "James Honeyman-Scott", "Pete Farndon", "Martin Chambers",
                       "Jimmy Scott", "Peter Farndon", "Christine Ellen Hynde", "Chrissy Hynde"],
    "Sheena Easton": ["Sheena Easton", "Sheena Shirley Orr"],
    "Adam and the Ants": ["Adam Ant", "Marco Pirroni", "Kevin Mooney", "Terry Lee Miall", "Chris Hughes",
                          "Stuart Goddard", "Stuart Leslie Goddard"],
    "The Go-Go's": ["Belinda Carlisle", "Jane Wiedlin", "Charlotte Caffey", "Kathy Valentine", "Gina Schock"],
    "James Ingram": ["James Ingram"],
    "Luther Vandross": ["Luther Vandross", "Luther Ronzoni Vandross"],
    "Men at Work": ["Colin Hay", "Ron Strykert", "Greg Ham", "John Rees", "Jerry Speiser"],
    "Asia": ["John Wetton", "Steve Howe", "Geoff Downes", "Carl Palmer", "Geoffrey Downes"],
    "Jennifer Holliday": ["Jennifer Holliday"],
    "The Human League": ["Philip Oakey", "Susan Ann Sulley", "Joanne Catherall", "Philip Adrian Wright",
                         "Ian Burden", "Jo Callis", "Phil Oakey", "Phillip Oakey", "Adrian Wright",
                         "Susanne Sulley", "Susan Sulley"],
    "Stray Cats": ["Brian Setzer", "Lee Rocker", "Slim Jim Phantom", "James McDonnell", "Leon Drucker", "Jim Phantom"],
    "Culture Club": ["Boy George", "Jon Moss", "Mikey Craig", "Roy Hay", "George O'Dowd",
                     "George Alan O'Dowd", "Michael Craig"],
    "Big Country": ["Stuart Adamson", "Bruce Watson", "Tony Butler", "Mark Brzezicki"],
    "Eurythmics": ["Annie Lennox", "Dave Stewart", "David A. Stewart", "Ann Lennox"],
    "Men Without Hats": ["Ivan Doroschuk", "Stefan Doroschuk", "Colin Doroschuk", "Allan McCarthy"],
    "Musical Youth": ["Dennis Seaton", "Kelvin Grant", "Michael Grant", "Patrick Waite", "Frederick Waite", "Junior Waite"],
    "Cyndi Lauper": ["Cyndi Lauper", "Cindy Lauper", "Cynthia Lauper"],
    "Sheila E.": ["Sheila E.", "Sheila Escovedo", "Sheila Cecilia Escovedo"],
    "Frankie Goes to Hollywood": ["Holly Johnson", "Paul Rutherford", "Peter Gill", "Mark O'Toole", "Brian Nash", "Nasher"],
    "Corey Hart": ["Corey Hart", "Corey Mitchell Hart"],
    "The Judds": ["Naomi Judd", "Wynonna Judd", "Diana Ellen Judd", "Christina Ciminella"],
    "Sade": ["Sade Adu", "Stuart Matthewman", "Paul S. Denman", "Andrew Hale", "Folasade Adu", "Helen Adu", "Paul Denman"],
    "a-ha": ["Morten Harket", "Magne Furuholmen", "Paul Waaktaar-Savoy", "Paul Waaktaar", "Pål Waaktaar", "Pal Waaktaar"],
    "Freddie Jackson": ["Freddie Jackson"],
    "Katrina and the Waves": ["Katrina Leskanich", "Kimberley Rew", "Vince de la Cruz", "Alex Cooper", "Kimberly Rew"],
    "Julian Lennon": ["Julian Lennon", "John Charles Julian Lennon"],
    "Bruce Hornsby and the Range": ["Bruce Hornsby", "John Molo", "George Marinelli", "Joe Puerta",
                                    "David Mansfield", "Joseph Puerta"],
    "Glass Tiger": ["Alan Frew", "Sam Reid", "Al Connelly", "Wayne Parker", "Michael Hanson"],
    "Nu Shooz": ["John Smith", "Valerie Day"],
    "Simply Red": ["Mick Hucknall", "Michael Hucknall", "Michael James Hucknall"],
    "Timbuk3": ["Pat MacDonald", "Barbara K. MacDonald"],
    "Jody Watley": ["Jody Watley"],
    "Breakfast Club": ["Dan Gilroy", "Ed Gilroy", "Gary Burke", "Stephen Bray", "Paul Kauk", "Steve Bray"],
    "Cutting Crew": ["Nick Van Eede", "Kevin MacMichael", "Colin Farley", "Martin Beedle", "Frosty Beedle"],
    "Terence Trent D'Arby": ["Terence Trent D'Arby", "Sananda Maitreya", "Terence Trent Howard",
                             "Trent D'Arby", "Francesco Maitreya"],
    "Swing Out Sister": ["Corinne Drewery", "Andy Connell", "Martin Jackson", "Andrew Connell", "Corrine Drewery"],
    "Tracy Chapman": ["Tracy Chapman"],
    "Rick Astley": ["Rick Astley", "Richard Paul Astley"],
    "Toni Childs": ["Toni Childs"],
    "Take 6": ["Claude McKnight", "Mark Kibble", "Joel Kibble", "David Thomas", "Alvin Chea", "Cedric Dent",
               "Mervyn Warren", "Dave Thomas", "Joey Kibble"],
    "Vanessa Williams": ["Vanessa Williams", "Vanessa Lynn Williams"],
    "Neneh Cherry": ["Neneh Cherry", "Neneh Mariann Karlsson"],
    "Indigo Girls": ["Amy Ray", "Emily Saliers"],
    "Soul II Soul": ["Jazzie B", "Nellee Hooper", "Caron Wheeler", "Beresford Romeo", "Nellie Hooper"],
    "Tone Loc": ["Tone Loc", "Anthony Terrell Smith", "Anthony Smith"],
}

# Surnames distinctive enough to show a page is about the act even when the
# full name is not repeated (journalists write "Knopfler", "Setzer"). Used for
# the mention count only, never to attribute speech.
SURNAMES = {
    "Dire Straits": ["Knopfler", "Illsley"], "The Knack": ["Fieger", "Averre"],
    "The Pretenders": ["Hynde", "Honeyman-Scott", "Farndon"], "Adam and the Ants": ["Pirroni"],
    "The Go-Go's": ["Wiedlin", "Caffey", "Schock"], "Men at Work": ["Strykert", "Speiser"],
    "Asia": ["Wetton"], "The Human League": ["Oakey", "Sulley", "Catherall"],
    "Stray Cats": ["Setzer"], "Big Country": ["Brzezicki"], "Men Without Hats": ["Doroschuk"],
    "Sheila E.": ["Escovedo"], "Sade": ["Matthewman"], "a-ha": ["Harket", "Furuholmen", "Waaktaar"],
    "Katrina and the Waves": ["Leskanich"], "Simply Red": ["Hucknall"],
    "Cutting Crew": ["MacMichael", "Van Eede"], "Swing Out Sister": ["Drewery"], "Take 6": ["Kibble"],
    "Indigo Girls": ["Saliers"], "Soul II Soul": ["Jazzie"], "Eurythmics": ["Lennox"],
    "Jody Watley": ["Watley"], "Robbie Dupree": ["Dupree"],
}

# Act-name patterns: (regex, case-sensitive?). Acts whose name is an ordinary
# word or phrase must appear capitalised (Title Case or ALL CAPS).
ACT_NAME = {
    "Dire Straits": (r"dire straits", False), "Rickie Lee Jones": (r"rickie lee jones", False),
    "The Knack": (r"\b(?:[Tt]he )?(?:Knack|KNACK)\b", True), "Amy Holland": (r"amy holland", False),
    "Christopher Cross": (r"christopher cross", False), "Irene Cara": (r"irene cara", False),
    "Robbie Dupree": (r"robbie dupree", False), "The Pretenders": (r"\bpretenders\b", False),
    "Adam and the Ants": (r"adam (?:and|&) the ants|adam ant\b", False),
    "James Ingram": (r"james ingram", False), "Luther Vandross": (r"vandross", False),
    "Sheena Easton": (r"sheena easton", False), "The Go-Go's": (r"go[- ]?go['’]?s\b", False),
    "Asia": (r"\b(?:Asia|ASIA)\b", True), "Jennifer Holliday": (r"jennifer holliday", False),
    "Men at Work": (r"men at work", False), "Stray Cats": (r"\b(?:Stray Cats?|STRAY CATS?)\b", True),
    "The Human League": (r"human league", False),
    "Big Country": (r"\b(?:Big Country|BIG COUNTRY)\b", True), "Culture Club": (r"culture club", False),
    "Eurythmics": (r"eurythmics", False), "Men Without Hats": (r"men without hats", False),
    "Musical Youth": (r"musical youth", False), "Corey Hart": (r"corey hart", False),
    "Cyndi Lauper": (r"lauper", False),
    "Frankie Goes to Hollywood": (r"frankie goes to hollywood|\bFGTH\b", False),
    "Sheila E.": (r"sheila e\b|sheila escovedo", False),
    "The Judds": (r"\bthe judds\b|naomi judd|wynonna", False),
    "Freddie Jackson": (r"freddie jackson", False), "Julian Lennon": (r"julian lennon", False),
    "Katrina and the Waves": (r"katrina (?:and|&) the waves", False),
    "Sade": (r"(?<![Dd]e )\b(?:Sade|SADE)\b", True), "a-ha": (r"\ba-ha\b", False),
    "Bruce Hornsby and the Range": (r"hornsby", False), "Glass Tiger": (r"glass tiger", False),
    "Nu Shooz": (r"nu shooz", False), "Simply Red": (r"simply red", False), "Timbuk3": (r"timbuk ?3", False),
    "Breakfast Club": (r"breakfast club", False), "Cutting Crew": (r"cutting crew", False),
    "Jody Watley": (r"jody watley", False), "Swing Out Sister": (r"swing out sister", False),
    "Terence Trent D'Arby": (r"d['’]arby|sananda maitreya", False), "Rick Astley": (r"astley", False),
    "Take 6": (r"\b(?:Take (?:6|Six)|TAKE (?:6|SIX))\b", True), "Toni Childs": (r"toni childs", False),
    "Tracy Chapman": (r"tracy chapman", False), "Vanessa Williams": (r"vanessa williams", False),
    "Indigo Girls": (r"indigo girls", False), "Neneh Cherry": (r"neneh cherry", False),
    "Soul II Soul": (r"soul (?:ii|2) soul", False), "Tone Loc": (r"tone l[oō]c", False),
}

INDEX_RULES = [
    ("tag, category or author page", re.compile(
        r"/(?:tag|tags|category|categories|author|authors|topics?|label|labels)/[^/?#]+/?(?:page/\d+/?)?(?:[?#].*)?$", re.I)),
    ("paginated listing", re.compile(
        r"/page/\d+/?(?:[?#].*)?$|(?:theguardian\.com|nme\.com)/[^?]*[?&]page=\d+", re.I)),
    ("date archive page", re.compile(
        r"^https?://[^/]+(?:/[a-z-]+)?/(?:19|20)\d{2}(?:/\d{2})?/?(?:[?#].*)?$", re.I)),
    ("search page", re.compile(r"/search(?:/|\?|$)", re.I)),
    ("profile page", re.compile(r"/profile/", re.I)),
    ("back-issue catalogue", re.compile(r"back-issues-a-z", re.I)),
]
NOT_INDEX = re.compile(r"smashinginterviews\.com/interviews/", re.I)
FORUM = re.compile(r"forum\.|/forums?/|/topic/|/thread|showthread|viewtopic|reddit\.com/r/", re.I)
GARBLED = re.compile(r"â€|Ã[\u0080-¿]|Â[ -¿]")
LONE_A = re.compile(r"â(?![\u0080-¿€™œ“”‘’¦])")


def repair_lone_a(text):
    """Rebuild punctuation that kept only its first byte and shows as 'â'."""
    out = []
    for p in text.split("\n"):
        if "â" not in p:
            out.append(p)
            continue
        p = p.replace("â¦", "…")
        p = re.sub(r"(?<=[A-Za-z])â(?=[A-Za-z])", "’", p)                 # it’s, Cherry’s
        p = re.sub(r"(?<=\s)â(?=\d0s\b|\d\ds\b)", "’", p)                  # the ’80s
        p = re.sub(r"(?<=\s)â(?=\s)", "—", p)                              # spaced dash
        p = re.sub(r"(?:(?<=^)|(?<=[\s(\[—]))â(?=[A-Za-z0-9])", "“", p)   # opening quote
        p = re.sub(r"(?<=[.,!?;:])â(?=\s|$|[)\]—])", "”", p)               # closing after punctuation
        chars = list(p)
        for m in re.finditer(r"(?<=[A-Za-z0-9])â(?=\s|$|[.,;:!?)\]—])", p):
            before = "".join(chars[:m.start()])
            chars[m.start()] = "”" if before.count("“") > before.count("”") else "’"
        out.append("".join(chars))
    return "\n".join(out)


def page_type(url):
    u = str(url)
    if FORUM.search(u) and "/profile/" not in u.lower():
        return "forum thread"
    if NOT_INDEX.search(u):
        return ""
    for name, rx in INDEX_RULES:
        if rx.search(u):
            return name
    return ""


def blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or \
        str(v).strip().lower() in ("", "nan", "none", "nat")


def text_of(v):
    return "" if blank(v) else str(v)


def member_regex(names):
    parts = []
    for n in names:
        p = re.escape(n).replace("'", "['’]").replace("\\'", "['’]")
        parts.append(p)
    return re.compile(r"(?<![\w-])(?:" + "|".join(parts) + r")(?![\w-])", re.I)


ACT_RX = {a: re.compile(p, 0 if cs else re.I) for a, (p, cs) in ACT_NAME.items()}
MEM_RX = {a: member_regex(ns + SURNAMES.get(a, [])) for a, ns in MEMBERS.items()}


def excel_safe(frame, text_cap):
    out = frame.copy()
    if "Extracted Text" in out.columns:
        out["Extracted Text"] = out["Extracted Text"].map(lambda v: text_of(v)[:text_cap])
    for c in out.columns:
        if pd.api.types.is_object_dtype(out[c]) or pd.api.types.is_string_dtype(out[c]):
            out[c] = out[c].map(lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v)
    return out


def main():
    print(VERSION)
    with open(IN_JSON, encoding="utf-8") as f:
        df = pd.DataFrame(json.load(f))
    df = df.reset_index(drop=True)
    print(f"Loaded {len(df)} records from CLEAN4 JSON")
    with pd.ExcelFile(IN_XLSX) as xf:
        sheets = {sh: pd.read_excel(xf, sheet_name=sh) for sh in xf.sheet_names if sh != "Corpus"}
    missing_acts = sorted(set(df["Focal Act"].astype(str)) - set(MEMBERS))
    if missing_acts:
        raise SystemExit(f"Acts with no member list in this script: {missing_acts}")

    # 1. record ID: 1..N in CLEAN4's existing order
    if df.duplicated(["Focal Act", "URL"]).any():
        raise SystemExit("Two records share an act and URL; stopping without writing anything.")
    df.insert(0, "Record ID", range(1, len(df) + 1))

    # 4. language, read before any repair (the garbling does not change it)
    df["Language"] = [langid.classify(text_of(t)[:3000])[0] if text_of(t).strip() else ""
                      for t in df["Extracted Text"]]

    # 5. garbled characters, repaired before counting so repaired names count
    def note(i, what):
        old = text_of(df.at[i, "Notes"]).strip()
        df.at[i, "Notes"] = (old + "; " if old else "") + what

    ftfy_fixed, rebuilt = 0, 0
    for i in df.index:
        t = text_of(df.at[i, "Extracted Text"])
        new = t
        if GARBLED.search(t):
            f = ftfy.fix_encoding(t)
            if "†" not in t:        # a closing quote that lost its last byte
                f = f.replace("†", "”")
            if "⦠" not in t:        # an ellipsis that lost its middle byte
                f = f.replace("⦠", "…")
            if f != t:
                new = f
                ftfy_fixed += 1
                note(i, "garbled characters repaired (ftfy)")
        if df.at[i, "Language"] == "en" and len(LONE_A.findall(new)) >= 3:
            f = repair_lone_a(new)
            if f != new:
                new = f
                rebuilt += 1
                note(i, "lost punctuation rebuilt from lone 'â'")
        if new != t:
            df.at[i, "Extracted Text"] = new
            df.at[i, "Text Length"] = len(new)

    # 3. mention counts, and page type
    act_n, mem_n, kinds = [], [], []
    for _, r in df.iterrows():
        a = str(r["Focal Act"])
        both = text_of(r["Title / Headline"]) + "\n" + text_of(r["Extracted Text"])
        act_n.append(len(ACT_RX[a].findall(both)))
        mem_n.append(len(MEM_RX[a].findall(both)))
        kinds.append(page_type(r["URL"]))
    df["Act Name Mentions"] = act_n
    df["Member Mentions"] = mem_n
    df["Page Type"] = kinds

    # 2. turn set exclusion
    index_page = df["Page Type"].isin([name for name, _ in INDEX_RULES])
    never = (df["Act Name Mentions"] == 0) & (df["Member Mentions"] == 0)
    df["Turn Set Exclusion"] = ""
    df.loc[never, "Turn Set Exclusion"] = "never names the act or a member"
    df.loc[index_page, "Turn Set Exclusion"] = "index or listing page"

    # write
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, default=str)
    high = df["Date Confidence"].astype(str).str.lower() == "high"
    excl = df[high & (df["Turn Set Exclusion"] != "")].copy()
    review = excl[["Record ID", "Focal Act", "Turn Set Exclusion", "Page Type", "Temporal Class",
                   "Title / Headline", "URL", "Act Name Mentions", "Member Mentions"]].copy()
    review["Opening text"] = excl["Extracted Text"].map(lambda v: text_of(v)[:300])
    review = review.sort_values(["Turn Set Exclusion", "Focal Act"])
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        excel_safe(df, EXCEL_CELL_CAP).to_excel(xw, sheet_name="Corpus", index=False)
        for sh, d in sheets.items():
            excel_safe(d, REJECTED_TEXT_CAP if sh == "Rejected" else EXCEL_CELL_CAP).to_excel(
                xw, sheet_name=sh, index=False)
        excel_safe(review, EXCEL_CELL_CAP).to_excel(xw, sheet_name="Turn set exclusions", index=False)

    # summary
    turn_set = high & (df["Turn Set Exclusion"] == "")
    member_only = turn_set & (df["Act Name Mentions"] == 0) & (df["Member Mentions"] > 0)
    print(f"\n{'=' * 64}")
    print(f"  records                    : {len(df)}  ({df['Focal Act'].nunique()} acts)")
    print(f"  high date confidence       : {int(high.sum())}")
    print(f"  excluded from the turn set : {len(excl)}")
    for reason, n in excl["Turn Set Exclusion"].value_counts().items():
        print(f"      {reason:32} {n}")
    print(f"  TURN SET                   : {int(turn_set.sum())}")
    print(f"  garbled text repaired      : {ftfy_fixed}  (ftfy)")
    print(f"  lost punctuation rebuilt   : {rebuilt}  (lone 'â')")
    print(f"  non-English in turn set    : {int((turn_set & (df['Language'] != 'en')).sum())}")
    print(f"  forum threads in turn set  : {int((turn_set & (df['Page Type'] == 'forum thread')).sum())}  (kept, flagged in Page Type)")
    print(f"  member named, act never    : {int(member_only.sum())}  (check before counting as band voice)")
    print("\n  Turn set by temporal class:")
    for k, n in df.loc[turn_set, "Temporal Class"].value_counts().items():
        print(f"      {k:18} {n}")
    print("\n  Exclusions by act (top 10):")
    for a, n in excl["Focal Act"].value_counts().head(10).items():
        print(f"      {a:28} {n}")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}")


if __name__ == "__main__":
    main()
