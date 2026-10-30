"""
===============================================================================
harrison_corpus_v7.py: COLLECT THE ARTICLE CORPUS
===============================================================================
Step 1 (articles) of the replication of Harrison, Askin and Hagtvedt (2023),
Administrative Science Quarterly 68(1). It collects publicly available
articles about the 52 acts nominated for the Grammy for Best New Artist from
1980 to 1990: the 55 nominees less Robin Williams, The Blues Brothers and
Milli Vanilli, as in the paper (p. 104). Solo acts are collected too: the
paper sampled all 52 and set solo artists aside only at the analysis stage.

HOW IT SEARCHES
  The paper searched "with the name of each band and band member" (p. 104).
  The queries do the same and nothing more:
    each act       "{act}", "{act}" interview, "{act}" Q&A,
                   "{act}" oral history
    each member    "{member}", "{member}" interview, "{member}" "{act}"
    each act       "{act}" site:<domain>, for the music archives, fan sites
                   and mainstream outlets listed below
  No query carries a theme from the paper, a year or an OR operator, and
  every query runs to completion for every act: there is no cap and no
  stopping rule. Results come from Serper (Google results, three pages per
  query) and from the Wayback Machine's index for nine outlets.

WHAT IT KEEPS
  Each page is fetched and its text, title, date and author are extracted
  with trafilatura, with JSON-LD, meta tags and the byline as fallbacks.
  Clear-cut cases are rejected: a failed or stub extraction, no act or member
  name in the text, a review with no interview marker, a navigation page, an
  announcement stub, a future date. Judgement calls are flagged and kept:
  the act may not be the subject, a listicle, the act named but no member,
  a page that looks like an archive of several articles. Every rejection is
  logged with its reason. Paywalled archives are listed as leads and never
  fetched. Each record also carries the date the page states in its own
  metadata (JSON-LD, meta tags or a <time> tag); clean_dates_v7.py uses it
  to confirm or correct the extracted date.

FIVE BATCHES (10, 10, 10, 10 and 12 acts), one per Colab session
  Each batch writes harrison_v7_batch<N>.xlsx and .json (the JSON keeps the
  full text) and a checkpoint, so a stopped batch resumes where it stopped.
  When all five batches are in the folder, the run merges them into
  harrison_v7_MERGED.xlsx and .json, keeping one record per act and URL (the
  same page can be a record for two acts).

COLAB
  !pip install trafilatura openpyxl pandas tqdm requests -q
  Store the Serper key in Colab Secrets as SERPER_API_KEY, set BATCH below,
  and run:
    from google.colab import drive, userdata
    drive.mount('/content/drive')
    import os; os.environ["SERPER_API_KEY"] = userdata.get("SERPER_API_KEY")
    %run /content/drive/MyDrive/harrison_v7/harrison_corpus_v7.py
  Repeat with BATCH = 1, 2, 3, 4 and 5.
===============================================================================
"""

import json
import os
import random
import re
import time
import warnings
from datetime import datetime, date

import pandas as pd
import requests
import trafilatura
from tqdm import tqdm

warnings.filterwarnings("ignore")

# ═════════════════════════════════════════════════════════════════════════════
# CONFIG
# ═════════════════════════════════════════════════════════════════════════════
# Prefer a Colab secret / environment variable so the key never appears in the
# notebook, the file, or a screenshot:
#     from google.colab import userdata
#     import os; os.environ["SERPER_API_KEY"] = userdata.get("SERPER_API_KEY")
SERPER_API_KEY = os.environ.get("SERPER_API_KEY", "YOUR_KEY_HERE")

BATCH = 1                      # 1-5. Run one batch per session.
USE_DRIVE = True               # checkpoint to Google Drive
DRIVE_DIR = "/content/drive/MyDrive/harrison_v7"

RUN_STRATIFIED_DATES = False   # before:/after: returns 0 hits through Serper —
                               # verified empirically. Left in place in case a
                               # working date-filter syntax is found.
RUN_SITE_QUERIES = True        # targeted site: queries
SEGMENT_AGGREGATES = False     # True splits a page that looks like an archive
                               # of several articles into one record per
                               # piece. Off, because the splitter also cut
                               # long single interviews into pieces: such
                               # pages are kept whole and flagged instead.
MERGE_WHEN_COMPLETE = True     # once all five batches are in the folder,
                               # merge them into harrison_v7_MERGED

SERPER_NUM = 10                # tiers cap organic results near 10; num=20 gave
                               # FEWER results in testing. Depth comes from
                               # SERPER_PAGES instead.
SERPER_PAGES = 3               # pages per query (~30 results). Raise for more
                               # recall at proportional credit cost.
DELAY_SEARCH = (0.6, 1.2)
DELAY_FETCH = (1.2, 2.5)
MIN_TEXT = 280                 # chars; below this an extraction is a stub

ERA_HINT_MIN, ERA_HINT_MAX = 1975, 1999
PERIOD_PUBYEAR = 2005          # pub year < this ⇒ period record
CURRENT_YEAR = date.today().year

# ═════════════════════════════════════════════════════════════════════════════
# SAMPLING FRAME — all 52 acts (55 BNA nominees 1980-1990 minus 3 exclusions)
# Excluded per Harrison p.104: Robin Williams, The Blues Brothers, Milli Vanilli
#
# Each act's members at the time of the nomination, checked by hand. A solo
# act takes the artist's own name (and any stage name) as its members.
# ═════════════════════════════════════════════════════════════════════════════
ACTS = {
    # ── 1980 ─────────────────────────────────────────────────────────────────
    "Rickie Lee Jones":            (1980, ["Rickie Lee Jones"]),
    "Dire Straits":                (1980, ["Mark Knopfler", "David Knopfler",
                                           "John Illsley", "Pick Withers"]),
    "The Knack":                   (1980, ["Doug Fieger", "Berton Averre",
                                           "Prescott Niles", "Bruce Gary"]),
    # ── 1981 ─────────────────────────────────────────────────────────────────
    "Christopher Cross":           (1981, ["Christopher Cross"]),
    "Irene Cara":                  (1981, ["Irene Cara"]),
    "Robbie Dupree":               (1981, ["Robbie Dupree"]),
    "Amy Holland":                 (1981, ["Amy Holland"]),
    "The Pretenders":              (1981, ["Chrissie Hynde", "James Honeyman-Scott",
                                           "Pete Farndon", "Martin Chambers"]),
    # ── 1982 ─────────────────────────────────────────────────────────────────
    "Sheena Easton":               (1982, ["Sheena Easton"]),
    "Adam and the Ants":           (1982, ["Adam Ant", "Marco Pirroni",
                                           "Kevin Mooney", "Terry Lee Miall",
                                           "Chris Hughes"]),
    "The Go-Go's":                 (1982, ["Belinda Carlisle", "Jane Wiedlin",
                                           "Charlotte Caffey", "Kathy Valentine",
                                           "Gina Schock"]),
    "James Ingram":                (1982, ["James Ingram"]),
    "Luther Vandross":             (1982, ["Luther Vandross"]),
    # ── 1983 ─────────────────────────────────────────────────────────────────
    "Men at Work":                 (1983, ["Colin Hay", "Ron Strykert", "Greg Ham",
                                           "John Rees", "Jerry Speiser"]),
    "Asia":                        (1983, ["John Wetton", "Steve Howe",
                                           "Geoff Downes", "Carl Palmer"]),
    "Jennifer Holliday":           (1983, ["Jennifer Holliday"]),
    "The Human League":            (1983, ["Philip Oakey", "Susan Ann Sulley",
                                           "Joanne Catherall", "Philip Adrian Wright",
                                           "Ian Burden", "Jo Callis"]),
    "Stray Cats":                  (1983, ["Brian Setzer", "Lee Rocker",
                                           "Slim Jim Phantom"]),
    # ── 1984 ─────────────────────────────────────────────────────────────────
    "Culture Club":                (1984, ["Boy George", "Jon Moss",
                                           "Mikey Craig", "Roy Hay"]),
    "Big Country":                 (1984, ["Stuart Adamson", "Bruce Watson",
                                           "Tony Butler", "Mark Brzezicki"]),
    "Eurythmics":                  (1984, ["Annie Lennox", "Dave Stewart"]),
    "Men Without Hats":            (1984, ["Ivan Doroschuk", "Stefan Doroschuk",
                                           "Colin Doroschuk", "Allan McCarthy"]),
    "Musical Youth":               (1984, ["Dennis Seaton", "Kelvin Grant",
                                           "Michael Grant", "Patrick Waite",
                                           "Frederick Waite"]),
    # ── 1985 ─────────────────────────────────────────────────────────────────
    "Cyndi Lauper":                (1985, ["Cyndi Lauper"]),
    "Sheila E.":                   (1985, ["Sheila E.", "Sheila Escovedo"]),
    "Frankie Goes to Hollywood":   (1985, ["Holly Johnson", "Paul Rutherford",
                                           "Peter Gill", "Mark O'Toole",
                                           "Brian Nash"]),
    "Corey Hart":                  (1985, ["Corey Hart"]),
    "The Judds":                   (1985, ["Naomi Judd", "Wynonna Judd"]),
    # ── 1986 ─────────────────────────────────────────────────────────────────
    "Sade":                        (1986, ["Sade Adu", "Stuart Matthewman",
                                           "Paul S. Denman", "Andrew Hale"]),
    "a-ha":                        (1986, ["Morten Harket", "Magne Furuholmen",
                                           "Paul Waaktaar-Savoy"]),
    "Freddie Jackson":             (1986, ["Freddie Jackson"]),
    "Katrina and the Waves":       (1986, ["Katrina Leskanich", "Kimberley Rew",
                                           "Vince de la Cruz", "Alex Cooper"]),
    "Julian Lennon":               (1986, ["Julian Lennon"]),
    # ── 1987 ─────────────────────────────────────────────────────────────────
    "Bruce Hornsby and the Range": (1987, ["Bruce Hornsby", "John Molo",
                                           "George Marinelli", "Joe Puerta",
                                           "David Mansfield"]),
    "Glass Tiger":                 (1987, ["Alan Frew", "Sam Reid", "Al Connelly",
                                           "Wayne Parker", "Michael Hanson"]),
    "Nu Shooz":                    (1987, ["John Smith", "Valerie Day"]),
    "Simply Red":                  (1987, ["Mick Hucknall"]),
    "Timbuk3":                     (1987, ["Pat MacDonald", "Barbara K. MacDonald"]),
    # ── 1988 ─────────────────────────────────────────────────────────────────
    "Jody Watley":                 (1988, ["Jody Watley"]),
    "Breakfast Club":              (1988, ["Dan Gilroy", "Ed Gilroy", "Gary Burke",
                                           "Stephen Bray", "Paul Kauk"]),
    "Cutting Crew":                (1988, ["Nick Van Eede", "Kevin MacMichael",
                                           "Colin Farley", "Martin Beedle"]),
    "Terence Trent D'Arby":        (1988, ["Terence Trent D'Arby", "Sananda Maitreya"]),
    "Swing Out Sister":            (1988, ["Corinne Drewery", "Andy Connell",
                                           "Martin Jackson"]),
    # ── 1989 ─────────────────────────────────────────────────────────────────
    "Tracy Chapman":               (1989, ["Tracy Chapman"]),
    "Rick Astley":                 (1989, ["Rick Astley"]),
    "Toni Childs":                 (1989, ["Toni Childs"]),
    "Take 6":                      (1989, ["Claude McKnight", "Mark Kibble",
                                           "Joel Kibble", "David Thomas",
                                           "Alvin Chea", "Cedric Dent"]),
    "Vanessa Williams":            (1989, ["Vanessa Williams"]),
    # ── 1990 ─────────────────────────────────────────────────────────────────
    "Neneh Cherry":                (1990, ["Neneh Cherry"]),
    "Indigo Girls":                (1990, ["Amy Ray", "Emily Saliers"]),
    "Soul II Soul":                (1990, ["Jazzie B", "Nellee Hooper",
                                           "Caron Wheeler"]),
    "Tone Loc":                    (1990, ["Tone Loc", "Anthony Terrell Smith"]),
}

_ordered = sorted(ACTS.items(), key=lambda kv: (kv[1][0], kv[0]))
BATCHES = {
    1: [a for a, _ in _ordered[0:10]],
    2: [a for a, _ in _ordered[10:20]],
    3: [a for a, _ in _ordered[20:30]],
    4: [a for a, _ in _ordered[30:40]],
    5: [a for a, _ in _ordered[40:52]],
}

# ═════════════════════════════════════════════════════════════════════════════
# BLOCKLIST: sites with no attributable, primary or extractable interview text
# ═════════════════════════════════════════════════════════════════════════════
SKIP = {
    # commerce / streaming / ticketing
    "amazon.", "spotify.", "apple.com", "itunes.", "bandcamp.",
    "ticketmaster.", "livenation.", "ebay.",
    # social
    "facebook.", "twitter.", "x.com/", "instagram.", "tiktok.",
    "reddit.com", "pinterest.",
    # reference databases (not primary sources)
    "wikipedia.org", "discogs.", "setlist.fm", "last.fm", "rateyourmusic.",
    "imdb.com", "fandom.com", "encyclopedia.com",
    # lyrics
    "genius.", "azlyrics.", "lyrics.",
    # AI-generated encyclopedias and quote farms
    "grokipedia.com", "citatis.com",
    # content farms
    "moa.gov.so", "medium.com", "albumcoverhalloffame",
    # video and audio pages: the speech is not in the page text (the videos
    # are collected and transcribed separately, from video_links_v2.py on)
    "youtube.com/watch", "soundcloud.",
}

# ═════════════════════════════════════════════════════════════════════════════
# SITE-TARGETED DOMAINS: extra queries on top of the open search, never instead of it
# ═════════════════════════════════════════════════════════════════════════════
TIER1_ARCHIVES = [  # specialist and fan archives, rich in period material
    "muzines.co.uk", "musicmusingsandsuch.com", "songfacts.com",
    "rediscoverthe80s.com", "earofnewt.com", "superdeluxeedition.com",
    "wbssmedia.com", "recordcollectormag.com", "innerviews.org",
    "electricityclub.co.uk", "totparchive.co.uk",
]
TIER2_PRESS = [     # mainstream press
    "rollingstone.com", "theguardian.com", "nme.com", "billboard.com",
    "npr.org", "latimes.com", "loudersound.com",
]
# Not queried: pitchfork.com, vibe.com, essence.com and spin.com returned
# nothing usable in trial runs.

WAYBACK_OUTLETS = [
    "rollingstone.com", "nme.com", "spin.com", "billboard.com",
    "mtv.com", "vh1.com", "sonicnet.com", "latimes.com", "ew.com",
]

# ═════════════════════════════════════════════════════════════════════════════
# ARCHIVE LEADS: paywalled or subscription archives whose pages the search
# finds but which cannot be retrieved. They are listed, never fetched: their
# terms do not allow automated retrieval or text mining.
# ═════════════════════════════════════════════════════════════════════════════
LEAD_DOMAINS = {
    "rocksbackpages.com": "Rock's Backpages",
    "proquest.com": "ProQuest",
    "gale.com": "Gale",
    "newspapers.com": "Newspapers.com",
    "britishnewspaperarchive.co.uk": "British Newspaper Archive",
    "pressreader.com": "PressReader",
}


def is_lead(url: str) -> str:
    u = (url or "").lower()
    for d, label in LEAD_DOMAINS.items():
        if d in u:
            return label
    return ""



TALK_VERBS = [
    "interview", "q&a", "q and a", "talks", "speaks", "conversation",
    "recalls", "remembers", "explains", "tells", "discusses", "reflects",
    "looks back", "oral history", "in his own words", "in her own words",
    "in their own words",
]
TALK_RE = re.compile("|".join(re.escape(v) for v in TALK_VERBS), re.I)

# Whole-title nav pages: anchored, tolerant of separators (dashes and pipes)
NAV_WORDS = r"(artists?|artist main|news|videos?|photos?|reviews?|home|music|" \
            r"mtv music|youtube videos|the week in rock|gallery|archive|tour)"
JUNK_WHOLE_RE = re.compile(rf"^\s*{NAV_WORDS}(\s*[-–—|]\s*{NAV_WORDS})*\s*$", re.I)
# Nav titles often carry the site or act name as one segment, e.g.
# "Videos — News — Nu Shooz". Two or more nav words as separate segments is
# the signal, regardless of what the other segments contain.
JUNK_SEGMENTS_RE = re.compile(rf"(?<![a-z]){NAV_WORDS}(?![a-z])", re.I)


def _nav_segment_count(title: str) -> int:
    segs = re.split(r"\s*[-–—|]\s*", title.strip())
    return sum(1 for s in segs if JUNK_SEGMENTS_RE.fullmatch(s.strip() or "x") )


JUNK_STRING_RE = re.compile(r"news,\s*videos,\s*photos|breaking news today|"
                            r"\|\s*artists?\s*$", re.I)
# These three need a talk-verb exemption (they match anywhere in a title)
JUNK_LOOSE_RE = re.compile(r"\bannounces?\b|\bbio\s*\||\|\s*[^|]*\bcareer\s*$", re.I)

REVIEW_TITLE_RE = re.compile(r"\breviews?\b", re.I)
# Gig-review body signals (a review titled only with a venue name passes
# title-only detection)
GIG_BODY_RE = re.compile(
    r"(box office:|tickets?:\s*0\d|·\s*at\s+.{3,40},\s+(london|new york|"
    r"manchester|glasgow)|\btour dates?\b.*\btickets\b)", re.I)

LISTICLE_RE = re.compile(r"^\s*(top\s+)?\d{1,3}\s+(best|greatest|worst|essential|"
                         r"iconic|classic|underrated)\b|\b\d{1,3}\s+(songs|albums|"
                         r"tracks|moments|facts|things)\b", re.I)

AGGREGATE_RE = re.compile(r"category:\s*interviews?|interview archive|"
                          r"^\s*interviews?\s*$|press archive|articles? archive",
                          re.I)


def norm_apos(s: str) -> str:
    """Curly → straight, so "Guns N' Roses" matches across encodings."""
    return (s or "").replace("\u2019", "'").replace("\u2018", "'") \
                    .replace("\u201c", '"').replace("\u201d", '"')


def skippable(url: str) -> bool:
    u = (url or "").lower()
    return any(d in u for d in SKIP)


def slug_variants(name: str) -> list:
    """'The Go-Go's' → ['the-go-gos', 'the-gogos', 'thegogos', 'go-gos', 'gogos']
    Two forms are generated: one where internal punctuation collapses away
    ('gogos') and one where it becomes a hyphen ('go-gos'), because sites use
    both. With both forms, /bands/the-go-gos/ is found as well as /bands/gogos/."""
    base = norm_apos(name).lower()
    collapsed = re.sub(r"[^a-z0-9\s]", "", base).strip()          # go gos
    hyphenated = re.sub(r"[^a-z0-9]+", "-", base).strip("-")      # the-go-go-s
    hyphenated = re.sub(r"-s(-|$)", r"s\1", hyphenated)           # the-go-gos
    if not collapsed:
        return []
    words = collapsed.split()
    out = ["-".join(words), "".join(words), hyphenated]
    if len(words) > 1 and words[0] == "the":
        out += ["-".join(words[1:]), "".join(words[1:]),
                re.sub(r"^the-", "", hyphenated)]
    return [v for v in dict.fromkeys(out) if len(v) >= 4]


def url_segment_match(url: str, name: str) -> bool:
    """Whole path segments only: 'sade' matches /artists/sade, not
    /artists/marquis-de-sade. Hyphens do NOT count as boundaries, or the
    lookbehind would let 'de-sade' through."""
    path = (url or "").lower()
    for v in slug_variants(name):
        if re.search(rf"(?<![a-z0-9-]){re.escape(v)}(?![a-z0-9-])", path):
            return True
    return False


# ═════════════════════════════════════════════════════════════════════════════
# QUERY CONSTRUCTION — theme-neutral, one intent per query, no OR operators
# ═════════════════════════════════════════════════════════════════════════════
GENRE_TERMS = ["interview", "Q&A", "oral history"]


def build_queries(act: str, members: list) -> list:
    """Name-based only. Mirrors Harrison p.104: 'the name of each band and
    band member'. NO theme/construct terms, NO year terms."""
    qs = [f'"{act}"']
    qs += [f'"{act}" {g}' for g in GENRE_TERMS]

    for m in members:
        if m.strip().lower() == act.strip().lower():
            continue                       # solo acts: member == act
        qs.append(f'"{m}"')
        qs.append(f'"{m}" interview')
        qs.append(f'"{m}" "{act}"')        # disambiguation + recall

    if RUN_SITE_QUERIES:
        for d in TIER1_ARCHIVES + TIER2_PRESS:
            qs.append(f'"{act}" site:{d}')

    return list(dict.fromkeys(qs))


def build_stratified_queries(act: str) -> list:
    """Date-windowed pass using metadata operators, NOT year terms in the
    query body. Windows run to the present so retrospectives are not excluded
    (Harrison's oral histories span full careers)."""
    windows = [("1975-01-01", "1990-01-01"), ("1990-01-01", "2000-01-01"),
               ("2000-01-01", "2010-01-01"), ("2010-01-01", "2020-01-01"),
               ("2020-01-01", f"{CURRENT_YEAR + 1}-01-01")]
    return [f'"{act}" interview after:{a} before:{b}' for a, b in windows]


# ═════════════════════════════════════════════════════════════════════════════
# FILTERS — return (action, reason). KEEP / FLAG / REJECT.
# Judgement calls FLAG; only clear-cut cases REJECT.
# ═════════════════════════════════════════════════════════════════════════════
KEEP, FLAG, REJECT = "keep", "flag", "reject"


def is_junk_page(title: str) -> tuple:
    t = norm_apos(title)
    if JUNK_WHOLE_RE.match(t) or JUNK_STRING_RE.search(t):
        return REJECT, "nav/hub page"
    if _nav_segment_count(t) >= 2 and not TALK_RE.search(t):
        return REJECT, "nav/hub page (multi-segment)"
    if JUNK_LOOSE_RE.search(t) and not TALK_RE.search(t):
        return REJECT, "announcement/bio-stub title, no talk verb"
    return KEEP, ""


def is_review(title: str, text: str) -> tuple:
    t = norm_apos(title)
    if REVIEW_TITLE_RE.search(t) and not TALK_RE.search(t):
        return REJECT, "review in title, no interview marker"
    if GIG_BODY_RE.search(text[-1200:]) and not TALK_RE.search(t):
        return REJECT, "gig-review body signal (box office / venue listing)"
    return KEEP, ""


def is_listicle(title: str) -> tuple:
    if LISTICLE_RE.search(norm_apos(title)) and not TALK_RE.search(title):
        return FLAG, "listicle/roundup format"      # flagged, not rejected
    return KEEP, ""


def wrong_artist(text: str, act: str, members: list) -> tuple:
    """Full member name, whole-word. A first-name match also counts
    when the act name is also confirmed on the page (single-name performers,
    informal fan transcripts)."""
    body = norm_apos(text)
    act_present = re.search(rf"(?<!\w){re.escape(norm_apos(act))}(?!\w)", body, re.I)
    for m in members:
        mn = norm_apos(m)
        if re.search(rf"(?<!\w){re.escape(mn)}(?!\w)", body, re.I):
            return KEEP, ""
        first = mn.split()[0]
        if act_present and len(first) > 2 and \
           re.search(rf"(?<!\w){re.escape(first)}(?!\w)", body, re.I):
            return KEEP, ""
    if act_present:
        return FLAG, "act named but no member name matched"
    return REJECT, "no act or member name in text"


def focal_dominant(title: str, text: str, act: str, members: list) -> tuple:
    """FLAG rather than REJECT: titles often omit the act, and a rejected
    record is invisible."""
    t, body = norm_apos(title).lower(), norm_apos(text).lower()
    names = [act] + list(members)
    in_title = any(norm_apos(n).lower() in t for n in names)
    mentions = sum(body.count(norm_apos(n).lower()) for n in names)
    if in_title and mentions >= 4:
        return KEEP, ""
    if mentions >= 8:
        return KEEP, ""
    return FLAG, f"low focal dominance ({mentions} mentions, title={in_title})"


def future_dated(pub_date: str) -> tuple:
    y = pub_date[:4]
    if y.isdigit() and int(y) > CURRENT_YEAR:
        return REJECT, f"future publication date ({y})"
    return KEEP, ""


def is_aggregate_page(title: str, text: str) -> bool:
    """One URL holding many pieces (an interview archive or category page,
    or a very long page naming many years). See SEGMENT_AGGREGATES."""
    if AGGREGATE_RE.search(norm_apos(title)):
        return True
    if len(text) > 60000 and len(re.findall(r"\b(19[89]\d|20[0-2]\d)\b", text)) > 12:
        return True
    return False


def segment_aggregate(text: str, act: str, members: list) -> list:
    """Split on heading-like lines; keep segments naming the act or a member.
    Returns [] when structure is too messy — caller then flags instead."""
    parts = re.split(r"\n(?=[A-Z][^\n]{12,90}\n)", text)
    if len(parts) < 3:
        return []
    names = [norm_apos(n).lower() for n in [act] + list(members)]
    out = [p.strip() for p in parts
           if len(p.strip()) >= MIN_TEXT
           and any(n in norm_apos(p).lower() for n in names)]
    return out if len(out) >= 2 else []


# ═════════════════════════════════════════════════════════════════════════════
# METADATA EXTRACTION — layered fallbacks.
# trafilatura alone leaves gaps; each layer below catches a different failure
# mode. The layer that succeeded is recorded in "Metadata Source" so fill
# quality is auditable rather than mysterious.
# ═════════════════════════════════════════════════════════════════════════════
_MONTHS = ("january|february|march|april|may|june|july|august|september|"
           "october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|"
           "oct|nov|dec")

DATE_META_KEYS = [
    "article:published_time", "og:published_time", "og:article:published_time",
    "publishdate", "publish-date", "pubdate", "date", "dc.date", "dc.date.issued",
    "dcterms.created", "dcterms.date", "sailthru.date", "parsely-pub-date",
    "cxenseparse:recs:publishtime", "timestamp", "article:modified_time",
]

AUTHOR_META_KEYS = [
    "author", "article:author", "og:article:author", "dc.creator",
    "parsely-author", "sailthru.author", "byl", "byline",
]


def _iso(y, mo, d) -> str:
    try:
        y, mo, d = int(y), int(mo), int(d)
        if 1900 <= y <= CURRENT_YEAR + 1 and 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{mo:02d}-{d:02d}"
    except (TypeError, ValueError):
        pass
    return ""


def _month_num(name: str) -> int:
    name = name.lower()[:3]
    return {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
            "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11,
            "dec": 12}.get(name, 0)


def _dates_from_jsonld(html: str) -> str:
    """JSON-LD blocks carry datePublished on most modern CMS output."""
    for block in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        for key in ("datePublished", "dateCreated", "uploadDate", "dateModified"):
            mt = re.search(rf'"{key}"\s*:\s*"([^"]{{4,40}})"', block)
            if mt:
                v = mt.group(1)
                iso = re.match(r"(\d{4})-(\d{2})-(\d{2})", v)
                if iso:
                    got = _iso(*iso.groups())
                    if got:
                        return got
    return ""


def _dates_from_meta(html: str) -> str:
    for key in DATE_META_KEYS:
        for pat in (rf'<meta[^>]+(?:property|name|itemprop)=["\']{re.escape(key)}["\']'
                    rf'[^>]+content=["\']([^"\']{{4,40}})["\']',
                    rf'<meta[^>]+content=["\']([^"\']{{4,40}})["\']'
                    rf'[^>]+(?:property|name|itemprop)=["\']{re.escape(key)}["\']'):
            mt = re.search(pat, html, re.I)
            if mt:
                v = mt.group(1)
                iso = re.match(r"(\d{4})-(\d{2})-(\d{2})", v)
                if iso:
                    got = _iso(*iso.groups())
                    if got:
                        return got
    return ""


def _dates_from_time_tag(html: str) -> str:
    mt = re.search(r'<time[^>]+datetime=["\'](\d{4})-(\d{2})-(\d{2})', html, re.I)
    return _iso(*mt.groups()) if mt else ""


def _dates_from_url(url: str) -> str:
    """Very common CMS pattern: /2015/03/15/slug or /2015-03-15-slug"""
    mt = re.search(r"/(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:/|-|$)", url)
    if mt:
        got = _iso(*mt.groups())
        if got:
            return got
    mt = re.search(r"/(\d{4})/(\d{1,2})/", url)
    if mt:
        got = _iso(mt.group(1), mt.group(2), 1)
        if got:
            return got
    return ""


def _dates_from_text(text: str) -> str:
    """Visible bylines: 'Published 3 March 1986', 'March 3, 1986'."""
    head = text[:3000]
    mt = re.search(rf"\b(\d{{1,2}})\s+({_MONTHS})\.?\s+(\d{{4}})\b", head, re.I)
    if mt:
        got = _iso(mt.group(3), _month_num(mt.group(2)), mt.group(1))
        if got:
            return got
    mt = re.search(rf"\b({_MONTHS})\.?\s+(\d{{1,2}}),?\s+(\d{{4}})\b", head, re.I)
    if mt:
        got = _iso(mt.group(3), _month_num(mt.group(1)), mt.group(2))
        if got:
            return got
    return ""


def _date_from_wayback(url: str) -> str:
    """web.archive.org/web/YYYYMMDDhhmmss/... — the capture date is an UPPER
    BOUND on publication, not the publication date. Flagged as such."""
    mt = re.search(r"web\.archive\.org/web/(\d{4})(\d{2})(\d{2})", url)
    return _iso(*mt.groups()) if mt else ""


# ═════════════════════════════════════════════════════════════════════════════
# THE PAGE'S OWN DATE METADATA, kept beside the extracted date
# extract_date() below takes trafilatura's date first and reads the page's
# structured metadata only when trafilatura has none, which is rare. So the
# metadata date is also read on its own, from the same HTML and in its own
# order of authority: JSON-LD (datePublished, dateCreated, uploadDate), then
# meta tags, then a <time> tag. clean_dates_v7.py uses it to confirm or
# correct the extracted date.
# ═════════════════════════════════════════════════════════════════════════════
META_DATE_KEYS = [
    "article:published_time", "og:published_time", "og:article:published_time",
    "publishdate", "publish-date", "pubdate", "date", "dc.date",
    "dc.date.issued", "dcterms.created", "dcterms.date", "sailthru.date",
    "parsely-pub-date", "cxenseparse:recs:publishtime", "timestamp",
]


def _meta_from_jsonld(html):
    for block in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        for key in ("datePublished", "dateCreated", "uploadDate"):
            m = re.search(rf'"{key}"\s*:\s*"(\d{{4}})-(\d{{2}})-(\d{{2}})', block)
            if m:
                got = _iso(*m.groups())
                if got:
                    return got, f"json-ld {key}"
    return "", ""


def _meta_from_tags(html):
    for key in META_DATE_KEYS:
        for pat in (rf'<meta[^>]+(?:property|name|itemprop)=["\']{re.escape(key)}["\']'
                    rf'[^>]+content=["\'](\d{{4}})-(\d{{2}})-(\d{{2}})',
                    rf'<meta[^>]+content=["\'](\d{{4}})-(\d{{2}})-(\d{{2}})'
                    rf'[^"\']*["\'][^>]+(?:property|name|itemprop)=["\']{re.escape(key)}["\']'):
            m = re.search(pat, html, re.I)
            if m:
                got = _iso(*m.groups())
                if got:
                    return got, f"meta {key}"
    return "", ""


def _meta_from_time_tag(html):
    m = re.search(r'<time[^>]+datetime=["\'](\d{4})-(\d{2})-(\d{2})', html, re.I)
    if m:
        got = _iso(*m.groups())
        if got:
            return got, "<time> tag"
    return "", ""


def metadata_date(html):
    """The date the page states in its own metadata, and where it was found."""
    for fn in (_meta_from_jsonld, _meta_from_tags, _meta_from_time_tag):
        got, src = fn(html)
        if got:
            return got, src
    return "", "no metadata date"


def extract_date(traf_date: str, html: str, text: str, url: str) -> tuple:
    """Returns (iso_date, method). Ordered most → least reliable."""
    if traf_date:
        iso = re.match(r"(\d{4})-(\d{2})-(\d{2})", traf_date)
        if iso:
            got = _iso(*iso.groups())
            if got:
                return got, "trafilatura"
        if re.fullmatch(r"\d{4}", traf_date.strip()):
            return f"{traf_date.strip()}-01-01", "trafilatura (year only)"
    for fn, label in ((_dates_from_jsonld, "json-ld"),
                      (_dates_from_meta, "meta tag"),
                      (_dates_from_time_tag, "<time> tag")):
        got = fn(html)
        if got:
            return got, label
    got = _dates_from_url(url)
    if got:
        return got, "url path"
    got = _dates_from_text(text)
    if got:
        return got, "body text"
    got = _date_from_wayback(url)
    if got:
        return got, "wayback capture (UPPER BOUND, not publication date)"
    return "", "none"


def _author_from_jsonld(html: str) -> str:
    for block in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        mt = re.search(r'"author"\s*:\s*\{[^}]*"name"\s*:\s*"([^"]{2,60})"', block)
        if mt:
            return mt.group(1).strip()
        mt = re.search(r'"author"\s*:\s*"([^"]{2,60})"', block)
        if mt:
            return mt.group(1).strip()
    return ""


def _author_from_meta(html: str) -> str:
    for key in AUTHOR_META_KEYS:
        mt = re.search(rf'<meta[^>]+(?:property|name)=["\']{re.escape(key)}["\']'
                       rf'[^>]+content=["\']([^"\']{{2,60}})["\']', html, re.I)
        if mt:
            return mt.group(1).strip()
    return ""


def _author_from_text(text: str) -> str:
    # [ \t]+ not \s+ : \s crosses newlines and swallows the first body word
    # ("By Barney Hoskyns\n\nThe band..." → "Barney Hoskyns The").
    mt = re.search(r"\b[Bb]y[:\t ]+([A-Z][a-z]+(?:[ \t]+[A-Z][a-z'’-]+){1,2})",
                   text[:1200])
    if not mt:
        return ""
    name = re.sub(r"\s+", " ", mt.group(1)).strip()
    return name if 4 <= len(name) <= 60 else ""


def extract_author(traf_author: str, html: str, text: str) -> tuple:
    if traf_author:
        return traf_author.strip(), "trafilatura"
    for fn, label in ((_author_from_jsonld, "json-ld"),
                      (_author_from_meta, "meta tag"),
                      (_author_from_text, "byline text")):
        got = fn(html) if fn is not _author_from_text else fn(text)
        if got:
            return got, label
    return "", "none"



# ═════════════════════════════════════════════════════════════════════════════
# METADATA: assigned only on positive evidence; blank otherwise
# ═════════════════════════════════════════════════════════════════════════════
def source_type(title: str, text: str) -> str:
    t = norm_apos(title).lower()
    if re.search(r"\bq&a\b|\bq and a\b", t):
        return "Q&A interview"
    if "interview" in t or "oral history" in t:
        return "interview"
    if TALK_RE.search(t):
        return "interview"
    head = norm_apos(text)[:2000]
    if re.search(r"^\s*(Q|Interviewer|I):\s", head, re.M):
        return "Q&A interview"
    return ""                                   # blank, not a default


def speakers_found(text: str, members: list) -> str:
    body = norm_apos(text)
    hits = [m for m in members
            if re.search(rf"(?<!\w){re.escape(norm_apos(m))}(?!\w)", body, re.I)]
    return "; ".join(hits)                      # blank when none


def speaker_status(speakers: str, source_t: str) -> str:
    """Positive evidence only: a status is assigned when a member was actually
    matched in the text. Blank otherwise: no silent default."""
    if not speakers:
        return ""
    return "band/act member" if source_t else "band/act member (mentioned)"


def content_era_hint(text: str) -> str:
    years = re.findall(r"\b(19[789]\d)\b", norm_apos(text)[:1500])
    years = [y for y in years if ERA_HINT_MIN <= int(y) <= ERA_HINT_MAX]
    return min(years) if years else ""


def is_period_record(pub_date: str, era_hint: str) -> bool:
    y = pub_date[:4]
    if y.isdigit() and int(y) < PERIOD_PUBYEAR:
        return True
    return bool(era_hint)


# ═════════════════════════════════════════════════════════════════════════════
# RETRIEVAL
# ═════════════════════════════════════════════════════════════════════════════
HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/124.0 Safari/537.36")}


class SearchAborted(RuntimeError):
    """Raised when the search backend fails in a way that would silently
    produce an empty corpus (bad key, exhausted credits, sustained blocking).
    The run stops loudly rather than finish a batch with nothing."""


_serper_stats = {"ok": 0, "empty": 0, "consecutive_fail": 0}


def search_serper(query: str, pages: int = None) -> list:
    """Free/most tiers return ~10 organic results per call regardless of `num`,
    so depth comes from paging, not from a larger num."""
    pages = SERPER_PAGES if pages is None else pages
    out = []
    for page in range(1, pages + 1):
        out.extend(_serper_once(query, page))
    return list(dict.fromkeys(out))


def _serper_once(query: str, page: int = 1) -> list:
    if not SERPER_API_KEY or SERPER_API_KEY == "YOUR_KEY_HERE":
        raise SearchAborted("Set SERPER_API_KEY before running.")
    time.sleep(random.uniform(*DELAY_SEARCH))

    last_error = None
    for attempt in range(4):
        try:
            r = requests.post("https://google.serper.dev/search",
                              headers={"X-API-KEY": SERPER_API_KEY,
                                       "Content-Type": "application/json"},
                              json={"q": query, "num": SERPER_NUM, "page": page,
                                    "gl": "us", "hl": "en"}, timeout=25)
        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            time.sleep(3 * (attempt + 1))
            continue

        # Fatal: never silently continue on these.
        if r.status_code in (401, 403):
            raise SearchAborted(f"Serper auth failed ({r.status_code}). "
                                f"Check SERPER_API_KEY. Body: {r.text[:200]}")
        if r.status_code in (402, 429) and attempt == 3:
            raise SearchAborted(
                f"Serper returned {r.status_code} after retries "
                "(out of credits or rate limited). Stopping so the batch is "
                "not silently truncated. Checkpoint is saved — resume later.")
        if r.status_code in (402, 429):
            time.sleep(20 * (attempt + 1))
            continue
        if r.status_code != 200:
            last_error = f"HTTP {r.status_code}: {r.text[:200]}"
            time.sleep(3 * (attempt + 1))
            continue

        try:
            data = r.json()
        except Exception:
            return []

        links = [it.get("link", "") for sec in ("organic", "news")
                 for it in data.get(sec, []) if it.get("link")]
        if links:
            _serper_stats["ok"] += 1
            _serper_stats["consecutive_fail"] = 0
        else:
            _serper_stats["empty"] += 1
            _serper_stats["consecutive_fail"] += 1
            # A genuine zero-result query is normal; 40 in a row is not.
            if _serper_stats["consecutive_fail"] >= 40:
                raise SearchAborted(
                    "40 consecutive empty Serper responses — the backend is "
                    "almost certainly failing rather than the queries being "
                    "genuinely empty. Stopping. Checkpoint is saved.")
        return links

    _serper_stats["consecutive_fail"] += 1
    _serper_stats["last_error"] = last_error
    if last_error:
        print(f"    [serper] all retries failed: {last_error}")
    return []


def search_wayback(act: str, members: list, limit: int = 40) -> list:
    """CDX domain match + urlkey regex on act/member slugs."""
    variants = []
    for n in [act] + list(members):
        variants.extend(slug_variants(n))
    variants = list(dict.fromkeys(variants))[:8]
    if not variants:
        return []
    pattern = ".*(" + "|".join(re.escape(v) for v in variants) + ").*"
    urls = []
    for outlet in WAYBACK_OUTLETS:
        time.sleep(random.uniform(*DELAY_SEARCH))
        for attempt in range(3):
            try:
                r = requests.get("https://web.archive.org/cdx/search/cdx",
                                 params={"url": outlet, "matchType": "domain",
                                         "output": "json",
                                         "fl": "timestamp,original",
                                         "filter": ["statuscode:200",
                                                    f"urlkey:{pattern}"],
                                         "collapse": "urlkey",
                                         "limit": str(limit)}, timeout=45)
                if r.status_code == 200 and r.text.strip():
                    rows = r.json()
                    for ts, orig in rows[1:]:
                        urls.append(f"https://web.archive.org/web/{ts}id_/{orig}")
                break
            except Exception:
                time.sleep(2 * (attempt + 1))
    return urls


def fetch_and_extract(url: str) -> dict:
    """Retries transient failures, then tries two trafilatura strategies before
    giving up. Runs the full metadata fallback chain on the raw HTML, which
    trafilatura alone does not do."""
    html = None
    for attempt in range(3):
        time.sleep(random.uniform(*DELAY_FETCH))
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200 and r.text:
                html = r.text
                break
            if r.status_code in (403, 404, 410):
                return None                      # permanent, don't retry
        except Exception:
            pass
        time.sleep(2 * (attempt + 1))
    if not html:
        return None

    # Two extraction passes: recall-first, then precision as a fallback.
    d = None
    for kwargs in ({"favor_recall": True}, {"favor_precision": True}):
        try:
            raw = trafilatura.extract(html, output_format="json",
                                      with_metadata=True, include_comments=False,
                                      **kwargs)
            if raw:
                cand = json.loads(raw)
                if len((cand.get("text") or "").strip()) >= MIN_TEXT:
                    d = cand
                    break
        except Exception:
            continue
    if d is None:
        return None

    text = norm_apos((d.get("text") or "").strip())
    if len(text) < MIN_TEXT:
        return None

    pub_date, date_method = extract_date(
        (d.get("date") or "").strip(), html, text, url)
    author, author_method = extract_author(
        (d.get("author") or "").strip(), html, text)

    meta_date, meta_src = metadata_date(html)

    title = norm_apos((d.get("title") or "").strip())
    if not title:                                 # last-resort title
        mt = re.search(r"<title[^>]*>(.{3,200}?)</title>", html, re.S | re.I)
        title = norm_apos(mt.group(1).strip()) if mt else ""

    return {"title": title,
            "author": author,
            "date": pub_date,
            "site": (d.get("hostname") or d.get("sitename") or "").strip(),
            "text": text,
            "date_method": date_method,
            "author_method": author_method,
            "meta_date": meta_date,
            "meta_src": meta_src}


# ═════════════════════════════════════════════════════════════════════════════
# RECORD ASSEMBLY
# ═════════════════════════════════════════════════════════════════════════════
COLUMNS = ["Focal Act", "Grammy Year", "Title / Headline", "Outlet / Source Name",
           "URL", "Source Type", "Publication Date", "Interview / Event Date",
           "Speaker Name(s)", "Speaker Status", "Author / Interviewer",
           "Extracted Text", "Text Length", "Provenance", "Metadata Source",
           "Content Era Hint", "Period Record", "Review Flag", "Retrieved On",
           "Query", "Notes",
           "Metadata Date (verified)", "Metadata Date Source", "Fetch Status"]


def make_record(act, gyear, d, url, query, flags, text_override=None, note=""):
    text = text_override if text_override is not None else d["text"]
    era = content_era_hint(text)
    stype = source_type(d["title"], text)
    spk = speakers_found(text, ACTS[act][1])
    return {
        "Focal Act": act,
        "Grammy Year": gyear,
        "Title / Headline": d["title"],
        "Outlet / Source Name": d["site"] or (url.split("/")[2] if "//" in url else ""),
        "URL": url,
        "Source Type": stype,
        "Publication Date": d["date"],
        "Interview / Event Date": "",
        "Speaker Name(s)": spk,
        "Speaker Status": speaker_status(spk, stype),
        "Author / Interviewer": d["author"],
        "Extracted Text": text,
        "Text Length": len(text),
        "Provenance": "open-web",
        "Metadata Source": f"date:{d.get('date_method','?')}; "
                           f"author:{d.get('author_method','?')}",
        "Content Era Hint": era,
        "Period Record": is_period_record(d["date"], era),
        "Review Flag": "; ".join(flags),
        "Retrieved On": date.today().isoformat(),
        "Query": query,
        "Notes": note,
        "Metadata Date (verified)": d.get("meta_date", ""),
        "Metadata Date Source": d.get("meta_src", "no metadata date"),
        "Fetch Status": "200",
    }


def process_act(act: str) -> tuple:
    """URL dedup is PER ACT, not global: the same page can legitimately be a
    record for two different focal acts (e.g. a piece covering both). Global
    dedup would give it arbitrarily to whichever act ran first."""
    gyear, members = ACTS[act]
    queries = build_queries(act, members)
    if RUN_STRATIFIED_DATES:
        queries += build_stratified_queries(act)

    # 1. collect candidate URLs
    candidates, leads = {}, []
    for q in tqdm(queries, desc=f"{act[:22]:22s} search", leave=False):
        for u in search_serper(q):
            if not u or u in candidates:
                continue
            label = is_lead(u)
            if label:
                leads.append((act, label, u, q))     # capture, never scrape
                continue
            if not skippable(u):
                candidates[u] = q
    for u in search_wayback(act, members):
        if u and u not in candidates and not skippable(u):
            candidates[u] = "Wayback CDX"

    # 2. fetch, filter, assemble
    records, rejects = [], []
    seen = set()
    for url, q in tqdm(list(candidates.items()),
                       desc=f"{act[:22]:22s} fetch ", leave=False):
        if url in seen:
            continue
        d = fetch_and_extract(url)
        if d is None:
            rejects.append((act, "", url, q, "", "fetch/extract failed or under MIN_TEXT"))
            continue

        flags, rejected = [], None
        for action, reason in (
            is_junk_page(d["title"]),
            is_review(d["title"], d["text"]),
            future_dated(d["date"]),
            wrong_artist(d["text"], act, members),
        ):
            if action == REJECT:
                rejected = reason
                break
            if action == FLAG:
                flags.append(reason)

        if rejected:
            rejects.append((act, d["title"], url, q, str(len(d["text"])), rejected))
            continue

        for action, reason in (is_listicle(d["title"]),
                               focal_dominant(d["title"], d["text"], act, members)):
            if action == FLAG:
                flags.append(reason)

        if is_aggregate_page(d["title"], d["text"]):
            if SEGMENT_AGGREGATES:
                segs = segment_aggregate(d["text"], act, members)
                if segs:
                    for i, seg in enumerate(segs, 1):
                        records.append(make_record(
                            act, gyear, d, f"{url}#seg{i}", q, flags, seg,
                            note=f"segmented from aggregate page ({i}/{len(segs)})"))
                    seen.add(url)
                    continue
                flags.append("aggregate page, segmentation failed")
            else:
                flags.append("aggregate page, kept whole")

        records.append(make_record(act, gyear, d, url, q, flags))
        seen.add(url)

    return records, rejects, leads


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════
def quality_report(df):
    """Printed after every batch so metadata gaps are visible immediately,
    not discovered weeks later."""
    if df.empty:
        print("\n!! No records — investigate before running further batches.")
        return
    n = len(df)
    print(f"\n{'='*62}\nQUALITY REPORT — {n} records\n{'='*62}")

    print("\nField fill rates:")
    for col in ["Title / Headline", "Publication Date", "Author / Interviewer",
                "Source Type", "Speaker Name(s)", "Outlet / Source Name"]:
        filled = df[col].astype(str).str.strip().replace("nan", "").ne("").sum()
        bar = "#" * int(20 * filled / n)
        flag = "  <-- LOW" if filled / n < 0.75 else ""
        print(f"  {col:24} {filled:4}/{n} ({100*filled/n:3.0f}%) {bar:20}{flag}")

    print("\nHow publication dates were obtained:")
    methods = (df["Metadata Source"].astype(str)
               .str.extract(r"date:([^;]+)")[0].fillna("?"))
    for meth, cnt in methods.value_counts().items():
        note = "  (upper bound only)" if "wayback" in meth else ""
        print(f"  {meth:42} {cnt:4}{note}")

    undated = df[df["Publication Date"].astype(str).str.strip().eq("")]
    if len(undated):
        print(f"\n{len(undated)} records with NO date — top domains:")
        for dom, c in undated["Outlet / Source Name"].value_counts().head(5).items():
            print(f"  {dom:40} {c}")

    print(f"\nPeriod records: {int(df['Period Record'].sum())}/{n} "
          f"({100*df['Period Record'].sum()/n:.0f}%)")
    flagged = df["Review Flag"].astype(str).str.strip().ne("").sum()
    print(f"Flagged for review: {flagged}/{n} ({100*flagged/n:.0f}%)")
    if flagged:
        reasons = (df[df["Review Flag"].ne("")]["Review Flag"]
                   .str.split("; ").explode().value_counts())
        for r, c in reasons.head(6).items():
            print(f"    {r[:52]:52} {c}")
    print(f"\nMedian text length: {int(df['Text Length'].median())} chars")
    print("="*62)



# ═════════════════════════════════════════════════════════════════════════════
# MERGE: once all five batches are in the folder
# Reads each batch's JSON (full text; the .xlsx copies are cut at Excel's
# 32,767-character cell limit), keeps one record per (Focal Act, URL), checks
# that all 52 acts are present, and writes harrison_v7_MERGED.xlsx and .json.
# ═════════════════════════════════════════════════════════════════════════════
EXPECTED_ACTS = 52
EXCEL_CELL_CAP = 32000


def _nonblank(series):
    """Count non-empty values. Handles pandas' StringDtype, whose nulls are
    pd.NA rather than the string "nan"."""
    return (series.fillna("").astype(str).str.strip()
            .replace({"nan": "", "None": "", "<NA>": ""}).ne("").sum())


def _load_batch(dir_, i):
    """The batch's JSON (full text) if present, else its workbook."""
    js = os.path.join(dir_, f"harrison_v7_batch{i}.json")
    xl = os.path.join(dir_, f"harrison_v7_batch{i}.xlsx")
    if os.path.exists(js):
        with open(js, encoding="utf-8") as f:
            recs = json.load(f)
        df = pd.DataFrame(recs)
        src = "json (full text)"
    elif os.path.exists(xl):
        df = pd.read_excel(xl, sheet_name="Corpus")
        src = "xlsx (TEXT TRUNCATED at 32k)"
    else:
        print(f"  batch {i}: MISSING — neither .json nor .xlsx found")
        return None, None, None, None
    rej = lead = None
    if os.path.exists(xl):
        try:
            rej = pd.read_excel(xl, sheet_name="Rejects")
            lead = pd.read_excel(xl, sheet_name="Archive Leads")
        except Exception:
            pass
    print(f"  batch {i}: {len(df):5} records from {src}")
    return df, rej, lead, src


def merge_batches(dir_):
    """Combines the five batches into harrison_v7_MERGED.xlsx and .json in
    dir_, or says which batches are still missing and writes nothing."""
    missing = [i for i in sorted(BATCHES)
               if not (os.path.exists(os.path.join(dir_, f"harrison_v7_batch{i}.json"))
                       or os.path.exists(os.path.join(dir_, f"harrison_v7_batch{i}.xlsx")))]
    if missing:
        print(f"\nNot merging yet: batch(es) {missing} not found in {dir_}. The run that "
              "completes the last batch merges all five.")
        return None
    out_xlsx = os.path.join(dir_, "harrison_v7_MERGED.xlsx")
    out_json = os.path.join(dir_, "harrison_v7_MERGED.json")

    print("=" * 66)
    print("MERGING v7 BATCHES")
    print("=" * 66)

    corpora, rejects, leads = [], [], []
    for i in sorted(BATCHES):
        df, rej, lead, _ = _load_batch(dir_, i)
        if df is None:
            continue
        df["Batch"] = i
        corpora.append(df)
        if rej is not None and len(rej):
            rej = rej.copy(); rej["Batch"] = i; rejects.append(rej)
        if lead is not None and len(lead):
            lead = lead.copy(); lead["Batch"] = i; leads.append(lead)

    if not corpora:
        print("\nNo batches found. Check the folder and that Drive is mounted.")
        return None

    df = pd.concat(corpora, ignore_index=True)
    print(f"\nCombined before dedup: {len(df)} records")

    # ── dedup on (Focal Act, URL) ────────────────────────────────────────────
    before = len(df)
    df = df.drop_duplicates(subset=["Focal Act", "URL"], keep="first")
    print(f"After dedup on (Focal Act, URL): {len(df)}  [-{before - len(df)}]")

    cross = (df.groupby("URL")["Focal Act"].nunique() > 1).sum()
    print(f"URLs legitimately shared across acts: {cross} (kept — different focal acts)")

    # ── verify coverage ──────────────────────────────────────────────────────
    acts = sorted(df["Focal Act"].unique())
    print(f"\nActs present: {len(acts)} / {EXPECTED_ACTS}")
    if len(acts) != EXPECTED_ACTS:
        print("  !! ACT COUNT MISMATCH — check for a missing or partial batch")
    thin = df.groupby("Focal Act").size().sort_values()
    print("\nThinnest acts:")
    for a, n in thin.head(5).items():
        print(f"    {a:32} {n}")
    print("Richest acts:")
    for a, n in thin.tail(5).items():
        print(f"    {a:32} {n}")

    # ── sort: Grammy year → act → publication date ───────────────────────────
    df["Grammy Year"] = pd.to_numeric(df["Grammy Year"], errors="coerce")
    df = df.sort_values(["Grammy Year", "Focal Act", "Publication Date"],
                        na_position="last").reset_index(drop=True)

    # ── Summary ──────────────────────────────────────────────────────────────
    def truthy(s):
        return sum(1 for v in s if v is True or str(v).lower() == "true")

    summary = (df.groupby(["Focal Act", "Grammy Year"])
                 .apply(lambda g: pd.Series({
                     "Records": len(g),
                     "Dated": _nonblank(g["Publication Date"]),
                     "Period (flagged)": truthy(g["Period Record"]),
                     "Flagged for review": _nonblank(g["Review Flag"]),
                     "Median chars": int(g["Text Length"].median()),
                     "Batch": int(g["Batch"].iloc[0]),
                 }), include_groups=False)
                 .reset_index()
                 .sort_values(["Grammy Year", "Focal Act"]))

    # ── write ────────────────────────────────────────────────────────────────
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False)

    excel = df.copy()
    excel["Extracted Text"] = (excel["Extracted Text"].astype(str)
                               .str.slice(0, EXCEL_CELL_CAP))
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        summary.to_excel(xw, sheet_name="Summary", index=False)
        if rejects:
            pd.concat(rejects, ignore_index=True).to_excel(
                xw, sheet_name="Rejects", index=False)
        if leads:
            pd.concat(leads, ignore_index=True).to_excel(
                xw, sheet_name="Archive Leads", index=False)

    # ── report ───────────────────────────────────────────────────────────────
    n = len(df)
    pub = pd.to_datetime(df["Publication Date"], errors="coerce")
    print("\n" + "=" * 66)
    print(f"MERGED CORPUS: {n} records across {len(acts)} acts")
    print("=" * 66)
    for col in ["Title / Headline", "Publication Date", "Author / Interviewer",
                "Source Type", "Speaker Name(s)"]:
        filled = _nonblank(df[col])
        print(f"  {col:24} {filled:6}/{n} ({100*filled/n:3.0f}%)")
    print(f"\n  Period flagged        {truthy(df['Period Record']):6}/{n} "
          f"({100*truthy(df['Period Record'])/n:3.0f}%)")
    print(f"  ...genuinely pre-2005 {(pub.dt.year < 2005).sum():6}/{n} "
          f"({100*(pub.dt.year < 2005).sum()/n:3.0f}%)   <-- Period Record is "
          f"recomputed from the checked dates in clean_dates_v7.py")
    flagged = _nonblank(df["Review Flag"])
    print(f"  Flagged for review    {flagged:6}/{n} ({100*flagged/n:3.0f}%)")
    print(f"  Median text length    {int(df['Text Length'].median())} chars")
    print(f"\nWrote:\n  {out_xlsx}\n  {out_json}  (full untruncated text)")
    print("Next: dedupe_v7.py.")
    return df


def preflight():
    """Fail fast and loudly before a long run: verify the key works and that
    a known-good query actually returns results."""
    print("Preflight: testing Serper…")
    hits = search_serper('"Dire Straits" interview')
    if not hits:
        cause = _serper_stats.get("last_error")
        raise SearchAborted(
            "Preflight query returned no results — not starting the batch.\n"
            f"  Underlying cause: {cause or 'none recorded (empty 200 response)'}\n"
            "  If this is a connection/SSL error, it is a network problem, not "
            "your key. If it is HTTP 400, the request payload was rejected.")
    print(f"  OK ({len(hits)} results). Proceeding.\n")


def main():
    acts = BATCHES[BATCH]
    outdir = DRIVE_DIR if USE_DRIVE and os.path.isdir(DRIVE_DIR) else "."
    os.makedirs(outdir, exist_ok=True)
    ckpt = os.path.join(outdir, f"v7_batch{BATCH}_checkpoint.json")

    preflight()

    done, all_records, all_rejects, all_leads = {}, [], [], []
    if os.path.exists(ckpt):
        with open(ckpt) as f:
            state = json.load(f)
        done = state.get("done", {})
        all_records = state.get("records", [])
        all_rejects = state.get("rejects", [])
        all_leads = state.get("leads", [])
        print(f"Resuming: {len(done)} acts already complete.")

    print(f"=== v7 BATCH {BATCH}: {len(acts)} acts ===")
    for act in acts:
        if act in done:
            print(f"  skip {act} (done)")
            continue
        try:
            recs, rejs, leads = process_act(act)
        except SearchAborted as e:
            print(f"\n!! ABORTED during {act}: {e}")
            print("   Checkpoint saved. Fix the cause and re-run this batch.")
            break
        all_records.extend(recs)
        all_rejects.extend(rejs)
        all_leads.extend(leads)
        done[act] = len(recs)
        with open(ckpt, "w") as f:
            json.dump({"done": done, "records": all_records,
                       "rejects": all_rejects, "leads": all_leads}, f)
        warn = "  <-- LOW, check" if len(recs) < 5 else ""
        print(f"  {act}: {len(recs)} kept, {len(rejs)} rejected, "
              f"{len(leads)} leads{warn}")

    # outputs
    df = pd.DataFrame(all_records, columns=COLUMNS)
    excel = df
    if not df.empty:
        df = df.sort_values(["Grammy Year", "Focal Act", "Publication Date"],
                            na_position="last")
        excel = df.copy()
        excel["Extracted Text"] = excel["Extracted Text"].str.slice(0, 32000)

    xlsx = os.path.join(outdir, f"harrison_v7_batch{BATCH}.xlsx")
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        pd.DataFrame(all_rejects,
                     columns=["Focal Act", "Title", "URL", "Query",
                              "Text Length", "Reject Reason"]
                     ).to_excel(xw, sheet_name="Rejects", index=False)
        pd.DataFrame(all_leads,
                     columns=["Focal Act", "Archive", "URL", "Query"]
                     ).to_excel(xw, sheet_name="Archive Leads", index=False)
        summary = (df.groupby(["Focal Act", "Grammy Year"])
                     .agg(Records=("URL", "count"),
                          Period=("Period Record", "sum"),
                          Flagged=("Review Flag", lambda s: (s != "").sum()))
                     .reset_index() if not df.empty else pd.DataFrame())
        summary.to_excel(xw, sheet_name="Summary", index=False)

    with open(os.path.join(outdir, f"harrison_v7_batch{BATCH}.json"),
              "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False)

    print(f"\nBatch {BATCH} complete: {len(all_records)} records, "
          f"{len(all_rejects)} rejected, {len(all_leads)} archive leads.")
    print(f"Wrote {xlsx} (+ JSON sidecar with untruncated text).")
    quality_report(df)
    if MERGE_WHEN_COMPLETE:
        merge_batches(outdir)


if __name__ == "__main__":
    main()
