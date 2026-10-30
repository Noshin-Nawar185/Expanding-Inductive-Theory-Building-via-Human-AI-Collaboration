"""build_band_voice_jsonl.py: the band voice corpus, one file per act.
Step 2 (articles and videos). Version 1.4, 20 October 2026.

What it makes: the input for the GenAI coding steps (Step 3 onwards). For each act, every turn in which a member of the
act speaks in their own words, from the articles and the videos together, in date order (oldest first).

Reads, in the folder this script is in (Drive: MyDrive/harrison_v7):
  harrison_v7_oral_histories.json                                   the article oral histories (version 1.2)
  Video for Harrison/harrison_v7_video_oral_histories_quality.json  the video oral histories, with the quality measures
  harrison_v7_band_voice_fixes.json                                 the hand fixes from the cleaning round
Writes there (it never changes the files it reads):
  band_voice_by_act/<year>_<act>.jsonl   52 files, one per act, one turn per line
  harrison_v7_band_voice_corpus.xlsx     the same turns to read and check: Read me, Acts, one sheet per act, Left out
  harrison_v7_band_voice_acts.json       each act's file name and counts

Kept: turns by a member of the act, in their own words, whose speaker is known with high or medium confidence: members
at the time of the nomination (or the solo artist), marked 0 in the Other era column, and members who joined after the
nomination or had left before it, marked 1.
Left out (each listed in the Left out sheet, with the reason):
  1. video turns with crosstalk: half or more of their words spoken while someone else is talking;
  2. video turns whose wording is not confirmed: fewer than half of their words in both transcripts (almost all are
     Gemini's own text, where WhisperX heard nothing there or something different);
  3. copies of a turn already in the act's history (a republished article, a re-uploaded video, a quote used again):
     an identical turn of six words or more, or a turn of twelve words or more with at least 60% of its eight-word
     runs in one earlier turn. The earliest copy stays, and its Later copies counts the rest;
  4. turns dropped by hand in the cleaning round.
Corrected:
  - three video speaker names, in Speaker and in Before and After: Sade Adu (her name is the act's, so she had been
    read as the band as a group); Mark Brzezicki and Tony Butler (introduced by first name only in V1423);
  - one article (record 3491): the interviewer's questions, labelled 'Geeks of Doom:', cut from the end of Geoff
    Downes's answers;
  - runs of spaces and line breaks in Text, Before and After made one space;
  - by hand, in the cleaning round (harrison_v7_band_voice_fixes.json, made from harrison_v7_band_voice_fixes.xlsx):
    page furniture and editor's notes cut from article turns, and the dates of some videos checked. Each cut is made
    only where its passage is still in the turn, and any fix that cannot be made is reported and listed in the Hand
    fixes sheet. Before and After are left as in the source.
No other words are changed. The switches below change what is kept.

In Colab (any runtime), after Drive is mounted:
  %run "/content/drive/MyDrive/harrison_v7/build_band_voice_jsonl.py"
It takes a few minutes. A second run replaces the files it wrote before.
"""
import collections, datetime, hashlib, json, os, re, sys

VERSION = "build_band_voice_jsonl 1.4 (20 Oct 2026)"
MAIN_SET_ONLY = True        # only turns whose speaker is known with high or medium confidence
INCLUDE_OTHER_ERA = True    # also members who joined after the nomination or had left before it (Other era = 1)
DROP_CROSSTALK = True       # leave out video turns with half or more of their words in crosstalk
MIN_BOTH = 0.5              # leave out video turns with less than this share of their words in both transcripts
                            # (None keeps them all)
DROP_COPIES = True          # leave out later copies of a turn
EXACT_MIN_WORDS = 6         # a copy: an identical turn of at least this many words ...
COPY_MIN_WORDS = 12         # ... or a turn of at least this many words ...
COPY_SHARE = 0.6            # ... with at least this share of its runs of eight words in one earlier turn
SHINGLE = 8
CELL_CAP = 32000            # Excel holds at most 32,767 characters in a cell
LABEL_CUTS = {"3491": "Geeks of Doom:"}   # an interviewer's label left inside member turns: the text is cut there
APPLY_FIXES = True          # apply the hand fixes in FIXES (the file must be there); False builds without them

TEXT_OH = "harrison_v7_oral_histories.json"
VIDEO_OH = os.path.join("Video for Harrison", "harrison_v7_video_oral_histories_quality.json")
FIXES = "harrison_v7_band_voice_fixes.json"
OUT_DIR = "band_voice_by_act"
OUT_XLSX = "harrison_v7_band_voice_corpus.xlsx"
OUT_ACTS = "harrison_v7_band_voice_acts.json"

FIELDS = ["Order", "Turn ID", "Act", "Source", "Date", "Date precision", "Date from", "Years from nomination", "Period",
          "Speaker", "Speaker role", "Other era", "Text", "Before", "After", "Words", "Years mentioned", "Confidence", "Format",
          "Outlet", "Title", "Publication date", "URL", "Record ID", "Position in article", "Later copies",
          "Start", "Language", "Text from", "Crosstalk", "Words in both transcripts", "Check"]
EXPLAIN = {
    "Order": "The turn's place in the act's history, oldest first (1, 2, 3 ...). Sort by it, smallest to largest, to "
             "get back to date order.",
    "Turn ID": "The turn's ID in its corpus, to trace it back: article-turn (113-32) or video-turn (V1423-12).",
    "Act": "The Best New Artist nominee.",
    "Source": "'text' (an article) or 'video'.",
    "Date": "When the words were said: for an article, the interview date where the article gives it, otherwise the "
            "publication date; for a video, 1 January of its year (most videos are dated to the year only), or the "
            "day where a check by hand found it.",
    "Date precision": "How exact Date is: 'day', 'month' or 'year'.",
    "Date from": "Where the date comes from. For videos: the title, a check by hand, the description, or the upload "
                 "year (the latest the video can be from).",
    "Years from nomination": "The year of Date minus the year of the act's nomination (the Grammy ceremony year): "
                             "negative before it, 0 in it, positive after it.",
    "Period": "'before', 'nomination year' or 'after', from Years from nomination. It is when the account was given, "
              "not the time it describes.",
    "Speaker": "The member, named as on the act's line-up. 'band member, unspecified' where the source does not say "
               "which member speaks (in a video with a letter, A, B ..., that holds within that video only).",
    "Speaker role": "Member, Member (unidentified), The band as a group, or Other-era member.",
    "Other era": "0: a member of the act at the time of the nomination (or the solo artist). 1: a member who joined "
                 "after the nomination or had left before it.",
    "Text": "The turn's words, as said (runs of spaces and line breaks made one space).",
    "Before": "Up to 400 characters of what was said just before the turn, each part after its speaker ('/' between "
              "speakers). For an answer, usually the question.",
    "After": "Up to 400 characters of what was said just after the turn.",
    "Words": "The number of words in Text.",
    "Years mentioned": "Years named in the turn (1983, 1980s, '80s): a pointer to the time the account describes.",
    "Confidence": "How sure the speaker is: high or medium (turns of low confidence are not here); 'added by hand' "
                  "for Amy Holland's four interviews. For videos, provisional until the spot check.",
    "Format": "Articles: prose, labelled Q&A, unlabelled Q&A or first-person. Videos: video.",
    "Outlet": "The publication; for a video, the show where known, otherwise YouTube and the channel.",
    "Title": "The article's or the video's title.",
    "Publication date": "The article's publication date, or the video's upload date.",
    "URL": "The article, or the video opening two seconds before the turn.",
    "Record ID": "The article's or the video's ID.",
    "Position in article": "The turn's place in its article or video.",
    "Later copies": "How many later copies of this turn were left out (republished articles, re-uploaded videos, a "
                    "quote used again).",
    "Start": "Videos only: where the turn starts (minutes:seconds).",
    "Language": "Videos only: the turn's language (en, fr, ja ...). Turns not in English are as spoken.",
    "Text from": "Videos only: WhisperX (its words, matched with Gemini's), Gemini (Gemini's text) or WhisperX only.",
    "Crosstalk": "Videos only: the share of the turn's words spoken while someone else is talking, from 0 to 1. "
                 "Turns at 0.5 or more are left out.",
    "Words in both transcripts": "Videos only: the share of the turn's words that both transcripts have, word for "
                                 "word, from 0 to 1. Turns under 0.5 are left out.",
    "Check": "For the cleaning round: 'page furniture or editor's note?' (an advertisement, credit, address, footer "
             "or editor's note may be in the text); 'Gemini's text' (a video turn whose text is Gemini's, with at "
             "least half its words also in WhisperX's); 'a copy is dated ...' (a re-upload of this interview has "
             "another year, so one of the two dates is wrong). Each of these was looked at by hand in the cleaning "
             "round and also says 'checked by hand' (kept as it was) or 'text edited by hand' (page furniture or an "
             "editor's note cut; the Hand fixes sheet shows what).",
}
VIDEO_ONLY = ["Start", "Language", "Text from", "Crosstalk", "Words in both transcripts"]
NAME_FIXES = {("Sade", "Sade"): ("Sade Adu", "Member", "on-screen caption 'SADE' (V2571), and the same person elsewhere"),
              ("Big Country", "Mark"): ("Mark Brzezicki", "Member", "introduced as Mark in V1423"),
              ("Big Country", "Tony"): ("Tony Butler", "Member", "introduced and addressed as Tony in V1423")}
FURNITURE = re.compile(r"\(photo by|photo:|click here|newsletter|read more|advertisement|getty images|all rights "
                       r"reserved|written by\s*:|discover more from|sign up for|https?://|www\.|\(note:|editor'?s "
                       r"note|(?:^|\s)ed:\s|â€|�", re.I)
YEAR = re.compile(r"(?<![\d$£€.,])(?:19[4-9]\d|20[0-2]\d)(?:s\b|(?![\d,]))")   # as in the oral histories
DECADE = re.compile(r"(?<![\w])['’](?:[4-9]0|[0-2]0)s\b")
HAND_COLS = ["Fix", "Turn or video", "Act", "What changed", "Words before", "Words after", "Note", "Made"]


def script_folder():
    f = globals().get("__file__")
    return os.path.dirname(os.path.abspath(f)) if f else os.getcwd()


def one_space(s):
    return re.sub(r"\s+", " ", s or "").strip()


def norm_words(s):
    """Lower-case words without punctuation; apostrophes used as quotation marks are dropped."""
    s = (s or "").lower().replace("’", "'").replace("‘", "'")
    return [t for t in (w.strip("'") for w in re.sub(r"[^a-z0-9']+", " ", s).split()) if t]


def shingles(ws):
    return {hash(tuple(ws[i:i + SHINGLE])) for i in range(len(ws) - SHINGLE + 1)}


def slug(act):
    s = re.sub(r"[^a-z0-9]+", "_", act.lower().replace("'", "").replace("’", "")).strip("_")
    return s or "act"


def load(path):
    j = json.load(open(path, encoding="utf-8"))
    cols = j["columns"]
    return j, [dict(zip(cols, r)) for r in j["data"]]


def keep_row(r, source):
    voices = {"Band"} | ({"Other-era member"} if INCLUDE_OTHER_ERA else set())
    if r["Voice"] not in voices:
        return None
    if MAIN_SET_ONLY and r["Voice"] == "Band" and r["Main set"] != "yes":
        return None
    if MAIN_SET_ONLY and r["Voice"] == "Other-era member" and r["Confidence"] not in ("high", "medium"):
        return None
    if source == "video" and DROP_CROSSTALK and r.get("Set aside for crosstalk") == "yes":
        return "crosstalk"
    both = r.get("Words in both transcripts")
    if source == "video" and MIN_BOTH is not None and both is not None and both < MIN_BOTH:
        return "wording not confirmed"
    return "keep"


def label_fixes(vrows):
    """For each video with a corrected name: its old label and the full name, for Before and After."""
    out = collections.defaultdict(dict)
    for r in vrows:
        fix = NAME_FIXES.get((r["Act"], r["Speaker"]))
        if fix:
            out[r["Record ID"]][r["Speaker"]] = fix[0]
    return out


def relabel(s, labels):
    for old, new in labels.items():
        s = re.sub(r"(^|…| / )" + re.escape(old) + ": ", lambda m: m.group(1) + new + ": ", s)
    return s


def to_line(r, source, ctx=None):
    out = {}
    for k in FIELDS:
        if k == "Source":
            out[k] = source
        elif k in ("Check", "Order"):
            out[k] = "" if k == "Check" else None
        elif k in ("Text", "Before", "After"):
            out[k] = one_space(r.get(k))
        elif k == "Record ID":
            out[k] = str(r["Record ID"])
        elif k in VIDEO_ONLY:
            out[k] = r.get(k) if source == "video" else None
        else:
            out[k] = r.get(k)
    out["Other era"] = 1 if r["Voice"] == "Other-era member" else 0
    if out["Other era"]:
        out["Speaker role"] = "Other-era member"
    fix = NAME_FIXES.get((r["Act"], out["Speaker"])) if source == "video" else None
    if fix:
        out["Speaker"], out["Speaker role"] = fix[0], fix[1]
    if source == "video" and ctx and r["Record ID"] in ctx:
        for k in ("Before", "After"):
            out[k] = relabel(out[k], ctx[r["Record ID"]])
    cut = False
    lab = LABEL_CUTS.get(out["Record ID"]) if source == "text" else None
    if lab and lab in out["Text"]:
        out["Text"] = out["Text"][:out["Text"].index(lab)].strip()
        out["Words"] = len(out["Text"].split())
        cut = True
    return out, fix, cut


def order_key(x):
    rid = x["Record ID"]
    num = int(rid[1:]) if rid[:1] == "V" else int(rid)
    return (x["Date"], x["Publication date"] or "", 0 if x["Source"] == "text" else 1, num, x["Position in article"])


def checks(x):
    c = []
    if x["Source"] == "text" and FURNITURE.search(x["Text"]):
        c.append("page furniture or editor's note?")
    if x["Source"] == "video" and x["Text from"] == "Gemini":
        c.append("Gemini's text")
    return c


def years_mentioned(s):
    """Years named in a text, counted as the oral histories count them (for turns whose text is edited by hand)."""
    out = []
    for m in YEAR.finditer(s or ""):
        if m.group(0) not in out:
            out.append(m.group(0))
    for m in DECADE.finditer(s or ""):
        d = m.group(0)[1:3]
        y = ("19" if d[0] in "456789" else "20") + d + "s"
        if y not in out:
            out.append(y)
    return "; ".join(out)


def load_fixes(base):
    """The hand fixes: the turns looked at, the turn fixes by Turn ID and the video dates by Record ID."""
    p = os.path.join(base, FIXES)
    if not os.path.exists(p):
        return None
    j = json.load(open(p, encoding="utf-8"))
    return {"reviewed": set(j["reviewed"]), "turns": {t["Turn ID"]: t for t in j["turns"]},
            "videos": {v["Record ID"]: v for v in j["videos"]}}


def fix_date(x, v, nominated):
    """A video's date checked by hand: Date, and what follows from it."""
    x["Date"], x["Date precision"], x["Date from"] = v["date"], v["precision"], "checked by hand"
    y = int(v["date"][:4]) - int(nominated)
    x["Years from nomination"] = y
    x["Period"] = "before" if y < 0 else "nomination year" if y == 0 else "after"


def cut_text(before, after):
    """What a hand edit took out of a turn, for the Hand fixes sheet."""
    import difflib
    ops = difflib.SequenceMatcher(None, before, after, autojunk=False).get_opcodes()
    return " … ".join(before[i1:i2].strip() for op, i1, i2, j1, j2 in ops if op in ("delete", "replace"))


def apply_turn_fixes(kept, fixes, left, hand, done, edited):
    """Drop or edit the turns with a hand fix; the rest pass through unchanged."""
    out = []
    for x in kept:
        f = fixes["turns"].get(x["Turn ID"])
        if f is None:
            out.append(x)
            continue
        done.add(x["Turn ID"])
        if f.get("drop"):
            left.append({**x, "Why left out": "dropped by hand", "Copy of": "", "Share": None})
            hand.append({"Fix": "turn dropped", "Turn or video": x["Turn ID"], "Act": x["Act"], "What changed": x["Text"],
                         "Words before": x["Words"], "Words after": 0, "Note": f.get("note", ""), "Made": "yes"})
            continue
        text, missed = x["Text"], 0
        for old, new in f.get("edits", []):
            if text.count(old) == 1:
                text = text.replace(old, new)
            else:
                missed += 1
        n = len(f.get("edits", []))
        made = "yes" if not missed else ("no" if missed == n else "partly") + f" ({missed} of {n} passages are no longer in the turn)"
        hand.append({"Fix": "text edited", "Turn or video": x["Turn ID"], "Act": x["Act"],
                     "What changed": "cut: " + cut_text(x["Text"], text), "Words before": x["Words"],
                     "Words after": len(text.split()), "Note": f.get("note", ""), "Made": made})
        if text != x["Text"]:
            x["Text"], x["Words"], x["Years mentioned"] = text, len(text.split()), years_mentioned(text)
            edited.add(x["Turn ID"])
        out.append(x)
    return out


def id_key(s):
    a, _, b = str(s).partition("-")
    return (int(a.lstrip("V") or 0), int(b) if b.isdigit() else 0)


def write_xlsx(path, acts, by_act, left, readme, hand):
    try:
        import xlsxwriter
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "xlsxwriter"])
        import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name
    wb = xlsxwriter.Workbook(path, {"constant_memory": True, "strings_to_urls": False, "strings_to_formulas": False})
    bold = wb.add_format({"bold": True, "valign": "top"})
    title = wb.add_format({"bold": True, "font_size": 14})
    head = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "text_wrap": True, "valign": "top", "border": 1,
                          "border_color": "#BFBFBF"})
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})
    top = wb.add_format({"valign": "top"})
    datef = wb.add_format({"num_format": "yyyy-mm-dd", "valign": "top"})
    link = wb.add_format({"font_color": "#0563C1", "underline": 1, "valign": "top"})
    videof = wb.add_format({"bg_color": "#F2F2F2"})

    def cell(ws, i, j, k, v):
        if v is None:
            ws.write_blank(i, j, None, top)
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            ws.write_number(i, j, v, top)
        elif k in ("Date", "Publication date") and isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}", v):
            ws.write_datetime(i, j, datetime.datetime.strptime(v[:10], "%Y-%m-%d"), datef)
        elif k == "URL" and isinstance(v, str) and re.match(r"https?://", v) and len(v) <= 2000:
            ws.write_url(i, j, v, link, string=v)
        elif k in ("Text", "Before", "After", "What changed"):
            v = str(v)
            ws.write_string(i, j, v if len(v) <= CELL_CAP else v[:CELL_CAP - 50] + " … [cut here: in full in the JSONL]", wrap)
        else:
            ws.write_string(i, j, str(v), top)

    rm = wb.add_worksheet("Read me")
    rm.set_column(0, 0, 30)
    rm.set_column(1, 1, 120)
    rm.write(0, 0, "Band voice by act: what each member said, from articles and videos, oldest first", title)
    r = 2
    for k, v in readme:
        rm.write(r, 0, k, bold)
        rm.write(r, 1, v, wrap)
        r += 1
    r += 1
    rm.write(r, 0, "THE COLUMNS", bold)
    r += 1
    for k in FIELDS:
        rm.write(r, 0, k, bold)
        rm.write(r, 1, EXPLAIN[k], wrap)
        r += 1

    names, used = {}, {"read me", "acts", "left out", "hand fixes"}
    for a in acts:
        s = re.sub(r"[\[\]:*?/\\]", "", a["Act"]).strip("'")[:31]
        base, n = s, 2
        while s.casefold() in used:
            s, n = f"{base[:28]} {n}", n + 1
        used.add(s.casefold())
        names[a["Act"]] = s
    acols = ["Act", "Solo or group", "Nominated", "Result", "File", "Turns", "Words", "Text turns", "Video turns",
             "Text words", "Video words", "Other-era turns", "Before", "Nomination year", "After", "Left out for crosstalk",
             "Left out, wording not confirmed", "Left out as copies", "Dropped by hand", "Sheet"]
    ws = wb.add_worksheet("Acts")
    for j, k in enumerate(acols):
        ws.set_column(j, j, {"Act": 24, "File": 34}.get(k, 11))
        ws.write(0, j, k, head)
    ws.freeze_panes(1, 1)
    for i, a in enumerate(acts, 1):
        for j, k in enumerate(acols):
            if k == "Sheet":
                ws.write_url(i, j, f"internal:'{names[a['Act']].replace(chr(39), chr(39) * 2)}'!A1", link, string="open")
            else:
                cell(ws, i, j, k, a[k])
    ws.autofilter(0, 0, len(acts), len(acols) - 1)

    widths = {"Order": 7, "Turn ID": 10, "Act": 14, "Source": 7, "Date": 11, "Date precision": 8, "Date from": 14,
              "Years from nomination": 9, "Period": 12, "Speaker": 22, "Speaker role": 14, "Other era": 7, "Text": 80, "Before": 40,
              "After": 40, "Words": 7, "Years mentioned": 11, "Confidence": 9, "Format": 12, "Outlet": 18,
              "Title": 28, "Publication date": 11, "URL": 28, "Record ID": 8, "Position in article": 8,
              "Later copies": 7, "Start": 7, "Language": 7, "Text from": 10, "Crosstalk": 8,
              "Words in both transcripts": 9, "Check": 22}
    src = FIELDS.index("Source")
    for a in acts:
        ws = wb.add_worksheet(names[a["Act"]])
        rows = by_act.get(a["Act"], [])
        ws.write(0, 0, a["Act"], title)
        ws.write(1, 0, f"{a['Solo or group']} · nominated {a['Nominated']} · {a['Result']} · {a['Turns']:,} turns "
                       f"({a['Text turns']:,} from articles, {a['Video turns']:,} from videos) · {a['Words']:,} words "
                       f"· file {a['File']}")
        ws.write(2, 0, "In date order, oldest first. After sorting or filtering, sort by Order (smallest to largest) "
                       "to get back to it. Grey rows are from videos.")
        H0 = 3
        for j, k in enumerate(FIELDS):
            ws.set_column(j, j, widths.get(k, 12))
            ws.write(H0, j, k, head)
        ws.freeze_panes(H0 + 1, FIELDS.index("Text"))
        for i, x in enumerate(rows, H0 + 1):
            ws.set_row(i, 45)
            for j, k in enumerate(FIELDS):
                cell(ws, i, j, k, x[k])
        if rows:
            ws.autofilter(H0, 0, H0 + len(rows), len(FIELDS) - 1)
            col = xl_col_to_name(src)
            ws.conditional_format(H0 + 1, 0, H0 + len(rows), len(FIELDS) - 1,
                                  {"type": "formula", "criteria": f'=${col}{H0 + 2}="video"', "format": videof})

    lcols = ["Turn ID", "Act", "Source", "Date", "Speaker", "Words", "Why left out", "Copy of", "Share", "Text", "URL"]
    ws = wb.add_worksheet("Left out")
    for j, k in enumerate(lcols):
        ws.set_column(j, j, {"Text": 80, "Why left out": 22, "URL": 28, "Speaker": 22}.get(k, 11))
        ws.write(0, j, k, head)
    ws.freeze_panes(1, 0)
    for i, x in enumerate(left, 1):
        for j, k in enumerate(lcols):
            cell(ws, i, j, k, x[k])
    if left:
        ws.autofilter(0, 0, len(left), len(lcols) - 1)

    if hand:
        ws = wb.add_worksheet("Hand fixes")
        for j, k in enumerate(HAND_COLS):
            ws.set_column(j, j, {"Fix": 16, "Turn or video": 11, "Act": 20, "What changed": 90, "Note": 30,
                                 "Made": 16}.get(k, 9))
            ws.write(0, j, k, head)
        ws.freeze_panes(1, 0)
        for i, h in enumerate(hand, 1):
            ws.set_row(i, 45)
            for j, k in enumerate(HAND_COLS):
                cell(ws, i, j, k, h[k])
        ws.autofilter(0, 0, len(hand), len(HAND_COLS) - 1)
    wb.close()


def main(base=None):
    base = base or script_folder()
    print(VERSION)
    pt, pv = os.path.join(base, TEXT_OH), os.path.join(base, VIDEO_OH)
    for what, p in (("the article oral histories", pt), ("the video oral histories with the quality measures", pv)):
        if not os.path.exists(p):
            print(f"STOPPED: {what} not found at {p}. Nothing was written.")
            return None
    fixes = None
    if APPLY_FIXES:
        fixes = load_fixes(base)
        if fixes is None:
            print(f"STOPPED: the hand fixes not found at {os.path.join(base, FIXES)}. Put the file there, or set "
                  "APPLY_FIXES = False to build without them. Nothing was written.")
            return None
    tj, trows = load(pt)
    vj, vrows = load(pv)
    print(f"Read {len(trows):,} article rows and {len(vrows):,} video rows"
          + (f", and hand fixes for {len(fixes['turns'])} turns and {len(fixes['videos'])} videos." if fixes else "."),
          flush=True)
    nominated = {a["Act"]: a["Nominated"] for a in tj["acts"]}
    dated, hand, done, edited = {}, [], set(), set()

    # ---------------------------------------------------------------- keep the members' own words
    cand, left = collections.defaultdict(list), []
    n_fix, n_cut = collections.Counter(), 0
    ctx = label_fixes(vrows)
    for source, rows in (("text", trows), ("video", vrows)):
        for r in rows:
            k = keep_row(r, source)
            if k is None:
                continue
            x, fix, cut = to_line(r, source, ctx)
            v = fixes["videos"].get(x["Record ID"]) if fixes and source == "video" else None
            if v:
                dated.setdefault(x["Record ID"], (x["Date"], x["Date precision"]))
                fix_date(x, v, nominated[r["Act"]])
            if fix:
                n_fix[fix[0]] += 1
            n_cut += cut
            if k == "crosstalk":
                left.append({**x, "Why left out": "crosstalk", "Copy of": "", "Share": r.get("Crosstalk")})
                continue
            if k == "wording not confirmed":
                left.append({**x, "Why left out": k, "Copy of": "", "Share": r.get("Words in both transcripts")})
                continue
            if not x["Text"]:
                left.append({**x, "Why left out": "only the interviewer's words", "Copy of": "", "Share": None})
                continue
            cand[r["Act"]].append(x)

    # ---------------------------------------------------------------- oldest first; take out later copies
    by_act, copies_of = {}, collections.defaultdict(list)
    for act, rows in cand.items():
        rows.sort(key=order_key)
        kept, post, exact = [], collections.defaultdict(list), {}
        for x in rows:
            ws = norm_words(x["Text"])
            key = " ".join(ws)
            match = None
            if DROP_COPIES and len(ws) >= EXACT_MIN_WORDS and key in exact:
                match = (exact[key], 1.0)
            sh = shingles(ws) if (DROP_COPIES and len(ws) >= COPY_MIN_WORDS) else set()
            if match is None and sh:
                c = collections.Counter()
                for h in sh:
                    for i in post[h]:
                        c[i] += 1
                if c:
                    i, n = max(c.items(), key=lambda kv: (kv[1], -kv[0]))
                    if n / len(sh) >= COPY_SHARE:
                        match = (i, round(n / len(sh), 2))
            if match is not None:
                orig = kept[match[0]]
                orig["Later copies"] = (orig["Later copies"] or 0) + 1 + (x["Later copies"] or 0)
                copies_of[orig["Turn ID"]].append(x)
                left.append({**x, "Why left out": "copy", "Copy of": orig["Turn ID"], "Share": match[1]})
                continue
            kept.append(x)
            exact.setdefault(key, len(kept) - 1)
            for h in sh:
                post[h].append(len(kept) - 1)
        if fixes:
            kept = apply_turn_fixes(kept, fixes, left, hand, done, edited)
        for n, x in enumerate(kept, 1):
            x["Order"] = n
            c = checks(x)
            other = sorted({y["Date"][:4] for y in copies_of.get(x["Turn ID"], []) if y["Date"][:4] != x["Date"][:4]
                            and y["Source"] == "video" and x["Source"] == "video"})
            if other:
                c.append("a copy is dated " + ", ".join(other))
            if fixes and x["Turn ID"] in fixes["reviewed"]:
                c.append("text edited by hand" if x["Turn ID"] in edited else "checked by hand")
            x["Check"] = "; ".join(c)
        by_act[act] = kept

    # ---------------------------------------------------------------- the hand fixes that could not be made
    act_ix = {a["Act"]: i for i, a in enumerate(tj["acts"])}
    if fixes:
        for tid, f in fixes["turns"].items():
            if tid not in done:
                hand.append({"Fix": "turn dropped" if f.get("drop") else "text edited", "Turn or video": tid,
                             "Act": f.get("Act", ""), "What changed": "", "Words before": None, "Words after": None,
                             "Note": f.get("note", ""), "Made": "no: the turn is not in the corpus"})
        for rid, v in fixes["videos"].items():
            n = sum(1 for rows in by_act.values() for x in rows if x["Record ID"] == rid)
            was = dated.get(rid)
            what = "" if was is None else (f"{v['date'][:4]} confirmed" if was == (v["date"], v["precision"])
                                           else f"{was[0]} ({was[1]}) to {v['date']} ({v['precision']})")
            hand.append({"Fix": "video date checked", "Turn or video": rid, "Act": v.get("Act", ""), "What changed": what,
                         "Words before": None, "Words after": None, "Note": v.get("note", ""),
                         "Made": f"yes ({n} turns)" if was else "no: the video is not in the corpus"})
        rank = {"turn dropped": 0, "text edited": 1, "video date checked": 2}
        hand.sort(key=lambda h: (act_ix.get(h["Act"], 99), rank[h["Fix"]], id_key(h["Turn or video"])))
    not_made = [h for h in hand if not h["Made"].startswith("yes")]

    # ---------------------------------------------------------------- the acts, the files and the workbook
    out_dir = os.path.join(base, OUT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    acts = []
    left_by = collections.Counter((x["Act"], x["Why left out"]) for x in left)
    for a in tj["acts"]:
        act = a["Act"]
        rows = by_act.get(act, [])
        per = collections.Counter(x["Period"] for x in rows)
        fn = f"{a['Nominated']}_{slug(act)}.jsonl"
        acts.append({"Act": act, "Solo or group": a["Solo or group"], "Nominated": a["Nominated"], "Result": a["Result"],
                     "File": fn, "Turns": len(rows), "Words": sum(x["Words"] for x in rows),
                     "Text turns": sum(1 for x in rows if x["Source"] == "text"),
                     "Video turns": sum(1 for x in rows if x["Source"] == "video"),
                     "Text words": sum(x["Words"] for x in rows if x["Source"] == "text"),
                     "Video words": sum(x["Words"] for x in rows if x["Source"] == "video"),
                     "Other-era turns": sum(1 for x in rows if x["Other era"] == 1),
                     "Before": per["before"], "Nomination year": per["nomination year"], "After": per["after"],
                     "Left out for crosstalk": left_by[(act, "crosstalk")],
                     "Left out, wording not confirmed": left_by[(act, "wording not confirmed")],
                     "Left out as copies": left_by[(act, "copy")],
                     "Dropped by hand": left_by[(act, "dropped by hand")]})
        with open(os.path.join(out_dir, fn), "w", encoding="utf-8") as fh:
            for x in rows:
                fh.write(json.dumps(x, ensure_ascii=False) + "\n")
    missing = sorted(set(by_act) - {a["Act"] for a in acts})
    if missing:
        print(f"NOTE: acts in the turns but not in the act list, not written: {missing}")
    left.sort(key=lambda x: (x["Act"], order_key(x)))

    def n_left(why):
        xs = [x for x in left if x["Why left out"] == why]
        return len(xs), sum(x["Words"] or 0 for x in xs)
    tot, words = sum(a["Turns"] for a in acts), sum(a["Words"] for a in acts)
    (c1, w1), (c2, w2), (c3, w3), (c4, w4) = (n_left("crosstalk"), n_left("wording not confirmed"), n_left("copy"),
                                              n_left("dropped by hand"))
    lo = [f"{c1:,} video turns with crosstalk ({w1:,} words)",
          f"{c2:,} video turns whose wording is not confirmed ({w2:,} words)", f"{c3:,} copies ({w3:,} words)"]
    lo += [f"{c4:,} turns dropped by hand ({w4:,} words)"] if c4 else []
    summary = (f"{tot:,} turns and {words:,} words in {len(acts)} files: {sum(a['Text turns'] for a in acts):,} turns "
               f"from articles and {sum(a['Video turns'] for a in acts):,} from videos; "
               f"{sum(a['Other-era turns'] for a in acts):,} of the turns are by members from other eras. Left out: "
               + ", ".join(lo[:-1]) + " and " + lo[-1] + ".")
    he = [h for h in hand if h["Fix"] == "text edited" and h["Made"] != "no: the turn is not in the corpus"]
    hv = [h for h in hand if h["Fix"] == "video date checked" and h["Made"].startswith("yes")]
    n_cutw = sum((h["Words before"] or 0) - (h["Words after"] or 0) for h in he)
    n_day = sum(1 for h in hv if not h["What changed"].endswith("confirmed"))
    n_rev = sum(1 for rows in by_act.values() for x in rows if x["Check"].endswith("by hand"))
    kept_what = ("Turns by a member of the act, in their own words, whose speaker is known with high or medium "
                 "confidence" if MAIN_SET_ONLY else "All turns by a member of the act, in their own words") + (
        ": members at the time of the nomination (or the solo artist), marked 0 in Other era, and members who joined "
        "after the nomination or had left before it, marked 1" if INCLUDE_OTHER_ERA else
        ", for members at the time of the nomination (or the solo artist)")
    readme = [
        ("What this is", "The input for the GenAI coding steps. For each act, every turn in which a member of the act "
         "speaks in their own words, from the articles and the videos together, oldest first. Each act has a sheet "
         f"here and a JSONL file in {OUT_DIR}/ (the same turns, one per line, every text in full)."),
        ("An act's history in date order", "Each act's sheet, and each JSONL file, is already in date order. After "
         "sorting or filtering a sheet, sort by the Order column, smallest to largest, to get back to it. Order goes by "
         "Date (when the words were said), then the publication date, then the article or video, then the turn's "
         "place in it, so the turns of one interview stay together, in the order they were said. Most videos are "
         "dated to the year only (1 January), so they come before that year's dated articles."),
        ("Other ways to read it", "1. One member: filter Speaker, then sort by Order. 2. One interview: filter Record ID; "
         "its turns are in the order said. 3. Before, in or after the nomination year: filter Period. 4. Articles or "
         "videos only: filter Source. 5. Only the members at the time of the nomination: filter Other era to 0. Date is when the words were said, not the time they describe; Years mentioned "
         "gives the years named in the turn."),
        ("What is kept", kept_what + ". Interviewers, writers, other people, unknown speakers, second-hand reports"
         + (" and songs" if INCLUDE_OTHER_ERA else ", songs and members from other eras") + " are not here; they are in the "
         "oral histories, and Before and After show what was said around each turn."),
        ("What is left out", f"1. Video turns with crosstalk: half or more of their words spoken while someone else is "
         f"talking. 2. Video turns whose wording is not confirmed: fewer than half of their words in both transcripts "
         f"(almost all are Gemini's own text, where WhisperX heard nothing there or something different). 3. Copies of "
         f"a turn already in the act's history (a republished article, a re-uploaded video, a quote used again): an "
         f"identical turn of {EXACT_MIN_WORDS} words or more, or a turn of {COPY_MIN_WORDS} words or more with at "
         f"least {COPY_SHARE:.0%} of its {SHINGLE}-word runs in one earlier turn. The earliest copy stays, and its "
         "Later copies column counts the rest." + (f" 4. {c4} turns dropped by hand in the cleaning round (page "
         "furniture and the like; see the Hand fixes sheet)." if c4 else "") + " Every turn left out is in the Left "
         "out sheet, with the reason; Share "
         "there is the turn's crosstalk, its words in both transcripts, or, for a copy, how much of it is in the "
         "earlier turn."),
        ("What was corrected", "1. Three video speaker names, in Speaker and in Before and After: "
         + "; ".join(f"{v[0]} ({v[2]}, {n_fix[v[0]]} turns)" for v in NAME_FIXES.values()) + f". 2. One article "
         f"(record 3491): the interviewer's questions, labelled 'Geeks of Doom:', cut from the end of Geoff Downes's "
         f"answers ({n_cut} turns). 3. Runs of spaces and line breaks made one space."
         + (f" 4. By hand, in the cleaning round: page furniture or an editor's note cut from {len(he)} article turns "
            f"({n_cutw:,} words), and the dates of {len(hv)} videos checked ({n_day} now to the day, the rest confirmed "
            "as they were). Before and After are left as in the source." if fixes else "")
         + " No other words are changed."),
        ("Check column", "Marks turns to look at in the cleaning round: possible page furniture or an editor's note in "
         "an article's text; the few video turns whose text is Gemini's; and interviews whose re-upload has another "
         "year. Grey rows are from videos." + (" All of them were looked at by hand in the cleaning round, and each "
         "now also says 'checked by hand' (kept as it was) or 'text edited by hand'." if fixes else "")),
        ("Hand fixes", (f"From {FIXES} (made from harrison_v7_band_voice_fixes.xlsx). {len(fixes['reviewed'])} turns "
         f"marked Check were looked at by hand: {len(he)} had page furniture or an editor's note cut, {c4} were "
         f"dropped and the rest were kept as they were. The dates of {len(hv)} videos were checked. The Hand fixes "
         "sheet lists each fix and whether it was made" + (f"; {len(not_made)} could not be made." if not_made else
         ".")) if fixes else
         "None: this build was made without the hand fixes (APPLY_FIXES = False)."),
        ("Counts", summary),
        ("Version", VERSION),
    ]
    write_xlsx(os.path.join(base, OUT_XLSX), acts, by_act, left, readme, hand)
    with open(os.path.join(base, OUT_ACTS), "w", encoding="utf-8") as fh:
        json.dump({"about": {"version": VERSION, "notes": dict(readme), "fields": EXPLAIN}, "acts": acts}, fh,
                  ensure_ascii=False, indent=1)
    fp = hashlib.md5("".join(json.dumps(x, ensure_ascii=False) for a in acts for x in by_act.get(a["Act"], [])).encode()).hexdigest()
    print(summary)
    print("Speaker names corrected: " + ", ".join(f"{k} {n}" for k, n in n_fix.items())
          + f"; interviewer's questions cut from {n_cut} turns (record 3491)")
    if fixes:
        print(f"Hand fixes: {len(he)} turns edited ({n_cutw:,} words cut), {c4} dropped, {len(hv)} video dates checked "
              f"({n_day} now to the day); {n_rev:,} turns marked 'checked by hand' or 'text edited by hand' in Check")
        for h in not_made:
            print(f"NOTE: fix not made: {h['Fix']}, {h['Turn or video']} ({h['Act']}): {h['Made']}")
    print(f"Fingerprint of the rows: {fp} (the same inputs always give the same fingerprint)")
    print(f"Files: {OUT_DIR}/ ({len(acts)} JSONL files), {OUT_XLSX} "
          f"({os.path.getsize(os.path.join(base, OUT_XLSX)) / 1e6:.1f} MB), {OUT_ACTS}")
    return acts, by_act, left


if __name__ == "__main__":
    main()
