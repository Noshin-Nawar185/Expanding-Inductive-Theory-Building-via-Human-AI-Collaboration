"""build_video_oral_histories.py: chronological oral histories of the videos, one per act.
Step 2 (videos). Version 1.0, 18 October 2026.

Step 2 of the replication workflow, for the video corpus: a chronological oral history for each act,
built from the joined video turns, in the layout of the text corpus's oral histories (build_oral_histories.py 1.2).

Reads, in the folder this script is in (Drive: MyDrive/harrison_v7/Video for Harrison):
  harrison_v7_VIDEO_TURNS.json   the joined turns (merge_video_turns.py 1.3 or later)
  acts_frame.json                the acts' type, nomination year, result and line-ups
and writes there:
  harrison_v7_video_oral_histories.xlsx  for reading: a Read me, a Bands sheet (one row per act) and one sheet per act
  harrison_v7_video_oral_histories.json  for the GenAI steps: the same rows and columns for every act, plus the act details

Every turn of every video about the act is kept, oldest first (by the video's year, then its upload date); a video's
turns stay in the order they were said. The Voice column separates band voice (self), public narrative (other),
song, other-era members and unknown speakers.

In Colab, after Drive is mounted:
    %run "/content/drive/MyDrive/harrison_v7/Video for Harrison/build_video_oral_histories.py"
"""
import collections, datetime, hashlib, json, os, re, sys

VERSION = "video oral histories version 1.0 (18 Oct 2026; from merge_video_turns 1.3)"
CONTEXT = 400                     # characters of the neighbouring turns kept before and after each turn
CELL_CAP = 32000                  # Excel holds at most 32,767 characters in a cell


def script_folder():
    f = globals().get("__file__")
    return os.path.dirname(os.path.abspath(f)) if f else os.getcwd()


BASE = script_folder()
TURNS = os.path.join(BASE, "harrison_v7_VIDEO_TURNS.json")
FRAME = os.path.join(BASE, "acts_frame.json")
OUT = BASE

# the 29 columns of the text corpus's oral histories, then five of the video's own
COLUMNS = ["Date", "Years from nomination", "Period", "Voice", "Speaker", "Text", "Before", "After", "Date from",
           "Date precision", "Speaker role", "Reported member", "Words", "Years mentioned", "Confidence", "Main set",
           "Attribution rule", "Attribution evidence", "Outlet", "Title", "Article relevance", "Publication date", "Format",
           "URL",
           "Turn ID", "Record ID", "Position in article", "Later copies", "Act",
           "Start", "Video speaker label", "Language", "Text from", "Video flags"]

EXPLAIN = {
    "Date": "When the words were said: 1 January of the video's year (the selection's best guess; see Date from), "
            "since a video is dated only to the year.",
    "Date from": "How the year was found: in the title, by hand, beside a recording word in the description, or the "
                 "upload year (the latest the video can date from).",
    "Date precision": "'year' for every video.",
    "Years from nomination": "The video's year minus the year of the act's Best New Artist nomination (the Grammy "
                             "ceremony year): negative before it, 0 in it, positive after it.",
    "Period": "'before', 'nomination year' or 'after', by the same years.",
    "Voice": "'Band': a member of the act at the time of the nomination (for a solo artist, the artist), in their own "
             "words (band voice). 'Public narrative': interviewers, presenters, narrators and anyone else. 'Song': "
             "singing and song clips (never band voice). 'Other-era member': a member who joined after the nomination "
             "or had left before it. 'Unknown': speech whose speaker could not be told.",
    "Speaker": "Who speaks: members as named on the line-up list, named only on evidence in the video (said in it, "
               "or an on-screen caption); other people as the video names them. Otherwise the video's own label for the "
               "person, the same through the video ('band member, unspecified (A)', 'interviewer', 'unknown (B)'). A "
               "solo artist the video never names shows as their name in brackets.",
    "Speaker role": "Member, member (unidentified), the band as a group, other-era member, interviewer, narrator, other "
                    "person, or unknown.",
    "Reported member": "Not used for videos (second-hand speech is a text-corpus category).",
    "Text": "The turn's words: WhisperX's words where they line up with Gemini's, otherwise Gemini's text (turns not "
            "in English, which stay in their language, and members speaking under an interpreter). See Text from.",
    "Words": "Words in Text.",
    "Before": f"Up to {CONTEXT} characters of the turns just before (for an answer, usually the question), each after "
              "its speaker; ' / ' separates turns.",
    "After": f"Up to {CONTEXT} characters of the turns just after.",
    "Years mentioned": "Years said in the turn (1983, 1980s, '80s). A first pointer to the period an account "
                       "describes, found without AI; it is not the period itself.",
    "Confidence": "High, medium or low: how sure the speaker is (provisional until the spot check). Low: turns of "
                  "three words or fewer, stretches where Gemini gave nothing, labels to check across parts. High: turns "
                  "of 20 words or more whose words line up. Medium: the rest.",
    "Main set": "For band voice: 'yes' for high and medium confidence (the main set), 'no' for low.",
    "Attribution rule": "Where the speaker comes from: Gemini's label; Gemini's label, put back across parts; "
                        "WhisperX's voice (Gemini has nothing there); or guessed from WhisperX's voices (Gemini refused "
                        "the video).",
    "Attribution evidence": "For a named person, how the video names them (Gemini's evidence, with the time); "
                            "otherwise Gemini's description of the person.",
    "Outlet": "The show where known, otherwise YouTube and the channel.", "Title": "The video's title.",
    "Article relevance": "Not used for videos (every video was chosen for the act).",
    "Publication date": "The video's upload date.",
    "Format": "'video'.",
    "URL": "The video, opening two seconds before the turn.",
    "Turn ID": "The turn's ID in harrison_v7_VIDEO_TURNS (video-turn), for tracing it back.",
    "Record ID": "The video's ID (V1 to V8290, as in the selection).", "Position in article": "The turn's place in the video.",
    "Later copies": "0: re-uploads of the same interview are not yet detected for videos.",
    "Act": "The Best New Artist nominee.",
    "Start": "Where the turn starts in the video (minutes:seconds).",
    "Video speaker label": "The video's label for the person, kept the same through the video; joins the turn to the "
                           "People sheet of harrison_v7_VIDEO_TURNS.xlsx.",
    "Language": "The turn's language (en, fr, ja ...). Turns not in English are written as spoken.",
    "Text from": "WhisperX, Gemini, or WhisperX only (where Gemini gave nothing).",
    "Video flags": "Notes from the joining step: a label put back or to check across parts; mostly sung; times "
                   "estimated; speakers guessed.",
}

ROLE = {"band member": "Member", "band member, unspecified": "Member (unidentified)",
        "band collective": "The band as a group", "other-era member": "Other-era member",
        "interviewer": "Interviewer", "journalist": "Narrator", "other person": "Other person", "unknown": "Unknown",
        "ambiguous": "Unknown (ambiguous)"}

YEAR = re.compile(r"(?<![\d$£€.,])(?:19[4-9]\d|20[0-2]\d)(?:s\b|(?![\d,]))")
DECADE = re.compile(r"(?<![\w])['’](?:[4-9]0|[0-2]0)s\b")


def years_mentioned(s):
    out = []
    for m in YEAR.finditer(s or ""):
        y = m.group(0)
        if y not in out:
            out.append(y)
    for m in DECADE.finditer(s or ""):
        d = m.group(0)[1:3]
        y = ("19" if d[0] in "456789" else "20") + d + "s"
        if y not in out:
            out.append(y)
    return "; ".join(out)


def voice_of(t):
    st = t["Speaker type"] or ""
    if t["Band voice"] == "yes":
        return "Band"
    if t["Turn type"] == "song":
        return "Song"
    if st.startswith("other-era"):
        return "Other-era member"
    if st in ("unknown", "ambiguous"):
        return "Unknown"
    return "Public narrative"


def speaker_of(t):
    name, cands, st, lab = t["Speaker name"] or "", t["Speaker candidates"] or "", t["Speaker type"] or "", t["Video speaker label"] or ""
    if st == "unknown":
        return "Unknown"
    if st == "ambiguous":
        return f"Unclear ({cands})" if cands else "Unclear"
    if name == "band member, unspecified":
        return f"{lab} ({cands})" if cands else lab
    if name:
        return name
    if cands:
        return f"({cands})"
    return lab


def at_time(url, start):
    if not url or start is None:
        return url or ""
    return f"{url}{'&' if '?' in url else '?'}t={int(max(0, float(start) - 2))}s"     # two seconds early


def build_rows(turns):
    by_video = collections.defaultdict(list)
    for t in turns:
        by_video[t["Record ID"]].append(t)
    rows = collections.defaultdict(list)
    for vid, ts in by_video.items():
        ts.sort(key=lambda t: t["Turn order"])
        shown = [f"{speaker_of(t)}: {t['Turn text']}" for t in ts]
        for i, t in enumerate(ts):
            before = " / ".join(shown[max(0, i - 6):i])
            after = " / ".join(shown[i + 1:i + 7])
            if len(before) > CONTEXT:
                before = before[-CONTEXT:]
                before = "…" + before[before.index(" ") + 1:] if " " in before else before
            if len(after) > CONTEXT:
                after = after[:CONTEXT]
                after = after[:after.rindex(" ")] + "…" if " " in after else after
            year = int(t["Interview date"])
            gy = int(t["Grammy Year"])
            dy = year - gy
            bv = t["Band voice"] == "yes"
            df = (t["Interview date source"] or "").replace("video: ", "")
            row = {"Date": f"{year}-01-01", "Years from nomination": dy,
                   "Period": "before" if dy < 0 else ("nomination year" if dy == 0 else "after"),
                   "Voice": voice_of(t), "Speaker": speaker_of(t), "Text": t["Turn text"] or "", "Before": before,
                   "After": after, "Date from": df, "Date precision": "year",
                   "Speaker role": ROLE.get(t["Speaker type"], t["Speaker type"] or ""), "Reported member": "",
                   "Words": t["Words"], "Years mentioned": years_mentioned(t["Turn text"]), "Confidence": t["Confidence"],
                   "Main set": t["Main set"] if bv else "", "Attribution rule": t["Attribution rule"] or "",
                   "Attribution evidence": t["Attribution evidence"] or "", "Outlet": t["Outlet / Source Name"] or "",
                   "Title": t["Video title"] or "", "Article relevance": "", "Publication date": t["Publication Date"] or "",
                   "Format": "video", "URL": at_time(t["Video URL"], t["Start (s)"]), "Turn ID": t["Turn ID"],
                   "Record ID": vid, "Position in article": t["Turn order"], "Later copies": 0, "Act": t["Focal Act"],
                   "Start": t["Start"] or "", "Video speaker label": t["Video speaker label"] or "",
                   "Language": t["Language"] or "", "Text from": t["Text from"] or "", "Video flags": t["Video flags"] or ""}
            for act in [a.strip() for a in str(t["Focal Act"]).split(";") if a.strip()]:
                rows[act].append({**row, "Act": act})
    for act in rows:
        rows[act].sort(key=lambda x: (x["Date"], x["Publication date"], int(x["Record ID"][1:]), x["Position in article"]))
    return rows


def act_details(rows, frame):
    out = []
    for act in sorted(frame, key=lambda a: (frame[a]["year"], a.casefold())):
        f = frame[act]
        rs = rows.get(act, [])
        bv = [r for r in rs if r["Voice"] == "Band"]
        per = collections.Counter(r["Period"] for r in bv)
        dates = sorted(r["Date"] for r in bv)
        members = f.get("members") or ([act] if f["type"] == "Solo" else [])
        out.append({"Act": act, "Solo or group": f["type"], "Nominated": int(f["year"]),
                    "Result": "Winner" if f.get("winner") else "Nominee",
                    "Nomination-era members": ", ".join(members), "Other-era members": ", ".join(f.get("other_era") or []),
                    "Videos": len({r["Record ID"] for r in rs}), "Rows": len(rs), "Band-voice turns": len(bv),
                    "Before": per["before"], "Nomination year": per["nomination year"], "After": per["after"],
                    "Band-voice words": sum(r["Words"] for r in bv),
                    "Band-voice words in the main set": sum(r["Words"] for r in bv if r["Main set"] == "yes"),
                    "Public narrative words": sum(r["Words"] for r in rs if r["Voice"] == "Public narrative"),
                    "Song turns": sum(1 for r in rs if r["Voice"] == "Song"),
                    "Other-era turns": sum(1 for r in rs if r["Voice"] == "Other-era member"),
                    "Unknown turns": sum(1 for r in rs if r["Voice"] == "Unknown"),
                    "First band-voice year": dates[0][:4] if dates else "", "Last band-voice year": dates[-1][:4] if dates else "",
                    "Sheet": act if rs else ""})
    return out


def sheet_name(act, used):
    s = re.sub(r"[\[\]:*?/\\]", "", act).strip("'")[:31]
    base, i = s, 2
    while s.casefold() in used:
        s = f"{base[:28]} {i}"
        i += 1
    used.add(s.casefold())
    return s


README = [
    ("What this is", "Step 2 (videos) of the replication workflow: an oral history for each "
     "act, built from the joined video turns (harrison_v7_VIDEO_TURNS). Every turn of every video about the act is "
     "here, oldest first; a video's turns stay in the order they were said. The layout is that of the text corpus's "
     "oral histories (harrison_v7_oral_histories), with the same 29 columns and five of the video's own, so the two "
     "can be read side by side or stacked later."),
    ("How to use it", "Each act has its own sheet; the Bands sheet lists them with their counts and a link to each. "
     "Filter Voice to 'Band' for the band's own oral history, or to 'Public narrative' for the other layer. The Date "
     "column is a real date, so 'Sort Oldest to Newest' works; to keep each video's turns together, sort by Date, then "
     "Publication date, then Record ID, then Position in article. The URL opens the video at the turn."),
    ("Shading", "Band voice: white. Public narrative: pale grey. Song: pale green. Other-era members: pale blue. "
     "Unknown speakers: pale orange. Band voice outside the main set (low confidence): grey text."),
    ("Dates", "A video is dated to the year only (Date is 1 January of it): from its title for most videos, by hand "
     "for 50, and from the upload year for 118, which is the latest such a video can date from. Period describes when "
     "the account was given, not the period it describes; Years mentioned is a first pointer to the latter."),
    ("What is left out", "Nothing is left out of the joined turns: other-era members, unknown speakers and singing are "
     "kept and marked. The joining step itself left out lines WhisperX made up over silence or music, prompt echoes and "
     "transcription loops. Re-uploads of the same interview are not detected at this step, so Later copies is 0 "
     "(build_band_voice_jsonl.py leaves out such copies)."),
    ("Accuracy", "On the two videos checked by hand (V1158, V1319) the speaker was right for 84% of sentences, and "
     "band voice was 94% precise and 98% complete. A blind spot check of 211 sentences from 15 random stretches is "
     "described in the technical appendix."),
    ("Known limits", "Gemini refused three solo videos as a whole (Julian Lennon V2425, Jody Watley V3001, Irene Cara "
     "V342): their speakers are guessed from WhisperX's voices, at low confidence. 788 turns are not in English and are "
     "kept in their language. In long videos with several similar voices, which member speaks is less sure than "
     "whether a member speaks."),
    ("For the GenAI steps", "harrison_v7_video_oral_histories.json holds the same rows and columns for every act in "
     "one file ('columns', 'data'), plus the details of each act ('acts') and these notes ('about'). Unlike the "
     "workbook, it holds every text in full."),
    ("Version", VERSION),
]


def write_xlsx(path, rows, acts):
    try:
        import xlsxwriter
    except ImportError:                                  # not every Colab image has it
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "xlsxwriter"])
        import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name
    wb = xlsxwriter.Workbook(path, {"constant_memory": True, "strings_to_urls": False, "strings_to_formulas": False})
    bold = wb.add_format({"bold": True})
    title = wb.add_format({"bold": True, "font_size": 14})
    head = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "text_wrap": True, "valign": "top", "border": 1,
                          "border_color": "#BFBFBF"})
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})
    top = wb.add_format({"valign": "top"})
    datef = wb.add_format({"num_format": "yyyy-mm-dd", "valign": "top"})
    link = wb.add_format({"font_color": "#0563C1", "underline": 1, "valign": "top"})
    fills = {"Public narrative": "#EDEDED", "Song": "#E2EFDA", "Other-era member": "#DDEBF7", "Unknown": "#FCE4D6"}
    cf = {k: wb.add_format({"bg_color": v}) for k, v in fills.items()}
    greytext = wb.add_format({"font_color": "#808080"})

    rm = wb.add_worksheet("Read me")
    rm.set_column(0, 0, 26)
    rm.set_column(1, 1, 120)
    rm.write(0, 0, "Video oral histories: the Grammy Best New Artist nominees, 1980 to 1990", title)
    r = 2
    for k, v in README:
        rm.write(r, 0, k, bold)
        rm.write(r, 1, v, wrap)
        r += 1
    r += 1
    rm.write(r, 0, "THE COLUMNS", bold)
    r += 1
    for k in COLUMNS:
        rm.write(r, 0, k, bold)
        rm.write(r, 1, EXPLAIN[k], wrap)
        r += 1

    names, used = {}, {"read me", "bands"}
    for a in acts:
        if a["Sheet"]:
            names[a["Act"]] = sheet_name(a["Act"], used)

    bs = wb.add_worksheet("Bands")
    bcols = ["Act", "Solo or group", "Nominated", "Result", "Nomination-era members", "Other-era members", "Videos",
             "Rows", "Band-voice turns", "Before", "Nomination year", "After", "Band-voice words",
             "Band-voice words in the main set", "Public narrative words", "Song turns", "Other-era turns",
             "Unknown turns", "First band-voice year", "Last band-voice year", "Sheet"]
    widths = {"Act": 22, "Nomination-era members": 40, "Other-era members": 30, "Sheet": 10}
    for j, k in enumerate(bcols):
        bs.set_column(j, j, widths.get(k, 11))
        bs.write(0, j, k, head)
    bs.freeze_panes(1, 1)
    for i, a in enumerate(acts, 1):
        for j, k in enumerate(bcols):
            v = a[k]
            if k == "Sheet":
                if v:
                    q = names[a["Act"]].replace("'", "''")
                    bs.write_url(i, j, f"internal:'{q}'!A1", link, string="open")
                else:
                    bs.write(i, j, "no videos", top)
            elif k in ("Nomination-era members", "Other-era members"):
                bs.write(i, j, v, wrap)
            else:
                bs.write(i, j, v, top)
    bs.autofilter(0, 0, len(acts), len(bcols) - 1)

    cw = {"Date": 11, "Date from": 17, "Date precision": 9, "Years from nomination": 11, "Period": 14, "Voice": 15,
          "Speaker": 24, "Speaker role": 18, "Reported member": 9, "Text": 80, "Words": 7, "Before": 40, "After": 40,
          "Years mentioned": 12, "Confidence": 10, "Main set": 8, "Attribution rule": 28, "Attribution evidence": 30,
          "Outlet": 20, "Title": 30, "Article relevance": 9, "Publication date": 11, "Format": 8, "URL": 30,
          "Turn ID": 10, "Record ID": 8, "Position in article": 8, "Later copies": 7, "Act": 16, "Start": 8,
          "Video speaker label": 26, "Language": 8, "Text from": 12, "Video flags": 30}
    vcol, mcol = COLUMNS.index("Voice"), COLUMNS.index("Main set")
    H0 = 4                                      # header row of the table (0-based); data from H0 + 1
    for a in acts:
        if not a["Sheet"]:
            continue
        ws = wb.add_worksheet(names[a["Act"]])
        rs = rows[a["Act"]]
        ws.write(0, 0, a["Act"], title)
        ws.write(1, 0, f"{a['Solo or group']} · nominated {a['Nominated']} · {a['Result']} · nomination-era members: "
                       f"{a['Nomination-era members'] or '(none listed)'}"
                       + (f" · other-era members: {a['Other-era members']}" if a["Other-era members"] else ""))
        ws.write(2, 0, f"{a['Videos']:,} videos · {a['Rows']:,} rows · band voice: {a['Band-voice turns']:,} turns "
                       f"({a['Band-voice words']:,} words; {a['Before']} before the nomination year, "
                       f"{a['Nomination year']} in it, {a['After']:,} after) · public narrative: "
                       f"{a['Public narrative words']:,} words")
        for j, k in enumerate(COLUMNS):
            ws.set_column(j, j, cw.get(k, 12))
            ws.write(H0, j, k, head)
        ws.freeze_panes(H0 + 1, COLUMNS.index("Text"))      # the header row and Date to Speaker stay in view
        for i, row in enumerate(rs, H0 + 1):
            ws.set_row(i, 45)
            for j, k in enumerate(COLUMNS):
                v = row[k]
                if k == "Date" or (k == "Publication date" and v):
                    ws.write_datetime(i, j, datetime.datetime.strptime(v[:10], "%Y-%m-%d"), datef)
                elif k in ("Text", "Before", "After"):
                    v = (v or "").strip()
                    if len(v) > CELL_CAP:
                        v = v[:CELL_CAP - 50] + " … [cut here: the full text is in the JSON]"
                    ws.write_string(i, j, v, wrap)
                elif k == "URL" and v:
                    if len(v) <= 2000 and re.match(r"https?://", v):
                        ws.write_url(i, j, v, link, string=v)
                    else:
                        ws.write_string(i, j, v, top)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(i, j, v, top)
                else:
                    ws.write_string(i, j, "" if v is None else str(v), top)
        last = H0 + len(rs)
        ws.autofilter(H0, 0, last, len(COLUMNS) - 1)
        rng = f"A{H0 + 2}:{xl_col_to_name(len(COLUMNS) - 1)}{last + 1}"
        vref = f"${xl_col_to_name(vcol)}{H0 + 2}"
        mref = f"${xl_col_to_name(mcol)}{H0 + 2}"
        for k, fmt in cf.items():
            ws.conditional_format(rng, {"type": "formula", "criteria": f'={vref}="{k}"', "format": fmt})
        ws.conditional_format(rng, {"type": "formula", "criteria": f'=AND({vref}="Band",{mref}="no")', "format": greytext})
    wb.close()


def main(turns_path=None, frame_path=None, out=None):
    turns_path, frame_path, out = turns_path or TURNS, frame_path or FRAME, out or OUT
    print(VERSION)
    for what, p in (("harrison_v7_VIDEO_TURNS.json (from merge_video_turns.py 1.3 or later)", turns_path),
                    ("acts_frame.json", frame_path)):
        if not os.path.exists(p):
            print(f"STOPPED: {what} is not at {p}. Nothing was written.")
            return None
    tj = json.load(open(turns_path, encoding="utf-8"))
    frame = json.load(open(frame_path, encoding="utf-8"))
    turns = [dict(zip(tj["columns"], r)) for r in tj["data"]]
    rows = build_rows(turns)
    unknown_acts = sorted(set(rows) - set(frame))
    if unknown_acts:
        print(f"STOPPED: these acts are in the turns but not in acts_frame.json: {unknown_acts}. Nothing was written.")
        return None
    acts = act_details(rows, frame)
    allrows = [row for a in acts for row in rows.get(a["Act"], [])]
    about = {"version": VERSION, "what": README[0][1], "notes": {k: v for k, v in README[1:]}, "columns explained": EXPLAIN}
    with open(os.path.join(out, "harrison_v7_video_oral_histories.json"), "w", encoding="utf-8") as f:
        json.dump({"about": about, "acts": acts, "columns": COLUMNS,
                   "data": [[r[k] for k in COLUMNS] for r in allrows]}, f, ensure_ascii=False)
    write_xlsx(os.path.join(out, "harrison_v7_video_oral_histories.xlsx"), rows, acts)
    v = collections.Counter(r["Voice"] for r in allrows)
    fp = hashlib.md5(json.dumps([[r[k] for k in COLUMNS] for r in allrows], ensure_ascii=False).encode()).hexdigest()
    print(f"Fingerprint of the rows: {fp} (the same inputs always give the same fingerprint)")
    print(f"{len(acts)} acts, {sum(1 for a in acts if a['Band-voice turns'])} with band voice; {len(allrows):,} rows from "
          f"{len({r['Record ID'] for r in allrows})} videos: " + ", ".join(f"{k} {n:,}" for k, n in v.most_common()))
    for f in ("harrison_v7_video_oral_histories.json", "harrison_v7_video_oral_histories.xlsx"):
        print(f"  {f}: {os.path.getsize(os.path.join(out, f)) / 1e6:.1f} MB")
    return rows, acts


if __name__ == "__main__":
    main()
