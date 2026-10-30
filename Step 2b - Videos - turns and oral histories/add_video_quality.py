"""add_video_quality.py: measure the transcription quality of every video turn.
Step 2 (videos). Version 1.0, 19 October 2026.

Adds two measures of transcription quality to every video turn, and a flag for the turns to set aside:
  Crosstalk                  the share of the turn's words spoken while another voice is also speaking
                             (WhisperX marks each word it hears over another voice)
  Words in both transcripts  the share of the turn's words that the other transcript also has, word for word
                             (WhisperX's text against Gemini's, or Gemini's text against WhisperX's words)
  Set aside for crosstalk    'yes' when at least half of the turn's words are crosstalk (a provisional cut-off;
                             the spot check of the video speakers is described in the technical appendix)
Each video also gets the two measures over all its words, and the counts set aside (Videos sheet).
Nothing else changes: the joining is not run again, and every other column and row stays as it is.

Reads, in the folder this script is in (Drive: MyDrive/harrison_v7/Video for Harrison):
  harrison_v7_VIDEO_TURNS.json and .xlsx  the joined turns (merge_video_turns.py 1.3)
  video_turns/                            the joining step's file for each video (each turn's other transcript)
  video_transcripts/                      the WhisperX transcripts (each word's crosstalk mark)
  acts_frame.json, merge_video_turns.py and build_video_oral_histories.py (used as they are, not changed)
and writes there, under new names, leaving the current files as they are:
  harrison_v7_VIDEO_TURNS_quality.json and .xlsx            the turns table with the three new columns
  harrison_v7_video_oral_histories_quality.json and .xlsx   the oral histories with the three new columns

In Colab (any runtime; no GPU and no key needed), after Drive is mounted:
  import os; B = "/content/drive/MyDrive/harrison_v7/Video for Harrison"
  print(len(os.listdir(B + "/video_transcripts")), len(os.listdir(B + "/video_turns")))
  %run "/content/drive/MyDrive/harrison_v7/Video for Harrison/add_video_quality.py"
The listing first makes Drive show the whole folders. It takes a few minutes and can be run again at any time;
it only ever writes the four files ending in _quality.
"""
import bisect, collections, difflib, hashlib, importlib, json, os, sys

VERSION = "add_video_quality 1.0 (19 Oct 2026)"
CROSSTALK_CUT = 0.5      # a turn is set aside when at least this share of its words are crosstalk (provisional)
SLACK = 0.05             # seconds either side of a turn's start and end when picking WhisperX's words for it
NEW_COLS = ["Crosstalk", "Words in both transcripts", "Set aside for crosstalk"]
TAG = "_quality"         # added to the names of the new files


def script_folder():
    f = globals().get("__file__")
    return os.path.dirname(os.path.abspath(f)) if f else os.getcwd()


IN = {"turns": "harrison_v7_VIDEO_TURNS.json", "turns_xlsx": "harrison_v7_VIDEO_TURNS.xlsx",
      "per_video": "video_turns", "whisperx": "video_transcripts", "frame": "acts_frame.json",
      "merge": "merge_video_turns.py", "oh": "build_video_oral_histories.py"}
OUT = {"turns": "harrison_v7_VIDEO_TURNS" + TAG + ".json", "turns_xlsx": "harrison_v7_VIDEO_TURNS" + TAG + ".xlsx",
       "oh": "harrison_v7_video_oral_histories" + TAG + ".json",
       "oh_xlsx": "harrison_v7_video_oral_histories" + TAG + ".xlsx"}
ORIGINAL_OH = "harrison_v7_video_oral_histories.json"     # read only, to show that its rows are unchanged

EXPLAIN = {
    "Crosstalk": "The share of the turn's words spoken while another voice is also speaking, from 0 to 1. WhisperX "
                 "marks each word it hears over another voice; for a turn with Gemini's text, the words WhisperX heard "
                 "in the same stretch are used. Blank where WhisperX heard no spoken words there (only singing, or "
                 "nothing).",
    "Words in both transcripts": "The share of the turn's words that the other transcript also has, word for word and "
                 "in the same order, from 0 to 1: WhisperX's text against Gemini's for the turn, or Gemini's text "
                 "against WhisperX's words for the same stretch. Blank for 'WhisperX only' turns (Gemini has nothing "
                 "there). Low for turns not in English, where WhisperX often took down an interpreter.",
    "Set aside for crosstalk": f"'yes' when at least {CROSSTALK_CUT:.0%} of the turn's words are crosstalk, so who "
                 "speaks is less sure (a provisional cut-off until the spot check). For the GenAI steps, leave these "
                 "turns out of the band voice; they are kept here.",
}
VIDEO_EXPLAIN = ("The Videos sheet's last four columns: Crosstalk and Words in both transcripts over all the video's "
                 "words, and the turns and band-voice words set aside for crosstalk.")


def load_module(base, name):
    """merge_video_turns.py and build_video_oral_histories.py from this folder, used as they are."""
    if base not in sys.path:
        sys.path.insert(0, base)
    if name in sys.modules:
        return importlib.reload(sys.modules[name])
    return importlib.import_module(name)


def words(text, norm):
    return [n for n in (norm(w) for w in str(text or "").split()) if n]


def matched(a, b):
    """Words of a that b also has, in the same order."""
    if not a or not b:
        return 0
    return sum(m.size for m in difflib.SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks())


def whisperx_marks(path, mvt):
    """WhisperX's spoken (not sung) words in time order: their start times and whether each is crosstalk."""
    wx = json.load(open(path, encoding="utf-8"))
    ws = []
    for seg in wx.get("segments", []):
        for w in seg.get("words", []):
            if w.get("sung") or w.get("start") is None or not mvt.norm(mvt.clean(w.get("word"))):
                continue
            ws.append((float(w["start"]), bool(w.get("overlap"))))
    ws.sort()
    return [s for s, _ in ws], [o for _, o in ws]


def records(df):
    return df.astype(object).where(df.notna(), None).to_dict("records")


def fingerprint(rows):
    return hashlib.md5(json.dumps(rows, ensure_ascii=False).encode()).hexdigest()


def main(base=None):
    base = base or script_folder()
    print(VERSION)
    p = {k: os.path.join(base, v) for k, v in IN.items()}
    o = {k: os.path.join(base, v) for k, v in OUT.items()}
    missing = [f"{IN[k]}" for k in IN if not os.path.exists(p[k])]
    if missing:
        print("STOPPED: not found in " + base + ": " + ", ".join(missing) + ". Nothing was written.")
        return None
    assert not set(o.values()) & (set(p.values()) | {os.path.join(base, ORIGINAL_OH)}), "would overwrite an input"
    mvt = load_module(base, "merge_video_turns")
    ohb = load_module(base, "build_video_oral_histories")
    import pandas as pd

    # ---------------------------------------------------------------- the turns as they are
    tj = json.load(open(p["turns"], encoding="utf-8"))
    cols = tj["columns"]
    T = [dict(zip(cols, r)) for r in tj["data"]]
    by_video = collections.defaultdict(list)
    for t in T:
        by_video[t["Record ID"]].append(t)
    print(f"Read {len(T):,} turns from {len(by_video)} videos.", flush=True)

    # ---------------------------------------------------------------- the two measures, turn by turn
    V_extra, no_wx, no_pv, text_differs = {}, [], [], 0
    for n_done, (vid, ts) in enumerate(by_video.items(), 1):
        wp = os.path.join(p["whisperx"], vid + ".json")
        times, marks = whisperx_marks(wp, mvt) if os.path.exists(wp) else ([], [])
        if not times:
            no_wx.append(vid)
        pp = os.path.join(p["per_video"], vid + ".json")
        other = {}
        if os.path.exists(pp):
            for t in json.load(open(pp, encoding="utf-8"))["turns"]:
                other[f"{vid}-{t['turn']}"] = (t["text"], t["gemini_text"] if t["text_from"] == "WhisperX"
                                               else t["whisperx_text"])
        else:
            no_pv.append(vid)
        m_all = w_all = 0
        for t in ts:
            s, e = t["Start (s)"], t["End (s)"]
            cross = None
            if s is not None and e is not None and times:
                i = bisect.bisect_left(times, float(s) - SLACK)
                j = bisect.bisect_right(times, float(e) + SLACK)
                if j > i:
                    cross = round(sum(marks[i:j]) / (j - i), 3)
            agree = None
            if t["Text from"] != "WhisperX only" and t["Turn ID"] in other:
                own, oth = other[t["Turn ID"]]
                if (own or "") != (t["Turn text"] or ""):
                    text_differs += 1
                a = words(t["Turn text"], mvt.norm)
                if a:
                    m = matched(a, words(oth, mvt.norm))
                    agree = round(m / len(a), 3)
                    m_all += m
                    w_all += len(a)
            t["Crosstalk"] = cross
            t["Words in both transcripts"] = agree
            t["Set aside for crosstalk"] = "yes" if cross is not None and cross >= CROSSTALK_CUT else ""
        out = [t for t in ts if t["Set aside for crosstalk"] == "yes"]
        V_extra[vid] = {"Crosstalk": round(sum(marks) / len(marks), 3) if marks else None,
                        "Words in both transcripts": round(m_all / w_all, 3) if w_all else None,
                        "Turns set aside for crosstalk": len(out),
                        "Band-voice words set aside for crosstalk": sum(t["Words"] for t in out if t["Band voice"] == "yes")}
        if n_done % 100 == 0:
            print(f"  {n_done} of {len(by_video)} videos measured", flush=True)
    if no_wx:
        print(f"  no WhisperX words found for {len(no_wx)} videos (Crosstalk left blank): {no_wx[:10]}")
    if no_pv:
        print(f"  no joining file for {len(no_pv)} videos (Words in both transcripts left blank): {no_pv[:10]}")
    if text_differs:
        print(f"  NOTE: {text_differs} turns differ from their joining file's text; the turns table's text was used.")

    # ---------------------------------------------------------------- what it finds
    bv = [t for t in T if t["Band voice"] == "yes"]
    bw = sum(t["Words"] for t in bv)
    out_b = [t for t in bv if t["Set aside for crosstalk"] == "yes"]
    out_bm = [t for t in out_b if t["Main set"] == "yes"]
    src_w = collections.Counter()
    src_m = collections.Counter()
    for t in T:
        if t["Words in both transcripts"] is not None:
            n = len(words(t["Turn text"], mvt.norm))
            src_w[t["Text from"]] += n
            src_m[t["Text from"]] += t["Words in both transcripts"] * n
    cross_v = [v["Crosstalk"] for v in V_extra.values() if v["Crosstalk"] is not None]
    heavy = sum(1 for x in cross_v if x > 0.10)
    summary = (f"Set aside for crosstalk (at least {CROSSTALK_CUT:.0%} of a turn's words): "
               f"{sum(1 for t in T if t['Set aside for crosstalk'] == 'yes'):,} turns, of which {len(out_b):,} band-voice "
               f"turns with {sum(t['Words'] for t in out_b):,} words ({sum(t['Words'] for t in out_b) / bw:.1%} of band "
               f"voice; {sum(t['Words'] for t in out_bm):,} of them in the main set), in "
               f"{len({t['Record ID'] for t in out_b})} videos.")
    agree_txt = ("Words in both transcripts: " + ", ".join(
        f"{k}'s text {src_m[k] / src_w[k]:.1%}" for k in ("WhisperX", "Gemini") if src_w[k]) + ".")
    cross_txt = (f"Crosstalk: {heavy} of {len(cross_v)} videos ({heavy / max(1, len(cross_v)):.0%}) have more than 10% "
                 f"of their words spoken over each other.")
    print(cross_txt)
    print(agree_txt)
    print(summary)

    notes = {"Quality measures": f"Added by {VERSION} as the last three columns of Turns, with nothing else changed. "
                                 + " ".join(f"{k}: {v}" for k, v in EXPLAIN.items()),
             "Quality, what it finds": f"{cross_txt} {agree_txt} {summary}",
             "Quality, Videos sheet": VIDEO_EXPLAIN}

    # ---------------------------------------------------------------- the turns table, under its new name
    cols_q = cols + NEW_COLS
    with open(o["turns"], "w", encoding="utf-8") as fh:
        json.dump({"about": {"version": tj["about"]["version"] + "; quality measures by " + VERSION,
                             "notes": {**tj["about"].get("notes", {}), **notes}},
                   "columns": cols_q, "data": [[t[c] for c in cols_q] for t in T]}, fh, ensure_ascii=False)
    print("Writing the turns workbook ...", flush=True)
    sh = pd.read_excel(p["turns_xlsx"], sheet_name=["Videos", "People", "Label repairs"])
    V = records(sh["Videos"])
    for v in V:
        v.update(V_extra.get(str(v["ID"]), {k: None for k in next(iter(V_extra.values()))}))
    P, R = records(sh["People"]), records(sh["Label repairs"])
    mvt.README = list(mvt.README) + list(notes.items())
    mvt.write_xlsx(o["turns_xlsx"], T, V, P, R)

    # ---------------------------------------------------------------- the oral histories, under their new name
    print("Writing the oral histories ...", flush=True)
    frame = json.load(open(p["frame"], encoding="utf-8"))
    cols_oh = list(ohb.COLUMNS)
    rows = ohb.build_rows(T)
    q = {t["Turn ID"]: {k: t[k] for k in NEW_COLS} for t in T}
    for rs in rows.values():
        for r in rs:
            r.update(q[r["Turn ID"]])
    acts = ohb.act_details(rows, frame)
    for a in acts:
        a["Band-voice words set aside for crosstalk"] = sum(
            r["Words"] for r in rows.get(a["Act"], []) if r["Voice"] == "Band" and r["Set aside for crosstalk"] == "yes")
    allrows = [r for a in acts for r in rows.get(a["Act"], [])]
    same = None
    po = os.path.join(base, ORIGINAL_OH)
    if os.path.exists(po):
        same = fingerprint(json.load(open(po, encoding="utf-8"))["data"]) == fingerprint([[r[k] for k in cols_oh] for r in allrows])
    ohb.COLUMNS = cols_oh + NEW_COLS
    ohb.EXPLAIN = {**ohb.EXPLAIN, **EXPLAIN}
    readme = []
    for k, v in ohb.README:
        if k == "For the GenAI steps":
            v = v.replace("harrison_v7_video_oral_histories.json", OUT["oh"])
        if k == "Version":
            v = f"{ohb.VERSION}; quality measures by {VERSION}"
        readme.append((k, v))
    vi = [k for k, _ in readme].index("Version")
    readme[vi:vi] = [("Quality measures", f"Added by {VERSION} as the last three columns, with nothing else changed. "
                                          f"{summary} The cut-off is provisional until the spot check.")]
    ohb.README = readme
    about = {"version": f"{ohb.VERSION}; quality measures by {VERSION}", "what": readme[0][1],
             "notes": {k: v for k, v in readme[1:]}, "columns explained": ohb.EXPLAIN}
    with open(o["oh"], "w", encoding="utf-8") as fh:
        json.dump({"about": about, "acts": acts, "columns": ohb.COLUMNS,
                   "data": [[r[k] for k in ohb.COLUMNS] for r in allrows]}, fh, ensure_ascii=False)
    ohb.write_xlsx(o["oh_xlsx"], rows, acts)

    # ---------------------------------------------------------------- report
    if same is True:
        print(f"Check: the {len(cols_oh)} original columns of the oral histories are unchanged in all "
              f"{len(allrows):,} rows (the same fingerprint as {ORIGINAL_OH}).")
    elif same is False:
        print(f"CHECK FAILED: the original columns differ from {ORIGINAL_OH}; is it from the same turns table?")
    print("Done. New files (the current files are not changed):")
    for k in ("turns", "turns_xlsx", "oh", "oh_xlsx"):
        print(f"  {OUT[k]}: {os.path.getsize(o[k]) / 1e6:.1f} MB")
    return T, V, rows, acts


if __name__ == "__main__":
    main()
