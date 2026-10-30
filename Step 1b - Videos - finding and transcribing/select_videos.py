"""select_videos.py: helpers for select_videos_v2.py (version 1 of the video selection, 15 Oct 2026).

select_videos_v2.py uses this file's speaker check (is the act or a member the one talking, judged from the title)
and its tribute filter, and through it the act list, name lists and title rules of video_links_v2.py, which must
sit in the same folder. Run select_videos_v2.py, not this file, to choose the videos.

Version 1's own rule, replaced by version 2's: talk videos (interview, Q&A, panel, workshop, award or speech) whose TITLE shows the act or a
nomination-era member speaking, in English, 1 to 60 minutes long; per act, videos from around the
nomination (within 3 years, dated by title or description) first, then other dated ones, then undated,
best search rank first within each group, up to CAP minutes per act.
Run on its own (python select_videos.py [links.xlsx] [out.xlsx]) it writes that selection to
harrison_v7_VIDEO_SELECTION_v1.xlsx.
"""
import json, os, re, sys, types
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, "harrison_v7_VIDEO_LINKS.xlsx")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(DIR, "harrison_v7_VIDEO_SELECTION_v1.xlsx")
CAP = float(os.environ.get("CAP_MIN", "60"))
MAX_ONE = 60.0
WINDOW = 3

code = open(os.path.join(HERE, "video_links_v2.py"), encoding="utf-8").read().replace('if __name__ == "__main__":\n    df = main()', '')
vl = types.ModuleType("vl"); exec(compile(code, "vl", "exec"), vl.__dict__)

TALK = ["Interview", "Q&A", "Panel discussion", "Workshop or masterclass", "Award or speech"]
TRIBUTE = re.compile(r"\blegacy\b|\bin memoriam\b|\bfuneral\b|\bobituary\b|\btribute to\b|\bcelebrating the life\b", re.I)
TWO_DIGIT = re.compile(r"(?<!\d)\d{1,2}[-/.]\d{1,2}[-/.](\d{2})(?!\d)|['’](\d{2})(?!\d)")


def names_in(act, text):
    pos = []
    m = vl.ACT_RX[act].search(text)
    if m:
        pos.append(m.start())
    rx, lookup = vl.NAMES[act]
    if rx:
        for mm in rx.finditer(text):
            member, era, _ = lookup.get(vl._key(mm.group(0)), (None, None, False))
            if era == "nomination":
                pos.append(mm.start())
    return sorted(pos)


def speaker_ok(act, title, evidence):
    """The act or a member is the one talking: named before the talk word ('Big Country interview',
    'Sade talks'), or right after it ('Interview with Nu Shooz', 'Dick Clark Interviews Men at Work')."""
    t = vl.strip_songs(act, vl.FROM_SOURCE.sub(" ", title))[0]
    pos = names_in(act, t)
    if not pos:
        return False, "the act and its members are not named in the title"
    if evidence.startswith("title: the act or a member") or evidence.startswith("talk show:"):
        return True, ""
    kw = None
    for _, rx in vl.TITLE_KEPT:
        m = rx.search(t)
        if m and (kw is None or m.start() < kw.start()):
            kw = m
    if kw is None:
        return False, "no talk word in the title"
    if pos[0] < kw.start():
        return True, ""
    after = t[kw.end():]
    w = re.match(r"\s*(?:with|of|to|from|by|[:\-–—|])?\s*(?:the\s+)?", after, re.I)
    rest = after[w.end():] if w else after
    if names_in(act, rest[:60]) and min(names_in(act, rest[:60])) <= 2:
        return True, ""
    return False, "someone else may be talking (the act is named after the talk word)"


def year_guess(row):
    """Best-guess year, with two-digit years in the title ('3-15-89', "'85") when no four-digit year was found."""
    y, src = row["Year (best guess)"], str(row["Year from"])
    if src.startswith("upload"):
        m = TWO_DIGIT.search(str(row["Title"]))
        if m:
            yy = int(m.group(1) or m.group(2))
            if 70 <= yy <= 99:
                up = int(str(row["Upload date"])[:4]) if str(row["Upload date"])[:4].isdigit() else 2100
                if 1900 + yy <= up:
                    return 1900 + yy, "two-digit year in title"
    return y, src


def main():
    L = pd.read_excel(SRC, "Links")
    rows = []
    for i, r in L.iterrows():
        act = r["Act"]
        reason = ""
        ev = str(r["Type evidence"] or "")
        lang = str(r["Language"]) if pd.notna(r["Language"]) else ""
        y, ysrc = year_guess(r)
        dy = (y - r["Grammy year"]) if pd.notna(y) else None
        dated = not str(ysrc).startswith("upload")
        if r["Type"] not in TALK:
            reason = f"type: {r['Type']}"
        elif not (ev.startswith("title") or ev.startswith("talk show:")):
            reason = "type decided from the description only"
        elif lang[:2] not in ("", "en", "zx", "na"):
            reason = f"not in English ({lang})"
        elif r["Minutes"] < 1:
            reason = "under a minute"
        elif r["Minutes"] > MAX_ONE:
            reason = f"over {MAX_ONE:.0f} minutes"
        elif pd.notna(r["Other-era members named"]) and str(r["Other-era members named"]).strip() \
                and not (pd.notna(r["Members named"]) and str(r["Members named"]).strip()):
            reason = "only an other-era member is named"
        elif TRIBUTE.search(str(r["Title"])):
            reason = "about the act's legacy or a death, not the act talking"
        else:
            ok, why = speaker_ok(act, str(r["Title"]), ev)
            if not ok:
                reason = why
        tier = None
        if not reason:
            tier = 1 if (dated and dy is not None and abs(dy) <= WINDOW) else (2 if dated else 3)
        rows.append({**r.to_dict(), "Year (best guess)": y, "Year from": ysrc,
                     "Years from nomination": dy, "_eligible": not reason, "_reason": reason, "_tier": tier})
    D = pd.DataFrame(rows)
    order = {a: k for k, a in enumerate(vl.ACTS)}
    D["_rank"] = D["Best rank"].fillna(999)
    D["_absdy"] = D["Years from nomination"].abs().fillna(99)
    sel_minutes, chosen = {}, {}
    for act in vl.ACTS:
        E = D[(D["Act"] == act) & D["_eligible"]].sort_values(["_tier", "_absdy", "_rank"])
        used = 0.0
        for idx, r in E.iterrows():
            if used + r["Minutes"] > CAP:
                continue
            used += r["Minutes"]
            chosen[idx] = True
        sel_minutes[act] = used
    D["Selected"] = ["yes" if chosen.get(i) else "no" for i in D.index]
    D["Why"] = ""
    for i, r in D.iterrows():
        if r["Selected"] == "yes":
            D.at[i, "Why"] = {1: f"around the nomination (within {WINDOW} years)", 2: "dated, further from the nomination",
                              3: "undated (upload year only)"}[r["_tier"]]
        elif r["_eligible"]:
            D.at[i, "Why"] = f"eligible, but over the act's {CAP:.0f}-minute cap"
        else:
            D.at[i, "Why"] = r["_reason"]
    return D, sel_minutes


COLS = ["Act", "Grammy year", "Why", "Type", "Title", "Channel", "URL", "Minutes", "Year (best guess)", "Year from",
        "Years from nomination", "Show or event", "Members named", "Upload date", "Search source", "Best rank",
        "Notes", "Video ID"]


def write_book(D, path):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    order = {a: k for k, a in enumerate(vl.ACTS)}
    D = D.copy()
    D["_o"] = D["Act"].map(order)
    S = D[D["Selected"] == "yes"].sort_values(["_o", "_tier", "Year (best guess)", "_rank"])
    R = D[D["Selected"] != "yes"].sort_values(["_o", "Why", "_rank"])
    by = []
    for a in vl.ACTS:
        s_ = S[S["Act"] == a]
        by.append({"Act": a, "Grammy year": vl.ACTS[a]["year"], "Videos": len(s_),
                   "Minutes": round(s_["Minutes"].sum(), 1),
                   "Around the nomination (min)": round(s_[s_["_tier"] == 1]["Minutes"].sum(), 1),
                   "Further from it (min)": round(s_[s_["_tier"] == 2]["Minutes"].sum(), 1),
                   "Undated (min)": round(s_[s_["_tier"] == 3]["Minutes"].sum(), 1),
                   "Eligible, not chosen": int(((D["Act"] == a) & D["_eligible"] & (D["Selected"] != "yes")).sum())})
    B = pd.DataFrame(by)
    tot_h = S["Minutes"].sum() / 60
    readme = [
        ["Video selection", f"15 Oct 2026, from harrison_v7_VIDEO_LINKS.xlsx (video_links_v2). {len(S)} videos, "
                            f"{tot_h:.1f} hours, every act represented."],
        ["Rule", f"Talk videos only (interview, Q&A, panel, workshop, award or speech) whose title shows the act or a "
                 f"nomination-era member talking, in English, 1 to {MAX_ONE:.0f} minutes long. Per act, up to {CAP:.0f} "
                 f"minutes: first videos from around the nomination (within {WINDOW} years, dated by a year in the title "
                 f"or description), then other dated videos, then undated ones; best search rank first within each."],
        ["Why around the nomination first", "The text corpus has little band voice from those years (most of it is from "
                                           "long after), and period TV interviews are what video adds."],
        ["Left for later", "Documentaries (narrated, mostly others' voices), Unclear (check) rows (mostly performances of "
                           "lesser-known songs), videos typed from the description only, non-English videos, and anything "
                           "over the cap. Other candidates lists each with the reason."],
        ["To change it", "The transcription step reads the To transcribe sheet. Delete a row to drop a video; copy a row "
                         "in from Other candidates to add one."],
        ["Options", "A 45-minute cap gives about 37 hours; 60 minutes about 51 hours; 90 minutes about 74 hours."],
        ["Hours", f"Around the nomination {S[S['_tier']==1]['Minutes'].sum()/60:.1f} h; further from it "
                  f"{S[S['_tier']==2]['Minutes'].sum()/60:.1f} h; undated {S[S['_tier']==3]['Minutes'].sum()/60:.1f} h."],
    ]
    head = PatternFill("solid", fgColor="1F4E5F")
    widths = {"Act": 18, "Grammy year": 8, "Why": 30, "Type": 14, "Title": 60, "Channel": 22, "URL": 34, "Minutes": 8,
              "Year (best guess)": 9, "Year from": 20, "Years from nomination": 10, "Show or event": 16,
              "Members named": 22, "Upload date": 11, "Search source": 12, "Best rank": 7, "Notes": 30, "Video ID": 13}
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(readme).to_excel(xw, sheet_name="Read me", index=False, header=False)
        S[COLS].to_excel(xw, sheet_name="To transcribe", index=False)
        B.to_excel(xw, sheet_name="By act", index=False)
        R[COLS].to_excel(xw, sheet_name="Other candidates", index=False)
        for name, ws in xw.sheets.items():
            if name == "Read me":
                ws.column_dimensions["A"].width = 26
                ws.column_dimensions["B"].width = 120
                for row in ws.iter_rows():
                    for c in row:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
                    row[0].font = Font(bold=True)
                continue
            for j, c in enumerate(ws[1], start=1):
                c.fill, c.font = head, Font(bold=True, color="FFFFFF")
                c.alignment = Alignment(wrap_text=True, vertical="center")
                ws.column_dimensions[get_column_letter(j)].width = widths.get(c.value, 14)
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            if name in ("To transcribe", "Other candidates"):
                heads = [c.value for c in ws[1]]
                col = {h: heads.index(h) + 1 for h in heads}
                for row in ws.iter_rows(min_row=2):
                    for c in row:
                        c.alignment = Alignment(vertical="top", wrap_text=c.column in (col["Title"], col["Notes"], col["Why"]))
                    u = row[col["URL"] - 1]
                    if isinstance(u.value, str) and u.value.startswith("http"):
                        u.hyperlink, u.font = u.value, Font(color="1155CC", underline="single")
                    row[col["Upload date"] - 1].number_format = "yyyy-mm-dd"
                    row[col["Minutes"] - 1].number_format = "0.0"
    return S, B


if __name__ == "__main__":
    D, sm = main()
    write_book(D, OUT)
    S = D[D["Selected"] == "yes"]
    uniq = S.drop_duplicates("Video ID")
    print(f"CAP {CAP:.0f} min/act: {len(S)} rows, {len(uniq)} videos, {uniq['Minutes'].sum()/60:.1f} h; "
          f"around the nomination {S[S['_tier']==1]['Minutes'].sum()/60:.1f} h; acts with none: "
          f"{[a for a, m in sm.items() if m == 0]}")
    print(D[D["_eligible"]].groupby("_tier")["Minutes"].agg(["count", "sum"]))
    print(D["Why"].value_counts().head(15))
