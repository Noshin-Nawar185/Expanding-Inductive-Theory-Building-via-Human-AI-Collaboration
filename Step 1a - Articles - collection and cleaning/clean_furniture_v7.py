"""
===============================================================================
clean_furniture_v7.py: REMOVE TABLES, NAVIGATION SHELLS AND LISTING SITES
===============================================================================
Runs after clean_junk_v7.py. Reads harrison_v7_CLEAN.json and writes
harrison_v7_CLEAN2.json and .xlsx.

WHY A SECOND PASS
  clean_junk_v7.py rejects records with NO sentence-ending punctuation at
  all. That is too narrow. A Wayback capture of a Rolling Stone discography
  table contains 321 pipe characters and three full stops, so it passes, but
  it is a table, not an article:

      |  |  |  |  |  |
      |  | Tone Loc |
      | The Essentials | Dec 2002 |  |

  Extracted text carries "|" where the original page had table cells, and long
  runs of very short lines where it had a navigation menu. Neither appears in
  prose, so both are reliable structural signals.

WHAT IT DOES
  1. REJECTS unambiguous furniture: pipe characters above 5% of all text, or
     more than 75% of lines very short. These are stripped tables and
     navigation shells ("Skip to main content / Open menu / Search", or
     "HOME NEWS REVIEWS PHOTOS VIDEOS POLITICS RADIO SHOP SUBSCRIBE").

  2. REJECTS totparchive.co.uk as a domain. It is queried as a specialist
     archive because a small trial returned only period records, but that was
     an artefact: a TV archive is period by definition, since it is an
     episode index. At full scale it returned about 380 Top of the Pops
     listings, all by one author, median 1,722 characters, only 11 with any
     quoted speech. They are schedule entries: a date, a presenter, a
     performer list.

  3. FLAGS the borderline band rather than deleting it: 2% to 5% pipe
     characters, 60% to 75% short lines, or five or more navigation phrases.
     Here sit pages with real content beneath a navigation header, such as a
     Billboard biography wrapped in menu chrome. Deleting these on a ratio
     would lose real material, so they stay, marked "possible site furniture"
     in Review Flag.

ALSO FLAGGED, NOT CHANGED
  Paginated articles. Rolling Stone splits long pieces across /3/, /4/, /7/
  and each page is collected as a separate record. They are parts of one
  interview, and merging them is a different operation from deduplication,
  so this pass only marks them ("paginated article — consider merging").

Nothing is deleted outright: rejections go to the "Rejected" sheet with
reasons.
===============================================================================
"""

import json
import os
import re
from urllib.parse import urlparse

import pandas as pd

DIR = "/content/drive/MyDrive/harrison_v7"
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN.json")
IN_XLSX = os.path.join(DIR, "harrison_v7_CLEAN.xlsx")
OUT_JSON = os.path.join(DIR, "harrison_v7_CLEAN2.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_CLEAN2.xlsx")
EXCEL_CELL_CAP = 32000

# Reject outright above these; flag between the flag and reject thresholds.
PIPE_REJECT, PIPE_FLAG = 0.05, 0.02
SHORTLINE_REJECT, SHORTLINE_FLAG = 0.75, 0.60
SHORT_LINE_CHARS = 25

# Domains that produced listings rather than articles at scale.
DOMAIN_REJECT = {
    "totparchive.co.uk": "TV episode listings, not interviews (380 records, "
                         "11 with any quoted speech)",
}

NAV_RE = re.compile(
    r"(?i)\b(skip to (main )?content|open menu|click to expand|sign in|log in|"
    r"subscribe|newsletter|cookie|privacy policy|terms of use|"
    r"all rights reserved|artist main|albums & songs)\b")


def pipe_ratio(s):
    s = str(s or "")
    return s.count("|") / max(len(s), 1)


def short_line_ratio(s):
    lines = [l.strip() for l in str(s or "").split("\n")]
    lines = [l for l in lines if l]
    if not lines:
        return 1.0
    return sum(1 for l in lines if len(l) < SHORT_LINE_CHARS) / len(lines)


def load():
    if os.path.exists(IN_JSON):
        with open(IN_JSON, encoding="utf-8") as f:
            df = pd.DataFrame(json.load(f))
        print(f"Loaded {len(df)} records from JSON (full text)")
    else:
        df = pd.read_excel(IN_XLSX, sheet_name="Corpus")
        print(f"Loaded {len(df)} records from XLSX (truncated — JSON preferred)")
    return df


def main():
    df = load().reset_index(drop=True)
    n0 = len(df)
    text = df["Extracted Text"].astype(str)

    pipes = text.map(pipe_ratio)
    shorts = text.map(short_line_ratio)
    navs = text.map(lambda s: len(NAV_RE.findall(str(s))))
    domain = df["URL"].astype(str).map(
        lambda u: urlparse(u).netloc.replace("www.", "").lower())

    rejected = []

    def reject(mask, reason):
        nonlocal df, text, pipes, shorts, navs, domain
        if not mask.any():
            return
        for i in df.index[mask]:
            rec = df.loc[i].to_dict()
            rec["Rejection Reason"] = reason
            rejected.append(rec)
        keep = ~mask
        df, text = df[keep], text[keep]
        pipes, shorts, navs, domain = pipes[keep], shorts[keep], navs[keep], domain[keep]

    # ── 1. domain rejections ─────────────────────────────────────────────────
    for dom, why in DOMAIN_REJECT.items():
        m = domain.eq(dom)
        print(f"\n{dom}: {int(m.sum())} records")
        reject(m, f"domain rejected — {why}")

    # ── 2. unambiguous structural furniture ──────────────────────────────────
    m = pipes > PIPE_REJECT
    print(f"pipe density >{PIPE_REJECT:.0%}:       {int(m.sum())}")
    reject(m, f"stripped HTML table — over {PIPE_REJECT:.0%} pipe characters")

    m = shorts > SHORTLINE_REJECT
    print(f"short lines >{SHORTLINE_REJECT:.0%}:        {int(m.sum())}")
    reject(m, f"navigation shell — over {SHORTLINE_REJECT:.0%} very short lines")

    # ── 3. flag the borderline band, do not delete ───────────────────────────
    border = ((pipes > PIPE_FLAG) | (shorts > SHORTLINE_FLAG) | (navs >= 5))
    existing = df.get("Review Flag", pd.Series("", index=df.index)).fillna("").astype(str)
    df["Review Flag"] = [
        (e + "; " if e.strip() else "") + "possible site furniture" if b else e
        for e, b in zip(existing, border)]
    print(f"\nflagged as possible furniture (kept): {int(border.sum())}")

    # ── 4. flag paginated articles ───────────────────────────────────────────
    page_re = re.compile(r"/(\d{1,2})/?$")
    base = df["URL"].astype(str).map(lambda u: page_re.sub("/", u))
    ispage = df["URL"].astype(str).map(lambda u: bool(page_re.search(u)))
    grouped = base.groupby([df["Focal Act"], base]).transform("size") > 1
    multi = ispage & grouped
    df["Review Flag"] = [
        (f + "; " if str(f).strip() else "") + "paginated article — consider merging"
        if m else f for f, m in zip(df["Review Flag"], multi)]
    print(f"flagged as paginated parts (kept):    {int(multi.sum())}")

    df = df.sort_values(["Grammy Year", "Focal Act", "Publication Date"],
                        na_position="last").reset_index(drop=True)

    # ── write ────────────────────────────────────────────────────────────────
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, default=str)

    excel = df.copy()
    excel["Extracted Text"] = excel["Extracted Text"].astype(str).str.slice(0, EXCEL_CELL_CAP)
    rej = pd.DataFrame(rejected)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        if len(rej):
            r2 = rej.copy()
            r2["Extracted Text"] = r2["Extracted Text"].astype(str).str.slice(0, 1500)
            r2.to_excel(xw, sheet_name="Rejected", index=False)
        (df.groupby(["Focal Act", "Grammy Year"]).size().reset_index(name="Records")
           .sort_values(["Grammy Year", "Focal Act"])
           .to_excel(xw, sheet_name="Summary", index=False))

    print(f"\n{'='*60}")
    print(f"{n0} -> {len(df)} records  ({len(rej)} rejected, {100*len(rej)/n0:.1f}%)")
    print(f"acts remaining: {df['Focal Act'].nunique()}")
    if len(rej):
        print("\nrejected by reason:")
        for k, v in rej["Rejection Reason"].value_counts().items():
            print(f"   {v:5}  {k[:64]}")
    pa = df.groupby("Focal Act").size()
    print(f"\nper act — thinnest {pa.min()} ({pa.idxmin()}), "
          f"median {int(pa.median())}, richest {pa.max()} ({pa.idxmax()})")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}  (full text — use downstream)")


if __name__ == "__main__":
    main()
