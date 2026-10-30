"""
===============================================================================
clean_dates_v7.py: CHECK AND CORRECT PUBLICATION DATES
===============================================================================
Runs after dedupe_v7.py. Reads harrison_v7_DEDUPED.json (or .xlsx) and writes
harrison_v7_DATED.json and .xlsx.

WHY THIS IS NEEDED
  The dates extracted at collection go wrong in three ways:
    * for some pages trafilatura returns only the current year, which is
      stored as 1 January of that year (2026-01-01 in the collection for this
      study);
    * some pages carry the crawl date instead of the publication date;
    * many dates sit on 1 January of some year: the year may be right, but
      the day was filled in.
  Many of these can be recovered rather than blanked. muzines.co.uk hosts
  scanned 1980s magazines and puts the real date in the title, for example
  "Keyboards With Dire Straits (Sound On Sound, Sep 1986)". So the script
  gathers all the evidence first, marks the problems second, and decides only
  once, with everything in view.

HOW A DATE IS DECIDED, in order of precedence
  1. A citation in the title, "(Publication, Mon YYYY)": high confidence.
  2. A year next to a publication phrase in the first 4,000 characters of the
     body ("originally published", "first appeared", "reprinted from" ...):
     high.
  3. The extracted date, when the URL path (/YYYY/MM/) gives the same year:
     high.
  4. The URL path, when the extracted date is missing or disqualified: medium.
  5. The extracted date, marked as a conflict, when the URL path gives a
     different year: low.
  6. No date (rejected), when the extracted date is disqualified and there is
     nothing else. A date is disqualified when it equals the retrieval date,
     is 2026-01-01, is later than the page's Wayback capture, or is in the
     future.
  7. Otherwise the extracted date alone: medium, or low when it is 1 January.
  Then the date the page states in its own metadata (JSON-LD, meta tags or a
  <time> tag, recorded by harrison_corpus_v7.py) is checked against every
  date except those from rules 1 and 2, which are explicit statements and
  outrank machine metadata. A matching metadata date confirms the date; a
  different one replaces it. Either way the date becomes day-precise and high
  confidence.

NEW COLUMNS
  Publication Date (raw)               the collected value, never overwritten
  Publication Date (pre-verification)  the date before the metadata check
  Date Precision           day | month | year | none
  Date Confidence          high | medium | low | rejected
  Date Method              which rule decided it
  Date Source Publication  e.g. "Sound On Sound", where recovered from a title
  Content Era              the year the CONTENT dates from, where a title
                           citation or a publication phrase states it (not
                           the webpage date)
  Temporal Class           years from the nomination to the content (or the
                           date): pre-nomination (before it), contemporaneous
                           (0 to 3), near-term (4 to 10), retrospective (11 or
                           more), unknown
  Period Record            recomputed: the content or the date is from 1999
                           or earlier

The workbook also lists every changed date ("Dates changed") and every date
corrected from page metadata ("Dates corrected").

Run in Colab after mounting Drive.
===============================================================================
"""

import json
import os
import re
from datetime import date

import pandas as pd

DIR = "/content/drive/MyDrive/harrison_v7"
IN_JSON = os.path.join(DIR, "harrison_v7_DEDUPED.json")
IN_XLSX = os.path.join(DIR, "harrison_v7_DEDUPED.xlsx")
OUT_JSON = os.path.join(DIR, "harrison_v7_DATED.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_DATED.xlsx")
EXCEL_CELL_CAP = 32000

TODAY = pd.Timestamp(date.today())
MIN_YEAR, MAX_YEAR = 1975, TODAY.year

# A record counts as "period" if its CONTENT dates from this year or earlier.
# 1999 keeps the whole pre-2000 music press in scope while excluding modern
# retrospectives. Change PERIOD_YEAR for a different boundary.
PERIOD_YEAR = 1999

# Dates set by an explicit statement in the title or body: the page metadata
# check never overwrites them.
AUTHORITATIVE = {"title citation", "body publication context"}

MON = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
          "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}

# ═════════════════════════════════════════════════════════════════════════════
# STEP 2 — EVIDENCE EXTRACTORS.  Each is read-only.
# ═════════════════════════════════════════════════════════════════════════════
TITLE_RE = re.compile(rf"\(([^,()]{{2,45}}),\s*{MON}\.?\s*((?:19|20)\d{{2}})\)", re.I)
URL_RE = re.compile(r"/((?:19|20)\d{2})/(\d{1,2})(?:/(\d{1,2}))?(?=/|$|[-_.])")
CTX = (r"(originally published|first published|first appeared|reprinted from|"
       r"from the archives?|this article appeared)")
BODY_MY_RE = re.compile(rf"{CTX}[^.\n]{{0,60}}?{MON}\.?\s*((?:19|20)\d{{2}})", re.I)
BODY_Y_RE = re.compile(rf"{CTX}[^.\n]{{0,80}}?((?:19|20)\d{{2}})", re.I)
WB_RE = re.compile(r"web\.archive\.org/web/(\d{4})(\d{2})(\d{2})")


def _ok_year(y):
    return MIN_YEAR <= y <= MAX_YEAR


def ev_title(title):
    """(Publication, Mon YYYY) — an explicit human-authored citation."""
    m = TITLE_RE.search(str(title or ""))
    if not m:
        return None
    y = int(m.group(3))
    if not _ok_year(y):
        return None
    return {"y": y, "m": MONTHS[m.group(2).lower()[:3]], "d": None,
            "pub": m.group(1).strip()}


def ev_url(url):
    """/YYYY/MM/ or /YYYY/MM/DD/ in the path — a CMS convention."""
    m = URL_RE.search(str(url or ""))
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    if not (_ok_year(y) and 1 <= mo <= 12):
        return None
    dd = m.group(3)
    d_ = int(dd) if dd and 1 <= int(dd) <= 31 else None
    return {"y": y, "m": mo, "d": d_}


def ev_body(text):
    """A year sitting next to an explicit publication-context phrase.
    Deliberately strict — this is the fix for the era-hint problem, where any
    loose year in the opening lines was treated as evidence."""
    s = str(text or "")[:4000]
    m = BODY_MY_RE.search(s)
    if m:
        y = int(m.group(3))
        if _ok_year(y):
            return {"y": y, "m": MONTHS[m.group(2).lower()[:3]], "d": None}
    m = BODY_Y_RE.search(s)
    if m:
        y = int(m.group(2))
        if _ok_year(y):
            return {"y": y, "m": None, "d": None}
    return None


def ev_wayback(url):
    """Capture timestamp — an UPPER BOUND on publication, never a date."""
    m = WB_RE.search(str(url or ""))
    if not m:
        return None
    try:
        return pd.Timestamp(f"{m.group(1)}-{m.group(2)}-{m.group(3)}")
    except Exception:
        return None


def to_ts(ev):
    """Evidence dict -> timestamp, filling missing month/day with 1."""
    if not ev:
        return None
    try:
        return pd.Timestamp(year=ev["y"], month=ev.get("m") or 1,
                            day=ev.get("d") or 1)
    except Exception:
        return None


def _text(v):
    """A cell as a string, with empty, None and NaN all as ""."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    v = str(v).strip()
    return "" if v.lower() in ("nan", "nat", "none", "<na>") else v


def precision_of(ev):
    if not ev:
        return "none"
    if ev.get("d"):
        return "day"
    if ev.get("m"):
        return "month"
    return "year"


# ═════════════════════════════════════════════════════════════════════════════
def main():
    if os.path.exists(IN_JSON):
        with open(IN_JSON, encoding="utf-8") as f:
            df = pd.DataFrame(json.load(f))
        print(f"Loaded {len(df)} records from JSON (full text)")
    else:
        df = pd.read_excel(IN_XLSX, sheet_name="Corpus")
        print(f"Loaded {len(df)} records from XLSX (text truncated)")
    n = len(df)

    # ── STEP 1: preserve the original ────────────────────────────────────────
    df["Publication Date (raw)"] = df["Publication Date"]
    raw = pd.to_datetime(df["Publication Date"], errors="coerce")

    # ── STEP 2: gather evidence, change nothing ──────────────────────────────
    e_title = df["Title / Headline"].map(ev_title)
    e_url = df["URL"].map(ev_url)
    e_body = df["Extracted Text"].map(ev_body)
    e_wb = df["URL"].map(ev_wayback)
    print("\nSTEP 2 — evidence found:")
    for lbl, s in [("title citation", e_title), ("URL path", e_url),
                   ("body context", e_body), ("wayback capture", e_wb)]:
        print(f"   {lbl:18} {s.notna().sum():5}")

    # ── STEP 3: mark disqualifying signals ───────────────────────────────────
    f_crawl = raw.dt.date.astype(str) == df["Retrieved On"].astype(str)
    f_jan26 = raw.dt.date.astype(str) == "2026-01-01"
    f_afterwb = e_wb.notna() & (raw > e_wb)
    f_future = raw > TODAY
    disq = f_crawl | f_jan26 | f_afterwb | f_future
    print("\nSTEP 3 — disqualifying flags:")
    for lbl, s in [("= retrieval date", f_crawl), ("= 2026-01-01", f_jan26),
                   ("after wayback capture", f_afterwb), ("in the future", f_future)]:
        print(f"   {lbl:22} {int(s.sum()):5}")
    print(f"   {'ANY':22} {int(disq.sum()):5}")

    # ── STEP 4-6: adjudicate once, with precedence ───────────────────────────
    dates, prec, conf, meth, srcpub = [], [], [], [], []
    for i in range(n):
        t, u, b, w = e_title.iat[i], e_url.iat[i], e_body.iat[i], e_wb.iat[i]
        r, bad = raw.iat[i], bool(disq.iat[i])

        if t:                                   # explicit citation wins
            dates.append(to_ts(t)); prec.append(precision_of(t))
            conf.append("high"); meth.append("title citation")
            srcpub.append(t["pub"])
        elif b:                                 # explicit body statement
            dates.append(to_ts(b)); prec.append(precision_of(b))
            conf.append("high"); meth.append("body publication context")
            srcpub.append("")
        elif u and pd.notna(r) and not bad and r.year == u["y"]:
            dates.append(r)                     # two sources agree
            prec.append("day"); conf.append("high")
            meth.append("extracted, corroborated by URL"); srcpub.append("")
        elif u and (bad or pd.isna(r)):         # URL rescues a bad date
            dates.append(to_ts(u)); prec.append(precision_of(u))
            conf.append("medium"); meth.append("URL path"); srcpub.append("")
        elif u and pd.notna(r) and r.year != u["y"]:
            dates.append(r)                     # sources conflict — keep, flag
            prec.append("day"); conf.append("low")
            meth.append("CONFLICT: extracted vs URL disagree"); srcpub.append("")
        elif bad:                               # wrong, nothing to fall back on
            dates.append(pd.NaT); prec.append("none"); conf.append("rejected")
            meth.append("rejected: " + ("crawl date" if f_crawl.iat[i] else
                        "2026-01-01 fallback" if f_jan26.iat[i] else
                        "after wayback capture" if f_afterwb.iat[i] else "future"))
            srcpub.append("")
        elif pd.notna(r):                       # trafilatura alone
            jan1 = (r.month == 1 and r.day == 1)
            dates.append(r)
            prec.append("year" if jan1 else "day")
            conf.append("low" if jan1 else "medium")
            meth.append("extracted only" + (" (year-only, day inferred)" if jan1 else ""))
            srcpub.append("")
        else:
            dates.append(pd.NaT); prec.append("none"); conf.append("rejected")
            meth.append("no date found"); srcpub.append("")

    clean = pd.Series(dates, index=df.index)
    df["Publication Date"] = clean.dt.strftime("%Y-%m-%d").where(clean.notna(), "")

    # ── STEP 5: the page's own metadata date ─────────────────────────────────
    # harrison_corpus_v7.py records the date each page states in its own
    # metadata. It confirms or replaces every date except those set by a title
    # citation or a body publication context.
    df["Publication Date (pre-verification)"] = df["Publication Date"]
    corrected = [False] * n
    has_meta = "Metadata Date (verified)" in df.columns
    if has_meta:
        stored_all = df["Publication Date"].tolist()
        meta_all = [_text(v) for v in df["Metadata Date (verified)"]]
        for i in range(n):
            stored = str(stored_all[i] or "")[:10]
            meta = meta_all[i]
            if meth[i] in AUTHORITATIVE or not meta:
                continue                            # kept as decided above
            if meta == stored:                      # confirmed
                prec[i], conf[i] = "day", "high"
                meth[i] = meth[i] + " + metadata confirmed"
            else:                                   # corrected
                clean.iat[i] = pd.to_datetime(meta, errors="coerce")
                stored_all[i] = meta
                prec[i], conf[i] = "day", "high"
                meth[i] = "corrected from page metadata"
                corrected[i] = True
        df["Publication Date"] = stored_all
    else:
        print("\nNo page metadata dates in the input "
              "(column 'Metadata Date (verified)'): dates are not checked "
              "against page metadata.")

    df["Date Precision"] = prec
    df["Date Confidence"] = conf
    df["Date Method"] = meth
    df["Date Source Publication"] = srcpub

    # ── STEP 8: rebuild derived fields ───────────────────────────────────────
    # Content Era: when the CONTENT dates from — only where there is an
    # explicit citation or publication context. NOT any loose year.
    era = []
    for i in range(n):
        t, b = e_title.iat[i], e_body.iat[i]
        era.append(t["y"] if t else (b["y"] if b else ""))
    df["Content Era"] = era

    gy = pd.to_numeric(df["Grammy Year"], errors="coerce")
    eff_year = pd.Series([e if e else (d.year if pd.notna(d) else None)
                          for e, d in zip(era, clean)], index=df.index)
    delta = pd.to_numeric(eff_year, errors="coerce") - gy
    # NOTE: pd.cut(...).astype(str) PRESERVES NaN rather than producing the
    # string "nan", so a .replace("nan", ...) never fires. Convert explicitly.
    tc = pd.cut(delta, bins=[-100, -0.5, 3.5, 10.5, 100],
                labels=["pre-nomination", "contemporaneous",
                        "near-term", "retrospective"])
    df["Temporal Class"] = pd.Series(tc, index=df.index).astype(object).where(
        tc.notna(), "unknown").astype(str)

    df["Period Record"] = [
        bool((e and e <= PERIOD_YEAR) or (pd.notna(d) and d.year <= PERIOD_YEAR))
        for e, d in zip(era, clean)]

    # The collection script's era hint fires on any year from 1975 to 1999
    # near the top of an article, so it flags about half the corpus. "Content
    # Era" above replaces it; the old column is renamed so it is not used by
    # mistake.
    if "Content Era Hint" in df.columns:
        df = df.rename(columns={"Content Era Hint": "Content Era Hint (superseded)"})

    # ── STEP 7: validation ───────────────────────────────────────────────────
    print("\nSTEP 7 — validation:")
    kept = ~pd.Series(corrected, index=df.index)
    bad_wb = (kept & e_wb.notna() & clean.notna() & (clean > e_wb)).sum()
    bad_fut = (kept & clean.notna() & (clean > TODAY)).sum()
    print(f"   dates after their wayback capture: {bad_wb}  (must be 0)")
    print(f"   dates in the future:               {bad_fut}  (must be 0)")
    if has_meta:
        cor = ~kept
        m_wb = (cor & e_wb.notna() & clean.notna() & (clean > e_wb)).sum()
        m_fut = (cor & clean.notna() & (clean > TODAY)).sum()
        print(f"   dates taken from page metadata that are after the wayback "
              f"capture: {m_wb}, in the future: {m_fut}  (listed in the "
              f"'Dates corrected' sheet)")
    blank_but_precise = ((clean.isna()) & (pd.Series(prec) != "none")).sum()
    kept_but_none = ((clean.notna()) & (pd.Series(prec) == "none")).sum()
    print(f"   blank date with a precision label: {blank_but_precise}  (must be 0)")
    print(f"   dated record with no precision:    {kept_but_none}  (must be 0)")
    nan_tc = (df["Temporal Class"].isna() |
              df["Temporal Class"].astype(str).isin(["nan", "NaN"])).sum()
    print(f"   temporal class left as NaN:        {nan_tc}  (must be 0)")
    tit = clean[[m == "title citation" for m in meth]]
    if len(tit):
        print(f"   title-recovered dates, median year: {int(tit.dt.year.median())} "
              f"(expect ~1980s)")

    # ── write ────────────────────────────────────────────────────────────────
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, default=str)
    excel = df.copy()
    excel["Extracted Text"] = excel["Extracted Text"].astype(str).str.slice(0, EXCEL_CELL_CAP)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        chg = df[df["Publication Date"] != df["Publication Date (raw)"].astype(str)]
        cols = ["Focal Act", "Title / Headline", "URL", "Publication Date (raw)",
                "Publication Date", "Date Precision", "Date Confidence",
                "Date Method", "Date Source Publication"]
        chg[cols].to_excel(xw, sheet_name="Dates changed", index=False)
        if has_meta:
            ch = df[pd.Series(corrected, index=df.index)]
            ch[["Focal Act", "Title / Headline", "URL",
                "Publication Date (pre-verification)", "Publication Date",
                "Metadata Date Source", "Temporal Class"]].to_excel(
                xw, sheet_name="Dates corrected", index=False)

    # ── report ───────────────────────────────────────────────────────────────
    print(f"\n{'='*62}\nRESULT\n{'='*62}")
    print("Confidence:")
    for k, v in pd.Series(conf).value_counts().items():
        print(f"   {k:10} {v:6} ({100*v/n:4.1f}%)")
    print("Precision:")
    for k, v in pd.Series(prec).value_counts().items():
        print(f"   {k:10} {v:6} ({100*v/n:4.1f}%)")
    print("Method:")
    for k, v in pd.Series(meth).value_counts().head(8).items():
        print(f"   {v:6}  {k}")
    if has_meta:
        src = df["Metadata Date Source"].map(_text)
        print("\nPage metadata check:")
        print(f"   no metadata date      {int((src == 'no metadata date').sum()):6}")
        print(f"   page unavailable      {int(src.isin(['page unavailable', 'not fetched']).sum()):6}")
        print(f"   dates confirmed       {sum(1 for m in meth if m.endswith('metadata confirmed')):6}")
        print(f"   dates corrected       {sum(corrected):6}")
    print(f"\nContent Era present: {sum(1 for e in era if e)}")
    print("Temporal class:")
    for k, v in df["Temporal Class"].value_counts().items():
        print(f"   {k:18} {v:6}")
    print(f"\nPeriod Record: {int(df['Period Record'].sum())}")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}")


if __name__ == "__main__":
    main()
