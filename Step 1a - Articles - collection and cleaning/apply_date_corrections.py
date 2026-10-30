"""
===============================================================================
apply_date_corrections.py: APPLY THE HAND-CHECKED DATES
===============================================================================
Runs after clean_furniture_v7.py. Reads harrison_v7_CLEAN2.json and
body_context_dates_to_check.xlsx (the hand check), and writes
harrison_v7_CLEAN4.json and .xlsx. (CLEAN3 and CLEAN5 were one-off repairs of
the corpus collected for this study and are not part of this sequence; see
the README.)

WHY THIS EXISTS
  Rule 2 of clean_dates_v7.py dates a record from a phrase such as
  "originally published" followed by a year. It misfires: "first appeared in
  October 1984 on the group's debut album" is a song appearing on a record,
  not an article being published. Every record dated this way was checked by
  hand against the source, and the decisions are in
  body_context_dates_to_check.xlsx.

WHAT IT APPLIES

  1. CORRECTED DATES ("To check" sheet, Corrected Date filled). Many of these
     dates are found in neither the extracted date nor the page metadata, so
     they could only come from reading the source. Marked high confidence,
     method "hand-verified".

  2. CLEARED DATES (Corrected Date left blank). The date could not be
     determined from the source. It is removed rather than left wrong, and
     the confidence set to rejected.

  3. REJECTIONS ("Rejected" sheet). Wrong-entity matches found during the
     check (a Corey Hart interview filed under Adam and the Ants, a Djokovic
     tennis report under Take 6, Songfacts pages for other artists) and tag
     or archive index pages. Moved to the Rejected sheet.

  4. REPLACEMENT ARTICLES (URLs named in the Notes column). Where a tag or
     archive page was in the corpus instead of the article it pointed to, the
     article is fetched and added. A note with "instead" removes the
     original; "as well" or "also" keeps both.

  5. TAG PAGE EXPANSION (TAG_PAGES below). A tag page is an index of links,
     not an article. For earofnewt.com/tag/mark-knopfler/ the individual
     articles are fetched and added, keeping only those that mention the act
     or a member.

MATCHING
  On (Focal Act, URL) together, never URL alone: the same page can be a
  record for two acts. And never on row number, because the order of the
  records can change from one step to the next. Rows of the sheet with no
  match in the corpus are reported at the end of the run and skipped.

DATE FORMATS ACCEPTED
  A full date (1984-04-12) becomes day precision. A bare year (2016) becomes
  year precision (Excel stores those as plain numbers, which would otherwise
  be misread as 1970). "March 1986" or "1986-03" becomes month precision.
===============================================================================
"""

import json
import os
import re
import time
from datetime import date, datetime

import pandas as pd
import requests
import trafilatura

DIR = "/content/drive/MyDrive/harrison_v7"
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN2.json")
IN_XLSX = os.path.join(DIR, "harrison_v7_CLEAN2.xlsx")
REVIEW = os.path.join(DIR, "body_context_dates_to_check.xlsx")
OUT_JSON = os.path.join(DIR, "harrison_v7_CLEAN4.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_CLEAN4.xlsx")
EXCEL_CELL_CAP = 32000
PERIOD_YEAR = 1999

# Tag pages to expand into the articles they link to: (url, focal act)
TAG_PAGES = [
    ("https://earofnewt.com/tag/mark-knopfler/", "Dire Straits"),
]
MAX_PER_TAG_PAGE = 10

HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/124.0 Safari/537.36")}
MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"])}


def parse_corrected(v):
    """Returns (iso_date, precision) or (None, None) if blank."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None, None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime("%Y-%m-%d"), "day"
    s = str(v).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return None, None
    # bare year — Excel stores these as plain numbers
    m = re.fullmatch(r"(\d{4})(\.0)?", s)
    if m and 1900 <= int(m.group(1)) <= 2030:
        return f"{int(m.group(1)):04d}-01-01", "year"
    # "1986-03" — month and year only. Must be checked BEFORE the general
    # date parser, which would silently turn it into the 1st and call that
    # day precision.
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-01", "month"
    # "March 1986" / "Mar 1986"
    m = re.fullmatch(r"([A-Za-z]{3,9})\.?\s+(\d{4})", s)
    if m and m.group(1).lower()[:3] in MONTHS:
        return f"{int(m.group(2)):04d}-{MONTHS[m.group(1).lower()[:3]]:02d}-01", "month"
    try:
        return pd.to_datetime(s).strftime("%Y-%m-%d"), "day"
    except Exception:
        return None, None


def fetch_article(url):
    try:
        time.sleep(1.5)
        r = requests.get(url, headers=HEADERS, timeout=30)
        if r.status_code != 200 or not r.text:
            return None, None, f"HTTP {r.status_code}"
        txt = trafilatura.extract(r.text, favor_recall=True, include_comments=False)
        if not txt or len(txt.strip()) < 280:
            return None, r.text, "extraction too short"
        meta = trafilatura.extract_metadata(r.text)
        title = getattr(meta, "title", "") or ""
        d = getattr(meta, "date", "") or ""
        return {"text": txt.strip(), "title": title, "date": d}, r.text, "ok"
    except Exception as e:
        return None, None, type(e).__name__


def date_from_url(u):
    m = re.search(r"/((?:19|20)\d{2})/(\d{1,2})/(\d{1,2})/", str(u))
    if m:
        y, mo, dd = map(int, m.groups())
        if 1 <= mo <= 12 and 1 <= dd <= 31:
            return f"{y:04d}-{mo:02d}-{dd:02d}"
    return ""


def load_corpus():
    if os.path.exists(IN_JSON):
        with open(IN_JSON, encoding="utf-8") as f:
            df = pd.DataFrame(json.load(f))
        print(f"Loaded {len(df)} records from JSON (full text)")
    else:
        df = pd.read_excel(IN_XLSX, sheet_name="Corpus")
        print(f"Loaded {len(df)} records from XLSX (text truncated)")
    return df.reset_index(drop=True)


def main():
    df = load_corpus()
    n0 = len(df)

    chk = pd.read_excel(REVIEW, sheet_name="To check")
    try:                                   # the Rejected sheet has no header row
        rej = pd.read_excel(REVIEW, sheet_name="Rejected", header=None)
        rej.columns = list(chk.columns)[:len(rej.columns)]
    except Exception:
        rej = pd.DataFrame(columns=chk.columns)
    print(f"Review sheet: {len(chk)} to check, {len(rej)} rejected\n")

    key = lambda a, u: (str(a).strip(), str(u).strip())
    idx = {key(df.at[i, "Focal Act"], df.at[i, "URL"]): i for i in df.index}

    applied, cleared, missing, rejected_rows, drop = 0, 0, [], [], []
    superseded = []

    def note_verdict(note, url):
        """A note naming a replacement URL says what happens to the original.
        'instead' means the original goes; 'as well' means both stay."""
        n = note.lower()
        if not re.search(r"https?://", n):
            return "keep", ""
        if "instead" in n:
            return "reject", "superseded by the article it pointed to"
        if "as well" in n or "also" in n:
            return "keep", ""
        if re.search(r"/(tag|tags|category|author|archives?)/", str(url), re.I):
            return "reject", ("index page superseded by the article it pointed to "
                              "(note did not say 'instead' — assumed from the URL)")
        return "keep", ""

    # ── 1 & 2: corrections and clearances ────────────────────────────────────
    for _, r in chk.iterrows():
        k = key(r["Focal Act"], r["URL"])
        if k not in idx:
            missing.append(k)
            continue
        i = idx[k]
        iso, prec = parse_corrected(r.get("Corrected Date"))
        note = str(r.get("Notes") or "").strip()
        if note.lower() in ("nan", "none"):
            note = ""

        verdict, why = note_verdict(note, r["URL"])
        if verdict == "reject":
            rec = df.loc[i].to_dict()
            rec["Rejection Reason"] = f"hand-checked: {why}"
            rejected_rows.append(rec)
            drop.append(i)
            superseded.append((str(r["Focal Act"]), str(r["Title / Headline"])[:44], why))
            continue

        if iso:
            df.at[i, "Publication Date"] = iso
            df.at[i, "Date Precision"] = prec
            df.at[i, "Date Confidence"] = "high"
            df.at[i, "Date Method"] = "hand-verified"
            applied += 1
        else:
            df.at[i, "Publication Date"] = ""
            df.at[i, "Date Precision"] = "none"
            df.at[i, "Date Confidence"] = "rejected"
            df.at[i, "Date Method"] = ("hand-checked: date not determinable"
                                       + (f" ({note})" if note else ""))
            cleared += 1
        if note:
            old = str(df.at[i, "Notes"] or "").strip()
            if old.lower() in ("nan", "none"):
                old = ""
            df.at[i, "Notes"] = (old + "; " if old else "") + note

    # ── 3: rejections ────────────────────────────────────────────────────────
    for _, r in rej.iterrows():
        k = key(r["Focal Act"], r["URL"])
        if k not in idx:
            missing.append(k)
            continue
        i = idx[k]
        rec = df.loc[i].to_dict()
        rec["Rejection Reason"] = ("hand-checked: wrong entity or index page "
                                   + str(r.get("Notes") or "")).strip()
        rejected_rows.append(rec)
        drop.append(i)

    # ── 4: replacement articles named in Notes ───────────────────────────────
    to_fetch = []
    for _, r in pd.concat([chk, rej], ignore_index=True).iterrows():
        for u in re.findall(r"https?://\S+", str(r.get("Notes") or "")):
            to_fetch.append((u.rstrip(".,;"), str(r["Focal Act"]).strip(), r))

    # ── 5: tag pages expanded into their articles ────────────────────────────
    for tag_url, act in TAG_PAGES:
        print(f"expanding tag page: {tag_url}")
        _, html, status = fetch_article(tag_url)
        if not html:
            print(f"   could not fetch ({status})")
            continue
        host = re.match(r"https?://[^/]+", tag_url).group(0)
        links = re.findall(rf'href="({re.escape(host)}/(?:19|20)\d{{2}}/\d{{2}}/\d{{2}}/[^"#?]+)"', html)
        links = list(dict.fromkeys(links))[:MAX_PER_TAG_PAGE]
        print(f"   found {len(links)} article links")
        src = df[df["Focal Act"] == act]
        template = src.iloc[0] if len(src) else df.iloc[0]
        for u in links:
            to_fetch.append((u, act, template))

    # fetch and add
    added = 0
    existing = set(df["URL"].astype(str))
    for u, act, template in to_fetch:
        if u in existing:
            print(f"   already in corpus, skipped: {u[:70]}")
            continue
        art, _, status = fetch_article(u)
        if not art:
            print(f"   FAILED ({status}): {u[:70]}")
            continue
        names = [act] + [s.strip() for s in
                         str(template.get("Speaker Name(s)") or "").split(";") if s.strip()]
        if not any(n.lower() in art["text"].lower() for n in names if len(n) > 2):
            print(f"   no mention of {act}, skipped: {u[:60]}")
            continue
        rec = {c: "" for c in df.columns}
        rec.update({
            "Focal Act": act,
            "Grammy Year": template.get("Grammy Year"),
            "Title / Headline": art["title"] or u,
            "URL": u,
            "Outlet / Source Name": re.sub(r"^https?://(www\.)?", "", u).split("/")[0],
            "Extracted Text": art["text"],
            "Text Length": len(art["text"]),
            "Provenance": "open-web",
            "Retrieved On": date.today().isoformat(),
            "Notes": "added during hand-check: article behind a tag/index page",
            "Batch": template.get("Batch", ""),
        })
        d_ = art["date"] or date_from_url(u)
        if d_:
            rec["Publication Date"] = str(d_)[:10]
            rec["Date Precision"] = "day"
            rec["Date Confidence"] = "high"
            rec["Date Method"] = "page metadata / URL path (added record)"
        else:
            rec["Date Precision"] = "none"
            rec["Date Confidence"] = "rejected"
            rec["Date Method"] = "no date found (added record)"
        df.loc[len(df)] = rec
        existing.add(u)
        added += 1
        print(f"   added: {art['title'][:52]}")

    # ── finalise ─────────────────────────────────────────────────────────────
    df = df.drop(index=drop)
    clean = pd.to_datetime(df["Publication Date"], errors="coerce")
    era = pd.to_numeric(df.get("Content Era", pd.Series(index=df.index, dtype=object)),
                        errors="coerce")
    gy = pd.to_numeric(df["Grammy Year"], errors="coerce")
    eff = era.fillna(pd.Series(clean.dt.year, index=df.index))
    tc = pd.cut(eff - gy, bins=[-100, -0.5, 3.5, 10.5, 100],
                labels=["pre-nomination", "contemporaneous", "near-term", "retrospective"])
    df["Temporal Class"] = pd.Series(tc, index=df.index).astype(object).where(
        tc.notna(), "unknown").astype(str)
    df["Period Record"] = [
        bool((pd.notna(e) and e <= PERIOD_YEAR) or (pd.notna(d) and d.year <= PERIOD_YEAR))
        for e, d in zip(era, clean)]
    df = df.sort_values(["Grammy Year", "Focal Act", "Publication Date"],
                        na_position="last").reset_index(drop=True)

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, default=str)
    excel = df.copy()
    excel["Extracted Text"] = excel["Extracted Text"].astype(str).str.slice(0, EXCEL_CELL_CAP)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        if rejected_rows:
            rr = pd.DataFrame(rejected_rows)
            rr["Extracted Text"] = rr["Extracted Text"].astype(str).str.slice(0, 1500)
            rr.to_excel(xw, sheet_name="Rejected", index=False)

    print(f"\n{'='*60}")
    print(f"  dates corrected : {applied}")
    print(f"  dates cleared   : {cleared}")
    print(f"  records rejected: {len(rejected_rows)}")
    print(f"  articles added  : {added}")
    if superseded:
        print(f"\n  originals removed because a note named a replacement:")
        for act, title, why in superseded:
            print(f"     {act[:18]:18} {title}")
            print(f"        {why}")
    if missing:
        print(f"\n  !! {len(missing)} sheet rows had no match in the corpus:")
        for a, u in missing[:8]:
            print(f"     {a[:20]:20} {u[:64]}")
    print(f"\n{n0} -> {len(df)} records, {df['Focal Act'].nunique()} acts")
    print(f"hand-verified dates now in corpus: "
          f"{(df['Date Method'].astype(str) == 'hand-verified').sum()}")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}")


if __name__ == "__main__":
    main()
