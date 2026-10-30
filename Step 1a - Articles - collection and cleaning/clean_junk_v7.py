"""
===============================================================================
clean_junk_v7.py: REMOVE CLEAR-CUT JUNK AND NEAR-DUPLICATES
===============================================================================
Runs after clean_dates_v7.py. Reads harrison_v7_DATED.json (full text) and
writes harrison_v7_CLEAN.json and .xlsx.

WHY THE JSON, NOT THE XLSX
  Excel cuts a cell at 32,767 characters, so the workbook truncates the
  longest records, and these are disproportionately the long, substantial
  interviews. Working from the JSON means text comparisons use the real text.

WHAT IT REMOVES
  1. Film and TV databases: rottentomatoes, letterboxd, themoviedb, tvguide
     and similar. These are filmography listings ("Mark Knopfler Movies & TV
     Shows List") with no prose and nobody speaking. Some are films that
     merely share a name: "The Big Country" is the 1958 western, "The
     Breakfast Club" the John Hughes film.
  2. No prose at all: pages with no sentence-ending punctuation (navigation
     menus, link lists, site furniture).
  3. Orphans: records whose text never names the act and that list no
     member in Speaker Name(s).
  4. Near-duplicates. dedupe_v7.py catches identical text and identical
     canonical URLs. This catches the same article published on two regional
     editions of a site, where the addresses carry different article IDs and
     the text differs by a few characters (one version has a "Recommended"
     box, say). It compares the first 500 normalised characters within each
     act: two versions of one article share an opening, different articles
     almost never do. Paginated listicles stay separate.

WHAT IT DELIBERATELY DOES NOT REMOVE
  Records that never name the ACT but do name a member. These are
  legitimate: a Chrissie Hynde solo interview need never say "The
  Pretenders", and rejecting on the act name alone would delete some of the
  best material in the corpus.
  Records that mention the act only once. Whether such a record is about the
  act is a judgement call, made later and record by record (finalise_v7.py
  and clean7_v7.py), with a reason recorded for each exclusion.

ALSO
  It blanks implausible dates before 1975 (no act in the frame existed then;
  among them are Unix epoch artefacts, where a null timestamp became
  1970-01-01). The records stay; only the date is rejected, and the derived
  columns (Temporal Class, Period Record) are rebuilt to match.

Nothing is deleted outright: every removed record goes to the "Rejected"
sheet with its reason.
===============================================================================
"""

import hashlib
import json
import os
import re

import pandas as pd

DIR = "/content/drive/MyDrive/harrison_v7"
IN_JSON = os.path.join(DIR, "harrison_v7_DATED.json")
IN_XLSX = os.path.join(DIR, "harrison_v7_DATED.xlsx")
OUT_JSON = os.path.join(DIR, "harrison_v7_CLEAN.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_CLEAN.xlsx")
EXCEL_CELL_CAP = 32000

MIN_PLAUSIBLE_YEAR = 1975      # no act in the frame existed before this
PERIOD_YEAR = 1999
PREFIX_CHARS = 500             # shorter risks colliding on shared boilerplate

FILM_TV_DB = ["rottentomatoes", "letterboxd", "themoviedb", "tvguide.com",
              "allmovie", "metacritic", "justwatch", "fandango", "moviefone"]


def norm_text(s):
    s = str(s or "").lower()
    s = re.sub(r"\s+", " ", s)
    return re.sub(r"[^a-z0-9 ]", "", s).strip()


def sentence_count(s):
    return len(re.findall(r'[.!?]["\')\]]?\s', str(s or "")))


def load():
    if os.path.exists(IN_JSON):
        with open(IN_JSON, encoding="utf-8") as f:
            df = pd.DataFrame(json.load(f))
        print(f"Loaded {len(df)} records from JSON (full text)")
    else:
        df = pd.read_excel(IN_XLSX, sheet_name="Corpus")
        print(f"Loaded {len(df)} records from XLSX")
        print("  !! text truncated at 32k — the JSON is preferred")
    return df


def main():
    df = load().reset_index(drop=True)
    n0 = len(df)
    text = df["Extracted Text"].astype(str)
    url = df["URL"].astype(str)
    rejected = []

    def reject(mask, reason):
        nonlocal df, text, url
        for i in df.index[mask]:
            rec = df.loc[i].to_dict()
            rec["Rejection Reason"] = reason
            rejected.append(rec)
        df = df[~mask]
        text, url = text[~mask], url[~mask]

    # ── 1. film and TV databases ─────────────────────────────────────────────
    m = url.str.contains("|".join(FILM_TV_DB), case=False, na=False)
    print(f"\nfilm/TV database pages:   {int(m.sum())}")
    reject(m, "film/TV database listing — no prose, no speaker")

    # ── 2. no prose at all ───────────────────────────────────────────────────
    m = text.map(sentence_count) == 0
    print(f"no prose (site furniture): {int(m.sum())}")
    reject(m, "no prose — navigation/site furniture only")

    # ── 3. orphans: neither act nor any member named ─────────────────────────
    act_named = pd.Series(
        [str(df.at[i, "Focal Act"]).lower() in str(df.at[i, "Extracted Text"]).lower()
         for i in df.index], index=df.index)
    has_speaker = (df["Speaker Name(s)"].fillna("").astype(str).str.strip().ne(""))
    m = (~act_named) & (~has_speaker)
    print(f"orphans (no act, no member): {int(m.sum())}")
    reject(m, "neither focal act nor any member named in text")

    # ── 4. near-duplicates ───────────────────────────────────────────────────
    norm = text.map(norm_text)
    prefix = norm.str[:PREFIX_CHARS].map(
        lambda s: hashlib.md5(s.encode()).hexdigest() if len(s) >= PREFIX_CHARS else None)
    tmp = pd.DataFrame({"act": df["Focal Act"], "p": prefix}, index=df.index)
    m = tmp["p"].notna() & tmp.duplicated(subset=["act", "p"], keep="first")
    print(f"near-duplicates:           {int(m.sum())}")
    reject(m, f"near-duplicate — identical first {PREFIX_CHARS} characters")

    # ── 5. implausible dates: blank the DATE, keep the RECORD ────────────────
    pub = pd.to_datetime(df["Publication Date"], errors="coerce")
    bad = pub.notna() & (pub.dt.year < MIN_PLAUSIBLE_YEAR)
    print(f"\nimplausible pre-{MIN_PLAUSIBLE_YEAR} dates blanked: {int(bad.sum())} "
          f"(records kept)")
    for i in df.index[bad]:
        df.at[i, "Publication Date"] = ""
        df.at[i, "Date Precision"] = "none"
        df.at[i, "Date Confidence"] = "rejected"
        df.at[i, "Date Method"] = (f"rejected: implausible pre-{MIN_PLAUSIBLE_YEAR} "
                                   f"date (was {str(pub[i])[:10]})")

    # ── rebuild derived columns from the cleaned dates ───────────────────────
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

    # ── report ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"{n0} -> {len(df)} records  ({len(rej)} rejected, {100*len(rej)/n0:.1f}%)")
    print(f"acts remaining: {df['Focal Act'].nunique()}")
    if len(rej):
        print("\nrejected by reason:")
        for k, v in rej["Rejection Reason"].value_counts().items():
            print(f"   {v:5}  {k}")
    pa = df.groupby("Focal Act").size()
    print(f"\nper act — thinnest {pa.min()} ({pa.idxmin()}), "
          f"median {int(pa.median())}, richest {pa.max()} ({pa.idxmax()})")
    print(f"period records: {int(df['Period Record'].sum())}")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}  (full text — use this downstream)")


if __name__ == "__main__":
    main()
