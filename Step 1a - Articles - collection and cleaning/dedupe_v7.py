"""
===============================================================================
dedupe_v7.py: REMOVE DUPLICATE RECORDS
===============================================================================
Runs after harrison_corpus_v7.py. Reads harrison_v7_MERGED.json (full text;
the .xlsx is the fallback) and writes harrison_v7_DEDUPED.json and .xlsx.

Removes records that are the same page reached by a different URL.

TWO LAYERS, within each focal act:

  Layer 1 — canonical URL.  Most "duplicates" are one page wearing different
    clothes:
      * Wayback captures      web/20120202135021id_/nme.com/x
                              web/20120202084758id_/nme.com/x   -> same page
      * locale variants       nme.com/en_asia/x  vs  nme.com/en_au/x
      * tracking tokens       daddario.com/x?srsltid=AAA  vs  ?srsltid=BBB
    Canonicalising reveals the shared underlying address.

  Layer 2 — normalised text.  Catches genuine same-content-different-source
    cases that survive layer 1.

DEDUPLICATION IS WITHIN ACT ONLY. The same page can legitimately be a record
for two different focal acts (e.g. a Top of the Pops episode featuring both
Simply Red and Soul II Soul). Those are preserved.

QUERY STRINGS ARE NOT DROPPED WHOLESALE. Parameters like id=, p=, aid= identify
distinct content: in the corpus collected for this study, 336 URLs use id=.
Only known tracking parameters are stripped. Dropping all query strings would
have wrongly merged 101 groups of genuinely different pages.

Nothing is silently deleted: removed rows go to a "Removed - duplicates" sheet
with the reason and the URL they duplicated.

Run in Colab after mounting Drive.
===============================================================================
"""

import json
import os
import re

import pandas as pd
from urllib.parse import urlparse, parse_qsl, urlencode

DIR = "/content/drive/MyDrive/harrison_v7"
IN_JSON = os.path.join(DIR, "harrison_v7_MERGED.json")     # full text
IN_XLSX = os.path.join(DIR, "harrison_v7_MERGED.xlsx")     # fallback (truncated)
OUT_JSON = os.path.join(DIR, "harrison_v7_DEDUPED.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_DEDUPED.xlsx")
EXCEL_CELL_CAP = 32000

# Parameters that carry no content — safe to strip.
TRACKING = {
    "srsltid", "fbclid", "gclid", "msclkid", "ref", "referrer", "source",
    "mc_cid", "mc_eid", "_ga", "igshid", "spm", "share", "amp",
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "utm_id", "utm_name",
}
# NOTE: id, p, page, pageid, aid, story, vid are NOT stripped — they identify
# distinct content on several of the collected sites.


def canonical_url(u: str) -> str:
    """Reduce a URL to the underlying page it addresses."""
    u = str(u or "")
    # 1. unwrap Wayback:  web.archive.org/web/<timestamp>id_/<real url>
    m = re.search(r"web\.archive\.org/web/\d+(?:id_)?/(https?://.+)$", u)
    if m:
        u = m.group(1)
    try:
        p = urlparse(u)
    except Exception:
        return u.lower()
    # 2. host: lowercase, drop www. and any port
    host = p.netloc.lower().replace("www.", "").split(":")[0]
    # 3. path: drop locale segments, drop trailing slash
    path = re.sub(r"/(en_[a-z]{2,4}|en-[a-z]{2})(?=/)", "", p.path, flags=re.I)
    path = path.rstrip("/")
    # 4. query: keep content params, drop tracking, sort for stable ordering.
    #    keep_blank_values=True is essential: some sites put the content id in
    #    a valueless parameter (forum.songfacts.com/index.php?/topic/7193-...),
    #    and without it parse_qsl discards that, collapsing every thread on the
    #    site into one URL.
    kept = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True)
            if k.lower() not in TRACKING]
    q = urlencode(sorted(kept))
    # 5. fragment: normally an in-page anchor and safely dropped, BUT our
    #    pipeline appends #seg1, #seg2 ... when it splits an aggregate page
    #    into one record per interview. Those mark genuinely distinct content
    #    (172 records from 21 source pages), so preserve them.
    seg = ""
    if re.fullmatch(r"seg\d+", p.fragment or "", flags=re.I):
        seg = f"#{p.fragment.lower()}"
    return (f"{host}{path}?{q}{seg}" if q else f"{host}{path}{seg}").lower()


def normalise_text(s: str) -> str:
    """Lowercase, collapse whitespace, strip punctuation — absorbs cosmetic
    differences (curly vs straight quotes, stray line breaks) without
    touching the words."""
    s = str(s or "").lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s.strip()


def load():
    if os.path.exists(IN_JSON):
        with open(IN_JSON, encoding="utf-8") as f:
            df = pd.DataFrame(json.load(f))
        print(f"Loaded {len(df)} records from JSON (full text)")
    elif os.path.exists(IN_XLSX):
        df = pd.read_excel(IN_XLSX, sheet_name="Corpus")
        print(f"Loaded {len(df)} records from XLSX")
        print("  !! text is truncated at 32k in this file — prefer the JSON")
    else:
        raise FileNotFoundError(f"No merged corpus found in {DIR}")
    return df


def main():
    df = load()
    n0 = len(df)

    df["_canon_url"] = df["URL"].map(canonical_url)
    df["_norm_text"] = df["Extracted Text"].map(normalise_text)

    # Keep-which-copy: prefer a record with a date, then an author, then the
    # longest text. Sorting first means keep="first" retains the best copy.
    def has(col):
        return (df[col].fillna("").astype(str).str.strip()
                .replace({"nan": "", "None": "", "<NA>": ""}).ne("").astype(int))

    df["_rank_date"] = -has("Publication Date")
    df["_rank_auth"] = -has("Author / Interviewer")
    df["_rank_len"] = -pd.to_numeric(df["Text Length"], errors="coerce").fillna(0)
    df = df.sort_values(["_rank_date", "_rank_auth", "_rank_len"]).reset_index(drop=True)

    removed = []

    # ── Layer 1: canonical URL, within act ───────────────────────────────────
    dup1 = df.duplicated(subset=["Focal Act", "_canon_url"], keep="first")
    kept_map = (df[~dup1].set_index(["Focal Act", "_canon_url"])["URL"].to_dict())
    for _, r in df[dup1].iterrows():
        rec = r.to_dict()
        rec["Removal Reason"] = "duplicate: same canonical URL"
        rec["Duplicate Of"] = kept_map.get((r["Focal Act"], r["_canon_url"]), "")
        removed.append(rec)
    df = df[~dup1]
    print(f"\nLayer 1 (canonical URL): removed {int(dup1.sum())}")

    # ── Layer 2: normalised text, within act ─────────────────────────────────
    dup2 = df.duplicated(subset=["Focal Act", "_norm_text"], keep="first")
    kept_map2 = (df[~dup2].set_index(["Focal Act", "_norm_text"])["URL"].to_dict())
    for _, r in df[dup2].iterrows():
        rec = r.to_dict()
        rec["Removal Reason"] = "duplicate: identical text, different URL"
        rec["Duplicate Of"] = kept_map2.get((r["Focal Act"], r["_norm_text"]), "")
        removed.append(rec)
    df = df[~dup2]
    print(f"Layer 2 (normalised text): removed {int(dup2.sum())}")

    # ── restore order, drop helper columns ───────────────────────────────────
    df["Grammy Year"] = pd.to_numeric(df["Grammy Year"], errors="coerce")
    df = (df.sort_values(["Grammy Year", "Focal Act", "Publication Date"],
                         na_position="last")
            .drop(columns=[c for c in df.columns if c.startswith("_")])
            .reset_index(drop=True))

    rem = pd.DataFrame(removed)
    if len(rem):
        rem = rem.drop(columns=[c for c in rem.columns if c.startswith("_")])

    # ── write ────────────────────────────────────────────────────────────────
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False)

    excel = df.copy()
    excel["Extracted Text"] = excel["Extracted Text"].astype(str).str.slice(0, EXCEL_CELL_CAP)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        excel.to_excel(xw, sheet_name="Corpus", index=False)
        if len(rem):
            r2 = rem.copy()
            r2["Extracted Text"] = r2["Extracted Text"].astype(str).str.slice(0, 2000)
            r2.to_excel(xw, sheet_name="Removed - duplicates", index=False)
        summary = (df.groupby(["Focal Act", "Grammy Year"]).size()
                     .reset_index(name="Records")
                     .sort_values(["Grammy Year", "Focal Act"]))
        summary.to_excel(xw, sheet_name="Summary", index=False)

    # ── report ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"{n0} -> {len(df)} records  ({n0 - len(df)} removed, "
          f"{100*(n0-len(df))/n0:.1f}%)")
    print(f"Acts still present: {df['Focal Act'].nunique()}")
    cross = (df.groupby("Extracted Text")["Focal Act"].nunique() > 1).sum()
    print(f"Cross-act shared texts preserved: {cross}")
    if len(rem):
        print("\nRemoved by reason:")
        for k, v in rem["Removal Reason"].value_counts().items():
            print(f"   {v:5}  {k}")
        print("\nMost affected acts:")
        for a, v in rem["Focal Act"].value_counts().head(6).items():
            print(f"   {v:5}  {a}")
    print(f"\nWrote:\n  {OUT_XLSX}\n  {OUT_JSON}  (full text)")


if __name__ == "__main__":
    main()
