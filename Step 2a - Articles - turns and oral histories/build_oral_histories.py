"""build_oral_histories.py: chronological oral histories, one per act.
Step 2 (articles) of the replication workflow. Runs after turns_v7_stage4to9.py.

Reads the turns table (stage 3 version 3, stages 4 to 9 version 1.3), its records table, the stage 1 text, CLEAN7
(for each article's relevance) and harrison_v7_oral_histories_supplement.json (Amy Holland's interviews, added by
hand), with the acts' details (type, result, members) written into this file, and writes:
  harrison_v7_oral_histories.xlsx  for reading: a Read me, a Bands sheet (one row per act) and one sheet per act
  harrison_v7_oral_histories.json  for the GenAI steps: the same rows and columns for every act, plus the act details

Every turn of every article about the act is kept, in date order (oldest first; an article's turns stay in reading
order). Republished copies are dropped: the earliest copy stays and counts its later copies. The Voice column separates
band voice (self), second-hand self narrative, public narrative (other), other-era members and unknown speakers.

In Colab: put this file and the supplement in MyDrive/harrison_v7, next to CLEAN7 and the files stages 1 to 9 wrote, and run
    %run /content/drive/MyDrive/harrison_v7/build_oral_histories.py
"""
import collections, datetime, json, os, re, sys

VERSION = "oral histories version 1.2 (15 Oct 2026; from turns_v7 stages 4-9 version 1.3, stage 3 version 3; with Amy Holland's interviews added by hand; Sade a band)"
DIR = sys.argv[1] if len(sys.argv) > 1 else "/content/drive/MyDrive/harrison_v7"
OUT = sys.argv[2] if len(sys.argv) > 2 else DIR
IN = {"turns": "harrison_v7_turns_final.json", "records": "harrison_v7_records_final.json",
      "text": "harrison_v7_TURNS_stage1.json", "corpus": "harrison_v7_CLEAN7.json",
      "supplement": "harrison_v7_oral_histories_supplement.json"}
CONTEXT = 400                     # characters of text kept before and after each turn
CELL_CAP = 32000                  # Excel holds at most 32,767 characters in a cell

COLUMNS = ["Date", "Years from nomination", "Period", "Voice", "Speaker", "Text", "Before", "After", "Date from",
           "Date precision", "Speaker role", "Reported member", "Words", "Years mentioned", "Confidence", "Main set",
           "Attribution rule", "Attribution evidence", "Outlet", "Title", "Article relevance", "Publication date", "Format",
           "URL",
           "Turn ID", "Record ID", "Position in article", "Later copies", "Act"]

EXPLAIN = {
    "Date": "When the words were said: the year the text states for them ('told NME in 1983'), or else the article's "
            "publication date. A date known only to the month or the year is written as its first day.",
    "Date from": "'stated in the text' or 'publication date'.",
    "Date precision": "Whether the date is known to the day, the month or the year.",
    "Years from nomination": "The year of the date minus the year of the act's Best New Artist nomination (the "
                             "Grammy ceremony year): negative before it, 0 in it, positive after it.",
    "Period": "'before', 'nomination year' or 'after', by the same years.",
    "Voice": "'Band': a member of the act at the time of the nomination, in their own words (band voice). "
             "'Second-hand': the writer reporting a member's words without quotation marks. 'Public narrative': the "
             "writer's own words, interviewers' questions and comments, and other people's quotations. 'Other-era "
             "member': a member who joined after the nomination or had left before it. 'Unknown': speech whose "
             "speaker the rules could not tell.",
    "Speaker": "Who speaks: members as named on the line-up list; other people as the text names them; the writer is "
               "the byline, or the outlet when there is none. For an unidentified member or an ambiguous speaker, the "
               "people it could be follow in brackets.",
    "Speaker role": "Member, member (unidentified), the band as a group, other-era member, writer, interviewer, other "
                    "person, writer reporting a member, or unknown.",
    "Reported member": "For second-hand rows: the member whose words the writer reports.",
    "Text": "The turn's own words, verbatim. Attribution clauses between its parts and Q&A labels are left out; they "
            "are in Attribution evidence. In the workbook, the six turns longer than 32,000 characters (all the "
            "writer's narration) are cut and marked; the JSON holds them in full.",
    "Words": "Words in Text.",
    "Before": f"Up to {CONTEXT} characters of the article just before the turn (for an answer, usually the "
              "question). ' / ' marks a paragraph break.",
    "After": f"Up to {CONTEXT} characters of the article just after the turn.",
    "Years mentioned": "Years written in the turn (1983, 1980s, '80s). A first pointer to the period an account "
                       "describes, found without AI; it is not the period itself.",
    "Confidence": "High, medium, low or unknown: how sure the rules are of the speaker (technical appendix, Step 2, stage 9); "
                  "'added by hand' for the rows added by hand (see the Read me).",
    "Main set": "For band voice: 'yes' for high and medium confidence (the main set), 'no' for low.",
    "Attribution rule": "The rule that gave the speaker.",
    "Attribution evidence": "The words that triggered the rule (for example 'says Day').",
    "Outlet": "The source.", "Title": "The article's title.",
    "Article relevance": "How much the article is about the act (CLEAN7): 'named in title', 'named 2+ times', 'named once' "
                         "or 'not found'. Articles that name the act once are often lists, round-ups or pieces about "
                         "other artists; filter them out for a stricter set.",
    "Publication date": "The article's publication date.",
    "Format": "How the speaker is marked in the article: prose, labelled Q&A, unlabelled Q&A or first-person account.",
    "URL": "The article's page.",
    "Turn ID": "The turn's ID in the turns table (record-turn), for tracing a quotation back to its source.",
    "Record ID": "The article's ID in the corpus.", "Position in article": "The turn's place in the article.",
    "Later copies": "How many later copies of this turn (republished articles) were dropped.",
    "Act": "The Best New Artist nominee.",
}


# The acts: type and result from grammy_nominees.xlsx (55 nominees, 3 excluded); members from the name lists
# (harrison_v7_NAME_LISTS.json, nomination-era and other-era).
ACTS = {'Adam and the Ants': {'members': ['Adam Ant', 'Chris Hughes', 'Gary Tibbs', 'Marco Pirroni', 'Terry Lee Miall'],
                       'other-era': ['Andy Warren',
                                     'Dave Barbarossa',
                                     'Johnny Bivouac',
                                     'Jordan',
                                     'Kevin Mooney',
                                     'Leigh Gorman',
                                     'Lester Square',
                                     'Mark Ryan',
                                     'Matthew Ashman',
                                     'Paul Flanagan'],
                       'type': 'Band',
                       'winner': False},
 'Amy Holland': {'members': ['Amy Holland'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Asia': {'members': ['Carl Palmer', 'Geoff Downes', 'John Wetton', 'Steve Howe'],
          'other-era': ['Al Pitrelli', 'Greg Lake', 'John Payne', 'Mandy Meyer', 'Pat Thrall'],
          'type': 'Band',
          'winner': False},
 'Big Country': {'members': ['Bruce Watson', 'Mark Brzezicki', 'Stuart Adamson', 'Tony Butler'],
                 'other-era': ['Alan Wishart',
                               'Clive Parker',
                               'Pat Ahern',
                               'Pete Wishart',
                               'Jamie Watson',
                               'Mike Peters',
                               'Derek Forbes',
                               'Simon Hough',
                               'Tommie Paxton',
                               'Chris Stones',
                               'Reece Dobbin'],
                 'type': 'Band',
                 'winner': False},
 'Breakfast Club': {'members': ['Dan Gilroy',
                                'Ed Gilroy',
                                'Gary Burke',
                                'Paul Kauk',
                                'Randy Jackson',
                                'Stephen Bray'],
                    'other-era': ['Angie Smit', 'Madonna'],
                    'type': 'Band',
                    'winner': False},
 'Bruce Hornsby and the Range': {'members': ['Bruce Hornsby',
                                             'David Mansfield',
                                             'George Marinelli',
                                             'Joe Puerta',
                                             'John Molo',
                                             'Peter Harris'],
                                 'other-era': ['J. T. Thomas'],
                                 'type': 'Band',
                                 'winner': True},
 'Christopher Cross': {'members': ['Christopher Cross'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Corey Hart': {'members': ['Corey Hart'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Culture Club': {'members': ['Boy George', 'Jon Moss', 'Mikey Craig', 'Roy Hay'],
                  'other-era': ['Jon Suede'],
                  'type': 'Band',
                  'winner': True},
 'Cutting Crew': {'members': ['Colin Farley', 'Kevin MacMichael', 'Martin Beedle', 'Nick Van Eede', 'Tony Moore'],
                  'other-era': [],
                  'type': 'Band',
                  'winner': False},
 'Cyndi Lauper': {'members': ['Cyndi Lauper'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Dire Straits': {'members': ['David Knopfler', 'John Illsley', 'Mark Knopfler', 'Pick Withers'],
                  'other-era': ['Alan Clark', 'Guy Fletcher', 'Hal Lindes', 'Jack Sonni', 'Terry Williams'],
                  'type': 'Band',
                  'winner': False},
 'Eurythmics': {'members': ['Annie Lennox', 'Dave Stewart'], 'other-era': [], 'type': 'Duo', 'winner': False},
 'Frankie Goes to Hollywood': {'members': ['Brian Nash',
                                           'Holly Johnson',
                                           "Mark O'Toole",
                                           'Paul Rutherford',
                                           'Peter Gill'],
                               'other-era': ['Ambrose Reynolds', "Jed O'Toole"],
                               'type': 'Band',
                               'winner': False},
 'Freddie Jackson': {'members': ['Freddie Jackson'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Glass Tiger': {'members': ['Al Connelly', 'Alan Frew', 'Michael Hanson', 'Sam Reid', 'Wayne Parker'],
                 'other-era': [],
                 'type': 'Band',
                 'winner': False},
 'Indigo Girls': {'members': ['Amy Ray', 'Emily Saliers'], 'other-era': [], 'type': 'Duo', 'winner': False},
 'Irene Cara': {'members': ['Irene Cara'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'James Ingram': {'members': ['James Ingram'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Jennifer Holliday': {'members': ['Jennifer Holliday'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Jody Watley': {'members': ['Jody Watley'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Julian Lennon': {'members': ['Julian Lennon'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Katrina and the Waves': {'members': ['Alex Cooper', 'Katrina Leskanich', 'Kimberley Rew', 'Vince de la Cruz'],
                           'other-era': ['Bob Jakins'],
                           'type': 'Band',
                           'winner': False},
 'Luther Vandross': {'members': ['Luther Vandross'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Men Without Hats': {'members': ['Allan McCarthy', 'Colin Doroschuk', 'Ivan Doroschuk', 'Stefan Doroschuk'],
                      'other-era': ['Al Gunn',
                                    'Bruce Murphy',
                                    'Dave Hill',
                                    'Denis Toupin',
                                    'Félix Matte',
                                    'Heidi Garcia',
                                    'Jean-Marc Pisapia',
                                    'John Gurrin',
                                    'John Kastner',
                                    'Johnny Souranis',
                                    'Jérémie Arrobas',
                                    'Lenny Pinkas',
                                    'Lysanne Thibodeau',
                                    'Marika Tjelios',
                                    'Michel Langevin',
                                    'Pete Seabrooke',
                                    'Pierre Taillefer',
                                    'Richard Samson',
                                    'Roman Martyn',
                                    'Tracy Howe'],
                      'type': 'Band',
                      'winner': False},
 'Men at Work': {'members': ['Colin Hay', 'Greg Ham', 'Jerry Speiser', 'John Rees', 'Ron Strykert'],
                 'other-era': [],
                 'type': 'Band',
                 'winner': True},
 'Musical Youth': {'members': ['Dennis Seaton', 'Frederick Waite', 'Kelvin Grant', 'Michael Grant', 'Patrick Waite'],
                   'other-era': ['Frederick Waite Sr.'],
                   'type': 'Band',
                   'winner': False},
 'Neneh Cherry': {'members': ['Neneh Cherry'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Nu Shooz': {'members': ['John Smith', 'Valerie Day'], 'other-era': [], 'type': 'Duo', 'winner': False},
 'Rick Astley': {'members': ['Rick Astley'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Rickie Lee Jones': {'members': ['Rickie Lee Jones'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Robbie Dupree': {'members': ['Robbie Dupree'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Sade': {'members': ['Andrew Hale', 'Paul S. Denman', 'Sade Adu', 'Stuart Matthewman'],
          'other-era': ['Dave Early', 'Paul Anthony Cooke'],
          'type': 'Band',          # grammy_nominees.xlsx says Solo; corrected 15 Oct 2026 (four members; a band in Harrison et al.)
          'winner': True},
 'Sheena Easton': {'members': ['Sheena Easton'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Sheila E.': {'members': ['Sheila E.'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Simply Red': {'members': ['Chris Joyce',
                            'Fritz McIntyre',
                            'Ian Kirkham',
                            'Mick Hucknall',
                            'Sylvan Richardson',
                            'Tim Kellett',
                            'Tony Bowers'],
                'other-era': ['Aziz Ibrahim', 'David Fryman', 'Gota Yashiki', 'Heitor Pereira', 'Shaun Ward'],
                'type': 'Band',
                'winner': False},
 'Soul II Soul': {'members': ['Aitch B', 'Caron Wheeler', 'Daddae Harvey', 'Jazzie B', 'Nellee Hooper', 'Simon Law'],
                  'other-era': ['Doreen Waddell',
                                'Kym Mazelle',
                                'Lamya Al-Mugheiry',
                                'Marcia Lewis',
                                'Rose Windross',
                                'Victoria Wilson-James'],
                  'type': 'Collective',
                  'winner': False},
 'Stray Cats': {'members': ['Brian Setzer', 'Lee Rocker', 'Slim Jim Phantom'],
                'other-era': ['Bob Beecher', 'Gary Setzer', 'Tommy Byrnes'],
                'type': 'Band',
                'winner': False},
 'Swing Out Sister': {'members': ['Andy Connell', 'Corinne Drewery', 'Martin Jackson'],
                      'other-era': [],
                      'type': 'Band/Duo',
                      'winner': False},
 'Take 6': {'members': ['Alvin Chea',
                        'Cedric Dent',
                        'Claude McKnight',
                        'David Thomas',
                        'Mark Kibble',
                        'Mervyn Warren'],
            'other-era': ['Joel Kibble'],
            'type': 'Group',
            'winner': False},
 "Terence Trent D'Arby": {'members': ["Terence Trent D'Arby"], 'other-era': [], 'type': 'Solo', 'winner': False},
 "The Go-Go's": {'members': ['Belinda Carlisle',
                             'Charlotte Caffey',
                             'Gina Schock',
                             'Jane Wiedlin',
                             'Kathy Valentine'],
                 'other-era': ['Elissa Bello', 'Margot Olavarria', 'Paula Jean Brown'],
                 'type': 'Band',
                 'winner': False},
 'The Human League': {'members': ['Ian Burden',
                                  'Jo Callis',
                                  'Joanne Catherall',
                                  'Philip Adrian Wright',
                                  'Philip Oakey',
                                  'Susan Ann Sulley'],
                      'other-era': ['Ian Craig Marsh',
                                    'Jim Russell',
                                    'Martyn Ware',
                                    'Neil Sutton',
                                    'Russell Dennett'],
                      'type': 'Band',
                      'winner': False},
 'The Judds': {'members': ['Naomi Judd', 'Wynonna Judd'], 'other-era': [], 'type': 'Duo', 'winner': False},
 'The Knack': {'members': ['Berton Averre', 'Bruce Gary', 'Doug Fieger', 'Prescott Niles'],
               'other-era': ['Billy Ward', 'Pat Torpey', 'Phil Jost', "Stephen 'Mac' McNally"],
               'type': 'Band',
               'winner': False},
 'The Pretenders': {'members': ['Chrissie Hynde', 'James Honeyman-Scott', 'Martin Chambers', 'Pete Farndon'],
                    'other-era': ['Adam Seymour',
                                  'Andy Hobson',
                                  'Blair Cunningham',
                                  'Gerry Mcilduff',
                                  'Johnny Marr',
                                  'Malcolm Foster',
                                  'Robbie McIntosh',
                                  'T. M. Stevens'],
                    'type': 'Band',
                    'winner': False},
 'Timbuk3': {'members': ['Barbara K. MacDonald', 'Pat MacDonald'],
             'other-era': ['Courtney Audain', 'Wally Ingram'],
             'type': 'Duo',
             'winner': False},
 'Tone Loc': {'members': ['Tone Loc'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Toni Childs': {'members': ['Toni Childs'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'Tracy Chapman': {'members': ['Tracy Chapman'], 'other-era': [], 'type': 'Solo', 'winner': True},
 'Vanessa Williams': {'members': ['Vanessa Williams'], 'other-era': [], 'type': 'Solo', 'winner': False},
 'a-ha': {'members': ['Magne Furuholmen', 'Morten Harket', 'Paul Waaktaar-Savoy'],
          'other-era': [],
          'type': 'Band',
          'winner': False}}


def load():
    tj = json.load(open(os.path.join(DIR, IN["turns"]), encoding="utf-8"))
    recs = {r["Record ID"]: r for r in json.load(open(os.path.join(DIR, IN["records"]), encoding="utf-8"))}
    text = {x["Record ID"]: x["Normalised Text"] for x in json.load(open(os.path.join(DIR, IN["text"]), encoding="utf-8"))}
    for r in json.load(open(os.path.join(DIR, IN["corpus"]), encoding="utf-8")):
        if r["Record ID"] in recs:
            recs[r["Record ID"]]["Relevance"] = r.get("Relevance") or ""
    sp = os.path.join(DIR, IN["supplement"])
    if not os.path.exists(sp):
        raise SystemExit(f"{IN['supplement']} is missing: put it in {DIR} next to this script (it holds Amy Holland's "
                         "interviews, added by hand). Nothing was written.")
    sup = json.load(open(sp, encoding="utf-8"))
    if sup["columns"] != COLUMNS:
        raise SystemExit(f"{IN['supplement']} has other columns than this script. Nothing was written.")
    return tj, recs, text, sup


def key_name(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().casefold().strip()
    s = re.sub(r"^the\s+", "", s)
    return re.sub(r"[^a-z0-9]", "", s)                  # 'Tone Lōc' and 'Tone Loc' match


def context(t, start, end):
    b = t[max(0, start - CONTEXT):start]
    a = t[end:end + CONTEXT]
    if start - CONTEXT > 0 and " " in b:
        b = "…" + b[b.index(" ") + 1:]
    if end + CONTEXT < len(t) and " " in a:
        a = a[:a.rindex(" ")] + "…"
    tidy = lambda s: re.sub(r"[ \t]*\n\s*", " / ", s).strip()
    return tidy(b), tidy(a)


YEAR = re.compile(r"(?<![\d$£€.,])(?:19[4-9]\d|20[0-2]\d)(?:s\b|(?![\d,]))")
DECADE = re.compile(r"(?<![\w])['’](?:[4-9]0|[0-2]0)s\b")


def years_mentioned(s):
    out = []
    for m in YEAR.finditer(s):
        y = m.group(0)
        if y not in out:
            out.append(y)
    for m in DECADE.finditer(s):
        d = m.group(0)[1:3]
        y = ("19" if d[0] in "456789" else "20") + d + "s"
        if y not in out:
            out.append(y)
    return "; ".join(out)


def voice_of(t):
    st = t["Speaker type"] or ""
    if t["Band voice"] == "yes":
        return "Band"
    if t["Turn type"] == "second-hand":
        return "Second-hand"
    if t["Turn type"] == "speech" and st.startswith("other-era"):
        return "Other-era member"
    if t["Turn type"] == "speech" and st in ("unknown", "ambiguous"):
        return "Unknown"
    return "Public narrative"


ROLE = {"band member": "Member", "band member, unspecified": "Member (unidentified)",
        "band collective": "The band as a group", "other-era member": "Other-era member",
        "other-era member, unspecified": "Other-era member (unidentified)", "interviewer": "Interviewer",
        "other person": "Other person", "unknown": "Unknown", "ambiguous": "Unknown (ambiguous)"}


def role_of(t):
    if t["Turn type"] == "narration":
        return "Writer"
    if t["Turn type"] == "second-hand":
        return "Writer reporting a member"
    if t["Speaker type"] == "journalist":
        return "Writer"
    return ROLE.get(t["Speaker type"], t["Speaker type"])


def speaker_of(t):
    name, cands, st = t["Speaker name"] or "", t["Speaker candidates"] or "", t["Speaker type"] or ""
    if st in ("unknown",) and t["Turn type"] == "speech":
        return "Unknown"
    if st == "ambiguous":
        return f"Unclear ({cands})" if cands else "Unclear"
    if name and cands and "unspecified" in name:
        return f"{name} ({cands})"
    return name or (f"({cands})" if cands else "")


def build_rows(tj, recs, text):
    c = {k: i for i, k in enumerate(tj["columns"])}
    data = tj["data"]
    later = collections.Counter(r[c["Repeat of"]] for r in data if r[c["Repeat of"]])
    rows = collections.defaultdict(list)
    for r in data:
        if r[c["Repeat of"]]:
            continue
        t = {k: r[i] for k, i in c.items()}
        rid = t["Record ID"]
        idate = str(t["Interview date"])
        iso = idate if len(idate) == 10 else f"{idate[:4]}-01-01"
        Y = int(t["Grammy Year"])
        dy = int(iso[:4]) - Y
        before, after = context(text[rid], t["Span start"], t["Span end"])
        rec = recs.get(rid, {})
        bv = t["Band voice"] == "yes"
        row = {"Date": iso,
               "Date from": "stated in the text" if t["Interview date source"] == "stated in the text" else "publication date",
               "Date precision": t["Interview date precision"] or "",
               "Years from nomination": dy, "Period": "before" if dy < 0 else ("nomination year" if dy == 0 else "after"),
               "Voice": voice_of(t), "Speaker": speaker_of(t), "Speaker role": role_of(t),
               "Reported member": t["Reported member"] or "", "Text": t["Turn text"], "Words": t["Words"],
               "Before": before, "After": after, "Years mentioned": years_mentioned(t["Turn text"]),
               "Confidence": t["Confidence"], "Main set": t["Main set"] if bv else "",
               "Attribution rule": t["Attribution rule"] or "", "Attribution evidence": t["Attribution evidence"] or "",
               "Outlet": t["Outlet / Source Name"] or "", "Title": rec.get("Title / Headline") or "",
               "Article relevance": rec.get("Relevance") or "",
               "Publication date": t["Publication Date"] or "", "Format": t["Route"] or "", "URL": rec.get("URL") or "",
               "Turn ID": t["Turn ID"], "Record ID": rid, "Position in article": t["Turn order"],
               "Later copies": later.get(t["Turn ID"], 0), "Act": t["Focal Act"]}
        rows[t["Focal Act"]].append(row)
    for act in rows:
        rows[act].sort(key=lambda x: (x["Date"], x["Publication date"], x["Record ID"], x["Position in article"]))
    return rows


def act_details(rows, grammy_year):
    out = []
    for act in sorted(ACTS, key=lambda a: (grammy_year.get(a, 0), a.casefold())):
        info = ACTS[act]
        rs = rows.get(act, [])
        bv = [r for r in rs if r["Voice"] == "Band"]
        per = collections.Counter(r["Period"] for r in bv)
        dates = sorted(r["Date"] for r in bv)
        out.append({"Act": act, "Solo or group": info["type"], "Nominated": int(grammy_year[act]),
                    "Result": "Winner" if info["winner"] else "Nominee",
                    "Nomination-era members": ", ".join(info["members"]), "Other-era members": ", ".join(info["other-era"]),
                    "Articles": len({r["Record ID"] for r in rs}), "Rows": len(rs), "Band-voice turns": len(bv),
                    "Before": per["before"], "Nomination year": per["nomination year"], "After": per["after"],
                    "Band-voice words": sum(r["Words"] for r in bv),
                    "Second-hand turns": sum(1 for r in rs if r["Voice"] == "Second-hand"),
                    "Public narrative words": sum(r["Words"] for r in rs if r["Voice"] == "Public narrative"),
                    "Other-era turns": sum(1 for r in rs if r["Voice"] == "Other-era member"),
                    "Unknown turns": sum(1 for r in rs if r["Voice"] == "Unknown"),
                    "First band-voice date": dates[0] if dates else "", "Last band-voice date": dates[-1] if dates else "",
                    "Sheet": act if bv else ""})
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
    ("What this is", "Step 2 (articles) of the replication workflow: an oral history for each act, built from the "
     "turns table. Every turn of every article about the act is here, oldest first; an article's turns stay in "
     "reading order. The Voice column separates the act's own words (band voice) from what others say about it "
     "(public narrative), so the two can be read together or apart."),
    ("How to use it", "Each act has its own sheet; the Bands sheet lists them with their counts and a link to each. "
     "Filter Voice to 'Band' for the band's own oral history, or to 'Public narrative' for the other layer. The Date "
     "column is a real date, so 'Sort Oldest to Newest' works; to keep each article's turns together, sort by Date, "
     "then Record ID, then Position in article. The full text of a long cell shows in the formula bar."),
    ("Shading", "Band voice: white. Second-hand: pale yellow. Public narrative: pale grey. Other-era members: pale "
     "blue. Unknown speakers: pale orange. Band voice outside the main set (low confidence): grey text."),
    ("What is left out", "Later copies of republished articles: the earliest copy of a turn stays, dated by its first "
     "appearance, and Later copies counts the rest. Nothing else is left out: other-era members and unknown speakers "
     "are kept and marked, and every article about the act is in, including those that name it only once (Article "
     "relevance; filter them out for a stricter set)."),
    ("Added by hand", "Amy Holland has no band voice in the corpus, so four interviews found on 15 October (NJArts.net, "
     "2017; The Spokesman-Review, 2017; Goldendale Sentinel, 2017; VintageRock.com, 2020) were added from "
     "harrison_v7_oral_histories_supplement.json: 136 rows, 66 of them band voice, split into turns by hand with the "
     "coding rules of the validation. Their Confidence is 'added by hand' and their Attribution rule starts 'added by "
     "hand'; their Record IDs start at 90001. The supplement keeps each article's text, source and checksum."),
    ("Corrections", "Sade is a band here. grammy_nominees.xlsx lists Sade as Solo, but the line-up has four "
     "nomination-era members (Sade Adu, Stuart Matthewman, Andrew Hale, Paul S. Denman) and Harrison et al. treat Sade "
     "as a band."),
    ("Dates", "The date is when the words were said, which is usually the publication date: about 98% of band-voice "
     "turns carry no year of their own in the text. Most accounts were given long after the nomination, so 'Period' "
     "describes when an account was given, not the period it describes. Years mentioned is a first pointer to the "
     "latter."),
    ("Accuracy", "The turns were validated against hand codes for 21 held-out records (15 October): of the text "
     "labelled as band voice, 96% is band voice in the hand codes, with the right member in every case; 83% of the "
     "hand-coded band voice is found. See harrison_v7_validation_results.xlsx."),
    ("For the GenAI steps", "harrison_v7_oral_histories.json holds the same rows and columns for every act in one "
     "file ('columns', 'data'), plus the details of each act ('acts') and these notes ('about'). Unlike the workbook, "
     "it holds every text in full."),
    ("Version", VERSION),
]


def write_xlsx(path, rows, acts, note_for=None):
    try:
        import xlsxwriter
    except ImportError:                                  # not every Colab image has it
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "xlsxwriter"])
        import xlsxwriter
    wb = xlsxwriter.Workbook(path, {"constant_memory": True, "strings_to_urls": False, "strings_to_formulas": False})
    bold = wb.add_format({"bold": True})
    title = wb.add_format({"bold": True, "font_size": 14})
    head = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "text_wrap": True, "valign": "top", "border": 1,
                          "border_color": "#BFBFBF"})
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})
    top = wb.add_format({"valign": "top"})
    datef = wb.add_format({"num_format": "yyyy-mm-dd", "valign": "top"})
    link = wb.add_format({"font_color": "#0563C1", "underline": 1, "valign": "top"})
    fills = {"Second-hand": "#FFF2CC", "Public narrative": "#EDEDED", "Other-era member": "#DDEBF7",
             "Unknown": "#FCE4D6"}
    cf = {k: wb.add_format({"bg_color": v}) for k, v in fills.items()}
    greytext = wb.add_format({"font_color": "#808080"})

    rm = wb.add_worksheet("Read me")
    rm.set_column(0, 0, 26)
    rm.set_column(1, 1, 120)
    rm.write(0, 0, "Oral histories: the Grammy Best New Artist nominees, 1980 to 1990", title)
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

    names, used = {}, set()
    for a in acts:
        if a["Sheet"]:
            names[a["Act"]] = sheet_name(a["Act"], used | {"read me", "bands"})
            used.add(names[a["Act"]].casefold())

    bs = wb.add_worksheet("Bands")
    bcols = ["Act", "Solo or group", "Nominated", "Result", "Nomination-era members", "Other-era members", "Articles",
             "Rows", "Band-voice turns", "Before", "Nomination year", "After", "Band-voice words", "Second-hand turns",
             "Public narrative words", "Other-era turns", "Unknown turns", "First band-voice date",
             "Last band-voice date", "Sheet"]
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
                    bs.write(i, j, (note_for or {}).get(a["Act"], "no band voice"), top)
            elif k in ("First band-voice date", "Last band-voice date") and v:
                bs.write_datetime(i, j, datetime.datetime.strptime(v, "%Y-%m-%d"), datef)
            elif k in ("Nomination-era members", "Other-era members"):
                bs.write(i, j, v, wrap)
            else:
                bs.write(i, j, v, top)
    bs.autofilter(0, 0, len(acts), len(bcols) - 1)

    cw = {"Date": 11, "Date from": 17, "Date precision": 10, "Years from nomination": 11, "Period": 15, "Voice": 15,
          "Speaker": 22, "Speaker role": 18, "Reported member": 15, "Text": 80, "Words": 7, "Before": 40, "After": 40,
          "Years mentioned": 12, "Confidence": 10, "Main set": 8, "Attribution rule": 28, "Attribution evidence": 22,
          "Outlet": 20, "Title": 30, "Article relevance": 14, "Publication date": 11, "Format": 13, "URL": 30, "Turn ID": 10, "Record ID": 8,
          "Position in article": 8, "Later copies": 8, "Act": 16}
    vcol = COLUMNS.index("Voice")
    mcol = COLUMNS.index("Main set")
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
        ws.write(2, 0, f"{a['Articles']:,} articles · {a['Rows']:,} rows · band voice: {a['Band-voice turns']:,} turns "
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
                if k == "Date":
                    ws.write_datetime(i, j, datetime.datetime.strptime(v, "%Y-%m-%d"), datef)
                elif k == "Publication date" and v:
                    ws.write_datetime(i, j, datetime.datetime.strptime(v, "%Y-%m-%d"), datef)
                elif k in ("Text", "Before", "After"):
                    v = v.strip()                       # the JSON keeps the text exactly as in the turns table
                    if len(v) > CELL_CAP:
                        v = v[:CELL_CAP - 50] + " … [cut here: the full text is in the JSON]"
                    ws.write_string(i, j, v, wrap)
                elif k == "URL" and v:
                    if len(v) <= 2000 and re.match(r"https?://", v):
                        ws.write_url(i, j, v, link, string=v)
                    else:
                        ws.write_string(i, j, v, top)
                elif isinstance(v, (int, float)):
                    ws.write_number(i, j, v, top)
                else:
                    ws.write_string(i, j, str(v), top)
        last = H0 + len(rs)
        ws.autofilter(H0, 0, last, len(COLUMNS) - 1)
        rng = f"A{H0 + 2}:{xlsxwriter.utility.xl_col_to_name(len(COLUMNS) - 1)}{last + 1}"
        vref = f"${xlsxwriter.utility.xl_col_to_name(vcol)}{H0 + 2}"
        mref = f"${xlsxwriter.utility.xl_col_to_name(mcol)}{H0 + 2}"
        for k, fmt in cf.items():
            ws.conditional_format(rng, {"type": "formula", "criteria": f'={vref}="{k}"', "format": fmt})
        ws.conditional_format(rng, {"type": "formula", "criteria": f'=AND({vref}="Band",{mref}="no")', "format": greytext})
    wb.close()


def main():
    tj, recs, text, sup = load()
    rows = build_rows(tj, recs, text)
    added = [dict(zip(COLUMNS, r)) for r in sup["data"]]
    for r in added:                                    # rows added by hand (Amy Holland's interviews)
        rows[r["Act"]].append(r)
    for act in {r["Act"] for r in added}:
        rows[act].sort(key=lambda x: (x["Date"], x["Publication date"], x["Record ID"], x["Position in article"]))
    c = {k: i for i, k in enumerate(tj["columns"])}
    grammy_year = {r[c["Focal Act"]]: r[c["Grammy Year"]] for r in tj["data"]}
    missing = sorted(set(grammy_year) ^ set(ACTS))
    if missing:
        raise SystemExit(f"These acts are in the turns table or the act list but not both: {missing}")
    acts = act_details(rows, grammy_year)
    allrows = [row for a in acts for row in rows.get(a["Act"], [])]
    about = {"version": VERSION, "what": README[0][1], "notes": {k: v for k, v in README[1:]},
             "columns explained": EXPLAIN}
    with open(os.path.join(OUT, "harrison_v7_oral_histories.json"), "w", encoding="utf-8") as f:
        json.dump({"about": about, "acts": acts, "columns": COLUMNS,
                   "data": [[r[k] for k in COLUMNS] for r in allrows]}, f, ensure_ascii=False)
    write_xlsx(os.path.join(OUT, "harrison_v7_oral_histories.xlsx"), rows, acts)
    v = collections.Counter(r["Voice"] for r in allrows)
    import hashlib
    fp = hashlib.md5(json.dumps([[r[k] for k in COLUMNS] for r in allrows], ensure_ascii=False).encode()).hexdigest()
    fa = hashlib.md5(json.dumps(acts, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    print(VERSION)
    print(f"Fingerprint of the rows: {fp}; of the acts: {fa} (the same inputs always give the same fingerprints)")
    print(f"{len(acts)} acts, {sum(1 for a in acts if a['Sheet'])} with band voice; {len(allrows):,} rows: "
          + ", ".join(f"{k} {n:,}" for k, n in v.most_common()))
    for f in ("harrison_v7_oral_histories.json", "harrison_v7_oral_histories.xlsx"):
        print(f"  {f}: {os.path.getsize(os.path.join(OUT, f)) / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
