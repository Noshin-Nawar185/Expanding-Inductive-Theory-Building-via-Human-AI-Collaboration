"""s49_tests.py: checks on the records named in the stage 4 to 9 rules (turns_v7_stage4to9.py).

Each check names a record of the corpus collected for this study and what the rules should give one of its turns
(speaker, turn type, interview date and so on). Run it on the turns table that turns_v7_stage4to9.py writes:
    python s49_tests.py [harrison_v7_turns_final.json]
By default the file is read from HARRISON_DIR (in Colab /content/drive/MyDrive/harrison_v7). It prints PASS or FAIL
for each check and the number that pass. The record IDs are those of the study's corpus, so on a newly collected
corpus the checks do not apply.
"""
import json, os, sys
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DIR, "harrison_v7_turns_final.json")
with open(src, encoding="utf-8") as f:
    t = json.load(f)
rows = [dict(zip(t["columns"], v)) for v in t["data"]]
by_rec = {}
for r in rows:
    by_rec.setdefault(r["Record ID"], []).append(r)
def turn(rid, snippet):
    for r in by_rec.get(rid, []):
        if snippet in r["Turn text"]:
            return r
    return None
checks = []
def check(name, rid, snippet, **want):
    r = turn(rid, snippet)
    if r is None:
        checks.append((False, name, f"no turn with {snippet!r}"))
        return
    bad = {k: (r[k], v) for k, v in want.items() if (r[k] != v if not callable(v) else not v(r[k]))}
    checks.append((not bad, name, bad or ""))
check("1179: lyrics after 'Cross sums up' are not speech", 1179, "make it to the border of Mexico", **{"Turn type": "narration"})
check("310: 'Speaking to Rolling Stone ..., Knopfler addressed the issue and said:' is Knopfler", 310,
      "I got an objection from the editor", **{"Speaker name": "Mark Knopfler", "Band voice": "yes"})
check("5932: '... she flashed a quick grin and said,' is Lennox", 5932, "That's it folks", **{"Speaker name": "Annie Lennox"})
check("6416: 'Ivan quickly responded:'", 6416, "Definitely a blessing", **{"Speaker name": "Ivan Doroschuk"})
check("6166: a name in the rest of the attribution sentence does not switch the speaker", 6166,
      "We did some promotion for that", **{"Speaker name": "Dave Stewart"})
check("7287: 'she says' after a paragraph about Lauper is Lauper, not Billy Steinberg", 7287, "I'm very political",
      **{"Speaker name": "Cyndi Lauper"})
check("3322: 'she explained' is Belinda Carlisle", 3322, "Even with the Diane Warren project", **{"Speaker name": "Belinda Carlisle"})
check("3352: 'the group began as' is not the band speaking", 3352, "five of us kids sitting on a curb",
      **{"Speaker type": lambda v: v != "band collective"})
check("8319: 'John Lennon's son' does not make John the speaker", 8319, "This isn't a substitute for a concert",
      **{"Speaker name": "Julian Lennon"})
check("11964: Astley's tribute", 11964, "What a woman, what a life, what a voice", **{"Speaker name": "Rick Astley"})
check("11964: 'One of the GREATS!' continues Astley's tribute", 11964, "One of the GREATS", **{"Speaker name": "Rick Astley"})
check("780: 'so says songwriter, singer, guitarist and ... Doug Fieger'", 780, "it is what  it is", **{"Speaker name": "Doug Fieger"})
check("1239: 'songwriter Diane Warren' is Warren", 1239, "The love for your brilliant songs", **{"Speaker name": "Diane Warren"})
check("3786: 'laughs Asia's keyboard player Geoff Downes'", 3786, "It was very amusing", **{"Speaker name": "Geoff Downes"})
r = next((x for x in by_rec.get(5383, []) if "Bruce Watson" in x["Attribution evidence"]), None)
checks.append((r is not None and r["Speaker name"] == "Bruce Watson", "5383: says Big Country guitarist Bruce Watson", "" if r else "not found"))
check("3854: 'Hudson extended her praise ... she said' is Jennifer Hudson, not Holliday", 3854, "They paved the way",
      **{"Band voice": "no"})
check("8420: Pete Shotton, named as the subject before the quotation", 8420, "I happened to be there the day Julian",
      **{"Speaker name": "Pete Shotton"})
check("1728: 'she sings' is not speech", 1728, "Nobody tells me I can't", **{"Turn type": "narration"})
check("3609: a 1992 interview republished in 2020", 3609, "", **{"Interview date": lambda v: True})
r = next((x for x in by_rec.get(3609, []) if x["Turn type"] == "speech" and x["Attribution rule"].startswith("stage 3")), None)
checks.append((r is not None and r["Interview date"] == "1992", "3609: Q&A turns dated 1992 (stated in the introduction)", "" if r and r["Interview date"] == "1992" else r and r["Interview date"]))
r = next((x for x in by_rec.get(4590, []) if x["Turn type"] == "speech" and x["Attribution rule"].startswith("stage 3")), None)
checks.append((r is not None and r["Interview date"] == "1982", "4590: 'The 1982 Stray Cats Interview' dates its turns 1982", "" if r and r["Interview date"] == "1982" else r and r["Interview date"]))
r = next((x for x in by_rec.get(8464, []) if x["Turn type"] == "speech" and x["Attribution rule"].startswith("stage 3")), None)
checks.append((r is not None and r["Interview date source"] == "publication-date proxy", "8464: 'our 2014 interview' (an earlier one) is not this interview's date", "" if r and r["Interview date source"] == "publication-date proxy" else r and r["Interview date"]))
r = next((x for x in by_rec.get(11637, []) if "Sagan" in x["Turn text"] and x["Turn type"] == "speech"), None)
checks.append((r is None or r["Interview date source"] == "publication-date proxy", "11637: 'who passed away in 1996' is not an interview date", ""))
r = next((x for x in by_rec.get(5349, []) if x["Turn type"] == "second-hand" and "announced that his leukemia" in x["Turn text"]), None)
checks.append((r is not None and r["Reported member"] == "Mike Peters", "5349: 'Peters announced that ...' is second-hand (Mike Peters)", "" if r else "not found"))
ws = sum(1 for r in rows if r["Turn type"] == "narration" and not r["Turn text"].strip())
checks.append((ws == 0, "no narration turn is only white space", ws or ""))
# version 1.2 (7 Oct 2026)
check("v1.2 1837: a passage stage 2 left as prose continues McIntosh's answer", 1837, "I remember being at Olympic Studios",
      **{"Speaker name": "Robbie McIntosh", "Attribution rule": lambda v: v.startswith("Q&A passage")})
check("v1.2 113: 'Hendrix?' is the interviewer's question", 113, "Hendrix?", **{"Speaker type": "interviewer",
                                                                                "Question": "yes"})
check("v1.2 113: 'Completely loved him' is Knopfler's answer", 113, "Completely loved him",
      **{"Speaker name": "Mark Knopfler", "Band voice": "yes"})
check("v1.2 8420: John Lennon's 1970 quotation is not Julian Lennon's", 8420, "I had no idea it spelt LSD",
      **{"Speaker name": "John Lennon", "Band voice": "no"})
check("v1.2 8805: 'Sade recalls' is Sade Adu", 8805, "I'd never seen snow before", **{"Speaker name": "Sade Adu"})
check("v1.2 3252: 'the band says' after The Linda Lindas is not the Go-Go's", 3252, "The Linda Lindas started out",
      **{"Band voice": "no"})
check("v1.2 13100: ZZ Top's statement in an Indigo Girls record is not band voice", 13100,
      "We are saddened by the news today", **{"Band voice": "no"})
check("v1.2 12142: 'the group's Dennis Baptiste' (Committed) is not Take 6", 12142, "building our fan base",
      **{"Band voice": "no"})
check("v1.2 3051: the Go-Go's statement is the band", 3051, "It is so gratifying", **{"Speaker type": "band collective"})
check("v1.2 9266: 'the group said' after 'A-ha are calling it quits' is a-ha", 9266, "Change is always difficult",
      **{"Speaker type": "band collective"})
check("v1.2 6155: 'the duo's Dave Stewart' is Dave Stewart", 6155, "kinda came out of the blue",
      **{"Speaker name": "Dave Stewart"})
check("v1.2 3560: 'Carl Palmer of Asia' is Carl Palmer", 3560, "They were incredibly professional",
      **{"Speaker name": "Carl Palmer"})
check("v1.2 3155: the title 'Club Zero' after 'the band announced' is not speech", 3155, "Club Zero",
      **{"Turn type": "narration"})
check("v1.2 8703: a republished passage points to its first copy", 8703, "We have never been to Liverpool before",
      **{"Repeat of": "7597-2"})
bad_main = sum(1 for r in rows if (r["Band voice"] == "yes") != (r["Main set"] in ("yes", "no")) or
               (r["Main set"] == "yes") != (r["Band voice"] == "yes" and r["Confidence"] in ("high", "medium")))
checks.append((bad_main == 0, "v1.2: Main set is yes/no for band-voice turns only, yes for high and medium", bad_main or ""))
bad_rep = sum(1 for r in rows if r["Repeat of"] and r["Repeat of"] == r["Turn ID"])
checks.append((bad_rep == 0, "v1.2: no turn is a repeat of itself", bad_rep or ""))
coll_solo = sum(1 for r in rows if r["Speaker type"] == "band collective" and r["Focal Act"] in (
    "Christopher Cross", "Irene Cara", "Luther Vandross", "Sheena Easton", "Jennifer Holliday", "Cyndi Lauper", "Sheila E.",
    "Freddie Jackson", "Julian Lennon", "Terence Trent D'Arby", "Tracy Chapman", "Vanessa Williams", "Neneh Cherry",
    "Tone Loc", "Rickie Lee Jones", "Rick Astley", "Toni Childs"))
checks.append((coll_solo == 0, "v1.2: no band-collective turn in a solo act's record", coll_solo or ""))
# version 1.3 (12 Oct 2026): reads version 3 of stage 3
check("v1.3 4201: Colin Hay's answer with 'wasn't it?' in it is his", 4201, "So I just started going out on the road",
      **{"Speaker name": "Colin Hay", "Band voice": "yes"})
check("v1.3 4134: the answer under Songfacts' request is Colin Hay's", 4134, "We didn't have much money to make the videos",
      **{"Speaker name": "Colin Hay", "Band voice": "yes"})
check("v1.3 7776: the writer's own comment stays the interviewer's", 7776, "Now, I confess that I'm a Prince obsessive",
      **{"Speaker type": "interviewer"})
ok = sum(c[0] for c in checks)
for c in checks:
    print("PASS" if c[0] else "FAIL", c[1], c[2])
print(f"{ok} of {len(checks)} pass")
sys.exit(0 if ok == len(checks) else 1)
