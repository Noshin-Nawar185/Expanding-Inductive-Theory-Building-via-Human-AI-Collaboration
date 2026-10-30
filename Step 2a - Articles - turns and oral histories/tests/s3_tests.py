"""s3_tests.py: checks on the records named in the stage 3 rules (turns_v7_stage3.py).

Each check names a record of the corpus collected for this study and the speaker the rules should give it (68
checks). Run it after turns_v7_stage3.py:
    python s3_tests.py [folder]
The folder (by default HARRISON_DIR, in Colab /content/drive/MyDrive/harrison_v7) must hold
harrison_v7_TURNS_stage3.json and harrison_v7_TURNS_stage1.json. It prints ok or FAIL for each check and the number
that pass. The record IDs are those of the study's corpus, so on a newly collected corpus the checks do not apply.
"""
import json, os, sys
from collections import Counter

DIR = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
with open(os.path.join(DIR, "harrison_v7_TURNS_stage3.json"), encoding="utf-8") as f:
    res = {x["Record ID"]: {"who": x["Interviewee"], "spans": x["Spans"]} for x in json.load(f)}

def main_speaker(rid, types=None):
    c = Counter((x["speaker_type"], x["speaker"]) for x in res[rid]["spans"]
                if x["kind"] == "turn" and x["speaker_type"] != "interviewer" and (types is None or x["speaker_type"] in types))
    return c.most_common(1)[0][0] if c else (None, None)

def label_of(rid, lab):
    return next(((x["speaker_type"], x["speaker"]) for x in res[rid]["spans"] if x["label"] == lab), None)

checks = [
    ("4430 Slim Jim Phantom narrates", main_speaker(4430), ("band member", "Slim Jim Phantom")),
    ("3180 Kathy Valentine, not Mary Lucia", res[3180]["who"]["who"], ["Kathy Valentine"]),
    ("176 John Illsley, not Pete Feenstra", res[176]["who"]["who"], ["John Illsley"]),
    ("9581 Bobby Hornsby is someone else", main_speaker(9581), ("other person", "Bobby Hornsby")),
    ("5358 Kirsten Adamson is someone else", res[5358]["who"]["kind"], "other"),
    ("3539 Geoff Downes writes on his own site", main_speaker(3539), ("band member", "Geoff Downes")),
    ("3542 a blogger's obituary is no member's", main_speaker(3542)[0] in ("unknown", "other person"), True),
    ("12511 the fan site's author", main_speaker(12511), ("other person", "Aurélie")),
    ("6903 a namesake's fundraiser", main_speaker(6903)[0] in ("unknown", "other person"), True),
    ("9048 a fan's essay about Sade", main_speaker(9048)[0] in ("unknown", "other person"), True),
    ("1124 CHRIS (Caller)", label_of(1124, "CHRIS")[0], "other person"),
    ("4326 Lee: is Lee Rocker", label_of(4326, "Lee"), ("band member", "Lee Rocker")),
    ("3509 Carl Palmer answers", main_speaker(3509), ("band member", "Carl Palmer")),
    ("11374 Corinne: inline", label_of(11374, "Corinne"), ("band member", "Corinne Drewery")),
    ("13424 Liam answers, is not the interviewer", label_of(13424, "Liam")[0], "other person"),
    ("7467 Paul Morley's Q&A is no member's", main_speaker(7467, {"band member"}), (None, None)),
    ("113 Knopfler met RC", res[113]["who"]["who"], ["Mark Knopfler"]),
    ("10342 MacDonald spoke from Austin", res[10342]["who"]["who"], ["Pat MacDonald"]),
    ("13972 Billboard caught up with Lōc", res[13972]["who"]["who"], ["Tone Loc"]),
    ("12497 Gary Younge interviews Tracy Chapman", res[12497]["who"]["who"], ["Tracy Chapman"]),
    ("10182 Mick Hucknall", main_speaker(10182), ("band member", "Mick Hucknall")),
    ("1829 James Honeyman-Scott", main_speaker(1829), ("band member", "James Honeyman-Scott")),
    ("3883 Jennifer Holliday", main_speaker(3883), ("band member", "Jennifer Holliday")),
    ("8570 Kimberley Rew", main_speaker(8570), ("band member", "Kimberley Rew")),
    ("3883 the NPR host", label_of(3883, "MICHEL MARTIN") and label_of(3883, "MICHEL MARTIN")[0], "interviewer"),
    ("6168 Annie Lennox and Dave Stewart", sorted(res[6168]["who"]["who"]), ["Annie Lennox", "Dave Stewart"]),
    ("5161 Bruce and Jamie Watson", sorted(res[5161]["who"]["who"]), ["Bruce Watson", "Jamie Watson"]),
    ("2043 JN is Johnny Normal", label_of(2043, "JN")[0], "interviewer"),
    ("4134 Songfacts: is the interviewer", label_of(4134, "Songfacts")[0], "interviewer"),
]
# version 2 (7 Oct 2026): checks added after the check of version 1's workbook
TEXT = {x["Record ID"]: x["Normalised Text"] for x in json.load(open(os.path.join(DIR, "harrison_v7_TURNS_stage1.json"), encoding="utf-8"))
        if x["Record ID"] in {9352, 4400, 8812, 11375, 358, 12842, 1144, 3517, 8494, 1262, 3273, 3047, 11209, 4449, 11184,
                              2043, 13988, 9622,
                              # version 3
                              4201, 4134, 3013, 8171, 7132, 582, 7776, 13264, 8251}}


def span_at(rid, snippet, k=0):
    pos = TEXT[rid].find(snippet)
    assert pos >= 0, (rid, snippet)
    x = next(x for x in res[rid]["spans"] if x["start"] <= pos + k < x["end"])
    return x


def who_at(rid, snippet):
    x = span_at(rid, snippet, len(snippet) // 2)
    return (x["speaker_type"], x["speaker"]) if x["kind"] == "turn" else (x["kind"], None)


def conf_of(rid, lab):
    return next((x["confidence"] for x in res[rid]["spans"] if x["label"] == lab), None)


checks += [
    ("9352 'Pål Waaktaar Savoy:' is Paul Waaktaar-Savoy", label_of(9352, "Pål Waaktaar Savoy"),
     ("band member", "Paul Waaktaar-Savoy")),
    ("4400 'Slim Jim:' is Slim Jim Phantom", label_of(4400, "Slim Jim"), ("band member", "Slim Jim Phantom")),
    ("4400 'RAMzine's Sean Lewis:' is the author", who_at(4400, "The pay off, for me at least"),
     ("interviewer", "Sean Lewis")),
    ("8812 'Paul:' is Paul S. Denman ('Paul S Denman' in the text)", label_of(8812, "Paul"),
     ("band member", "Paul S. Denman")),
    ("8026 'JUDD:' is Wynonna Judd at medium", (label_of(8026, "JUDD"), conf_of(8026, "JUDD")),
     (("band member", "Wynonna Judd"), "medium")),
    ("7288 'Lauper:' at medium", conf_of(7288, "Lauper"), "medium"),
    ("4201 'HAY:' at medium", conf_of(4201, "HAY"), "medium"),
    ("11375 'CD:' is Corinne Drewery", who_at(11375, "CD: I hated that title"), ("band member", "Corinne Drewery")),
    ("358 'Knopfler;' is Mark Knopfler", who_at(358, "Knopfler; Oh yeah, probably"), ("band member", "Mark Knopfler")),
    ("12842 'JLI was' is James Little's (JL)", who_at(12842, "JLI was. I talked to a lot of students"),
     ("other person", "JL")),
    ("13988 'PR:' is Peter Relic", who_at(13988, "PR: Well, you know, I knew Mike D"), ("other person", "Peter Relic")),
    ("1144 'Do you know what The Onion is?' is Cross's", who_at(1144, "Do you know what The Onion is?"),
     ("band member", "Christopher Cross")),
    ("3517 the unlabelled answer about Heat of the Moment", who_at(3517, "My big inspiration when I came out of school"),
     ("band member", "John Wetton")),
    ("8494 'Initially not.' is Julian Lennon's", who_at(8494, "Initially not. It was only when Mum"),
     ("band member", "Julian Lennon")),
    ("1262 caption: 'Thank you, Evan.' is Cross", who_at(1262, "- Thank you, Evan."), ("band member", "Christopher Cross")),
    ("1262 caption: 'I'm Evan Smith.' is the host", who_at(1262, "- I'm Evan Smith."), ("interviewer", "Evan Smith")),
    ("1262 caption: 'Yeah, I think online you can find a mockup' is Cross", who_at(1262, "I think online you can find a mockup"),
     ("band member", "Christopher Cross")),
    ("3273 caption: Kathy Valentine is the guest", res[3273]["who"]["who"], ["Kathy Valentine"]),
    ("3273 caption: 'I was 21 years old when I joined the band'", who_at(3273, "I was 21 years old when I joined the band"),
     ("band member", "Kathy Valentine")),
    ("3047 'Address:' is page furniture", who_at(3047, "Address: 1315 S Congress Ave"), ("non-speech", None)),
    ("11209 a photo caption is not an answer", who_at(11209, "Jody Watley at the entrance to the Soul Train tunnel")[0] !=
     "band member", True),
    ("4449 the closing note is not Lee Rocker's", who_at(4449, "For info about Lee Rocker, his solo projects")[0] !=
     "band member", True),
    ("11184 the journalist's own story is not Jody Watley's", who_at(11184, "We became fast and close friends")[0] !=
     "band member", True),
    ("9622 'AP:' is the author, not David Mansfield", who_at(9622, "The concert footage is spectacular"),
     ("interviewer", "Alan Paul")),
    ("154 the rule text reads 'interviewee named in the title'",
     any("named in the named" in (x["rule"] or "") for d in res.values() for x in d["spans"]), False),
    # version 3 (12 Oct 2026)
    ("v3 4201: Colin Hay's answer with 'wasn't it?' in it is his", who_at(4201, "So I just started going out on the road"),
     ("band member", "Colin Hay")),
    ("v3 4201: the paragraph after it is his too", who_at(4201, "So I was on the road. I was touring"),
     ("band member", "Colin Hay")),
    ("v3 4201: a question he reports in quotation marks", who_at(4201, "So he played on it. And it was just somewhat"),
     ("band member", "Colin Hay")),
    ("v3 4201: the real question is still the interviewer's", who_at(4201, "You've known Ringo personally")[0],
     "interviewer"),
    ("v3 4134: the answer under 'I'd like to know how you decided...' is Colin Hay's",
     who_at(4134, "We didn't have much money to make the videos"), ("band member", "Colin Hay")),
    ("v3 3013: the answer under 'Tell me a little about that song.'",
     who_at(3013, "I have a very strong connection to the first record"), ("band member", "Charlotte Caffey")),
    ("v3 8171: 'NJ:' run into the interviewer's line", who_at(8171, "Not for my type. Mine is the least contagious"),
     ("band member", "Naomi Judd")),
    ("v3 2043: '¯¿¼JN-' is the interviewer's label", who_at(2043, "And can you speak Italian now")[0], "interviewer"),
    ("v3 2043: '¯¿¼MP-' is Marco Pirroni's", who_at(2043, "Of course I like today's music"),
     ("band member", "Marco Pirroni")),
    ("v3 7132: 'CL–' is Cyndi Lauper's", who_at(7132, "It's a story about an outsider"), ("band member", "Cyndi Lauper")),
    ("v3 582: 'PKM.' is the interviewer's label", who_at(582, "There's a big difference between performing and writing")[0],
     "interviewer"),
    ("v3 7776: the writer's own comment stays the interviewer's",
     who_at(7776, "Now, I confess that I'm a Prince obsessive")[0], "interviewer"),
    ("v3 13264: a question about 'your dad' stays the interviewer's",
     who_at(13264, "anyone who has more than a passing knowledge of Mats Gustafsson")[0], "interviewer"),
    ("v3 8251: a host's long statement does not give the next paragraph to the guest",
     who_at(8251, "And as a musician — like for me, I started off sampling")[0], "interviewer"),
]

bad = 0
for name, got, want in checks:
    ok = got == want
    bad += not ok
    print(("ok  " if ok else "FAIL"), name, "" if ok else f"| got {got!r} want {want!r}")
print(len(checks) - bad, "of", len(checks), "pass")
sys.exit(1 if bad else 0)
