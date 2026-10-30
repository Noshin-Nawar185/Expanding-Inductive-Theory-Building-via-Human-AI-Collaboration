"""gemini_speakers.py: who speaks when in the videos, from Gemini.
Step 1 (videos). Version 1.1, 18 October 2026.

Who speaks when, from Gemini, for the video interviews (the step after the WhisperX run).
Gemini is given each video's YouTube link, so it sees the picture as well as hearing the sound. It returns timed
turns, each with a speaker label, a role and a short description of who is on screen, and it names a person only
on evidence. WhisperX keeps the words and their timings; a later merge step joins the two.

Which videos: by default, every video of an act that is not a solo artist (band, duo, group, collective) that has
a WhisperX transcript in video_transcripts, except the 14 videos dropped after the WhisperX run (no usable
speech, or not in English; DROPPED below, by V ID).
With --solo, the same for solo artists instead.
A solo artist's video gets its own prompt: the role "artist" for the act, and the label "artist, unconfirmed (A)"
for someone who seems to be the artist but is not named on evidence. The prompt for bands, duos and groups is,
word for word, that of version 1.0, which labelled their videos; version 1.1 adds --solo and the solo prompt.
Videos longer than 20 minutes are sent in parts of about 15 minutes, so that no reply is cut short; each part is
told the labels already given, so a person keeps the same label through the video.

For each video it writes, in the gemini_speakers folder next to this script:
  <ID>.json          every reply as returned, the model, the date and the token counts, and the turns
  <ID>.txt           a version to read
  gemini_log.csv     one row per video
  usage_ledger.csv   minutes of video sent each day, for the free tier's daily allowance

In Colab (any runtime; no GPU needed):
  Cell 1:  !pip install -q -U google-genai          (then restart the session)
  Cell 2:  from google.colab import drive; drive.mount('/content/drive')
           import os; print(len(os.listdir("/content/drive/MyDrive/harrison_v7/Video for Harrison/video_transcripts")))
  Cell 3:  %run "/content/drive/MyDrive/harrison_v7/Video for Harrison/gemini_speakers.py" --daily-hours 0 --workers 5
           (add --solo for the solo artists' videos)
Cell 2's listing makes Drive show the whole folder before the script looks at it.
The key is read from Colab Secrets (GEMINI_API_KEY, with Notebook access on) and never printed.
Safe to run again: a video already saved is skipped, and a video stopped halfway carries on from its next part.
On the free tier it stops by itself before the day's 8 hours of YouTube video are used; run the same cell the
next day to carry on.
Options:
  --list                show which videos would be sent, and how many hours, without sending anything
  --solo                solo artists' videos instead of those of bands, duos and groups
  --ids V1158,V1319     send only these videos (any act)
  --out gemini_retest   save into another folder next to this script
  --daily-hours 0       no daily limit (a key with billing turned on); the default is 8
  --workers 2           videos sent at the same time
  --max 5               send at most this many videos in this run
  --model gemini-3.8-flash
  --redo                send again videos already saved
"""
import csv, datetime, json, math, os, re, sys, threading, time
from concurrent.futures import ThreadPoolExecutor

VERSION = "gemini_speakers 1.1 (18 Oct 2026)"


def script_folder(path=None):
    """The Drive folder this script is in (when run from Drive with %run); otherwise MyDrive/harrison_v7."""
    if path is None:
        path = globals().get("__file__")
    if path and os.path.abspath(path).startswith("/content/drive/"):
        return os.path.dirname(os.path.abspath(path))
    return "/content/drive/MyDrive/harrison_v7"


# ------------------------------------------------------------------ settings
BASE = script_folder()
SELECTION = f"{BASE}/harrison_v7_VIDEO_SELECTION.xlsx"
TRANSCRIPTS = f"{BASE}/video_transcripts"   # the WhisperX transcripts; only videos with one are sent
OUT_DIR = f"{BASE}/gemini_speakers"
MODEL = "gemini-3.8-flash"
WORKERS = 2                 # videos sent at the same time
DAILY_HOURS = 8.0           # the free tier's YouTube allowance per day (Pacific time); 0 = no limit
ONE_PART_MAX_MIN = 20       # videos up to this long are sent whole
PART_MIN = 15               # longer videos are sent in parts of about this length
MIN_PART_MIN = 4            # a part cut short is split in two, but not below this length
MAX_OUTPUT_TOKENS = 65536   # Gemini 3.8 Flash's limit
RETRIES = 3                 # tries for a busy service, a time-out or a dropped connection
ONLY_IDS = []
REDO = False
LIST_ONLY = False
MAX_VIDEOS = 0
SOLO = False                # True: solo artists' videos instead of those of bands, duos and groups

# Each act's type (Band, Duo, Solo, ...) and nomination-era members, from the act frame.
ACT_TYPES = {'Dire Straits': 'Band',
 'Rickie Lee Jones': 'Solo',
 'The Knack': 'Band',
 'Amy Holland': 'Solo',
 'Christopher Cross': 'Solo',
 'Irene Cara': 'Solo',
 'Robbie Dupree': 'Solo',
 'The Pretenders': 'Band',
 'Adam and the Ants': 'Band',
 'James Ingram': 'Solo',
 'Luther Vandross': 'Solo',
 'Sheena Easton': 'Solo',
 "The Go-Go's": 'Band',
 'Asia': 'Band',
 'Jennifer Holliday': 'Solo',
 'Men at Work': 'Band',
 'Stray Cats': 'Band',
 'The Human League': 'Band',
 'Big Country': 'Band',
 'Culture Club': 'Band',
 'Eurythmics': 'Duo',
 'Men Without Hats': 'Band',
 'Musical Youth': 'Band',
 'Corey Hart': 'Solo',
 'Cyndi Lauper': 'Solo',
 'Frankie Goes to Hollywood': 'Band',
 'Sheila E.': 'Solo',
 'The Judds': 'Duo',
 'a-ha': 'Band',
 'Freddie Jackson': 'Solo',
 'Julian Lennon': 'Solo',
 'Katrina and the Waves': 'Band',
 'Sade': 'Band',
 'Bruce Hornsby and the Range': 'Band',
 'Glass Tiger': 'Band',
 'Nu Shooz': 'Duo',
 'Simply Red': 'Band',
 'Timbuk3': 'Duo',
 'Breakfast Club': 'Band',
 'Cutting Crew': 'Band',
 'Jody Watley': 'Solo',
 'Swing Out Sister': 'Band/Duo',
 "Terence Trent D'Arby": 'Solo',
 'Rick Astley': 'Solo',
 'Take 6': 'Group',
 'Toni Childs': 'Solo',
 'Tracy Chapman': 'Solo',
 'Vanessa Williams': 'Solo',
 'Indigo Girls': 'Duo',
 'Neneh Cherry': 'Solo',
 'Soul II Soul': 'Collective',
 'Tone Loc': 'Solo'}
ACT_MEMBERS = {'Dire Straits': ['David Knopfler', 'John Illsley', 'Mark Knopfler', 'Pick Withers'],
 'Rickie Lee Jones': [],
 'The Knack': ['Berton Averre', 'Bruce Gary', 'Doug Fieger', 'Prescott Niles'],
 'Amy Holland': [],
 'Christopher Cross': [],
 'Irene Cara': [],
 'Robbie Dupree': [],
 'The Pretenders': ['Chrissie Hynde', 'James Honeyman-Scott', 'Martin Chambers', 'Pete Farndon'],
 'Adam and the Ants': ['Adam Ant', 'Chris Hughes', 'Gary Tibbs', 'Marco Pirroni', 'Terry Lee Miall'],
 'James Ingram': [],
 'Luther Vandross': [],
 'Sheena Easton': [],
 "The Go-Go's": ['Belinda Carlisle', 'Charlotte Caffey', 'Gina Schock', 'Jane Wiedlin', 'Kathy Valentine'],
 'Asia': ['Carl Palmer', 'Geoff Downes', 'John Wetton', 'Steve Howe'],
 'Jennifer Holliday': [],
 'Men at Work': ['Colin Hay', 'Greg Ham', 'Jerry Speiser', 'John Rees', 'Ron Strykert'],
 'Stray Cats': ['Brian Setzer', 'Lee Rocker', 'Slim Jim Phantom'],
 'The Human League': ['Ian Burden',
                      'Jo Callis',
                      'Joanne Catherall',
                      'Philip Adrian Wright',
                      'Philip Oakey',
                      'Susan Ann Sulley'],
 'Big Country': ['Bruce Watson', 'Mark Brzezicki', 'Stuart Adamson', 'Tony Butler'],
 'Culture Club': ['Boy George', 'Jon Moss', 'Mikey Craig', 'Roy Hay'],
 'Eurythmics': ['Annie Lennox', 'Dave Stewart'],
 'Men Without Hats': ['Allan McCarthy', 'Colin Doroschuk', 'Ivan Doroschuk', 'Stefan Doroschuk'],
 'Musical Youth': ['Dennis Seaton', 'Frederick Waite', 'Kelvin Grant', 'Michael Grant', 'Patrick Waite'],
 'Corey Hart': [],
 'Cyndi Lauper': [],
 'Frankie Goes to Hollywood': ['Brian Nash', 'Holly Johnson', "Mark O'Toole", 'Paul Rutherford', 'Peter Gill'],
 'Sheila E.': [],
 'The Judds': ['Naomi Judd', 'Wynonna Judd'],
 'a-ha': ['Magne Furuholmen', 'Morten Harket', 'Paul Waaktaar-Savoy'],
 'Freddie Jackson': [],
 'Julian Lennon': [],
 'Katrina and the Waves': ['Alex Cooper', 'Katrina Leskanich', 'Kimberley Rew', 'Vince de la Cruz'],
 'Sade': ['Andrew Hale', 'Paul S. Denman', 'Sade Adu', 'Stuart Matthewman'],
 'Bruce Hornsby and the Range': ['Bruce Hornsby',
                                 'David Mansfield',
                                 'George Marinelli',
                                 'Joe Puerta',
                                 'John Molo',
                                 'Peter Harris'],
 'Glass Tiger': ['Al Connelly', 'Alan Frew', 'Michael Hanson', 'Sam Reid', 'Wayne Parker'],
 'Nu Shooz': ['John Smith', 'Valerie Day'],
 'Simply Red': ['Chris Joyce',
                'Fritz McIntyre',
                'Ian Kirkham',
                'Mick Hucknall',
                'Sylvan Richardson',
                'Tim Kellett',
                'Tony Bowers'],
 'Timbuk3': ['Barbara K. MacDonald', 'Pat MacDonald'],
 'Breakfast Club': ['Dan Gilroy', 'Ed Gilroy', 'Gary Burke', 'Paul Kauk', 'Randy Jackson', 'Stephen Bray'],
 'Cutting Crew': ['Colin Farley', 'Kevin MacMichael', 'Martin Beedle', 'Nick Van Eede', 'Tony Moore'],
 'Jody Watley': [],
 'Swing Out Sister': ['Andy Connell', 'Corinne Drewery', 'Martin Jackson'],
 "Terence Trent D'Arby": [],
 'Rick Astley': [],
 'Take 6': ['Alvin Chea', 'Cedric Dent', 'Claude McKnight', 'David Thomas', 'Mark Kibble', 'Mervyn Warren'],
 'Toni Childs': [],
 'Tracy Chapman': [],
 'Vanessa Williams': [],
 'Indigo Girls': ['Amy Ray', 'Emily Saliers'],
 'Neneh Cherry': [],
 'Soul II Soul': ['Aitch B', 'Caron Wheeler', 'Daddae Harvey', 'Jazzie B', 'Nellee Hooper', 'Simon Law'],
 'Tone Loc': []}
# Videos dropped in the WhisperX audit of 18 October 2026 (no usable speech, or not in English).
DROPPED = ['V314',
 'V1225',
 'V1970',
 'V1980',
 'V1986',
 'V1987',
 'V2226',
 'V2232',
 'V2240',
 'V2695',
 'V2822',
 'V3148',
 'V3152',
 'V3485']

LOG_COLS = ["ID", "Act", "Minutes", "Parts", "Status", "Seconds", "Turns", "Labels", "Named labels",
            "Turns not in English", "Prompt tokens", "Output tokens", "Thinking tokens", "Model version", "Note",
            "Finished"]
LEDGER_COLS = ["Date (Pacific)", "ID", "Part", "Minutes", "Sent at (UTC)"]

PROMPT = """This is a video from YouTube, titled "{title}". It features {act}, nominated for the Grammy for Best New Artist in {year}. {members}
{part_note}
Using both the sound and the picture, write down the conversation as speaker turns, in order. Start a new turn every time the person speaking changes, however short the turn: a one-word reply such as "Yes" is a turn of its own. Do not merge two people's words into one turn.

Give each different person a label, and keep that label for them through the whole video:
- Use a person's name only when there is evidence for it: they are introduced, addressed by name or name themselves in the sound; an on-screen caption names them; or they are visibly the same person as someone already named on such evidence. Do not name anyone from their face or voice alone, and do not guess. If the only evidence is the title or the information above, you may give the name, but say so.
- A band member you cannot name is "band member, unspecified (A)", the next different one "band member, unspecified (B)", and so on. Use the same letter every time that person speaks again.
- Anyone else you cannot name is "unknown (A)", "unknown (B)", and so on, in the same way.
Roles: "interviewer" for whoever asks the questions, "band member", "narrator" for an off-screen announcer or narrator, "lyrics" for singing or song clips, and "other" for anyone else.
If a turn is not in English, give its language and write it down as spoken, without translating it.

Return only JSON, in this form:
{{"speakers": [{{"label": "the label used in the turns", "name": "the person's name, or null", "name_source": "said in the video | on-screen caption | same person as one named earlier | title or information given | null", "role": "band member | interviewer | narrator | other", "description": "what they look like or sound like, briefly", "evidence": "how you know, with the time"}}],
 "turns": [{{"start": "MM:SS", "end": "MM:SS", "speaker": "a label from the speakers list", "role": "band member | interviewer | narrator | lyrics | other", "language": "en, or the language's code", "on_screen": "who is visibly speaking, briefly", "text": "what is said, as close to word for word as you can"}}]}}"""

SOLO_PROMPT = """This is a video from YouTube, titled "{title}". It features {act}, a solo artist nominated for the Grammy for Best New Artist in {year}.
{part_note}
Using both the sound and the picture, write down the conversation as speaker turns, in order. Start a new turn every time the person speaking changes, however short the turn: a one-word reply such as "Yes" is a turn of its own. Do not merge two people's words into one turn.

Give each different person a label, and keep that label for them through the whole video:
- Use a person's name only when there is evidence for it: they are introduced, addressed by name or name themselves in the sound; an on-screen caption names them; or they are visibly the same person as someone already named on such evidence. Do not name anyone from their face or voice alone, and do not guess. If the only evidence is the title or the information above, you may give the name, but say so.
- Someone who seems to be {act} but whom you cannot name on such evidence is "artist, unconfirmed (A)", the next different one "artist, unconfirmed (B)", and so on. Use the same letter every time that person speaks again.
- Anyone else you cannot name is "unknown (A)", "unknown (B)", and so on, in the same way.
Roles: "artist" for {act}, "interviewer" for whoever asks the questions, "narrator" for an off-screen announcer or narrator, "lyrics" for singing or song clips, and "other" for anyone else, including other musicians.
If a turn is not in English, give its language and write it down as spoken, without translating it.

Return only JSON, in this form:
{{"speakers": [{{"label": "the label used in the turns", "name": "the person's name, or null", "name_source": "said in the video | on-screen caption | same person as one named earlier | title or information given | null", "role": "artist | interviewer | narrator | other", "description": "what they look like or sound like, briefly", "evidence": "how you know, with the time"}}],
 "turns": [{{"start": "MM:SS", "end": "MM:SS", "speaker": "a label from the speakers list", "role": "artist | interviewer | narrator | lyrics | other", "language": "en, or the language's code", "on_screen": "who is visibly speaking, briefly", "text": "what is said, as close to word for word as you can"}}]}}"""

PART_NOTE = ("\nThis request covers only the part of the video from {start} to {end} (part {k} of {n}). Give all times as "
             "positions in the whole video, not in the part.{known}\n")


class Stop(Exception):
    pass


def stop(msg):
    print("\n" + "=" * 78 + "\nSTOPPED: " + msg + "\n" + "=" * 78)
    raise Stop(msg)


class DailyLimit(Exception):
    pass


# ------------------------------------------------------------------ helpers (no network)
def acts_of(row):
    return [a.strip() for a in str(row["Act"]).split(";") if a.strip()]


def is_group(row):
    """True for an act that is not a solo artist (band, duo, group, collective)."""
    return any(ACT_TYPES.get(a, "?") != "Solo" for a in acts_of(row))


def members_of(row):
    seen = []
    for a in acts_of(row):
        for m in ACT_MEMBERS.get(a, []):
            if m not in seen:
                seen.append(m)
    return seen


def mmss(s):
    s = int(round(s or 0))
    return f"{s // 3600}:{s % 3600 // 60:02d}:{s % 60:02d}" if s >= 3600 else f"{s // 60:02d}:{s % 60:02d}"


def seconds(stamp):
    """'MM:SS', 'H:MM:SS' or a number -> seconds (None if unreadable)."""
    if stamp is None:
        return None
    if isinstance(stamp, (int, float)):
        return float(stamp)
    parts = str(stamp).strip().split(":")
    try:
        nums = [float(p) for p in parts]
    except ValueError:
        return None
    total = 0.0
    for n in nums:
        total = total * 60 + n
    return total


def parse_json(text):
    """The JSON object in a reply, with or without a code fence; None when there is none."""
    if not text:
        return None
    t = text.strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", t, re.S)
    if m:
        t = m.group(1).strip()
    starts = [i for i in (t.find("{"), t.find("[")) if i >= 0]
    if not starts:
        return None
    t = t[min(starts):]
    try:
        return json.loads(t)
    except ValueError:
        pass
    tried = 0
    for end in range(len(t), 0, -1):          # tolerate a little trailing text after the JSON
        if t[end - 1] in "}]":
            tried += 1
            if tried > 200:
                break
            try:
                return json.loads(t[:end])
            except ValueError:
                continue
    return None


def plan_parts(total_s):
    """Whole video up to ONE_PART_MAX_MIN; otherwise equal parts of about PART_MIN minutes."""
    total_s = max(1.0, float(total_s))
    if total_s <= ONE_PART_MAX_MIN * 60:
        return [{"start": 0.0, "end": total_s, "clip": False}]
    n = math.ceil(total_s / (PART_MIN * 60))
    size = total_s / n
    return [{"start": round(i * size, 1), "end": round(total_s if i == n - 1 else (i + 1) * size, 1), "clip": True}
            for i in range(n)]


def split_part(part):
    """A part cut short, in two halves (None when it is already short)."""
    length = part["end"] - part["start"]
    if length < 2 * MIN_PART_MIN * 60:
        return None
    mid = round(part["start"] + length / 2, 1)
    return [{"start": part["start"], "end": mid, "clip": True}, {"start": mid, "end": part["end"], "clip": True}]


def known_note(speakers):
    if not speakers:
        return ""
    items = []
    for s in speakers:
        desc = s.get("description") or s.get("role") or ""
        items.append(f"{s['label']}" + (f" = {desc}" if desc else ""))
    return (" People already labelled earlier in this video (keep these labels for them; give new labels to anyone "
            "new): " + "; ".join(items) + ".")


def build_prompt(row, part=None, k=1, n=1, speakers=None):
    note = ""
    if n > 1 and part:
        note = PART_NOTE.format(start=mmss(part["start"]), end=mmss(part["end"]), k=k, n=n, known=known_note(speakers))
    if not is_group(row):           # a solo artist (version 1.1)
        return SOLO_PROMPT.format(title=row["Title"], act=" and ".join(acts_of(row)), year=int(row["Grammy year"]),
                                  part_note=note)
    members = members_of(row)
    m = f"Its members at the time were {', '.join(members)}." if members else ""
    return PROMPT.format(title=row["Title"], act=" and ".join(acts_of(row)), year=int(row["Grammy year"]),
                         members=m, part_note=note)


def clean(x):
    return "" if x is None else str(x).strip()


def turns_of(parsed):
    turns = parsed.get("turns") if isinstance(parsed, dict) else parsed if isinstance(parsed, list) else []
    out = []
    for t in turns or []:
        if not isinstance(t, dict):
            continue
        out.append({"start": seconds(t.get("start")), "end": seconds(t.get("end")),
                    "speaker": clean(t.get("speaker")) or "unknown", "role": clean(t.get("role")),
                    "language": clean(t.get("language")) or "en", "on_screen": clean(t.get("on_screen")),
                    "text": clean(t.get("text"))})
    return out


def speakers_of(parsed):
    out = []
    for s in (parsed.get("speakers") if isinstance(parsed, dict) else None) or []:
        if isinstance(s, dict) and clean(s.get("label") or s.get("name")):
            name = clean(s.get("name"))
            out.append({"label": clean(s.get("label") or s.get("name")),
                        "name": None if name.lower() in ("", "null", "none") else name,
                        "name_source": clean(s.get("name_source")) or None, "role": clean(s.get("role")),
                        "description": clean(s.get("description")), "evidence": clean(s.get("evidence"))})
    return out


def fix_times(turns, part):
    """For a part, times given from the start of the part are moved to the whole video. Returns the shift used."""
    starts = [t["start"] for t in turns if t["start"] is not None]
    if not part["clip"] or part["start"] < 60 or not starts:
        return 0.0
    length = part["end"] - part["start"]
    if max(starts) <= length + 60 and min(starts) < part["start"] - 60:
        for t in turns:
            for key in ("start", "end"):
                if t[key] is not None:
                    t[key] = round(t[key] + part["start"], 1)
        return part["start"]
    return 0.0


def merge_speakers(known, new):
    have = {s["label"] for s in known}
    for s in new:
        if s["label"] not in have:
            known.append(s)
            have.add(s["label"])
        else:
            old = next(k for k in known if k["label"] == s["label"])
            if not old.get("name") and s.get("name"):
                old.update({"name": s["name"], "name_source": s["name_source"], "evidence": s["evidence"]})
    return known


def is_named(speaker):
    return speaker.get("name") is not None


def write_json(path, doc):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def write_txt(path, doc):
    head = [f"{doc['id']} | {doc['act']} | {doc['model']} | {doc['status']}", doc["title"], doc["url"],
            f"{doc['minutes']:.1f} minutes, sent in {len(doc['parts'])} part(s)", "", "Speakers:"]
    for s in doc["speakers"]:
        named = f" = {s['name']} ({s['name_source']})" if s.get("name") else ""
        head.append(f"  {s['label']}{named}, {s['role']}: {s['description']}"
                    + (f" | evidence: {s['evidence']}" if s.get("evidence") else ""))
    body = [""]
    for t in doc["turns"]:
        lang = "" if t["language"].lower().startswith("en") else f" [{t['language']}]"
        body.append(f"[{mmss(t['start'])}-{mmss(t['end'])}] {t['speaker']} ({t['role']}){lang}: {t['text']}"
                    + (f"\n      on screen: {t['on_screen']}" if t["on_screen"] else ""))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(head + body) + "\n")


def pacific_today(now=None):
    now = now or datetime.datetime.now(datetime.timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        return now.astimezone(ZoneInfo("America/Los_Angeles")).date().isoformat()
    except Exception:
        return (now - datetime.timedelta(hours=8)).date().isoformat()


class Allowance:
    """Minutes of YouTube video sent per Pacific day, kept in usage_ledger.csv."""

    def __init__(self, path, hours, clock=None):
        self.path, self.limit, self.clock, self.lock = path, hours * 60, clock or pacific_today, threading.Lock()

    def used_today(self):
        if not os.path.exists(self.path):
            return 0.0
        today = self.clock()
        with open(self.path, newline="", encoding="utf-8") as f:
            return sum(float(r["Minutes"]) for r in csv.DictReader(f) if r["Date (Pacific)"] == today)

    def reserve(self, vid, part_no, minutes):
        """Records the minutes about to be sent; False (nothing recorded) when they would pass the daily limit."""
        with self.lock:
            if self.limit > 0 and self.used_today() + minutes > self.limit:
                return False
            new = not os.path.exists(self.path)
            with open(self.path, "a", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=LEDGER_COLS)
                if new:
                    w.writeheader()
                w.writerow({"Date (Pacific)": self.clock(), "ID": vid, "Part": part_no, "Minutes": round(minutes, 2),
                            "Sent at (UTC)": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")})
            return True


def error_kind(e):
    """'daily' (the day's allowance is used up), 'retry' (busy, time-out, dropped connection) or 'fail'."""
    msg = f"{type(e).__name__} {getattr(e, 'code', '')} {getattr(e, 'status', '')} {e}"
    if re.search(r"\b429\b|RESOURCE_EXHAUSTED|quota", msg, re.I):
        return "daily" if re.search(r"per ?day|daily|youtube", msg, re.I) else "retry"
    if re.search(r"\b(500|502|503|504)\b|overloaded|unavailable|timed? ?out|timeout|deadline|connection|"
                 r"remote ?protocol|reset by peer|INTERNAL", msg, re.I):
        return "retry"
    return "fail"


def read_selection():
    import pandas as pd
    if not os.path.exists(SELECTION):
        stop(f"The selection workbook is not at {SELECTION}. Upload harrison_v7_VIDEO_SELECTION.xlsx there.")
    T = pd.read_excel(SELECTION, "To transcribe")
    rows = [r for _, r in T.iterrows()]
    if ONLY_IDS:
        known = {str(r["ID"]) for r in rows}
        unknown = [i for i in ONLY_IDS if i not in known]
        if unknown:
            stop(f"These IDs are not in the To transcribe sheet: {unknown}.")
        return [r for r in rows if str(r["ID"]) in ONLY_IDS], {}
    if not os.path.isdir(TRANSCRIPTS):
        stop(f"The WhisperX transcripts are not at {TRANSCRIPTS}; only videos with one are sent.")
    dropped, other, n_dropped, no_tx, chosen = set(DROPPED), 0, 0, 0, []
    for r in rows:
        if is_group(r) == SOLO:         # not the kind of act asked for
            other += 1
            continue
        if str(r["ID"]) in dropped:
            n_dropped += 1
            continue
        if not os.path.exists(os.path.join(TRANSCRIPTS, f"{r['ID']}.json")):
            no_tx += 1
            continue
        chosen.append(r)
    return chosen, {"other": other, "dropped": n_dropped, "no transcript": no_tx}


def video_seconds(row):
    """The audio's length from the WhisperX transcript, or else the listed length."""
    p = os.path.join(TRANSCRIPTS, f"{row['ID']}.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return float(json.load(f)["audio_seconds"])
        except Exception:
            pass
    return float(row["Minutes"]) * 60


def secret(name):
    """A key from Colab Secrets (or the environment); never printed."""
    try:
        from google.colab import userdata
        value = userdata.get(name)
    except Exception:
        value = os.environ.get(name)
    return value or None


# ------------------------------------------------------------------ Gemini
def gemini_client(key):
    from google import genai
    return genai.Client(api_key=key)


def gemini_call(client, model, url, prompt, part, timeout_s):
    """One request: the YouTube link (clipped to the part when needed) and the prompt. Returns the reply's text,
    its token counts, the model version and why it stopped."""
    from google.genai import types
    vm = (types.VideoMetadata(start_offset=f"{int(part['start'])}s", end_offset=f"{int(math.ceil(part['end']))}s")
          if part["clip"] else None)
    video = types.Part(file_data=types.FileData(file_uri=url), video_metadata=vm)
    cfg = types.GenerateContentConfig(max_output_tokens=MAX_OUTPUT_TOKENS, response_mime_type="application/json",
                                      http_options=types.HttpOptions(timeout=int(timeout_s * 1000)))
    r = client.models.generate_content(model=model, contents=types.Content(role="user", parts=[video, types.Part(text=prompt)]),
                                       config=cfg)
    u = getattr(r, "usage_metadata", None)
    usage = {"prompt": getattr(u, "prompt_token_count", None), "output": getattr(u, "candidates_token_count", None),
             "thinking": getattr(u, "thoughts_token_count", None), "total": getattr(u, "total_token_count", None)}
    finish = None
    try:
        finish = str(r.candidates[0].finish_reason)
    except Exception:
        pass
    return getattr(r, "text", None), usage, getattr(r, "model_version", None), finish


# ------------------------------------------------------------------ one video
class Runner:
    def __init__(self, client_factory, allowance, out_dir, model, sleep=time.sleep):
        self.client_factory, self.allowance, self.out_dir, self.model, self.sleep = \
            client_factory, allowance, out_dir, model, sleep
        self.local, self.stopped, self.print_lock, self.log_lock = threading.local(), threading.Event(), threading.Lock(), threading.Lock()
        self.stop_reason = ""

    def client(self):
        if not hasattr(self.local, "client"):
            self.local.client = self.client_factory()
        return self.local.client

    def say(self, msg):
        with self.print_lock:
            print(msg, flush=True)

    def log(self, row):
        path = os.path.join(self.out_dir, "gemini_log.csv")
        with self.log_lock:
            new = not os.path.exists(path)
            with open(path, "a", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=LOG_COLS)
                if new:
                    w.writeheader()
                w.writerow({k: row.get(k, "") for k in LOG_COLS})

    def ask(self, vid, prompt, url, part):
        """The request with retries. Returns (text, usage, model_version, finish) or raises."""
        length_min = (part["end"] - part["start"]) / 60
        timeout_s = 600 + 90 * length_min
        wait = 30
        for attempt in range(RETRIES):
            if self.stopped.is_set():
                raise DailyLimit(self.stop_reason)
            try:
                return gemini_call(self.client(), self.model, url, prompt, part, timeout_s)
            except Exception as e:
                kind = error_kind(e)
                if kind == "daily":
                    raise DailyLimit(f"Gemini says the day's allowance is used up ({str(e)[:160]})")
                if kind == "fail" or attempt == RETRIES - 1:
                    if re.search(r"\b429\b|RESOURCE_EXHAUSTED", str(e)) and attempt == RETRIES - 1:
                        raise DailyLimit(f"Gemini kept refusing with a quota error ({str(e)[:160]})")
                    raise
                self.say(f"   {vid}: busy or slow ({type(e).__name__}: {str(e)[:100]}); trying again in {wait} s")
                self.sleep(wait)
                wait *= 2

    def run_video(self, row):
        vid, url = str(row["ID"]), str(row["URL"])
        path = os.path.join(self.out_dir, f"{vid}.json")
        total = video_seconds(row)
        doc = None
        if os.path.exists(path) and not REDO:
            with open(path, encoding="utf-8") as f:
                doc = json.load(f)
            if doc.get("status") in ("done", "unreadable"):
                return "skipped"
            doc["status"] = "in progress"     # carry on from the next part
        if doc is None or REDO:
            doc = {"version": VERSION, "id": vid, "act": "; ".join(acts_of(row)), "title": str(row["Title"]),
                   "url": url, "model": self.model, "minutes": round(total / 60, 2), "status": "in progress",
                   "plan": plan_parts(total), "parts": [], "speakers": [], "turns": [], "started": now_utc()}
        t0 = time.time()
        queue = [p for p in doc["plan"] if not completed(doc, p)]
        while queue:
            part = queue.pop(0)
            if self.stopped.is_set():
                break
            n, k = len(doc["plan"]), doc["plan"].index(part) + 1
            prompt = build_prompt(row, part, k, n, doc["speakers"])
            minutes = (part["end"] - part["start"]) / 60
            if not self.allowance.reserve(vid, k, minutes):
                self.stop_reason = "the day's allowance of YouTube video is used up"
                self.stopped.set()
                break
            s0 = time.time()
            try:
                text, usage, mver, finish = self.ask(vid, prompt, url, part)
            except DailyLimit as e:
                self.stop_reason = str(e)
                self.stopped.set()
                break
            except Exception as e:
                doc["parts"].append({**part, "ok": False, "error": f"{type(e).__name__}: {str(e)[:300]}",
                                     "seconds": round(time.time() - s0, 1), "date": now_utc()})
                doc["status"] = "failed"
                break
            parsed = parse_json(text)
            cut = finish is not None and "MAX_TOKENS" in finish.upper()
            if (parsed is None or cut) and split_part(part):
                halves = split_part(part)            # cut short: send the two halves instead
                i = doc["plan"].index(part)
                doc["plan"][i:i + 1] = halves
                queue = halves + queue
                doc["parts"].append({**part, "ok": False, "error": "reply cut short or unreadable; split in two",
                                     "raw": text, "usage": usage, "finish": finish, "seconds": round(time.time() - s0, 1),
                                     "date": now_utc()})
                write_json(path, doc)
                continue
            turns = turns_of(parsed) if parsed is not None else []
            shift = fix_times(turns, part)
            doc["speakers"] = merge_speakers(doc["speakers"], speakers_of(parsed) if parsed is not None else [])
            doc["turns"] += turns
            doc["parts"].append({**part, "ok": parsed is not None, "prompt": prompt, "raw": text, "usage": usage,
                                 "model_version": mver, "finish": finish, "shift": shift,
                                 "seconds": round(time.time() - s0, 1), "date": now_utc()})
            doc["model_version"] = mver or doc.get("model_version")
            write_json(path, doc)
        done = all(completed(doc, p) for p in doc["plan"])
        if done and doc["status"] != "failed":
            doc["status"] = "done" if all(d["ok"] for d in doc["parts"] if "prompt" in d) else "unreadable"
            doc["turns"].sort(key=lambda t: (t["start"] is None, t["start"] or 0))
            doc["finished"] = now_utc()
        write_json(path, doc)
        if doc["status"] in ("done", "unreadable", "failed"):
            write_txt(path[:-5] + ".txt", doc)
            u = [d.get("usage") or {} for d in doc["parts"]]
            labels = doc["speakers"]
            self.log({"ID": vid, "Act": doc["act"], "Minutes": doc["minutes"], "Parts": len(doc["plan"]),
                      "Status": doc["status"], "Seconds": round(time.time() - t0, 1), "Turns": len(doc["turns"]),
                      "Labels": len(labels), "Named labels": sum(1 for s in labels if is_named(s)),
                      "Turns not in English": sum(1 for t in doc["turns"] if not t["language"].lower().startswith("en")),
                      "Prompt tokens": sum(x.get("prompt") or 0 for x in u), "Output tokens": sum(x.get("output") or 0 for x in u),
                      "Thinking tokens": sum(x.get("thinking") or 0 for x in u), "Model version": doc.get("model_version") or "",
                      "Note": next((d.get("error", "") for d in doc["parts"] if not d.get("ok") and "prompt" not in d), ""),
                      "Finished": now_utc()})
        return doc["status"]


def completed(doc, part):
    """A part with a reply recorded (readable or not); a failed request or a split part does not count."""
    return any(d["start"] == part["start"] and d["end"] == part["end"] and "prompt" in d for d in doc["parts"])


def now_utc():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


# ------------------------------------------------------------------ run
def parse_args(argv):
    import argparse
    global ONLY_IDS, REDO, LIST_ONLY, MAX_VIDEOS, WORKERS, DAILY_HOURS, MODEL, OUT_DIR, SOLO
    ap = argparse.ArgumentParser(prog="gemini_speakers.py")
    ap.add_argument("--solo", action="store_true", help="solo artists' videos instead of bands', duos' and groups'")
    ap.add_argument("--ids", default="", help="only these videos, by ID, separated by commas")
    ap.add_argument("--out", default="", help="another output folder, next to this script")
    ap.add_argument("--daily-hours", type=float, default=None, help="daily limit in hours of video (0 = none)")
    ap.add_argument("--workers", type=int, default=None, help="videos sent at the same time")
    ap.add_argument("--max", type=int, default=0, help="send at most this many videos in this run")
    ap.add_argument("--model", default="", help="the Gemini model")
    ap.add_argument("--redo", action="store_true", help="send again videos already saved")
    ap.add_argument("--list", action="store_true", help="show the videos that would be sent, send nothing")
    a = ap.parse_args(argv)
    ONLY_IDS = [x.strip().upper() for x in re.split(r"[,\s]+", a.ids) if x.strip()]
    REDO, LIST_ONLY, MAX_VIDEOS, SOLO = a.redo, a.list, a.max, a.solo
    WORKERS = a.workers if a.workers else WORKERS
    DAILY_HOURS = a.daily_hours if a.daily_hours is not None else DAILY_HOURS
    MODEL = a.model or MODEL
    if a.out:
        OUT_DIR = os.path.join(BASE, a.out)


def main(client_factory=None, clock=None, sleep=time.sleep):
    print(VERSION)
    rows, why = read_selection()
    os.makedirs(OUT_DIR, exist_ok=True)
    allowance = Allowance(os.path.join(OUT_DIR, "usage_ledger.csv"), DAILY_HOURS, clock)

    def status_of(r):
        p = os.path.join(OUT_DIR, f"{r['ID']}.json")
        if not os.path.exists(p) or REDO:
            return "to do"
        with open(p, encoding="utf-8") as f:
            return "done" if json.load(f).get("status") in ("done", "unreadable") else "to do"
    todo = [r for r in rows if status_of(r) == "to do"]
    hours_all = sum(video_seconds(r) for r in rows) / 3600
    hours_todo = sum(video_seconds(r) for r in todo) / 3600
    if why:
        kind, other_kind = ("solo artists", "bands, duos and groups") if SOLO else ("bands, duos and groups", "solo artists")
        print(f"Chosen: {len(rows)} videos of {kind} ({hours_all:.1f} hours). Not sent: {why['other']} of "
              f"{other_kind}, {why['dropped']} dropped in the audit, {why['no transcript']} without a WhisperX transcript.")
    else:
        print(f"Chosen: {len(rows)} videos by ID ({hours_all:.1f} hours).")
    print(f"Already done: {len(rows) - len(todo)}. To do: {len(todo)} ({hours_todo:.1f} hours). Model {MODEL}; "
          f"{WORKERS} at a time; daily limit {'none' if DAILY_HOURS <= 0 else f'{DAILY_HOURS:g} hours'} "
          f"({allowance.used_today() / 60:.1f} hours used today, Pacific time).")
    if LIST_ONLY:
        for r in todo:
            print(f"  {r['ID']:7s} {'; '.join(acts_of(r))[:30]:30s} {video_seconds(r) / 60:6.1f} min  {str(r['Title'])[:70]}")
        return
    if MAX_VIDEOS:
        todo = todo[:MAX_VIDEOS]
    if not todo:
        print("Nothing to do.")
        return
    if client_factory is None:
        key = secret("GEMINI_API_KEY")
        if not key:
            stop("GEMINI_API_KEY is not in Colab Secrets (or Notebook access is off for this notebook).")
        client_factory = lambda: gemini_client(key)
    runner = Runner(client_factory, allowance, OUT_DIR, MODEL, sleep)
    counts, t0, done_n = {}, time.time(), [0]

    def one(r):
        if runner.stopped.is_set():
            return "not started"
        try:
            st = runner.run_video(r)
        except Exception as e:      # an unexpected error in one video does not stop the others
            st = "failed"
            runner.say(f"   {r['ID']}: {type(e).__name__}: {str(e)[:200]}")
        with runner.print_lock:
            done_n[0] += 1
            counts[st] = counts.get(st, 0) + 1
            p = os.path.join(OUT_DIR, f"{r['ID']}.json")
            info = ""
            if os.path.exists(p) and st in ("done", "unreadable"):
                with open(p, encoding="utf-8") as f:
                    d = json.load(f)
                info = f": {len(d['turns'])} turns, {len(d['speakers'])} labels, {len(d['plan'])} part(s)"
            print(f"[{done_n[0]}/{len(todo)}] {r['ID']} {'; '.join(acts_of(r))} ({video_seconds(r) / 60:.1f} min) "
                  f"{st}{info} | {allowance.used_today() / 60:.1f} h used today", flush=True)
        return st

    with ThreadPoolExecutor(max_workers=max(1, WORKERS)) as ex:
        list(ex.map(one, todo))
    print(f"\nFinished this run in {(time.time() - t0) / 60:.0f} min: "
          + ", ".join(f"{v} {k}" for k, v in sorted(counts.items())) + f". Files are in {OUT_DIR}.")
    if runner.stopped.is_set():
        print(f"Stopped early: {runner.stop_reason}. Run the same cell again tomorrow (after midnight Pacific time, "
              f"3 pm in Singapore until 1 November, then 4 pm) to carry on.")


if __name__ == "__main__":
    try:
        parse_args(sys.argv[1:])
        main()
    except Stop:
        pass
