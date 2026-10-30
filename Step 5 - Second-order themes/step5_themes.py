"""step5_themes.py, version 1.8 (30 October 2026)

What it does: runs Step 5, comparing cases for second-order themes, with the three AI agents of Step 3, each working
on its own from its own Step 3 codes, memo and memory, in eight rounds. Each agent keeps a running memo: its memory as
it ended Step 3 (its memos on the acts and the notes it kept, numbered M1, M2 ...), which every call of rounds 1 to 4
updates by adding entries (numbered on from Step 3's), revising them or dropping them. The script makes the changes,
so every later call, and Step 6, sees the memo as updated (in round 1 the acts run side by side, so their changes are
made once all the acts are done, in act order).
  0. timelines (no model calls): each act's events (from events_final.txt) and its turns in date order, each turn
     with its period (before, around or after the act's Best New Artist ceremony; "around" means within
     AROUND_MONTHS either side of it) and its days since each event. This stands in for Step 4 Part 2.
  1. within: each agent studies each act on its own, one call per act (Prompt 1): it reads the act's chronology
     (events and turns in date order, with its own Step 3 codes), drafts the act's themes and writes a summary of
     about 300 words. Each act starts from the running memo as Step 3 left it, so the acts run side by side
     (PARALLEL_ACTS at a time per agent, after a first act that writes the cache);
  2. groupings: from the summaries, each agent sets up its ways of sorting the acts into groups, as many as it judges
     useful and at least MIN_GROUPINGS (Prompt 2);
  3. comparisons: one call per grouping (Prompt 3): the agent compares the acts within and across the groups and
     drafts a column of themes, opening its folder (chronologies, codes, earlier work) where it needs to;
  4. final: each agent brings its theme columns together into one list and records every change (Prompt 4); each
     final theme lists the codes that carry it best and the earlier themes it comes from, which the merge follows
     back to their codes (those of an earlier theme that went into that final theme alone);
  5. memo part 1 (Prompt 5), before any agent sees another's work;
  6. merge (no model calls): the three agents' work in one workbook, without changing or resolving anything, and
     where their final themes draw on the same turns;
  7. lead: Agent A compares the three agents' work (Prompt 6), without choosing between them;
  8. memo parts 2 and 3 (Prompt 7), after reading the lead's comparison.
No round reconciles the three agents' themes. The performance (DV) data is not used in Step 5.

Agents (see MODELS): A = Claude Opus 5.5 (the lead, at low effort), B = Gemini 4 Argon and C = GPT-6.1 Sol (both at
medium effort), as in Step 3.

Reads, in MyDrive/harrison_v7 (it never changes them):
  band_voice_by_act/*.jsonl and harrison_v7_band_voice_acts.json   the corpus, as in Step 3
  step3_inputs/excluded_turns.csv and step3_inputs/<GUIDE_FILE>      the turns Step 3 left out, and the guide
  step3_<STEP3_RUN>/agent_A, agent_B, agent_C                        each agent's Step 3 work: codes/ (its codes by
                                                                     act), state.json (its labels and their
                                                                     definitions), memory.txt and memo_<letter>.md
  step5_inputs/prompt1.txt ... prompt7.txt                            the seven prompts, as plain text, with the
                                                                     {placeholders} as in the prompt document
  step5_inputs/events_final.txt                                       the final events dataset (E001 ...)
Writes, in MyDrive/harrison_v7/step5_<RUN>/:
  timelines/   for each act: <act>_timeline.txt (events and turns in date order, with periods) and
               <act>_days.csv (each turn's days since each event)
  agent_A, agent_B, agent_C   each agent's work: state.json (cost, calls, flags), within/<act>.json, groupings.json,
               comparison_<n>.json, final.json, memo_part1.md, memo_parts2and3.md, memo_<letter>.md (all three
               parts), memory.txt (the running memo as Step 5 left it, for Step 6) and memo_changes.csv (every entry
               added, revised or dropped, and in which call), folder/ (the files it can open), calls/ (every call: the
               changing part of the prompt, the reply, tokens and cost), prompts/ (the fixed parts), flags.jsonl
               (problems that remained)
  merged/      step5_combined.xlsx (codes by pass, themes, groupings, comparisons, revisions, summaries, overlap, the
               running memos' changes), the same as CSV files in merged/folder/, and comparison_by_lead.md
  reports/     pilot_report.md
  run_settings.json, run_log.txt
Everything is saved as it goes: a round that stops (Colab disconnects, you press Stop, the budget runs out) carries on
where it left off when you run it again, and nothing done is done twice.

HOW TO RUN, in Colab (any runtime; no GPU), with the same keys in Colab Secrets as Step 3:
  0. Upload this script to MyDrive/harrison_v7, and the seven prompts and events_final.txt to
     MyDrive/harrison_v7/step5_inputs. Step 3's folder (step3_full) stays where it is.
  1. !pip install -q "anthropic==2.3.0" "openai==3.37.0" "google-genai==2.33.0" tiktoken python-docx openpyxl
     If Colab asks you to restart the session, do so (Runtime > Restart session), then carry on from step 2.
  2. from google.colab import drive; drive.mount("/content/drive")
     %run -i "/content/drive/MyDrive/harrison_v7/step5_themes.py"
     Run this cell again after you change a prompt or a setting. For the full run, add a line under the %run line:
     RUN = "full"; ACTS = "all"; BUDGET_USD = ...   (from the pilot report's projection, with some room)
  3. check_inputs()      the corpus, Step 3's work, the events, the prompts, the guide, the keys (no model calls)
     dry_run()           writes the filled prompts to step5_<RUN>/dry_run, to read before spending anything
     test_models()       a few small calls per model: the keys, the JSON format, the cache and the folder tool
  4. run_all()           every round in order; or one at a time: run_timelines(), run_within(), run_groupings(),
                         run_comparisons(), run_final(), run_memo1(), run_merge(), run_lead(), run_memo23()
     pilot_report()      what the pilot produced, the checks, and the cost and hours of the full run
     status()            where each agent is, its cost and its flags, at any time
  Each round can be run again: it skips what is done and carries on. To stop, press Stop once and wait: the calls in
  progress finish and are saved first.
The pilot runs on PILOT_ACTS in its own folder (step5_pilot); the full run (RUN = "full", ACTS = "all") starts fresh in
step5_full.
"""
import collections, csv, datetime, hashlib, json, os, re, shutil, threading, time, traceback, unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed

VERSION = "step5_themes 1.8 (30 Oct 2026)"

if any(t.name.startswith("step5-agent") and t.is_alive() for t in threading.enumerate()):
    raise SystemExit("Agents from the last run are still finishing their calls in this session, so the script was "
                     "not loaded again. Wait a minute (status() shows where they are), then run this cell again.")

# ==================================================================================================== settings
BASE = "/content/drive/MyDrive/harrison_v7"
RUN = "pilot"                      # the run folder is step5_<RUN>; a new name starts a new run
ACTS = "pilot"                     # "pilot" (PILOT_ACTS), "all" (all 52, in nomination-year order) or a list of names
PILOT_ACTS = ["Irene Cara", "Toni Childs", "Tracy Chapman", "James Ingram", "Glass Tiger", "Men at Work", "The Judds",
              "The Go-Go's"]
STEP3_RUN = "full"                 # Step 3's run folder is step3_<STEP3_RUN>
STEP3_INPUTS = "step3_inputs"      # the guide and the excluded turns, as Step 3 used them
INPUT_DIR = "step5_inputs"         # the seven prompts and the events file
PROMPT_FILES = {n: f"prompt{n}.txt" for n in range(1, 8)}
GUIDE_FILE = "best_practices_guide.docx"
EXCLUDE_FILE = "excluded_turns.csv"
EVENTS_FILE = "events_final.txt"
CEREMONIES = {   # the Best New Artist ceremonies (TVmaze's list of Grammy ceremonies and Wikipedia)
    1980: "1980-02-27", 1981: "1981-02-25", 1982: "1982-02-24", 1983: "1983-02-23", 1984: "1984-02-28",
    1985: "1985-02-26", 1986: "1986-02-25", 1987: "1987-02-24", 1988: "1988-03-02", 1989: "1989-02-22",
    1990: "1990-02-21"}
AROUND_MONTHS = 12                 # "around" the nomination: within this many months either side of the ceremony

MODELS = {   # which model is each agent; effort is the thinking level (Claude at low, the other two at medium)
    "A": {"provider": "anthropic", "model": "claude-opus-5-5", "effort": "low", "max_output": 128000},
    "B": {"provider": "google", "model": "gemini-4-argon", "effort": "medium", "max_output": 65536},
    "C": {"provider": "openai", "model": "gpt-6.1-sol", "effort": "medium", "max_output": 128000},
}
PRICES = {   # US dollars per million tokens, as published on 25 October 2026
    "claude-opus-5-5": {"input": 4.00, "output": 20.00, "cache_read": 0.20, "cache_write_5m": 5.00,
                        "cache_write_1h": 8.00},
    "gemini-4-argon": {"input": 2.00, "output": 10.00, "cache_read": 0.10,
                       "from": {"2026-12-01": {"input": 4.00, "output": 20.00, "cache_read": 0.20}}},
    "gpt-6.1-sol": {"input": 2.00, "output": 10.00, "cache_read": 0.10, "cache_write": 2.50,
                    "long_input": 272000, "long_input_x": 2.0, "long_output_x": 1.5},
}
CONTEXT = {  # each model's context window, in its own tokens (the prompt plus room for max_output), 25 October 2026
    "claude-opus-5-5": 1000000, "gemini-4-argon": 1048576, "gpt-6.1-sol": 1050000}
SECRETS = {"anthropic": "ANTHROPIC_API_KEY", "google": "GEMINI_API_KEY", "openai": "OPENAI_API_KEY"}
BUDGET_USD = 150.0                 # no new call starts once the three agents' total cost in this run passes this (calls
                                   # in progress finish, so the total can pass it by their cost); for the full run, set
                                   # it from the pilot report's projection, with some room

PARALLEL_ACTS = 4                  # round 1: acts in progress at once for each agent (after the first, which writes
                                   # the cache); nothing carries from one act to the next
SUMMARY_WORDS = (240, 360)         # round 1: a summary outside this range is noted,
SUMMARY_ASK = (200, 400)           # and one outside this range is asked for again
LEFT_OUT_ASK = (0.10, 5)           # round 1: an act is asked for again when more than this share of its labels (and
                                   # more than this many) are in no theme and not listed as fitting none; the
                                   # reply names them. Fewer are noted
MAX_SUPPORT, MAX_CONTRARY = 5, 3   # Turn IDs per theme
MAX_THEME_CODES = 30               # rounds 3 and 4: codes listed per theme, those that carry it best (so that round 4's
                                   # reply stays well within Gemini's 65,536 output tokens at 52 acts; B's labels
                                   # average 76 characters); more is noted, not asked again
MIN_GROUPINGS, MIN_GROUP, MIN_PAIRS = 2, 3, 2  # groupings per agent (at least; the agent decides how many); acts per
                                               # group; pairs of each kind compared
MEMO_WORDS = {1: 1000, 2: 600, 3: 300}
LEAD_WORDS = 1500
WORDS_SLACK = 1.3                  # a memo or comparison this much over its length is asked for again, once
CONTEXT_ROOM = 0.8                 # an act whose prompt would fill more than this share of a model's context window
                                   # shows its uncoded turns as one line each (and, if still too long, its coded turns
                                   # with their codes only)
OUTPUT_ROOM = 40000                # tokens kept free for the reply when checking that a prompt fits
RETRIES_UNUSABLE = 2               # re-asks when a reply cannot be used (not JSON, parts missing)
RETRIES_PROBLEMS = 1               # re-asks when a reply can be used but a check finds problems; then it is flagged
TRANSIENT_RETRIES = 8              # after network or server errors (once the SDKs' own retries are used up), waits
FIRST_WAIT, MAX_WAIT = 30, 600     # 30 seconds, then twice as long each time up to 10 minutes, and tries again: about
                                   # 45 minutes in all, so that a provider's outage does not stop an agent overnight
OTHER_RETRIES = 4                  # other errors that are not fatal: tried again only 4 times (about 7 minutes)
MAX_TOOL_STEPS = 60                # rounds 3, 4, 5, 7 and 8: most times an agent may use its folder in one call,
READ_BUDGET = 800000               # and most characters it may read from it in all (about 200,000 tokens)
READ_LINES, READ_CHARS = 300, 60000                    # and in one read
GPT_CHUNK_TOKENS = 4000            # (as in Step 3: GPT caches long memories in chunks)
GPT_MAX_CHUNKS = 40
OUTCOME_WORDS = re.compile(         # round 2: a grouping whose name or dimension sorts on an outcome (not on how the
    r"\b(break ?ups?|broke up|split up|splits? of the (?:group|band)|disband\w*|reunions?|reunit\w*|re-?formed|"
    r"re-?formation|longevity|how long \w+ (?:act|acts|group|groups|band|bands) lasted|lasted|career length|"   # (the
    r"length of (?:career|careers|time together)|still (?:active|together|recording|performing)|one[- ]hit|"    # material
    r"(?:later|long[- ]term|eventual|subsequent|lasting|sustained|enduring) (?:commercial |critical |chart )?success|"
    r"(?<!early )(?<!initial )(?<!first )(?:commercial|chart|critical) (?:success|acclaim|failure|performance)|"
    r"chart (?:peaks?|positions?|performance|hits?)|hit[- ]?mak\w*|sellers?|best[- ]?sell\w*|sales|certifications?|"
    r"kept releasing|later careers?|how \w+ fared|(?:career|commercial|sales|chart) decline|"    # speaks of success
    r"decline of the (?:act|group|band)|faded)\b", re.I)                                         # and recognition)

# ==================================================================================================== basics
_G = globals()                     # with %run -i, loading the script again keeps the same locks and stop signal
PRINT_LOCK = _G.get("PRINT_LOCK") or threading.Lock()
COST_LOCK = _G.get("COST_LOCK") or threading.Lock()
FILE_LOCK = _G.get("FILE_LOCK") or threading.Lock()
STOP = _G.get("STOP") or threading.Event()
STOP_WHY = _G.get("STOP_WHY") or ["stopped by hand"]      # why STOP was set: by hand, or the budget
THREAD_PREFIX = "step5-agent"
COSTS, CACHED_CALLS = {}, {}
LETTERS = list(MODELS)
LEAD = "A"
PERIODS = ("before", "around", "after")


class AgentStop(Exception):
    """Stops one agent's run, with a message saying why; the other agents carry on."""


def stop_reason():
    return AgentStop(STOP_WHY[0])


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


_LOCAL = threading.local()         # each agent thread keeps the run folder it started with, whatever RUN is set to later


def rpath(*p):
    base, run = getattr(_LOCAL, "base", None) or BASE, getattr(_LOCAL, "run", None) or RUN
    return os.path.join(base, f"step5_{run}", *p)


def s3path(letter, *p):
    return os.path.join(getattr(_LOCAL, "base", None) or BASE, f"step3_{STEP3_RUN}", f"agent_{letter}", *p)


def say(*a):
    """Prints a line, and keeps it in step5_<RUN>/run_log.txt, so the log survives a lost Colab page."""
    line = " ".join(str(x) for x in a)
    with PRINT_LOCK:
        print(line, flush=True)
        try:
            if os.path.isdir(rpath()):
                with open(rpath("run_log.txt"), "a", encoding="utf-8") as f:
                    f.write(f"{now()}  {line}\n")
        except Exception:
            pass


def ensure(d):
    os.makedirs(d, exist_ok=True)
    return d


def write_text(path, s):
    ensure(os.path.dirname(path))
    tmp = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"     # one per thread: agents never share one
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(s)
    os.replace(tmp, path)


def write_json(path, obj):
    write_text(path, json.dumps(obj, ensure_ascii=False, indent=1))


def read_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def read_text(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8-sig") as f:
        return f.read()


def append_jsonl(path, obj):
    ensure(os.path.dirname(path))
    with FILE_LOCK, open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def read_jsonl(path):
    out = []
    if os.path.exists(path):
        with FILE_LOCK, open(path, encoding="utf-8") as f:
            for line in f:
                try:
                    out.append(json.loads(line))
                except Exception:
                    pass
    return out


def sha(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def short(s, n=160):
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s if len(s) <= n else s[:n - 1] + "…"


def nwords(s):
    return len(str(s or "").split())


_ENC = None


def est_tokens(s):
    """Tokens counted before a call: o200k_base where tiktoken can load it, otherwise characters / 4."""
    global _ENC
    if _ENC is None:
        try:
            import tiktoken
            _ENC = tiktoken.get_encoding("o200k_base")
        except Exception:
            _ENC = False
    if _ENC:
        try:
            return len(_ENC.encode(s, disallowed_special=()))
        except Exception:
            pass
    return len(s) // 4


def secret(name):
    """A key from Colab Secrets. Quotation marks, spaces and invisible characters copied with it are dropped (a key
    never has them); a key that still has other characters is refused with a clear message, and never printed."""
    v = None
    try:
        from google.colab import userdata
        v = userdata.get(name)
    except Exception:
        pass
    v = v or os.environ.get(name)
    if not v:
        raise AgentStop(f"the key {name} is not in Colab Secrets (or notebook access is off)")
    v = re.sub(r"[\s ​-‏⁠﻿\"'`‘’“”]", "", v)
    if not v or not v.isascii():
        raise AgentStop(f"the key {name} in Colab Secrets has characters a key never has, perhaps copied with it from "
                        f"a document; paste it again, straight from the provider's page")
    return v


def busy():
    """Agent threads still running, for example after Stop while their calls finish."""
    return sorted(t.name for t in threading.enumerate() if t.name.startswith(THREAD_PREFIX) and t.is_alive())

# ==================================================================================================== corpus


CORPUS = {"acts": None, "turns": {}, "all_turns": {}, "by_id": {}, "stem_of": {}, "excluded": {}}


def act_stem(a):
    return a["File"][:-6] if a["File"].endswith(".jsonl") else a["File"]


def read_exclusions():
    """The turns Step 3 left out: step3_inputs/excluded_turns.csv, a Turn ID per row (other columns are notes)."""
    p = os.path.join(BASE, STEP3_INPUTS, EXCLUDE_FILE)
    if not os.path.exists(p):
        return {}
    out = {}
    with open(p, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            tid = str(row.get("Turn ID") or "").strip()
            if tid:
                out[tid] = {k: (v or "") for k, v in row.items() if k}
    return out


def load_corpus():
    """The corpus, act by act in nomination-year order. turns: the turns Step 3 coded (without the excluded turns);
    all_turns: every turn; by_id: every turn by its Turn ID."""
    if CORPUS["acts"] is not None:
        return CORPUS
    j = read_json(os.path.join(BASE, "harrison_v7_band_voice_acts.json"))
    if j is None:
        raise SystemExit("STOPPED: harrison_v7_band_voice_acts.json not found in " + BASE)
    acts = j["acts"]
    acts = [acts[i] for i in sorted(range(len(acts)), key=lambda i: (int(acts[i]["Nominated"]), i))]
    excl = read_exclusions()
    turns, all_turns, by_id, stem_of = {}, {}, {}, {}
    for a in acts:
        p = os.path.join(BASE, "band_voice_by_act", a["File"])
        with open(p, encoding="utf-8") as f:
            rows = [json.loads(line) for line in f if line.strip()]
        rows.sort(key=lambda r: r["Order"])
        for r in rows:
            by_id[r["Turn ID"]] = r
            stem_of[r["Turn ID"]] = act_stem(a)
        all_turns[act_stem(a)] = rows
        turns[act_stem(a)] = [r for r in rows if r["Turn ID"] not in excl]
    CORPUS.update(acts=acts, turns=turns, all_turns=all_turns, by_id=by_id, stem_of=stem_of,
                  excluded={t: v for t, v in excl.items() if t in by_id})
    return CORPUS


def select_acts(acts=None):
    c = load_corpus()
    acts = ACTS if acts is None else acts
    if acts == "all":
        return list(c["acts"])
    names = PILOT_ACTS if acts == "pilot" else ([acts] if isinstance(acts, str) else list(acts))
    by_name = {a["Act"].casefold(): a for a in c["acts"]}
    out = []
    for n in names:
        if n.casefold() not in by_name:
            raise SystemExit(f"STOPPED: no act called {n!r} in the corpus.")
        out.append(by_name[n.casefold()])
    order = {act_stem(a): i for i, a in enumerate(c["acts"])}
    return sorted(out, key=lambda a: order[act_stem(a)])


def kind(a):
    return "solo artist" if str(a.get("Solo or group", "")).lower().startswith("solo") else "group"


def won(a):
    return str(a.get("Result", "")).lower().startswith("win")


def ceremony(a):
    y = int(a["Nominated"])
    if y not in CEREMONIES:
        raise SystemExit(f"STOPPED: no ceremony date for {y} in CEREMONIES ({a['Act']}).")
    return datetime.date.fromisoformat(CEREMONIES[y])


def fmt_date(t):
    d, p = (t.get("Date") or "").strip(), (t.get("Date precision") or "").lower()
    if not d:
        return "unknown"
    return d[:4] if p == "year" else d[:7] if p == "month" else d[:10]


def as_date(s, precision=None):
    """A date to place things in time: a date given only to the month is put on the 15th, one given only to the
    year on 1 July; None if there is no date."""
    s = str(s or "").strip()
    if not s:
        return None
    p = (precision or "").lower()
    try:
        if p == "year" or len(s) == 4:
            return datetime.date(int(s[:4]), 7, 1)
        if p == "month" or len(s) == 7:
            return datetime.date(int(s[:4]), int(s[5:7]), 15)
        return datetime.date.fromisoformat(s[:10])
    except ValueError:
        return None


def period_of(d, cer):
    if d is None:
        return "unknown"
    room = datetime.timedelta(days=round(AROUND_MONTHS * 365.25 / 12))
    if abs(d - cer) <= room:
        return "around"
    return "before" if d < cer else "after"


def turn_date(t):
    return as_date(t.get("Date"), t.get("Date precision"))

# ==================================================================================================== Step 3's work


S3 = {}


INVISIBLE = dict.fromkeys(map(ord, "\u00ad\u200b\u200c\u200d\u2060\ufeff"), None)   # soft hyphen, zero-width
                                   # characters: a model sometimes writes one inside a word (C did in pilot 3), which
                                   # would make a label it copied look like a code it never used


def label_key(label):
    return re.sub(r"\s+", " ", str(label).translate(INVISIBLE).casefold()).strip(" .,;:!?\"'“”‘’")


def step3(letter):
    """The agent's Step 3 work: its memo (all three parts), its memory as it ended Step 3, and its labels."""
    if letter not in S3:
        memo = read_text(s3path(letter, f"memo_{letter}.md"))
        if not memo or not memo.strip():
            raise AgentStop(f"Step 3's memo {s3path(letter, f'memo_{letter}.md')} is not there; is STEP3_RUN right?")
        memory = read_text(s3path(letter, "memory.txt"))
        if not memory or not memory.strip():
            raise AgentStop(f"Step 3's memory {s3path(letter, 'memory.txt')} is not there")
        st = read_json(s3path(letter, "state.json")) or {}
        labels = {}
        for x in (st.get("labels") or {}).values():
            labels[label_key(x["label"])] = {"label": x["label"], "from": x.get("from", ""),
                                             "definition": x.get("definition") or x.get("description", "")}
        S3[letter] = {"memo": memo.strip(), "memory": memory.strip(), "labels": labels, "codes": {},
                      "next": int(st.get("next") or 0)}       # the number Step 3's next memory entry would have had
    return S3[letter]


def act_codes(letter, stem):
    """The agent's Step 3 codes for an act, turn by turn: {Turn ID: {"looks_back", "data_problem", "codes": [...]}}."""
    s3 = step3(letter)
    if stem not in s3["codes"]:
        p = s3path(letter, "codes", f"codes_{stem}.csv")
        if not os.path.exists(p):
            raise AgentStop(f"Step 3's codes for {stem} are not there ({p})")
        out = collections.OrderedDict()
        with open(p, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                tid = (r.get("Turn ID") or "").strip()
                if not tid:
                    continue
                rec = out.setdefault(tid, {"looks_back": False, "data_problem": "", "codes": []})
                rec["looks_back"] = rec["looks_back"] or (r.get("Looks back") or "").strip().lower() == "yes"
                rec["data_problem"] = rec["data_problem"] or (r.get("Data problem") or "").strip()
                lab = (r.get("Label") or "").strip()
                if lab and lab not in ("(no code)", "(skipped)"):
                    rec["codes"].append({"label": lab, "from": (r.get("Label from") or "").strip(),
                                         "description": (r.get("Description") or "").strip(),
                                         "evidence": (r.get("Evidence") or "").strip(),
                                         "unsure": (r.get("Unsure") or "").strip().lower() == "yes"})
        s3["codes"][stem] = out
    return s3["codes"][stem]


def act_labels(letter, stem):
    """The labels the agent used in an act: {label_key: label}."""
    return {label_key(c["label"]): c["label"] for r in act_codes(letter, stem).values() for c in r["codes"]}


def run_labels(letter, acts):
    out = {}
    for a in acts:
        out.update(act_labels(letter, act_stem(a)))
    return out


def label_turns(letter, acts):
    """For each label the agent used in these acts, the turns it coded with it."""
    out = collections.defaultdict(set)
    for a in acts:
        for tid, r in act_codes(letter, act_stem(a)).items():
            for c in r["codes"]:
                out[label_key(c["label"])].add(tid)
    return out

# ==================================================================================================== events and timelines


EVENTS = {"by_act": None}
EVENT_NAMES = {"nomination": "Best New Artist nomination", "win": "Best New Artist win",
               "studio_album_release": "Studio album", "breakup": "Breakup", "reunion": "Reunion"}


def act_key(name):
    """An act's name for matching the events file to the corpus: accents, case, a leading "The" and punctuation aside
    ("The Go-Go's" and "Go-Go’s", "Tone Loc" and "Tone Lōc" match)."""
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode().casefold().strip()
    s = re.sub(r"^the\s+", "", s)
    return re.sub(r"[^a-z0-9]+", "", s)


def load_events():
    """The final events dataset, act by act, in date order."""
    if EVENTS["by_act"] is None:
        p = os.path.join(BASE, INPUT_DIR, EVENTS_FILE)
        if not os.path.exists(p):
            raise SystemExit(f"STOPPED: {p} not found.")
        by = collections.defaultdict(list)
        with open(p, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f, delimiter="\t"):
                if not (r.get("event_id") or "").strip():
                    continue
                by[act_key(r["act_name"])].append(
                    {"id": r["event_id"].strip(), "type": r["event_type"].strip(), "date": r["event_date_start"].strip(),
                     "precision": (r.get("date_precision") or "").strip(),
                     "description": (r.get("event_description") or "").strip(),
                     "uncertain": (r.get("status") or "").strip().lower() == "uncertain"})
        for v in by.values():
            v.sort(key=lambda e: (as_date(e["date"], e["precision"]) or datetime.date.max, e["id"]))
        EVENTS["by_act"] = dict(by)
    return EVENTS["by_act"]


def act_events(a):
    return load_events().get(act_key(a["Act"]), [])


def turns_in_order(a):
    """The act's turns (those Step 3 kept) in date order, partial dates placed as for the periods (the 15th of the
    month, 1 July); turns with the same date keep the corpus's order, interview by interview."""
    return sorted(load_corpus()["turns"][act_stem(a)], key=lambda t: turn_date(t) or datetime.date.max)


def event_line(e):
    what = e["description"] or EVENT_NAMES.get(e["type"], e["type"])
    return f"EVENT {e['id']}, {e['date'] or 'date unknown'}" + (" (date uncertain)" if e["uncertain"] else "") \
        + f": {what}"


def source_line(t):
    src = ["article" if t.get("Source") == "text" else "video"]
    src += [str(x).strip() for x in (t.get("Outlet"), t.get("Title")) if x and str(x).strip()]
    return "; ".join(src)


def code_line(c):
    tag = c["from"] or "data"
    if c["unsure"]:
        tag += "; unsure"
    return f"- {c['label']} [{tag}]: {c['description']} \"{c['evidence']}\""


def render_turn5(t, cer, rec, level):
    """A turn as Prompt 1 shows it. level 0: in full; level 1: a turn with no code on one line; level 2: also a turn
    with codes without its text (its codes keep their evidence)."""
    rec = rec or {"looks_back": False, "data_problem": "", "codes": []}
    per = period_of(turn_date(t), cer)
    if level >= 1 and not rec["codes"]:
        return f"Turn ID: {t['Turn ID']} ({t.get('Speaker') or 'unknown'}, {fmt_date(t)}, {per}): no code"
    lines = [f"Turn ID: {t['Turn ID']}", f"Speaker: {t.get('Speaker') or 'unknown'}", f"Date: {fmt_date(t)}",
             f"Period: {per}", f"Source: {source_line(t)}"]
    if rec["looks_back"]:
        lines.append("Looks back: yes")
    if rec["data_problem"]:
        lines.append(f"Data problem: {rec['data_problem']}")
    if level < 2:
        lines.append(f"Text: {t['Text']}")
    lines.append("Your codes:" if rec["codes"] else "Your codes: none")
    lines += [code_line(c) for c in rec["codes"]]
    return "\n".join(lines)


def chronology(a, codes=None, level=0):
    """The act's events and turns in date order (turns_in_order), each event placed before the first turn dated on or
    after it. With codes (an agent's Step 3 codes), each turn shows them; without, the timeline has no codes."""
    cer = ceremony(a)
    room = round(AROUND_MONTHS * 365.25 / 12)
    head = (f"The {AROUND_MONTHS} months either side of the ceremony on {cer.isoformat()} count as around it: from "
            f"{(cer - datetime.timedelta(days=room)).isoformat()} to {(cer + datetime.timedelta(days=room)).isoformat()}.")
    evs = list(act_events(a))
    out = [head]
    for t in turns_in_order(a):
        d = turn_date(t)
        while evs and d is not None and (as_date(evs[0]["date"], evs[0]["precision"]) or datetime.date.max) <= d:
            out.append(event_line(evs.pop(0)))
        if codes is None:
            per = period_of(d, cer)
            out.append(f"Turn ID: {t['Turn ID']}\nSpeaker: {t.get('Speaker') or 'unknown'}\nDate: {fmt_date(t)}\n"
                       f"Period: {per}\nSource: {source_line(t)}\nText: {t['Text']}")
        else:
            out.append(render_turn5(t, cer, codes.get(t["Turn ID"]), level))
    out += [event_line(e) for e in evs]
    return "\n\n".join(out)


def build_timelines(acts):
    """Step 4's timelines (no model calls): for each act, <act>_timeline.txt (events and turns in date order, with
    periods) and <act>_days.csv (each turn's days since each event: negative for an event after the turn)."""
    d = ensure(rpath("timelines"))
    for a in acts:
        stem = act_stem(a)
        write_text(os.path.join(d, f"{stem}_timeline.txt"), chronology(a) + "\n")
        evs, cer = act_events(a), ceremony(a)
        with open(os.path.join(d, f"{stem}_days.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Turn ID", "Date", "Date precision", "Period"]
                       + [f"Days since {e['id']} ({EVENT_NAMES.get(e['type'], e['type'])}, {e['date']})" for e in evs])
            for t in turns_in_order(a):
                td = turn_date(t)
                row = [t["Turn ID"], fmt_date(t), t.get("Date precision") or "", period_of(td, cer)]
                for e in evs:
                    ed = as_date(e["date"], e["precision"])
                    row.append((td - ed).days if td and ed else "")
                w.writerow(row)
    return d

# ==================================================================================================== prompts and guide


PROMPTS, GUIDE = {}, {"text": None}
PH = re.compile(r"\{(?![\"“])([^{}\n]+)\}")     # a placeholder; a JSON object in a format starts with {" instead
REQUIRED = {
    1: {"GUIDE", "LETTER", "MEMO3", "MEMORY", "ACT", "KIND", "RESULT", "YEAR", "DATE", "N_INTERVIEWS", "N_TURNS",
        "N_CODED", "CHRONO"},
    2: {"GUIDE", "LETTER", "MEMO3", "MEMORY", "ACTS"},
    3: {"GUIDE", "LETTER", "MEMO3", "MEMORY", "GNUM", "GTOTAL", "GNAME", "GDIM", "N_ACTS", "GROUPING", "ACTS"},
    4: {"GUIDE", "LETTER", "MEMO3", "MEMORY", "N_ACTS", "COLUMNS"},
    5: {"GUIDE", "LETTER", "MEMO3", "MEMORY", "FINAL"},
    6: {"GUIDE", "GROUPINGS3", "FINAL3", "OVERLAP"},
    7: {"GUIDE", "LETTER", "MEMO1", "MEMORY", "COMPARISON", "FINAL3"}}
BIG = ("GUIDE", "MEMO3", "MEMORY", "CHRONO", "ACTS", "GROUPING", "COLUMNS", "FINAL", "GROUPINGS3", "FINAL3", "OVERLAP",
       "MEMO1", "COMPARISON")


def ph_key(inner):
    s = inner.strip()
    low = s.lower()
    if "best practices" in low:
        return "GUIDE"
    if re.fullmatch(r"a\s*[,/]\s*b\s*(,\s*)?(or|/)\s*c", low):
        return "LETTER"
    if low.startswith("your memory") or low.startswith("your running memo"):   # before the memo: it says "your
        return "MEMORY"                                                            # Step 3 memo on each act"
    if "step 3 memo" in low:
        return "MEMO3"
    if "step 5 memo, part 1" in low or "memo part 1" in low:
        return "MEMO1"
    if low.startswith("the chronology"):
        return "CHRONO"
    if low.startswith("for each act"):
        return "ACTS"
    if low.startswith("the grouping:") or low.startswith("the grouping "):
        return "GROUPING"
    if "columns of themes" in low or "four columns" in low:
        return "COLUMNS"
    if low.startswith("your final themes"):
        return "FINAL"
    if "three agents' groupings" in low or "three agents’ groupings" in low:
        return "GROUPINGS3"
    if "three agents' final themes" in low or "three agents’ final themes" in low:
        return "FINAL3"
    if low.startswith("overlap figures"):
        return "OVERLAP"
    if "lead agent's comparison" in low or "lead agent’s comparison" in low:
        return "COMPARISON"
    if low == "name":
        return "ACT"
    if "group or solo" in low:
        return "KIND"
    if "won or nominated" in low:
        return "RESULT"
    if low == "year":
        return "YEAR"
    if low == "date":
        return "DATE"
    if low == "grouping":
        return "GNAME"
    if low == "dimension":
        return "GDIM"
    if low in ("number of groupings", "the number of groupings"):
        return "GTOTAL"
    return {"k": "N_INTERVIEWS", "m": "N_TURNS", "c": "N_CODED", "n": "GNUM", "N": "N_ACTS"}.get(s)


def keys_in(text):
    return [ph_key(m.group(1)) for m in PH.finditer(text)]


def fill(template, values):
    """Fills the {placeholders} in one pass, so that braces inside the filled text are never touched."""
    def sub(m):
        k = ph_key(m.group(1))
        return str(values[k]) if k in values else m.group(0)
    return PH.sub(sub, template)


def split_after(template, key):
    """The prompt up to and including a placeholder (what stays the same from call to call), and the rest."""
    ms = [m for m in PH.finditer(template) if ph_key(m.group(1)) == key]
    return template[:ms[0].end()], template[ms[0].end():]


def read_prompt(n):
    p = os.path.join(BASE, INPUT_DIR, PROMPT_FILES[n])
    if not os.path.exists(p):
        raise SystemExit(f"STOPPED: {p} not found.")
    with open(p, encoding="utf-8-sig") as f:
        s = f.read().replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in s.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    if lines and re.match(r"\s*Prompt\s*\d+\s*:", lines[0], re.I):
        lines.pop(0)
    return "\n".join(lines).strip("\n")


def check_prompt(n, s):
    problems = []
    found = [m.group(0) for m in PH.finditer(s)]
    ks = keys_in(s)
    unknown = [f for f, k in zip(found, ks) if k is None]
    if unknown:
        problems.append(f"placeholders the script does not know: {unknown}")
    missing = REQUIRED[n] - set(ks)
    if missing:
        problems.append(f"placeholders missing: {sorted(missing)}")
    extra = set(k for k in ks if k) - REQUIRED[n]
    if extra:
        problems.append(f"placeholders that belong to other prompts: {sorted(extra)}")
    for k in BIG:
        if ks.count(k) > 1:
            problems.append(f"{k} appears {ks.count(k)} times")
    if n == 1 and "MEMORY" in ks:
        pre, _ = split_after(s, "MEMORY")
        bad = set(k for k in keys_in(pre) if k) - {"GUIDE", "LETTER", "MEMO3", "MEMORY"}
        if bad:
            problems.append(f"these change from act to act, so they must come after the memory: {sorted(bad)}")
    if n in (1, 2, 3, 4) and "JSON" not in s:
        problems.append("no JSON format in the prompt")
    return problems


def load_prompts():
    if not PROMPTS:
        for n in PROMPT_FILES:
            s = read_prompt(n)
            pr = check_prompt(n, s)
            if pr:
                raise SystemExit(f"STOPPED: {PROMPT_FILES[n]}: " + "; ".join(pr))
            PROMPTS[n] = s
    return PROMPTS


def docx_text(path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    out = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[-1]
        if tag == "p":
            p = Paragraph(el, d)
            t = p.text.strip()
            if not t:
                continue
            style = (p.style.name or "").lower() if p.style is not None else ""
            if style.startswith("heading") or style == "title":
                lvl = re.findall(r"\d+", style)
                out.append("#" * min(int(lvl[0]) if lvl else 1, 6) + " " + t)
            elif "list" in style:
                out.append("- " + t)
            else:
                out.append(t)
        elif tag == "tbl":
            for row in Table(el, d).rows:
                cells = []
                for c in row.cells:
                    ct = " ".join(x.text.strip() for x in c.paragraphs if x.text.strip())
                    if not cells or ct != cells[-1]:
                        cells.append(ct)
                if any(cells):
                    out.append(" | ".join(cells))
    return "\n\n".join(out)


def load_guide():
    if GUIDE["text"] is None:
        p = os.path.join(BASE, STEP3_INPUTS, GUIDE_FILE)
        if not os.path.exists(p):
            raise SystemExit(f"STOPPED: the guide {p} not found (set GUIDE_FILE).")
        if p.lower().endswith(".docx"):
            g = docx_text(p)
        else:
            g = read_text(p).replace("\r\n", "\n").strip()
        if len(g.split()) < 500:
            raise SystemExit(f"STOPPED: the guide has only {len(g.split())} words; is it the right file?")
        GUIDE["text"] = g
    return GUIDE["text"]

# ==================================================================================================== run settings record


ROUNDS = {"within": 1, "groupings": 2, "comparisons": 3, "final": 4, "memo1": 5, "lead": 6, "memo23": 7}


def round_settings(r, acts):
    s = {"script": VERSION, "acts": [act_stem(a) for a in acts], "step3_run": STEP3_RUN}
    if r in ROUNDS:
        load_prompts()
        s.update(models=MODELS, guide_sha=sha(load_guide()), prompt_sha=sha(PROMPTS[ROUNDS[r]]))
    if r in ("timelines", "within"):
        s["events_sha"] = sha(read_text(os.path.join(BASE, INPUT_DIR, EVENTS_FILE)) or "")
        s["around_months"] = AROUND_MONTHS
    return s


def round_started(r):
    if r == "timelines":
        return False                           # the timelines are simply written again
    if r == "merge":
        return os.path.exists(rpath("merged", "done.json"))
    if r == "lead":
        return os.path.exists(rpath("merged", "comparison_by_lead.md"))
    for L in LETTERS:
        d = rpath(f"agent_{L}")
        if r == "within" and os.path.isdir(os.path.join(d, "within")) and os.listdir(os.path.join(d, "within")):
            return True
        if r == "groupings" and os.path.exists(os.path.join(d, "groupings.json")):
            return True
        if r == "comparisons" and any(fn.startswith("comparison_") for fn in (os.listdir(d) if os.path.isdir(d) else [])):
            return True
        if r == "final" and os.path.exists(os.path.join(d, "final.json")):
            return True
        if r == "memo1" and os.path.exists(os.path.join(d, "memo_part1.md")):
            return True
        if r == "memo23" and os.path.exists(os.path.join(d, "memo_parts2and3.md")):
            return True
    return False


def check_settings(r, acts, agents=None):
    """The settings a round started with are kept: once a round has work saved, it never carries on with other
    prompts, guide, models, acts or events. Before that, the settings are simply recorded again."""
    path = rpath("run_settings.json")
    rec = read_json(path, {})
    cur = round_settings(r, acts)
    old = rec.get(r)
    run = {"script": VERSION, "agents": list(agents or LETTERS), "time": now()}
    if old is None or not round_started(r):
        rec[r] = dict(cur, started=now(), runs=[run])
        write_json(path, rec)
        return
    diff = [k for k in cur if k != "script" and old.get(k) != cur[k]]
    if diff:
        raise SystemExit(f"STOPPED: round {r!r} of run {RUN!r} started with other settings: {', '.join(diff)}. Put "
                         f"back what it started with, or set RUN to a new name to start a new run.")
    old.setdefault("runs", []).append(run)
    write_json(path, rec)

# ==================================================================================================== agent state


def new_state(letter):
    return {"agent": letter, "cost": 0.0, "calls": 0, "flags": 0, "stopped": None}


def load_state(letter):
    path = rpath(f"agent_{letter}", "state.json")
    try:
        st = read_json(path)
    except Exception:                          # a damaged file: kept aside, and the counts rebuilt from the call logs
        st = None
        shutil.move(path, path[:-5] + f"_damaged_{datetime.datetime.now():%Y%m%d_%H%M%S}.json")
    if st is None:
        st = new_state(letter)
        sync_counters(st)
    for k, v in new_state(letter).items():
        st.setdefault(k, v)
    return st


def logged_calls(letter):
    d = rpath(f"agent_{letter}", "calls")
    return [c for fn in (sorted(os.listdir(d)) if os.path.isdir(d) else []) if fn.endswith(".jsonl")
            for c in read_jsonl(os.path.join(d, fn))]


def sync_counters(st):
    """After a crash, Drive can keep an older state.json than the call logs: where those are ahead, the number of
    calls and their cost are taken from them, so that call IDs never repeat and the budget counts every call."""
    L = st["agent"]
    calls = logged_calls(L)
    n = max([int(str(c.get("call", "")).rsplit("-", 1)[-1]) for c in calls
             if re.fullmatch(rf"{L}-\d+", str(c.get("call", "")))] + [0])
    cost = round(sum(c.get("cost", 0) or 0 for c in calls), 6)
    changed = False
    if n > st["calls"]:
        st["calls"], changed = n, True
    if cost > st["cost"] + 0.005:
        st["cost"], changed = cost, True
    if changed:
        say(f"[{L}] state.json was behind the agent's call logs, so its counts were taken from them: "
            f"{st['calls']} calls, ${st['cost']:.2f}.")
    return changed


STATE_LOCKS = collections.defaultdict(threading.Lock)


def save_state(st):
    with STATE_LOCKS[st["agent"]]:
        write_text(rpath(f"agent_{st['agent']}", "state.json"), json.dumps(st, ensure_ascii=False))


def total_cost():
    with COST_LOCK:
        if COSTS:
            return sum(COSTS.get(L, 0.0) for L in LETTERS)
    return sum(load_state(L)["cost"] for L in LETTERS)


def flag(st, round_name, act, items, call_id):
    """Keeps the problems that remained after the re-asks in flags.jsonl."""
    for p in items:
        append_jsonl(rpath(f"agent_{st['agent']}", "flags.jsonl"),
                     {"agent": st["agent"], "round": round_name, "act": act, "problem": short(p, 1000),
                      "call": call_id, "time": now()})
    with COST_LOCK:
        st["flags"] += len(items)

# ==================================================================================================== checks


QUOTES = dict.fromkeys(map(ord, "\"'“”‘’`´«»"), None)
ELLIPSIS = re.compile(r"[\[(]?\s*(?:\.\s*\.\s*\.|…)\s*[\])]?")
BRACKETED = re.compile(r"\[[^\[\]]*\]")
TID_RE = re.compile(r"\b(?:V\d+-\d+|\d{1,6}-\d+)\b")
QUOTE_RE = re.compile(r"“([^”\n]{1,500})”|\"([^\"\n]{1,500})\"")


def norm_words(s):
    s = unicodedata.normalize("NFKC", str(s).translate(INVISIBLE)).casefold().translate(QUOTES)
    return re.findall(r"\w+", s)


def turn_words(text):
    s = unicodedata.normalize("NFKC", str(text).translate(INVISIBLE)).casefold().translate(QUOTES)
    inside = bytearray(len(s) + 1)
    for m in BRACKETED.finditer(s):
        inside[m.start():m.end()] = b"\x01" * (m.end() - m.start())
    return [(m.group(0), bool(inside[m.start()])) for m in re.finditer(r"\w+", s)]


def _match_from(frag, tw, i):
    todo, seen = [(0, i)], set()
    while todo:
        k, j = todo.pop()
        if k == len(frag):
            return j
        if j >= len(tw) or (k, j) in seen:
            continue
        seen.add((k, j))
        word, optional = tw[j]
        if optional:
            todo.append((k, j + 1))
        if word == frag[k]:
            todo.append((k + 1, j + 1))
    return -1


def _seq_at(frag, tw, start=0):
    for i in range(start, len(tw)):
        if tw[i][0] == frag[0]:
            end = _match_from(frag, tw, i)
            if end >= 0:
                return end
    return -1


def evidence_in(ev, text):
    """A quotation is the turn's own words, in order: spacing, capitals, punctuation and quotation marks aside; '...'
    may join parts of the turn, and an editor's insertion in square brackets in the turn may be left out."""
    tw = turn_words(text)
    frags = [f for f in (norm_words(x) for x in ELLIPSIS.split(ev)) if f]
    pos = 0
    for f in frags:
        pos = _seq_at(f, tw, pos)
        if pos < 0:
            return False
    return bool(frags)


_NORM_CACHE = {}


def _norm_text(tid):
    if tid not in _NORM_CACHE:
        _NORM_CACHE[tid] = " ".join(norm_words(CORPUS["by_id"][tid]["Text"]))
    return _NORM_CACHE[tid]


_RECORDS = {"key": None, "ids": set()}


def corpus_records():
    """The record IDs of the corpus: a Turn ID is its article's or video's record ID, a hyphen and the turn's place."""
    by = CORPUS["by_id"]
    if _RECORDS["key"] != (id(by), len(by)):
        _RECORDS.update(key=(id(by), len(by)), ids={t.rsplit("-", 1)[0] for t in by})
    return _RECORDS["ids"]


def looks_like_tid(x):
    """A Turn ID-shaped string that is not in the corpus is a problem only when its record part is a record of the
    corpus, and it is not a span of years such as "1984-85"."""
    a, b = x.rsplit("-", 1)
    if a not in corpus_records():
        return False
    return not (a.isdigit() and len(a) == 4 and 1900 <= int(a) <= 2099
                and (len(b) <= 2 or (len(b) == 4 and 1900 <= int(b) <= 2099)))


def in_a_name(q, names):
    """Whether a quote is a label or a theme's name, or part of one: its words run on in one name, or, if it is
    shortened with '...' or '…', its parts do, in order. `names` are the names as " word word ... ", lower case."""
    frags = [" ".join(f) for f in (norm_words(x) for x in ELLIPSIS.split(q)) if f]
    if not frags:
        return False
    for n in names:
        pos = 0
        for f in frags:
            k = n.find(" " + f + " ", pos)
            if k < 0:
                break
            pos = k + len(f) + 1
        else:
            return True
    return False


def quote_problems(text, scope_ids, labels_ok=()):
    """The quotes in a memo, point by point: each must be word for word in a turn the point cites (or, if it cites
    none, in a turn of the acts in scope), unless it is a label or a theme's name, in full or in part (three words or
    more, or parts joined by '...'); each Turn ID must be in the corpus."""
    problems, ok, names = [], {label_key(x) for x in labels_ok}, None
    for pt in re.split(r"\n\s*(?=(?:\d+[.)]|[-*•])\s)", str(text)):
        ids = TID_RE.findall(pt)
        bad = sorted(set(i for i in ids if i not in CORPUS["by_id"] and looks_like_tid(i)))
        if bad:
            problems.append(f"Turn ID(s) that are not in the corpus: {', '.join(bad[:5])}.")
        cited = [i for i in dict.fromkeys(ids) if i in CORPUS["by_id"]]
        for mq in QUOTE_RE.finditer(pt):
            raw = mq.group(1) or mq.group(2) or ""
            q = raw.strip(" .,;:")
            words = norm_words(q)
            if len(q) < 6 or len(words) < 3 or label_key(q) in ok:
                continue
            if names is None:
                names = [" " + " ".join(norm_words(x)) + " " for x in labels_ok]
            if in_a_name(raw, names):              # raw: the strip above would take a final "..." off
                continue
            first = " ".join(norm_words(ELLIPSIS.split(q)[0])) or " ".join(words)
            pool = cited or [i for i in scope_ids if first in _norm_text(i)]
            if any(evidence_in(q, CORPUS["by_id"][i]["Text"]) for i in pool):
                continue
            problems.append(f'The quote "{short(q, 90)}" is not word for word in '
                            + (f"the turn(s) it cites ({', '.join(cited[:4])})." if cited else "any turn of the acts."))
    return problems


FENCE_RE = re.compile(r"```(?:json|JSON)?\s*(\{.*\})\s*```", re.S)


def parse_json(text):
    """The JSON object in a reply: the whole reply, a ```json block inside it, or the text from the first { to the
    last } (rounds 3 and 4 reply in text, after using their folder, and may put a sentence before the JSON)."""
    if text is None or not str(text).strip():
        return None, "the reply is empty"
    s = str(text).strip()
    cands = [s]
    m = FENCE_RE.search(s)
    if m:
        cands.append(m.group(1))
    i, j = s.find("{"), s.rfind("}")
    if 0 <= i < j:
        cands.append(s[i:j + 1])
    err = None
    for c in cands:
        try:
            obj = json.loads(c)
        except Exception as e:
            err = err or short(e, 200)
            continue
        if isinstance(obj, dict):
            return obj, None
        err = err or "the reply is JSON but not an object"
    return None, err


def _strs(v, split=None):
    """A list of strings from a list (or, as Gemini sometimes replies, one string; `split` then cuts it up)."""
    if isinstance(v, str):
        v = re.split(split, v) if split else [v]
    elif isinstance(v, dict):                  # {Turn ID or label: a note}: the keys are what counts
        v = list(v.keys())
    return [re.sub(r"\s+", " ", str(x)).strip() for x in (v or []) if str(x).strip()]


def _text(v):
    """A text from a string or, as Gemini sometimes replies, a list of strings."""
    if isinstance(v, (list, tuple)):
        v = " ".join(str(x).strip() for x in v if str(x).strip())
    return re.sub(r"[ \t]+", " ", str(v or "")).strip()


PERIOD_SPLIT = r"\s*(?:,|;|/|\||\band\b|\bor\b)\s*"
ACT_SPLIT = r"\s*(?:;|\n)\s*"


def _ids(v):
    out = []
    for x in _strs(v):
        out += TID_RE.findall(x) or [x]
    return list(dict.fromkeys(out))


ENTRY_RE = re.compile(r"\bM\s*(\d+)\b", re.I)


def check_memo(obj, known, tids):
    """The update to the running memo in a reply: entries to add, and entries to revise or drop by their number
    (known: the numbers of the entries the call saw). Returns the update cleaned, and notes: a number that is not in
    the memo, or a new entry whose Turn ID is not a turn in scope. Nothing here asks a reply again."""
    m = obj.get("memo")
    if isinstance(m, (list, str)):             # entries to add only
        m = {"add": m}
    if not isinstance(m, dict):
        m = {"add": obj.get("memory") or []}   # the older field name
    notes, revise, drop = [], [], []
    for r in m.get("revise") or []:
        if isinstance(r, dict):
            k, tx = ENTRY_RE.search(str(r.get("entry") or "")), _text(r.get("text"))
        else:                                  # "M12: the revised entry"
            k, tx = ENTRY_RE.match(str(r).strip()), re.sub(r"^\s*M\s*\d+\s*[:.)-]?\s*", "", str(r)).strip()
        if k and tx:
            revise.append({"entry": f"M{int(k.group(1))}", "text": tx})
    for r in _strs(m.get("drop"), r"\s*[,;]\s*"):
        k = ENTRY_RE.search(r)
        if k:
            drop.append(f"M{int(k.group(1))}")
    asked = [x["entry"] for x in revise] + drop
    unknown = sorted({x for x in asked if known is not None and x not in known}, key=lambda x: int(x[1:]))
    if unknown:
        notes.append(f"memo: entries to revise or drop that are not in the memo: {', '.join(unknown[:8])}.")
    add = _strs(m.get("add"))
    stray = [i for a in add for i in TID_RE.findall(a.split(")", 1)[0]) if i not in tids]
    if stray:
        notes.append(f"memo: new entries cite Turn IDs that are not turns in scope: {', '.join(stray[:6])}.")
    keep = (lambda x: x in known) if known is not None else (lambda x: True)
    return {"add": add, "revise": [x for x in revise if keep(x["entry"])],
            "drop": list(dict.fromkeys(x for x in drop if keep(x)))}, notes


def check_theme(th, labels, tids, acts=None, need_acts=False, where=""):
    """One theme: its codes must be the agent's own labels, its Turn IDs in the acts, at least two codes, at most
    MAX_SUPPORT and MAX_CONTRARY Turn IDs. Returns the theme cleaned (labels as the agent wrote them in Step 3, Turn
    IDs that exist, within the limits) and its problems."""
    probs = []
    name = re.sub(r"\s+", " ", str(th.get("theme") or "")).strip()
    tag = f'{where}theme "{short(name, 60)}"' if name else f"{where}a theme without a name"
    if not name:
        probs.append(f"{tag}: give every theme a name.")
    codes, unknown = [], []
    for c in _strs(th.get("codes")):
        k = label_key(c)
        if k in labels:
            codes.append(labels[k])
        else:
            unknown.append(c)
    codes = list(dict.fromkeys(codes))
    if unknown:
        probs.append(f"{tag}: these codes are not among your labels{' in this act' if acts is None else ''}, written "
                     f"exactly: {'; '.join(short(u, 50) for u in unknown[:6])}.")
    if len(codes) < 2:
        probs.append(f"{tag}: a theme draws on at least two of your codes.")
    sup, con = _ids(th.get("supporting")), _ids(th.get("contrary"))
    bad = [i for i in sup + con if i not in tids]
    if bad:
        probs.append(f"{tag}: these Turn IDs are not turns of {'this act' if acts is None else 'the acts'}: "
                     f"{', '.join(bad[:6])}.")
    sup, con = [i for i in sup if i in tids], [i for i in con if i in tids]
    if not sup:
        probs.append(f"{tag}: give at least one Turn ID that supports it.")
    if len(sup) > MAX_SUPPORT or len(con) > MAX_CONTRARY:
        probs.append(f"{tag}: give at most {MAX_SUPPORT} supporting and {MAX_CONTRARY} contrary Turn IDs.")
    per = [p.lower().strip(" .") for p in _strs(th.get("periods"), PERIOD_SPLIT)]
    per = [p for p in per if p]
    if any(p not in PERIODS for p in per):
        probs.append(f"{tag}: periods are \"before\", \"around\" or \"after\".")
    out = {"theme": name, "definition": _text(th.get("definition")), "boundaries": _text(th.get("boundaries")),
           "codes": codes, "periods": [p for p in dict.fromkeys(per) if p in PERIODS], "supporting": sup[:MAX_SUPPORT],
           "contrary": con[:MAX_CONTRARY]}
    if not out["definition"]:
        probs.append(f"{tag}: give its definition.")
    if acts is not None:
        names = {x.casefold(): x for x in acts}
        alist = _strs(th.get("acts"), ACT_SPLIT)
        unk = [x for x in alist if x.casefold() not in names]
        if unk:
            probs.append(f"{tag}: these are not act names as given: {'; '.join(unk[:5])}.")
        out["acts"] = list(dict.fromkeys(names[x.casefold()] for x in alist if x.casefold() in names))
        if need_acts and not out["acts"]:
            probs.append(f"{tag}: list the acts it covers.")
        tl = _strs(th.get("too_little_evidence"), ACT_SPLIT)
        out["too_little_evidence"] = [names[x.casefold()] for x in tl if x.casefold() in names] + \
                                     [x for x in tl if x.casefold() not in names]
    return out, probs


def check_within(obj, labels, tids, known=None):
    if not isinstance(obj, dict) or not isinstance(obj.get("themes"), list):
        return None, ['The reply has no "themes" list.'], [], []
    soft, themes = [], []
    if not obj["themes"]:
        soft.append("Give the act's themes: the list is empty.")
    for th in obj["themes"]:
        if not isinstance(th, dict):
            soft.append("Each theme must be an object in the JSON format given above.")
            continue
        t, p = check_theme(th, labels, tids)
        themes.append(t)
        soft += p
    unplaced, unk = [], []
    for c in _strs(obj.get("unplaced_codes")):
        (unplaced if label_key(c) in labels else unk).append(labels.get(label_key(c), c))
    summary = _text(obj.get("summary"))
    n = nwords(summary)
    notes = []
    if not SUMMARY_ASK[0] <= n <= SUMMARY_ASK[1]:
        soft.append(f"The summary has {n} words; write it in about 300 words ({SUMMARY_WORDS[0]} to "
                    f"{SUMMARY_WORDS[1]}).")
    elif not SUMMARY_WORDS[0] <= n <= SUMMARY_WORDS[1]:
        notes.append(f"The summary has {n} words (about 300 were asked for).")
    used = {label_key(c) for t in themes for c in t["codes"]} | {label_key(c) for c in unplaced}
    left = [labels[k] for k in labels if k not in used]
    if unk:
        notes.append(f"Codes listed as fitting no theme that are not labels of this act: {'; '.join(unk[:5])}.")
    if len(left) > LEFT_OUT_ASK[1] and len(left) > LEFT_OUT_ASK[0] * len(labels):
        soft.append(f"{len(left)} of your {len(labels)} codes in this act are in no theme and not in "
                    f"\"unplaced_codes\". Put each of them in every theme whose turns it supports or, if it fits none "
                    f"of your themes, in \"unplaced_codes\", written exactly as here: " + "; ".join(left) + ".")
    elif left:
        notes.append(f"{len(left)} of the act's {len(labels)} labels are in no theme and not listed as fitting none, "
                     f"e.g. {'; '.join(short(x, 40) for x in left[:5])}.")
    memo, mnotes = check_memo(obj, known, tids)
    out = {"themes": themes, "unplaced_codes": list(dict.fromkeys(unplaced)), "summary": summary, "memo": memo,
           "labels_left_out": left}
    return out, [], soft, notes + mnotes


def check_groupings(obj, act_names, known=None, tids=()):
    if not isinstance(obj, dict) or not isinstance(obj.get("groupings"), list):
        return None, ['The reply has no "groupings" list.'], [], []
    soft, out = [], []
    names = {x.casefold(): x for x in act_names}
    gs = [g for g in obj["groupings"] if isinstance(g, dict)]
    if len(gs) < MIN_GROUPINGS:
        soft.append(f"Set up at least {MIN_GROUPINGS} groupings; the reply has {len(gs)}.")
    partitions = []
    for i, g in enumerate(gs, 1):
        gname = str(g.get("grouping") or f"grouping {i}").strip()
        tag = f'Grouping "{short(gname, 50)}"'
        basis = str(g.get("basis") or "").strip().lower()
        if basis not in ("material", "facts"):
            soft.append(f'{tag}: its basis is "material" or "facts".')
        groups, seen, part = [], collections.Counter(), []
        for grp in g.get("groups") or []:
            if not isinstance(grp, dict):
                continue
            acts_in = _strs(grp.get("acts"), ACT_SPLIT)
            unk = [x for x in acts_in if x.casefold() not in names]
            if unk:
                soft.append(f"{tag}: these are not act names as given: {'; '.join(unk[:5])}.")
            acts_ok = [names[x.casefold()] for x in acts_in if x.casefold() in names]
            seen.update(acts_ok)
            gn = str(grp.get("group") or "").strip()
            if len(acts_ok) < MIN_GROUP and "cannot tell" not in gn.lower():
                soft.append(f'{tag}: the group "{short(gn, 40)}" has {len(acts_ok)} act(s); each group holds at least '
                            f'{MIN_GROUP} (other than "cannot tell").')
            groups.append({"group": gn, "definition": _text(grp.get("definition")), "acts": acts_ok})
            part.append(frozenset(acts_ok))
        real = [x for x in groups if "cannot tell" not in x["group"].lower()]
        if len(real) < 2:
            soft.append(f"{tag}: sort the acts into at least two groups.")
        missing = [n for n in act_names if n not in seen]
        twice = [n for n, k in seen.items() if k > 1]
        if missing:
            soft.append(f"{tag}: every act goes into one group; missing: {'; '.join(missing[:6])}.")
        if twice:
            soft.append(f"{tag}: these acts are in more than one group: {'; '.join(twice[:6])}.")
        hit = OUTCOME_WORDS.search(gname + " " + _text(g.get("dimension")))   # the name and dimension only: a group's
        if hit:                                                                 # definition may mention context
            soft.append(f'{tag} seems to sort the acts on an outcome ("{hit.group(0)}"): do not sort on breakups, '
                        f'reunions, how long an act lasted or later success.')
        partitions.append(frozenset(part))
        out.append({"grouping": gname, "dimension": _text(g.get("dimension")), "basis": basis,
                    "why": _text(g.get("why")), "groups": groups})
    if gs and not any(g["basis"] == "material" for g in out):
        soft.append("At least one grouping must come from the material (basis \"material\").")
    for i in range(len(partitions)):
        for j in range(i + 1, len(partitions)):
            if partitions[i] == partitions[j]:
                soft.append(f"Groupings {i + 1} and {j + 1} sort the acts in the same way; the groupings should cut "
                            f"across one another.")
    memo, mnotes = check_memo(obj, known, set(tids))
    return {"groupings": out, "memo": memo}, [], soft, mnotes


COMPARISON_KINDS = ("within a group", "across groups", "alike pair", "different pair", "poor fit")


KIND_PATTERNS = (     # what the pair looks like decides its kind, not what the comparison looks for in it: an
    ("alike pair", r"\balike pairs?\b|\bpairs? (?:that|which) looks? alike\b|^alike\b"),   # alike pair is
    ("different pair", r"\bdifferent pairs?\b|\bpairs? (?:that|which) looks? different\b|^different\b"),  # searched
    ("poor fit", r"\bpoor fit\b|\bfits? (?:its|their) groups? badly\b|^bad(?:ly)? fit"),     # for a difference
    ("across groups", r"^across\b|\bacross (?:the )?groups\b|^between (?:the )?groups\b"),
    ("within a group", r"^within\b|\bwithin (?:a|one|the same) group\b"))


def comparison_kind(k):
    """A comparison's kind: one of COMPARISON_KINDS, or the reply's own wording if it is none of them."""
    k = _text(k).lower().strip(" .")
    if k in COMPARISON_KINDS:
        return k
    for want, pat in KIND_PATTERNS:
        if re.search(pat, k):
            return want
    return k or "not given"


THEME_ID_RE = re.compile(r"\b([WG]\d+\.\d+)\b")


def from_ids(entries, id_names):
    """Earlier themes given by their IDs (W3.2, G1.4), as "act: theme" or "grouping n: theme", which is how the
    script records them; an entry that names a theme instead is kept as it is. Returns the entries and the IDs that
    are not in the agent's columns."""
    if not id_names:
        return entries, []
    out, unknown = [], []
    for e in entries:
        found = THEME_ID_RE.findall(e)
        if not found:
            out.append(e)
        for x in found:
            out.append(id_names.get(x, x))
            if x not in id_names:
                unknown.append(x)
    return list(dict.fromkeys(out)), unknown


def check_column(obj, labels, tids, act_names, final=False, all_labels=None, known=None, id_names=None):
    """Round 3 (a grouping's column) and round 4 (the final themes). In round 4, all_labels are the codes of the
    agent's four columns ({label_key: label}); those no final theme lists are recorded (not asked for: at 52 acts
    there can be thousands, more than a reply could list)."""
    if not isinstance(obj, dict) or not isinstance(obj.get("themes"), list):
        return None, ['The reply has no "themes" list (reply in the JSON format given above).'], [], []
    soft, themes, notes = [], [], []
    if not obj["themes"]:
        soft.append("The list of themes is empty.")
    for th in obj["themes"]:
        if not isinstance(th, dict):
            continue
        t, p = check_theme(th, labels, tids, act_names, need_acts=True)
        key = "comes_from" if final else "builds_on"
        t[key], unknown = from_ids(_strs(th.get(key), ACT_SPLIT), id_names)
        if unknown:
            notes.append(f'theme "{short(t["theme"], 60)}" gives IDs that are not in your columns: '
                         f'{", ".join(unknown[:6])}.')
        if len(t["codes"]) > MAX_THEME_CODES:
            notes.append(f'theme "{short(t["theme"], 60)}" lists {len(t["codes"])} codes; the prompt asks for at most '
                         f'{MAX_THEME_CODES}.')
        themes.append(t)
        soft += p
    memo, mnotes = check_memo(obj, known, tids)
    notes += mnotes
    out = {"themes": themes, "memo": memo}
    if final:
        revs = [r for r in (obj.get("revisions") or []) if isinstance(r, dict)]
        if not revs:
            soft.append("Record your decisions on the earlier themes in \"revisions\".")
        out["revisions"] = [{"change": _text(r.get("change")), "themes": from_ids(_strs(r.get("themes")), id_names)[0],
                             "result": _text(r.get("result")), "why": _text(r.get("why"))} for r in revs]
        used = {label_key(c) for t in themes for c in t["codes"]}
        cols = all_labels if all_labels is not None else {}
        out["column_codes_not_listed"] = [cols[k] for k in cols if k not in used]
    else:
        comps = [c for c in (obj.get("comparisons") or []) if isinstance(c, dict)]
        out["comparisons"] = [{"kind": comparison_kind(c.get("kind")), "compared": _text(c.get("compared")),
                               "why": _text(c.get("why")), "found": _text(c.get("found")),
                               "changed_or_open": _text(c.get("changed_or_open"))} for c in comps]
        kinds = collections.Counter(c["kind"] for c in out["comparisons"])
        if kinds["alike pair"] < MIN_PAIRS or kinds["different pair"] < MIN_PAIRS:
            soft.append(f'Record at least {MIN_PAIRS} pairs of acts that look alike (kind "alike pair") and '
                        f'{MIN_PAIRS} that look different (kind "different pair"); the reply has {kinds["alike pair"]} '
                        f'and {kinds["different pair"]}.')
    return out, [], soft, notes


def _obj(props, required=None):
    return {"type": "object", "additionalProperties": False, "required": required or list(props), "properties": props}


_S = {"type": "string"}
_LIST = {"type": "array", "items": _S}
_PERIODS = {"type": "array", "items": {"type": "string", "enum": list(PERIODS)}}
SCHEMA_TEST = _obj({"memory": _LIST})
_MEMO = _obj({"add": _LIST, "revise": {"type": "array", "items": _obj({"entry": _S, "text": _S})}, "drop": _LIST})
SCHEMA1 = _obj({"themes": {"type": "array", "items": _obj({
    "theme": _S, "definition": _S, "boundaries": _S, "codes": _LIST, "periods": _PERIODS, "supporting": _LIST,
    "contrary": _LIST})}, "unplaced_codes": _LIST, "summary": _S, "memo": _MEMO})
SCHEMA2 = _obj({"groupings": {"type": "array", "items": _obj({
    "grouping": _S, "dimension": _S, "basis": {"type": "string", "enum": ["material", "facts"]}, "why": _S,
    "groups": {"type": "array", "items": _obj({"group": _S, "definition": _S, "acts": _LIST})}})},
    "memo": _MEMO})

_SIZES = {}


def _size(text):
    """est_tokens, kept for texts seen before (the fixed part, the memory blocks, full chunks)."""
    k = hashlib.sha1(text.encode("utf-8")).hexdigest()
    if k not in _SIZES:
        _SIZES[k] = est_tokens(text)
    return _SIZES[k]


def stable_chunks(blocks, size=None, cap=None):
    """The memory blocks (one per act), joined in order into chunks of at least `size` tokens; the last chunk is
    the one still filling. The text is the same; only the cut points differ. As the memory only grows at the end, a
    full chunk is the same in every later call, and so is the place of the breakpoint after it. If there would be
    more than `cap` chunks, the size doubles: the cut points change once, and the memory is cached again once."""
    size, cap = size or GPT_CHUNK_TOKENS, cap or GPT_MAX_CHUNKS
    sizes = [_size(b) for b in blocks]
    while True:
        chunks, cur, n = [], [], 0
        for b, s in zip(blocks, sizes):
            cur.append(b)
            n += s
            if n >= size:
                chunks.append("".join(cur))
                cur, n = [], 0
        if cur:
            chunks.append("".join(cur))
        if len(chunks) <= cap:
            return chunks
        size *= 2


# ==================================================================================================== the three models


def _status(e):
    for a in ("status_code", "code", "status"):
        v = getattr(e, a, None)
        try:
            return int(v)
        except (TypeError, ValueError):
            continue
    return None


def is_fatal(e):
    return _status(e) in (400, 401, 403, 404, 413, 422) or isinstance(e, (TypeError, UnicodeError, AgentStop))


OUTAGE_NAMES = re.compile(r"Connection|Timeout|Overloaded|RateLimit|InternalServer|ServiceUnavailable|ServerError|"
                          r"Protocol|ReadError|WriteError|NetworkError|Unavailable|DeadlineExceeded")


def is_outage(e):
    """A network or server error such as a provider's outage causes, worth waiting through for TRANSIENT_RETRIES;
    any other error that is not fatal is tried again only OTHER_RETRIES times, as waiting will not cure it."""
    if _status(e) in (408, 409, 425, 429, 500, 502, 503, 504, 520, 522, 524, 529):
        return True
    if isinstance(e, (ConnectionError, TimeoutError)):
        return True
    if isinstance(e, RuntimeError) and re.search(r"stream ended|failed on OpenAI's side", str(e)):
        return True
    return any(OUTAGE_NAMES.search(c.__name__) for c in type(e).__mro__)


def price(model, day=None):
    """The prices that apply on a day, today unless given (a model's prices can change on a set date, as Gemini's do
    on 1 December)."""
    day = day or datetime.date.today().isoformat()
    p = dict(PRICES.get(model) or {})
    for d, newp in sorted((p.pop("from", None) or {}).items()):
        if day >= d:
            p.update(newp)
    return p


def price_changes(model):
    """The dates after today on which a model's prices change."""
    return sorted(d for d in (PRICES.get(model) or {}).get("from", {}) if d > datetime.date.today().isoformat())


def nice_date(d):
    dt = datetime.date.fromisoformat(d)
    return f"{dt.day} {dt:%B %Y}"


def cost_of(model, u, batch=False):
    """The cost of one response, from its usage; a response from a batch costs half."""
    p = price(model)
    if not p:
        return 0.0
    if "cache_write_1h" in p:                          # Claude
        c = (u["input"] * p["input"] + u["cache_read"] * p["cache_read"] + u["cache_write_5m"] * p["cache_write_5m"]
             + u["cache_write_1h"] * p["cache_write_1h"] + u["output"] * p["output"])
    elif "cache_write" in p:                           # GPT: the long-input rates apply to each response on its own
        long = u["input_total"] > p["long_input"]
        xi, xo = (p["long_input_x"], p["long_output_x"]) if long else (1.0, 1.0)
        c = (u["input"] * p["input"] + u["cache_read"] * p["cache_read"] + u["cache_write"] * p["cache_write"]) * xi \
            + u["output"] * p["output"] * xo
    else:                                              # Gemini
        c = u["input"] * p["input"] + u["cache_read"] * p["cache_read"] + u["output"] * p["output"]
    return c / 1e6 * (0.5 if batch else 1.0)


def empty_usage():
    return {"input_total": 0, "input": 0, "cache_read": 0, "cache_write": 0, "cache_write_5m": 0,
            "cache_write_1h": 0, "output": 0, "thinking": 0}


def sum_usage(us):
    tot = empty_usage()
    for u in us:
        for k in tot:
            tot[k] += u.get(k, 0) or 0
    return tot


def own_tokens(letter, n):
    """Roughly how many of its own tokens a model counts for n tokens counted here (o200k_base): in the pilot,
    Claude's tokenizer counted about 1.45 times as many, Gemini's and GPT's about 1.05 times."""
    return int(n * (1.45 if MODELS[letter]["provider"] == "anthropic" else 1.05))


def _int(x):
    try:
        return int(x or 0)
    except (TypeError, ValueError):
        return 0


def full_text(parts):
    """The whole prompt of a call, as one text (for the providers that take it so, and for counting)."""
    return (parts["fixed"] + "".join(parts.get("memory") or []) + parts.get("labels", "") + parts.get("act_memory", "")
            + parts.get("tail", ""))


TOOL_DOCS = {
    "list_files": "List the files in your folder, with the number of lines in each.",
    "read_file": f"Read a file in your folder from start_line (1 for the start of the file): whole lines, up to "
                 f"{READ_LINES} at a time. The reply says which lines it shows, how many the file has, and where to "
                 f"read on. One call can read about {READ_BUDGET:,} characters in all, in up to {MAX_TOOL_STEPS} "
                 f"uses of the tools, so choose what to read.",
}

# Each provider's json_call and tool_call add the usage of every response to the list `usages` as it arrives, so that
# responses already paid for are counted even if a later step fails. They return the reply's text, whether it is
# complete (ok), why the model stopped, and the model version.


class ClaudeAgent:
    def __init__(self, cfg, letter):
        import anthropic
        self.cfg, self.letter = cfg, letter
        self.client = anthropic.Anthropic(api_key=secret(SECRETS["anthropic"]), max_retries=6, timeout=1800)

    def _usage(self, msg):
        u = msg.usage
        cc = getattr(u, "cache_creation", None)
        w1 = _int(getattr(cc, "ephemeral_1h_input_tokens", 0)) if cc else 0
        w5 = _int(getattr(cc, "ephemeral_5m_input_tokens", 0)) if cc else 0
        cw = _int(u.cache_creation_input_tokens)
        if cc is None:
            w5 = cw
        cr, inp = _int(u.cache_read_input_tokens), _int(u.input_tokens)
        return {"input_total": inp + cr + cw, "input": inp, "cache_read": cr, "cache_write": cw, "cache_write_5m": w5,
                "cache_write_1h": w1, "output": _int(u.output_tokens), "thinking": 0}

    def _content(self, parts, cache):
        """The fixed part, the memory (a block per act), the labels list as it was at the act's start, then what
        changes: the act's notes and the batch. Cache breakpoints, each kept for an hour (four at most): the fixed
        part, the last two memory blocks (where the last act's cache entry ends, and this act's) and the labels."""
        content = [{"type": "text", "text": parts["fixed"]}]
        if cache:
            content[0]["cache_control"] = {"type": "ephemeral", "ttl": "1h"}
        mb = parts["memory"]
        for i, b in enumerate(mb):
            blk = {"type": "text", "text": b}
            if cache and i >= len(mb) - 2:
                blk["cache_control"] = {"type": "ephemeral", "ttl": "1h"}
            content.append(blk)
        if parts.get("labels", "").strip():
            content.append({"type": "text", "text": parts["labels"]})
            if cache:
                content[-1]["cache_control"] = {"type": "ephemeral", "ttl": "1h"}
        for k in ("act_memory", "tail"):
            if parts.get(k, "").strip():
                content.append({"type": "text", "text": parts[k]})
        return content

    def _stream(self, **kw):
        with self.client.messages.stream(model=self.cfg["model"], max_tokens=self.cfg["max_output"], **kw) as s:
            return s.get_final_message()

    def json_call(self, parts, schema, cache, usages=None):
        usages = [] if usages is None else usages
        oc = {"format": {"type": "json_schema", "schema": schema}}
        if self.cfg.get("effort"):
            oc["effort"] = self.cfg["effort"]
        msg = self._stream(messages=[{"role": "user", "content": self._content(parts, cache)}], output_config=oc)
        usages.append(self._usage(msg))
        text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        return {"text": text, "usages": usages, "stop": msg.stop_reason,
                "ok": msg.stop_reason in ("end_turn", "stop_sequence"), "model": getattr(msg, "model", "")}

    def tool_call(self, parts, tools, usages=None):
        usages = [] if usages is None else usages
        tdefs = [{"name": "list_files", "description": TOOL_DOCS["list_files"],
                  "input_schema": {"type": "object", "properties": {}}},
                 {"name": "read_file", "description": TOOL_DOCS["read_file"],
                  "input_schema": {"type": "object", "properties": {"name": {"type": "string"},
                                                                    "start_line": {"type": "integer"}},
                                   "required": ["name"]}}]
        # one breakpoint only, set by the API at the end of each request (top-level cache_control): a block with its
        # own one-hour breakpoint as well would clash with it, which the API refuses
        msgs = [{"role": "user", "content": self._content(parts, False)}]
        kw = {"output_config": {"effort": self.cfg["effort"]}} if self.cfg.get("effort") else {}
        steps = stopping = 0
        while True:
            msg = self._stream(messages=msgs, tools=tdefs, cache_control={"type": "ephemeral"}, **kw)
            usages.append(self._usage(msg))
            calls = [b for b in msg.content if getattr(b, "type", "") == "tool_use"]
            if msg.stop_reason != "tool_use" or not calls:
                text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
                return {"text": text, "usages": usages, "stop": msg.stop_reason, "tool_steps": steps,
                        "ok": msg.stop_reason in ("end_turn", "stop_sequence"), "model": getattr(msg, "model", "")}
            if STOP.is_set():                          # after Stop, the tools tell the agent to write its reply
                stopping += 1                          # now, so that the call ends with a reply that is kept; it
                if stopping > 3:                       # is cut off only if the agent keeps on reading
                    raise stop_reason()
            steps += 1
            if steps > MAX_TOOL_STEPS + 10:            # past the limit, every tool use is told to write now
                raise AgentStop("the call kept using its folder after the reading limit")
            msgs.append({"role": "assistant", "content": msg.content})
            msgs.append({"role": "user", "content": [
                {"type": "tool_result", "tool_use_id": b.id, "content": tools(b.name, dict(b.input or {}))}
                for b in calls]})


class GPTAgent:
    def __init__(self, cfg, letter):
        import openai
        self.cfg, self.letter, self.openai, self.streaming = cfg, letter, openai, True
        self.client = openai.OpenAI(api_key=secret(SECRETS["openai"]), max_retries=6, timeout=1800)

    def _usage(self, r):
        u = getattr(r, "usage", None)
        if u is None:
            return empty_usage()
        d = getattr(u, "input_tokens_details", None)
        total = _int(u.input_tokens)
        cr = _int(getattr(d, "cached_tokens", 0)) if d else 0
        cw = _int(getattr(d, "cache_write_tokens", 0)) if d else 0
        od = getattr(u, "output_tokens_details", None)
        return {"input_total": total, "input": max(total - cr - cw, 0), "cache_read": cr, "cache_write": cw,
                "cache_write_5m": 0, "cache_write_1h": 0, "output": _int(u.output_tokens),
                "thinking": _int(getattr(od, "reasoning_tokens", 0)) if od else 0}

    def _stream(self, kw):
        """Streams the response, so that a long call never sits silent on the connection, and returns the final
        response, complete or stopped early. If OpenAI asks the organisation to verify before it streams this model,
        the calls go on without streaming. A response that failed on OpenAI's side is raised as an error, so that it is
        tried again after a wait, like any other server error."""
        final = None
        if self.streaming:
            try:
                with self.client.responses.stream(**kw) as s:
                    for ev in s:
                        if getattr(ev, "type", "") in ("response.completed", "response.incomplete", "response.failed"):
                            final = ev.response
                if final is None:
                    raise RuntimeError("the response stream ended without a final response")
            except self.openai.BadRequestError as e:
                if not re.search(r"verif", str(e), re.I):
                    raise
                self.streaming = False
                say(f"[{self.letter}] OpenAI asks for organisation verification to stream {self.cfg['model']}; "
                    f"carrying on without streaming.")
        if not self.streaming:
            final = self.client.responses.create(**kw)
        if getattr(final, "status", "") == "failed":
            raise RuntimeError(f"the response failed on OpenAI's side: {short(getattr(final, 'error', None), 200)}")
        return final

    def _result(self, r, usages, steps=0):
        inc = getattr(r, "incomplete_details", None)
        status = getattr(r, "status", "") or ""
        return {"text": r.output_text or "", "usages": usages, "ok": status == "completed", "tool_steps": steps,
                "stop": status + (f" ({getattr(inc, 'reason', '')})" if inc else ""), "model": getattr(r, "model", "")}

    def json_call(self, parts, schema, cache, usages=None):
        """With cache, a breakpoint after the fixed part and after every chunk of the memory (stable_chunks). GPT
        reuses a cached prefix only when the breakpoints before it are where they were (in the pilot, breakpoints
        that moved on with each batch meant that the memory was written to the cache again on every call), so the
        breakpoints stay put: a chunk, once full, never changes, and each call adds to the last chunk or starts a
        new one. Each call writes at most the last chunk or two to the cache and reads the rest."""
        usages = [] if usages is None else usages
        content = [{"type": "input_text", "text": parts["fixed"]}]
        mb = stable_chunks(parts["memory"]) if cache else parts["memory"]
        for b in mb:
            content.append({"type": "input_text", "text": b})
        n_cached = len(content)
        if parts.get("labels", "").strip():           # the labels list as it was at the act's start: cached too
            content.append({"type": "input_text", "text": parts["labels"]})
            n_cached += 1
        for k in ("act_memory", "tail"):
            if parts.get(k, "").strip():
                content.append({"type": "input_text", "text": parts[k]})
        if cache:
            for i in range(n_cached):
                content[i]["prompt_cache_breakpoint"] = {"mode": "explicit"}
        kw = dict(model=self.cfg["model"], input=[{"role": "user", "content": content}],
                  max_output_tokens=self.cfg["max_output"], store=False,
                  prompt_cache_options={"mode": "explicit"}, prompt_cache_key=f"step5-{RUN}-{self.letter}",
                  text={"format": {"type": "json_schema", "name": "reply", "schema": schema, "strict": True}})
        if self.cfg.get("effort"):
            kw["reasoning"] = {"effort": self.cfg["effort"]}
        r = self._stream(kw)
        usages.append(self._usage(r))
        return self._result(r, usages)

    def tool_call(self, parts, tools, usages=None):
        usages = [] if usages is None else usages
        tdefs = [{"type": "function", "name": "list_files", "description": TOOL_DOCS["list_files"], "strict": True,
                  "parameters": {"type": "object", "properties": {}, "required": [], "additionalProperties": False}},
                 {"type": "function", "name": "read_file", "description": TOOL_DOCS["read_file"], "strict": True,
                  "parameters": {"type": "object", "properties": {"name": {"type": "string"},
                                                                  "start_line": {"type": "integer"}},
                                 "required": ["name", "start_line"], "additionalProperties": False}}]
        text = full_text(parts)
        base = dict(model=self.cfg["model"], max_output_tokens=self.cfg["max_output"], tools=tdefs, store=True,
                    prompt_cache_key=f"step5-{RUN}-{self.letter}")
        if self.cfg.get("effort"):
            base["reasoning"] = {"effort": self.cfg["effort"]}
        r = self._stream(dict(base, input=[{"role": "user", "content": [{"type": "input_text", "text": text}]}]))
        usages.append(self._usage(r))
        steps = stopping = 0
        while True:
            calls = [o for o in (r.output or []) if getattr(o, "type", "") == "function_call"]
            if not calls or getattr(r, "status", "") != "completed":
                return self._result(r, usages, steps)
            if STOP.is_set():                          # after Stop, the tools tell the agent to write its reply
                stopping += 1                          # now, so that the call ends with a reply that is kept; it
                if stopping > 3:                       # is cut off only if the agent keeps on reading
                    raise stop_reason()
            steps += 1
            if steps > MAX_TOOL_STEPS + 10:            # past the limit, every tool use is told to write now
                raise AgentStop("the call kept using its folder after the reading limit")
            outs = []
            for c in calls:
                try:
                    args = json.loads(c.arguments or "{}")
                except Exception:
                    args = {}
                outs.append({"type": "function_call_output", "call_id": c.call_id, "output": tools(c.name, args)})
            r = self._stream(dict(base, previous_response_id=r.id, input=outs))
            usages.append(self._usage(r))


class GeminiAgent:
    def __init__(self, cfg, letter):
        from google import genai
        from google.genai import types
        self.cfg, self.letter, self.types, self.json_mode = cfg, letter, types, True
        self.client = genai.Client(api_key=secret(SECRETS["google"]), http_options=types.HttpOptions(
            timeout=30 * 60 * 1000, retry_options=types.HttpRetryOptions(attempts=6)))

    def _usage(self, um):
        if um is None:
            return empty_usage()
        total, cr = _int(um.prompt_token_count), _int(um.cached_content_token_count)
        th = _int(um.thoughts_token_count)
        return {"input_total": total, "input": max(total - cr, 0), "cache_read": cr, "cache_write": 0,
                "cache_write_5m": 0, "cache_write_1h": 0, "output": _int(um.candidates_token_count) + th,
                "thinking": th}

    def _config(self, **kw):
        t = self.types
        if self.cfg.get("effort"):
            kw["thinking_config"] = t.ThinkingConfig(thinking_level=self.cfg["effort"])
        return t.GenerateContentConfig(**kw)

    def json_call(self, parts, schema, cache, usages=None):
        usages = [] if usages is None else usages
        text = full_text(parts)
        for attempt in (1, 2):
            kw = {"max_output_tokens": self.cfg["max_output"]}
            if self.json_mode:
                kw["response_mime_type"] = "application/json"
            try:
                chunks, um, finish, model = [], None, None, ""
                for ch in self.client.models.generate_content_stream(model=self.cfg["model"], contents=text,
                                                                     config=self._config(**kw)):
                    if ch.candidates and ch.candidates[0].content and ch.candidates[0].content.parts:
                        chunks += [p.text for p in ch.candidates[0].content.parts
                                   if getattr(p, "text", None) and not getattr(p, "thought", False)]
                    if ch.candidates and ch.candidates[0].finish_reason:
                        finish = ch.candidates[0].finish_reason
                    um = ch.usage_metadata or um
                    model = getattr(ch, "model_version", None) or model
                usages.append(self._usage(um))
                fin = getattr(finish, "name", str(finish))
                return {"text": "".join(chunks), "usages": usages, "stop": fin, "ok": fin == "STOP", "model": model}
            except Exception as e:
                if attempt == 1 and self.json_mode and _status(e) == 400 and re.search(r"mime|json", str(e), re.I):
                    self.json_mode = False
                    say(f"[{self.letter}] JSON mode refused by {self.cfg['model']}; carrying on without it.")
                    continue
                raise

    def tool_call(self, parts, tools, usages=None):
        usages = [] if usages is None else usages
        t = self.types
        decls = [t.FunctionDeclaration(name="list_files", description=TOOL_DOCS["list_files"],
                                       parameters_json_schema={"type": "object", "properties": {}}),
                 t.FunctionDeclaration(name="read_file", description=TOOL_DOCS["read_file"],
                                       parameters_json_schema={"type": "object", "properties": {
                                           "name": {"type": "string"}, "start_line": {"type": "integer"}},
                                           "required": ["name"]})]
        # one request without streaming returns at most 32,768 tokens; past that, the SDK continues by itself up to
        # max_output_tokens (automatic continuation), so a long memo still comes back whole
        config = self._config(max_output_tokens=self.cfg["max_output"], tools=[t.Tool(function_declarations=decls)],
                              automatic_function_calling=t.AutomaticFunctionCallingConfig(disable=True))
        text = full_text(parts)
        contents = [t.Content(role="user", parts=[t.Part(text=text)])]
        steps = stopping = 0
        while True:
            r = self.client.models.generate_content(model=self.cfg["model"], contents=contents, config=config)
            usages.append(self._usage(r.usage_metadata))
            calls = r.function_calls or []
            cand = r.candidates[0] if r.candidates else None
            if not calls:
                fin = getattr(getattr(cand, "finish_reason", None), "name", "")
                body = "".join(p.text for p in (cand.content.parts if cand and cand.content and cand.content.parts
                                                else []) if getattr(p, "text", None) and not getattr(p, "thought", False))
                return {"text": body, "usages": usages, "stop": fin, "ok": fin == "STOP", "tool_steps": steps,
                        "model": getattr(r, "model_version", "")}
            if STOP.is_set():                          # after Stop, the tools tell the agent to write its reply
                stopping += 1                          # now, so that the call ends with a reply that is kept; it
                if stopping > 3:                       # is cut off only if the agent keeps on reading
                    raise stop_reason()
            steps += 1
            if steps > MAX_TOOL_STEPS + 10:            # past the limit, every tool use is told to write now
                raise AgentStop("the call kept using its folder after the reading limit")
            contents.append(cand.content)
            contents.append(t.Content(role="user", parts=[
                t.Part(function_response=t.FunctionResponse(id=c.id, name=c.name,
                                                            response={"result": tools(c.name, dict(c.args or {}))}))
                for c in calls]))


PROVIDERS = {"anthropic": ClaudeAgent, "openai": GPTAgent, "google": GeminiAgent}
_CLIENTS = {}


def provider(letter):
    if letter not in _CLIENTS:
        cfg = MODELS[letter]
        _CLIENTS[letter] = PROVIDERS[cfg["provider"]](cfg, letter)
    return _CLIENTS[letter]


def _book(st, kind_, parts, rec, usages, est, secs, batch=False):
    """Adds a call's cost to the agent's, saves the state, and keeps the record of the call: in calls/, one file per
    round (round 2: one per act), with the changing part of the prompt and the reply; the fixed part in prompts/.
    batch_api is True for a reply that came back from a provider's batch, at half price (round 2's "batch" is the
    number of the batch of turns)."""
    letter = st["agent"]
    model = MODELS[letter]["model"]
    c = sum(cost_of(model, u, batch) for u in usages)
    with COST_LOCK:
        st["cost"] = round(st["cost"] + c, 6)
        st["calls"] += 1
        COSTS[letter] = st["cost"]
        call_id = f"{letter}-{st['calls']:05d}"
    save_state(st)
    rec = dict(rec, call=call_id, agent=letter, kind=kind_, time=now(), seconds=secs, script=VERSION, model=model,
               effort=MODELS[letter].get("effort"), tokens_counted_before=est, usage=sum_usage(usages),
               last_output=(usages[-1].get("output", 0) if usages else 0),     # the reply itself (thinking included),
               responses=len(usages), cost=round(c, 6), batch_api=batch,     # which is what max_output limits
               fixed_sha=sha(parts["fixed"]), tail=parts["tail"], memory_blocks=len(parts["memory"]))
    append_jsonl(rpath(f"agent_{letter}", "calls", f"{rec.get('log', kind_)}.jsonl"), rec)
    fixed_path = rpath(f"agent_{letter}", "prompts", f"fixed_{rec['fixed_sha'][:12]}.txt")
    if not os.path.exists(fixed_path):
        write_text(fixed_path, parts["fixed"])
    return call_id, c


def call_model(st, kind_, parts, schema=None, tools=None, cache=False, meta=None):
    """One model call: the budget check, waits and retries after network or server errors (which Stop cuts short),
    the cost, the log line and the record of the call. Returns the result."""
    letter, meta = st["agent"], dict(meta or {})
    prov = provider(letter)
    if total_cost() >= BUDGET_USD:
        STOP_WHY[0] = f"the budget of ${BUDGET_USD:,.0f} is spent"
        STOP.set()                             # stops the other agents too
        raise stop_reason()
    est = est_tokens(full_text(parts))
    wait = FIRST_WAIT
    for attempt in range(TRANSIENT_RETRIES + 1):
        if STOP.is_set():
            raise stop_reason()
        t0, usages = time.time(), []
        try:
            res = prov.tool_call(parts, tools, usages) if tools else prov.json_call(parts, schema, cache, usages)
            break
        except Exception as e:
            if usages:                 # responses received before the failure (in a memo's tool loop) are paid for
                _book(st, kind_, parts, dict(meta, ok=False, error=short(e, 400)), usages, est,
                      round(time.time() - t0, 1))
            if isinstance(e, AgentStop):
                raise
            if is_fatal(e) or attempt >= (TRANSIENT_RETRIES if is_outage(e) else OTHER_RETRIES):
                raise AgentStop(f"{MODELS[letter]['model']} returned an error: {short(e, 400)}")
            say(f"[{letter}] {type(e).__name__}: {short(e, 160)}; trying again in {wait} s")
            if STOP.wait(wait):
                raise stop_reason()
            wait = min(wait * 2, MAX_WAIT)
    secs = round(time.time() - t0, 1)
    u = sum_usage(res["usages"])
    miss = False
    if cache and MODELS[letter]["provider"] != "google":   # Gemini caches by itself, when it can
        with COST_LOCK:
            CACHED_CALLS[letter] = CACHED_CALLS.get(letter, 0) + 1
            first = CACHED_CALLS[letter] == 1
        # what the call should have read from the cache at least: the fixed part and the memory, but for its newest
        # part (Claude: the last batch's block; GPT: the chunk still filling, and one that may have just filled)
        mem = parts["memory"]
        older = stable_chunks(mem)[:-2] if MODELS[letter]["provider"] == "openai" else mem[:-1]
        miss = not first and u["cache_read"] < 0.8 * (_size(parts["fixed"]) + sum(_size(b) for b in older))
    call_id, c = _book(st, kind_, parts, dict(meta, model_version=res.get("model", ""), stop=res["stop"],
                                               ok=res["ok"], cache_miss=miss, tool_steps=res.get("tool_steps", 0),
                                               reply=res["text"]), res["usages"], est, secs)
    res.update(call_id=call_id, cost=c, usage=u)
    n = len(res["usages"])
    say(f"[{letter}] {meta.get('label', kind_)}: ~{est:,} tokens in" + (f", {n} responses (file reads)" if n > 1 else "")
        + f", {u['output']:,} out, {u['cache_read']:,} from the cache, {secs:.0f} s, ${c:.3f} (agent "
          f"${st['cost']:.2f}, all ${total_cost():.2f})" + ("  CACHE MISS" if miss else ""))
    return res


BRIEFLY = {   # what a reply cut off at the model's output limit is asked to shorten, by round
    "within": "one sentence for each definition and boundary, and the summary in about 300 words",
    "groupings": "one sentence for each definition and reason",
    "comparisons": "a few sentences for each comparison, and one sentence for each definition and boundary",
    "final": ("one sentence for each definition, boundary and reason, and one revision record for all the themes that "
              "share a decision"),
}


def cut_off(res, round_name):
    """The problem for a reply that was cut off or stopped early. One cut off at the model's output limit is asked for
    again more briefly (asking for the same reply again would cut it off again)."""
    stop = str(res.get("stop") or "")
    if re.search(r"max_?(?:output_)?tokens", stop, re.I):
        return (f"The reply was cut off at the model's output limit ({stop}): it was longer than you can write in one "
                f"reply. Write it again with every part the prompt asks for, but more briefly: "
                f"{BRIEFLY.get(round_name, 'within the lengths the prompt gives')}.")
    return f"The reply was cut off or stopped early ({stop})."


PREVIOUS_CHARS = 120000                # the most of a previous reply shown again when it is asked for again


def problem_note(problems, unusable, tail_hint, previous=None):
    """What a re-ask adds to the prompt: the problems and, for a reply that could be used, the reply itself, so that
    the agent corrects it rather than starting again (and, in rounds 3 to 8, need not read its folder again)."""
    head = ("Your previous reply to this prompt could not be used, for these reasons:" if unusable
            else "Your previous reply to this prompt had these problems:")
    lst = "\n".join(f"{i}. {p}" for i, p in enumerate(problems[:20], 1))
    more = f"\n(and {len(problems) - 20} more)" if len(problems) > 20 else ""
    prev = ""
    if previous and previous.strip() and not unusable:
        p = previous.strip()
        if len(p) > PREVIOUS_CHARS:
            p = p[:PREVIOUS_CHARS] + "\n[... the rest of the reply is not shown]"
        prev = f"\n\nYour previous reply was:\n<previous_reply>\n{p}\n</previous_reply>"
    return f"\n\n{head}\n{lst}{more}{prev}\n\nReply again in full, {tail_hint}, correcting them."


def ask_until_usable(st, kind_, parts, schema, check, meta, hint, soft_retries=None, tools=None):
    """Asks, checks, and asks again: up to RETRIES_UNUSABLE times when the reply cannot be used, and RETRIES_PROBLEMS
    times (or soft_retries) when it can be used but has problems. Returns the best usable reply (what to keep,
    unusable, problems, notes, result), or stops the agent. After Stop, a usable reply already received is kept
    rather than lost."""
    tail0, tries = parts["tail"], []
    left_u, left_p = RETRIES_UNUSABLE, (RETRIES_PROBLEMS if soft_retries is None else soft_retries)
    while not STOP.is_set():
        try:
            res = call_model(st, kind_, parts, schema=schema, tools=tools, cache=meta.get("cache", False),
                             meta=dict(meta, attempt=len(tries) + 1))
        except AgentStop:
            if STOP.is_set() and any(not x[1] for x in tries):
                break
            raise
        obj, err = parse_json(res["text"])
        if not res["ok"]:
            out, hard, soft, notes = None, [cut_off(res, meta.get("round", kind_))], [], []
        elif obj is None:
            out, hard, soft, notes = None, [f"The reply is not valid JSON ({err})."], [], []
        else:
            out, hard, soft, notes = check(obj)
        tries.append((out, hard, soft, notes, res))
        if hard:
            say(f"[{st['agent']}] {meta['label']}: reply not usable: {short(hard[0], 200)}")
        elif soft:
            say(f"[{st['agent']}] {meta['label']}: {len(soft)} problem(s): {short(soft[0], 200)}")
        if not hard and not soft:
            break
        if (hard and left_u == 0) or (not hard and left_p == 0):
            break
        if hard:
            left_u -= 1
        else:
            left_p -= 1
        parts = dict(parts, tail=tail0 + problem_note(hard or soft, bool(hard), hint, None if hard else res["text"]))
    usable = [x for x in tries if not x[1]]
    if not usable:
        if STOP.is_set():
            raise stop_reason()
        raise AgentStop(f"{meta['label']}: no usable reply after {len(tries)} tries; last problem: "
                        f"{short(tries[-1][1][0], 300)}. Run the round again.")
    return min(usable, key=lambda x: (len(x[2]), len((x[0] or {}).get("labels_left_out") or [])))


def run_agents(fn, letters, *args):
    """Runs fn(letter, *args) for the agents at the same time. Stop (in Colab) lets the calls in progress finish."""
    left = busy()
    if left:
        say(f"Not started: agents from the last run are still finishing their calls ({', '.join(left)}). Wait a "
            f"minute and run this again; status() shows where each agent is.")
        return {}
    STOP_WHY[0] = "stopped by hand"
    STOP.clear()
    ensure(rpath())
    letters, results = [L for L in (letters or LETTERS)], {}
    states = {}
    for L in LETTERS:                          # counts that a stale state.json would get wrong, from the logs
        st = load_state(L)
        if os.path.isdir(rpath(f"agent_{L}")) and sync_counters(st):
            save_state(st)
        states[L] = st
    with COST_LOCK:
        COSTS.clear()
        CACHED_CALLS.clear()
        for L in LETTERS:
            COSTS[L] = states[L]["cost"]
    for L in list(letters):                    # keys and clients first, here in the main thread (Colab Secrets)
        try:
            provider(L)
        except Exception as e:
            say(f"[{L}] NOT STARTED: {e}")
            results[L] = f"not started: {e}"
            letters.remove(L)
    if not letters:
        return results
    ex = ThreadPoolExecutor(max_workers=len(letters), thread_name_prefix=THREAD_PREFIX)
    futs = {ex.submit(_guard, fn, L, (BASE, RUN), *args): L for L in letters}
    try:
        for f in as_completed(futs):
            results[futs[f]] = f.result()
    except KeyboardInterrupt:
        STOP.set()
        say("Stopping: the calls in progress will finish and be saved first. Please wait ...")
        for f in futs:
            try:
                results[futs[f]] = f.result()
            except Exception:
                pass
    finally:
        ex.shutdown(wait=True)
    return results


def _guard(fn, letter, where, *args):
    _LOCAL.base, _LOCAL.run = where
    try:
        return _guarded(fn, letter, *args)
    finally:
        _LOCAL.base = _LOCAL.run = None


def _guarded(fn, letter, *args):
    try:
        r = fn(letter, *args)
        return r if isinstance(r, str) else "done"     # a round may say what it is waiting for instead
    except AgentStop as e:
        st = load_state(letter)
        st["stopped"] = f"{now()}: {e}"
        save_state(st)
        say(f"[{letter}] STOPPED: {e}")
        return f"stopped: {e}"
    except Exception as e:
        tb = traceback.format_exc()
        append_jsonl(rpath(f"agent_{letter}", "errors.jsonl"), {"time": now(), "error": repr(e), "traceback": tb})
        st = load_state(letter)
        st["stopped"] = f"{now()}: error {e!r}"
        save_state(st)
        say(f"[{letter}] ERROR: {e!r} (the traceback is in agent_{letter}/errors.jsonl)")
        return f"error: {e!r}"


def folder_tools(folder):
    """The agent's folder tools for one call. Past READ_BUDGET characters or MAX_TOOL_STEPS uses, every further use
    is answered with a note to write the reply now, so that a call never stops half-way for reading too much."""
    files = sorted(f for f in os.listdir(folder) if not f.startswith("."))
    cache, used = {}, {"uses": 0, "chars": 0}

    def lines_of(name):
        if name not in cache:
            with open(os.path.join(folder, name), encoding="utf-8-sig") as fh:
                cache[name] = fh.read().splitlines()
        return cache[name]

    def run(name, args):
        used["uses"] += 1
        if STOP.is_set():
            return ("The run is being stopped. Do not use the tools again: write your reply now, from what you have "
                    "read and your memory.")
        if used["uses"] > MAX_TOOL_STEPS or used["chars"] >= READ_BUDGET:
            return (f"You have reached this call's reading limit (about {READ_BUDGET:,} characters, or "
                    f"{MAX_TOOL_STEPS} uses of the tools). Do not use the tools again: write your reply now, from "
                    f"what you have read and your memory.")
        out = _run(name, args)
        used["chars"] += len(out)
        return out

    def _run(name, args):
        try:
            if name == "list_files":
                return "\n".join(f"{f} ({len(lines_of(f))} lines)" for f in files)
            if name == "read_file":
                fname = str(args.get("name") or "").strip().split("/")[-1]
                if fname not in files:
                    return f"There is no file called {fname!r} in your folder. Use list_files to see the files."
                ls = lines_of(fname)
                s = max(1, _int(args.get("start_line") or 1))
                if s > len(ls):
                    return f"{fname} has {len(ls)} lines; there is nothing from line {s}."
                out, size, e = [], 0, s - 1
                for i in range(s - 1, min(len(ls), s - 1 + READ_LINES)):
                    line = ls[i]
                    if out and size + len(line) + 1 > READ_CHARS:
                        break
                    if len(line) > READ_CHARS:
                        line = line[:READ_CHARS] + " … [this line is longer; it is cut here]"
                    out.append(line)
                    size += len(line) + 1
                    e = i + 1
                more = f"\n(Read on from line {e + 1}.)" if e < len(ls) else "\n(End of the file.)"
                return f"{fname}, lines {s} to {e} of {len(ls)}:\n" + "\n".join(out) + more
            return f"Unknown tool {name!r}."
        except Exception as ex:
            return f"The file could not be read: {ex}"
    return run


def part_heading(n):
    """A line that starts with the heading Part n (with # or ** around it, or not)."""
    return re.compile(rf"^\s*(?:#+\s*)?\**\s*Part\s+{n}\b", re.I | re.M)


PART_RE = {n: part_heading(n) for n in (1, 2, 3)}
PART1_LINE = re.compile(r"^\s*(?:#+\s*)?\**\s*Part\s+1\s*\**\s*:?\s*\**\s*$", re.I | re.M)   # a bare Part 1 heading


def ask_text(st, kind_, parts, tools, check, meta):
    """Memo rounds: asks and checks the headings, and asks once more if a check fails. Keeps the better complete
    reply; a reply that was cut off is never kept (the agent stops, and running the round again asks again)."""
    tail0, tries = parts["tail"], []
    for attempt in (1, 2):
        if STOP.is_set():
            break
        res = call_model(st, kind_, parts, tools=tools, meta=dict(meta, attempt=attempt))
        probs = check(res["text"]) if res["ok"] else [cut_off(res, meta.get("round", kind_))]
        tries.append((res, probs))
        if not probs:
            break
        say(f"[{st['agent']}] {meta['label']}: {probs[0]}")
        parts = dict(parts, tail=tail0 + problem_note(probs, not res["ok"], "as the prompt asks",
                                                      res["text"] if res["ok"] else None))
    good = [x for x in tries if x[0]["ok"] and x[0]["text"].strip()]
    if not good:
        raise AgentStop(f"{meta['label']}: no complete reply; " + (STOP_WHY[0] if STOP.is_set() else
                                                                   "run the round again"))
    return min(good, key=lambda x: len(x[1]))

# ==================================================================================================== Step 5's files


def within_path(letter, stem):
    return rpath(f"agent_{letter}", "within", f"{stem}.json")


def groupings_path(letter):
    return rpath(f"agent_{letter}", "groupings.json")


def comparison_path(letter, n):
    return rpath(f"agent_{letter}", f"comparison_{n}.json")


def final_path(letter):
    return rpath(f"agent_{letter}", "final.json")


def n_groupings(letter):
    """How many groupings the agent set up (0 before round 2): the agent decides, at least MIN_GROUPINGS."""
    g = read_json(groupings_path(letter))
    return len(g["groupings"]) if g else 0


def theme_text(t, with_acts=False, tid=None):
    lines = [f"Theme: {t['theme']}"] + ([f"ID: {tid}"] if tid else []) + [f"Definition: {t['definition']}"]
    if t.get("boundaries"):
        lines.append(f"Boundaries: {t['boundaries']}")
    lines.append("Codes: " + "; ".join(t["codes"]))
    if with_acts and t.get("acts"):
        lines.append("Acts: " + "; ".join(t["acts"]))
    if t.get("periods"):
        lines.append("Periods: " + ", ".join(t["periods"]))
    lines.append("Supporting Turn IDs: " + (", ".join(t["supporting"]) or "none"))
    if t.get("contrary"):
        lines.append("Contrary Turn IDs: " + ", ".join(t["contrary"]))
    if t.get("too_little_evidence"):
        lines.append("Too little evidence in: " + "; ".join(t["too_little_evidence"]))
    if t.get("builds_on"):
        lines.append("Builds on: " + "; ".join(t["builds_on"]))
    if t.get("comes_from"):
        lines.append("Comes from: " + "; ".join(t["comes_from"]))
    return "\n".join(lines)


def theme_ids(letter, acts):
    """Short IDs for an agent's earlier themes, as rounds 3 and 4 show them, so that a reply can name them briefly:
    W<i>.<k> for the k-th within-case theme of the i-th act of the run (in nomination order), G<n>.<k> for the k-th
    theme of grouping n. Returns {ID: the theme as "act: theme" or "grouping n: theme"}, which is how the script
    records it."""
    ids = {}
    for i, a in enumerate(acts, 1):
        w = read_json(within_path(letter, act_stem(a))) or {}
        for k, t in enumerate(w.get("themes", []), 1):
            ids[f"W{i}.{k}"] = f"{a['Act']}: {t['theme']}"
    for n in range(1, n_groupings(letter) + 1):
        c = read_json(comparison_path(letter, n)) or {}
        for k, t in enumerate(c.get("themes", []), 1):
            ids[f"G{n}.{k}"] = f"grouping {n}: {t['theme']}"
    return ids


def grouping_block(g, n):
    lines = [f"Grouping {n}: {g['grouping']}", f"Dimension: {g['dimension']}", f"Basis: {g['basis']}",
             f"Why: {g['why']}", "Groups:"]
    for grp in g["groups"]:
        lines.append(f"- {grp['group']}: {grp['definition']}\n  Acts: {'; '.join(grp['acts'])}")
    return "\n".join(lines)


# ==================================================================================================== the running memo


MEMO_LINE = re.compile(r"^(M\d+) \(([^)]*)\): (.*)$")
ADD_RE = re.compile(r"^\s*(?:M\d+\s*)?\(([^)]*)\)\s*:?\s*(.*)$", re.S)   # an entry the agent numbered itself is
                                                                         # read without its number


def memo_base(letter):
    """The running memo as Step 3 left it: the text of its memory, its numbered entries, and the next number."""
    s3 = step3(letter)
    entries = {}
    for line in s3["memory"].splitlines():
        m = MEMO_LINE.match(line.strip())
        if m:
            entries[m.group(1)] = {"id": m.group(1), "where": m.group(2), "text": m.group(3), "step": 3}
    nxt = max([s3.get("next") or 0] + [int(k[1:]) + 1 for k in entries] + [1])
    return s3["memory"], entries, nxt


def memo_updates(letter, acts, upto, n=None):
    """The changes the agent made to its running memo before the call at stage `upto` ("within", "groupings",
    "compare" with the grouping n, "final", "end"), in the order they are made: round 1 by act, in act order (the
    acts run side by side, so each act starts from the memo as Step 3 left it), then the groupings, the comparisons
    one by one and the final themes."""
    if upto == "within":
        return []
    out = [(f"studying {a['Act']} on its own", (read_json(within_path(letter, act_stem(a))) or {}).get("memo") or {})
           for a in acts]
    if upto == "groupings":
        return out
    out.append(("setting up your groupings", (read_json(groupings_path(letter)) or {}).get("memo") or {}))
    last = (n - 1) if upto == "compare" else n_groupings(letter)
    for k in range(1, last + 1):
        out.append((f"comparing the acts in grouping {k}", (read_json(comparison_path(letter, k)) or {}).get("memo")
                    or {}))
    if upto in ("compare", "final"):
        return out
    out.append(("bringing your themes together", (read_json(final_path(letter)) or {}).get("memo") or {}))
    return out


def running_memo(letter, acts, upto, n=None):
    """The running memo as the call at stage `upto` sees it: Step 3's memory with the Step 5 changes made, in order.
    Returns its text, its entries (dropped ones marked), and the log of changes."""
    base, entries, nxt = memo_base(letter)
    if upto == "within":
        return base, entries, []
    entries = {k: dict(v) for k, v in entries.items()}
    added, log = [], []
    for where, upd in memo_updates(letter, acts, upto, n):
        for r in upd.get("revise", []):
            e = entries.get(r["entry"])
            if e and not e.get("dropped"):
                log.append({"where": where, "change": "revised", "entry": e["id"], "step": e["step"],
                            "before": f"({e['where']}): {e['text']}", "after": r["text"]})
                m = ADD_RE.match(r["text"])
                e["where"], e["text"] = (m.group(1), m.group(2).strip()) if m else (e["where"], r["text"])
                e["revised"] = where
        for k in upd.get("drop", []):
            e = entries.get(k)
            if e and not e.get("dropped"):
                e["dropped"] = where
                log.append({"where": where, "change": "dropped", "entry": k, "step": e["step"],
                            "before": f"({e['where']}): {e['text']}", "after": ""})
        for a in upd.get("add", []):
            m = ADD_RE.match(a)
            k, nxt = f"M{nxt}", nxt + 1
            entries[k] = {"id": k, "where": m.group(1).strip() if m else "general",
                          "text": (m.group(2) if m else a).strip(), "step": 5, "added": where}
            added.append(k)
            log.append({"where": where, "change": "added", "entry": k, "step": 5, "before": "",
                        "after": f"({entries[k]['where']}): {entries[k]['text']}"})
    lines = []                                 # Step 3's text, with its entries revised or dropped in Step 5
    for line in base.splitlines():
        m = MEMO_LINE.match(line.strip())
        e = entries.get(m.group(1)) if m else None
        if e and e.get("dropped"):
            continue
        if e and e.get("revised"):
            line = f"{e['id']} ({e['where']}): {e['text']} [revised in Step 5, {e['revised']}]"
        lines.append(line)
    groups = collections.OrderedDict()
    for k in added:
        e = entries[k]
        if not e.get("dropped"):
            groups.setdefault(e["added"], []).append(
                f"{k} ({e['where']}): {e['text']}" + (f" [revised, {e['revised']}]" if e.get("revised") else ""))
    text = "\n".join(lines)
    text += "\n\nYour entries from Step 5:\n" + ("\n\n".join(f"From {w}:\n" + "\n".join(v) for w, v in groups.items())
                                               if groups else "none yet.")
    return text, entries, log


def memory_for(letter, acts, stage, n=None):
    return running_memo(letter, acts, stage, n)[0]


def memo_known(letter, acts, stage, n=None):
    """The numbers of the entries the call at this stage sees, which its changes may refer to."""
    return {k for k, e in running_memo(letter, acts, stage, n)[1].items() if not e.get("dropped")}


def save_running_memo(letter, acts):
    """After round 4: the running memo as Step 5 leaves it (memory.txt, for Step 6), and every change (memo_changes.csv)."""
    text, entries, log = running_memo(letter, acts, "end")
    write_text(rpath(f"agent_{letter}", "memory.txt"), text + "\n")
    write_csv(rpath(f"agent_{letter}", "memo_changes.csv"), ["Call", "Change", "Entry", "From step", "Before", "After"],
              [[x["where"], x["change"], x["entry"], x["step"], x["before"], x["after"]] for x in log])
    return log


def act_counts(a):
    kept = load_corpus()["turns"][act_stem(a)]
    n_text = sum(1 for t in kept if t.get("Source") == "text")
    return n_text, len(kept) - n_text


def acts_block_groupings(letter, acts):
    out = []
    for a in acts:
        w = read_json(within_path(letter, act_stem(a)))
        n_text, n_video = act_counts(a)
        themes = "\n".join(f"- {t['theme']}: {t['definition']}" for t in w["themes"]) or "- none"
        out.append(f"Act: {a['Act']} ({kind(a)}); Best New Artist {'won' if won(a) else 'nominated'} at the "
                   f"{a['Nominated']} ceremony; turns from articles: {n_text}, from video: {n_video}\n"
                   f"Your summary: {w['summary']}\nYour themes for this act:\n{themes}")
    return "\n\n".join(out)


def acts_block_compare(letter, acts, g):
    group_of = {x: grp["group"] for grp in g["groups"] for x in grp["acts"]}
    out = []
    for i, a in enumerate(acts, 1):
        w = read_json(within_path(letter, act_stem(a)))
        themes = "\n\n".join(theme_text(t, tid=f"W{i}.{k}") for k, t in enumerate(w["themes"], 1)) or "none"
        out.append(f"Act: {a['Act']}\nGroup in this grouping: {group_of.get(a['Act'], 'none')}\n"
                   f"Your summary: {w['summary']}\nYour within-case themes for this act:\n\n{themes}")
    return "\n\n========\n\n".join(out)


def columns_block(letter, acts):
    parts = ["WITHIN-CASE THEMES, BY ACT"]
    for i, a in enumerate(acts, 1):
        w = read_json(within_path(letter, act_stem(a)))
        parts.append(f"== {a['Act']}\n\n" + ("\n\n".join(theme_text(t, tid=f"W{i}.{k}")
                                                         for k, t in enumerate(w["themes"], 1)) or "none"))
    g = read_json(groupings_path(letter))
    for n, grp in enumerate(g["groupings"], 1):
        c = read_json(comparison_path(letter, n))
        parts.append(f"GROUPING {n}: {grp['grouping']} (sorting on {grp['dimension']})\n\n"
                     + ("\n\n".join(theme_text(t, True, tid=f"G{n}.{k}") for k, t in enumerate(c["themes"], 1))
                        or "none"))
    return "\n\n".join(parts)


def final_block(letter):
    f = read_json(final_path(letter))
    revs = "\n".join(f"- {r['change']}: {'; '.join(r['themes'])} -> {r['result']} ({r['why']})" for r in f["revisions"])
    n = len(f["themes"])
    return (f"Your {n} final theme{'' if n == 1 else 's'}:\n\n" + "\n\n".join(theme_text(t, True) for t in f["themes"])
            + "\n\nYour revision records:\n" + (revs or "- none"))


def build_folder(letter, acts):
    """What the agent can open in rounds 3, 4 and 5: each act's chronology with its codes, its codes alone, its labels,
    and its Step 5 work so far."""
    folder = ensure(rpath(f"agent_{letter}", "folder"))
    labels = step3(letter)["labels"]
    uses, acts_of = collections.Counter(), collections.defaultdict(set)
    for a in acts:
        stem = act_stem(a)
        p = os.path.join(folder, f"chronology_{stem}.txt")
        if not os.path.exists(p):
            write_text(p, chronology(a, act_codes(letter, stem), 0) + "\n")
        src = s3path(letter, "codes", f"codes_{stem}.csv")
        if not os.path.exists(os.path.join(folder, f"codes_{stem}.csv")):
            shutil.copyfile(src, os.path.join(folder, f"codes_{stem}.csv"))
        for r in act_codes(letter, stem).values():
            for c in r["codes"]:
                uses[label_key(c["label"])] += 1
                acts_of[label_key(c["label"])].add(a["Act"])
    with open(os.path.join(folder, "labels_summary.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Label", "Data or own", "Times used", "Acts", "Definition"])
        for k, n in sorted(uses.items(), key=lambda kv: (-kv[1], kv[0])):
            x = labels.get(k, {})
            w.writerow([x.get("label", k), x.get("from", ""), n, "; ".join(sorted(acts_of[k])), x.get("definition", "")])
    ws = [(i, a, read_json(within_path(letter, act_stem(a)))) for i, a in enumerate(acts, 1)]
    write_text(os.path.join(folder, "within_case_themes.txt"), "\n\n".join(
        f"== {a['Act']}\n\n" + "\n\n".join(theme_text(t, tid=f"W{i}.{k}") for k, t in enumerate(x["themes"], 1))
        for i, a, x in ws if x) + "\n")
    write_text(os.path.join(folder, "summaries.txt"), "\n\n".join(
        f"== {a['Act']}\n{x['summary']}" for _, a, x in ws if x) + "\n")
    g = read_json(groupings_path(letter))
    if g:
        write_text(os.path.join(folder, "groupings.txt"),
                   "\n\n".join(grouping_block(x, i) for i, x in enumerate(g["groupings"], 1)) + "\n")
        for n in range(1, len(g["groupings"]) + 1):
            c = read_json(comparison_path(letter, n))
            if c:
                comps = "\n\n".join(f"Compared: {x['compared']}\nWhy: {x['why']}\nFound: {x['found']}\n"
                                    f"Changed or left open: {x['changed_or_open']}" for x in c["comparisons"])
                write_text(os.path.join(folder, f"grouping_{n}_comparisons_and_themes.txt"),
                           f"GROUPING {n}: {g['groupings'][n - 1]['grouping']}\n\nCOMPARISONS\n\n{comps}\n\nTHEMES\n\n"
                           + "\n\n".join(theme_text(t, True, tid=f"G{n}.{k}") for k, t in enumerate(c["themes"], 1))
                           + "\n")
    if os.path.exists(final_path(letter)):
        write_text(os.path.join(folder, "final_themes.txt"), final_block(letter) + "\n")
    return folder

# ==================================================================================================== round 0: timelines


def run_timelines(acts=None):
    """Step 4's timelines, for the acts: events and turns in date order, with periods and days since each event."""
    sel = select_acts(acts)
    load_events()
    no_events = [a["Act"] for a in sel if not act_events(a)]
    if no_events:
        print(f"PROBLEM: no events in {EVENTS_FILE} for: {', '.join(no_events)}")
    d = build_timelines(sel)
    say(f"Timelines written for {len(sel)} act(s) in {d}.")
    return d

# ==================================================================================================== round 1: within each case


def fits(text, letter):
    """Whether an agent's prompt fits its model's context window, with room for the reply."""
    return own_tokens(letter, est_tokens(text)) + OUTPUT_ROOM <= CONTEXT_ROOM * CONTEXT[MODELS[letter]["model"]]


def within_parts(letter, a, fixed, post):
    stem = act_stem(a)
    codes = act_codes(letter, stem)
    kept = load_corpus()["turns"][stem]
    for level in (0, 1, 2):
        vals = {"LETTER": letter, "ACT": a["Act"], "KIND": kind(a), "RESULT": "won" if won(a) else "nominated",
                "YEAR": a["Nominated"],
                "DATE": ceremony(a).isoformat(), "N_INTERVIEWS": len({t["Record ID"] for t in kept}),
                "N_TURNS": len(kept), "N_CODED": sum(1 for t in kept if (codes.get(t["Turn ID"]) or {}).get("codes")),
                "CHRONO": chronology(a, codes, level)}
        tail = fill(post, vals)
        if fits(fixed + tail, letter):
            return {"fixed": fixed, "memory": [], "tail": tail}, level
    raise AgentStop(f"{a['Act']}: the act does not fit the models' context windows even with its codes only")


def within_fixed(letter):
    s3 = step3(letter)
    pre, post = split_after(PROMPTS[1], "MEMORY")
    return fill(pre, {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"], "MEMORY": s3["memory"]}), post


def _within(letter, acts):
    st = load_state(letter)
    st["stopped"] = None
    save_state(st)
    fixed, post = within_fixed(letter)
    known = memo_known(letter, acts, "within")          # every act starts from the memo as Step 3 left it
    todo = [a for a in acts if not os.path.exists(within_path(letter, act_stem(a)))]
    if not todo:
        say(f"[{letter}] round 1 is already done for these acts.")
        return
    where = (_LOCAL.base, _LOCAL.run)

    def one(a):
        _LOCAL.base, _LOCAL.run = where
        if STOP.is_set():
            raise stop_reason()
        stem = act_stem(a)
        parts, level = within_parts(letter, a, fixed, post)
        labels = act_labels(letter, stem)
        tids = {t["Turn ID"] for t in load_corpus()["all_turns"][stem]}
        out, hard, soft, notes, res = ask_until_usable(
            st, "within", parts, SCHEMA1, lambda o: check_within(o, labels, tids, known),
            {"label": f"{a['Act']} (within)", "cache": True, "log": "within", "act": stem, "round": "within"},
            "in the JSON format given above")
        write_json(within_path(letter, stem), dict(out, act=a["Act"], stem=stem, level=level, call=res["call_id"],
                                                   problems=soft, notes=notes))
        flag(st, "within", stem, soft + notes, res["call_id"])
        save_state(st)
        say(f"[{letter}] {a['Act']}: {len(out['themes'])} themes, summary {nwords(out['summary'])} words"
            + (" (uncoded turns shown on one line)" if level == 1 else " (coded turns without text)" if level == 2
               else "") + (f"; {len(soft)} problem(s) flagged" if soft else ""))

    one(todo[0])                               # the first act writes the cache for the rest
    if len(todo) > 1:
        errors = []
        with ThreadPoolExecutor(max_workers=max(1, PARALLEL_ACTS), thread_name_prefix=f"{THREAD_PREFIX}-{letter}") as ex:
            futs = [ex.submit(one, a) for a in todo[1:]]
            for f in as_completed(futs):
                try:
                    f.result()
                except AgentStop as e:
                    errors.append(e)
        if errors:
            raise errors[0]
    say(f"[{letter}] round 1 done: {len(acts)} act(s).")


def run_within(acts=None, agents=None):
    """Round 1: each agent studies each act on its own (Prompt 1)."""
    load_corpus(), load_prompts(), load_guide(), load_events()
    sel = select_acts(acts)
    no_events = [a["Act"] for a in sel if not act_events(a)]
    if no_events:                              # every act has at least its nomination: a name the events file spells
        raise SystemExit(f"STOPPED: {EVENTS_FILE} has no events for {', '.join(no_events)}. Check that the act names "
                         f"match the corpus.")      # differently from the corpus
    check_settings("within", sel, agents)
    build_timelines(sel)
    return run_agents(_within, agents, sel)

# ==================================================================================================== round 2: groupings


def _groupings(letter, acts):
    st = load_state(letter)
    if os.path.exists(groupings_path(letter)):
        say(f"[{letter}] round 2 is already done.")
        return
    missing = [a["Act"] for a in acts if not os.path.exists(within_path(letter, act_stem(a)))]
    if missing:
        raise AgentStop(f"round 1 is not finished for {len(missing)} act(s), e.g. {missing[0]}")
    st["stopped"] = None
    s3 = step3(letter)
    text = fill(PROMPTS[2], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"],
                             "MEMORY": memory_for(letter, acts, "groupings"), "ACTS": acts_block_groupings(letter, acts)})
    names, known, tids = [a["Act"] for a in acts], memo_known(letter, acts, "groupings"), all_tids(acts)
    out, hard, soft, notes, res = ask_until_usable(
        st, "groupings", {"fixed": text, "memory": [], "tail": ""}, SCHEMA2,
        lambda o: check_groupings(o, names, known, tids),
        {"label": "round 2 (groupings)", "log": "groupings", "round": "groupings"}, "in the JSON format given above")
    write_json(groupings_path(letter), dict(out, call=res["call_id"], problems=soft, notes=notes))
    flag(st, "groupings", "", soft + notes, res["call_id"])
    save_state(st)
    say(f"[{letter}] round 2 done: " + "; ".join(f"{g['grouping']} ({g['basis']}, "
                                                    f"{len(g['groups'])} groups)" for g in out["groupings"]))


def run_groupings(acts=None, agents=None):
    """Round 2: each agent sets up three groupings of the acts (Prompt 2)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("groupings", sel, agents)
    return run_agents(_groupings, agents, sel)

# ==================================================================================================== round 3: comparisons


def all_tids(acts):
    return {t["Turn ID"] for a in acts for t in load_corpus()["all_turns"][act_stem(a)]}


def _comparisons(letter, acts):
    st = load_state(letter)
    g = read_json(groupings_path(letter))
    if not g:
        raise AgentStop("round 2 (groupings) is not done for this agent")
    if all(os.path.exists(comparison_path(letter, n)) for n in range(1, len(g["groupings"]) + 1)):
        say(f"[{letter}] round 3 is already done.")
        return
    st["stopped"] = None
    s3 = step3(letter)
    labels, tids, names = run_labels(letter, acts), all_tids(acts), [a["Act"] for a in acts]
    for n, grp in enumerate(g["groupings"], 1):
        if os.path.exists(comparison_path(letter, n)):
            continue
        if STOP.is_set():
            raise stop_reason()
        tools = folder_tools(build_folder(letter, acts))
        known = memo_known(letter, acts, "compare", n)
        id_names = {k: v for k, v in theme_ids(letter, acts).items() if k.startswith("W")}
        text = fill(PROMPTS[3], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"],
                                 "MEMORY": memory_for(letter, acts, "compare", n), "GNUM": n,
                                 "GTOTAL": len(g["groupings"]), "GNAME": grp["grouping"],
                                 "GDIM": grp["dimension"], "N_ACTS": len(acts), "GROUPING": grouping_block(grp, n),
                                 "ACTS": acts_block_compare(letter, acts, grp)})
        out, hard, soft, notes, res = ask_until_usable(
            st, "comparison", {"fixed": text, "memory": [], "tail": ""}, None,
            lambda o: check_column(o, labels, tids, names, known=known, id_names=id_names),
            {"label": f"round 3 (grouping {n})", "log": "comparisons", "round": "comparisons", "grouping": n},
            "in the JSON format given above", tools=tools)
        write_json(comparison_path(letter, n), dict(out, n=n, grouping=grp["grouping"], call=res["call_id"],
                                                    problems=soft, notes=notes))
        flag(st, f"comparison {n}", "", soft + notes, res["call_id"])
        save_state(st)
        say(f"[{letter}] grouping {n} ({grp['grouping']}): {len(out['comparisons'])} comparisons, "
            f"{len(out['themes'])} themes")
    say(f"[{letter}] round 3 done.")


def run_comparisons(acts=None, agents=None):
    """Round 3: each agent compares the acts within each of its groupings (Prompt 3)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("comparisons", sel, agents)
    return run_agents(_comparisons, agents, sel)

# ==================================================================================================== round 4: final themes


def _final(letter, acts):
    st = load_state(letter)
    if os.path.exists(final_path(letter)):
        say(f"[{letter}] round 4 is already done.")
        return
    k = n_groupings(letter)
    missing = [n for n in range(1, k + 1) if not os.path.exists(comparison_path(letter, n))]
    if not os.path.exists(groupings_path(letter)) or missing:
        raise AgentStop("round 3 (comparisons) is not finished for this agent")
    st["stopped"] = None
    s3 = step3(letter)
    labels, tids, names = run_labels(letter, acts), all_tids(acts), [a["Act"] for a in acts]
    in_columns = {}                            # the codes of the four columns: those no final theme lists are recorded
    for a in acts:
        for t in read_json(within_path(letter, act_stem(a)))["themes"]:
            in_columns.update({label_key(c): c for c in t["codes"]})
    for n in range(1, k + 1):
        for t in read_json(comparison_path(letter, n))["themes"]:
            in_columns.update({label_key(c): c for c in t["codes"]})
    tools = folder_tools(build_folder(letter, acts))
    known = memo_known(letter, acts, "final")
    text = fill(PROMPTS[4], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"],
                             "MEMORY": memory_for(letter, acts, "final"), "N_ACTS": len(acts),
                             "COLUMNS": columns_block(letter, acts)})
    out, hard, soft, notes, res = ask_until_usable(
        st, "final", {"fixed": text, "memory": [], "tail": ""}, None,
        lambda o: check_column(o, labels, tids, names, final=True, all_labels=in_columns, known=known,
                               id_names=theme_ids(letter, acts)),
        {"label": "round 4 (final themes)", "log": "final", "round": "final"}, "in the JSON format given above",
        tools=tools)
    write_json(final_path(letter), dict(out, call=res["call_id"], problems=soft, notes=notes))
    flag(st, "final", "", soft + notes, res["call_id"])
    save_state(st)
    log = save_running_memo(letter, acts)      # the running memo as Step 5 leaves it, for Step 6
    say(f"[{letter}] running memo: {sum(x['change'] == 'added' for x in log)} entries added, "
        f"{sum(x['change'] == 'revised' for x in log)} revised and {sum(x['change'] == 'dropped' for x in log)} "
        f"dropped in Step 5; saved in agent_{letter}/memory.txt for Step 6.")
    say(f"[{letter}] round 4 done: {len(out['themes'])} final themes, {len(out['revisions'])} revision records.")


def run_final(acts=None, agents=None):
    """Round 4: each agent brings its themes together into one list (Prompt 4)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("final", sel, agents)
    return run_agents(_final, agents, sel)

# ==================================================================================================== round 5: memo part 1


def theme_names(letters, acts):
    """Names a memo or the lead's comparison may quote without the quote being in a turn: the agents' themes (final,
    column and within-case), groupings, groups and dimensions."""
    out = []
    for L in letters:
        f = read_json(final_path(L)) or {}
        out += [t["theme"] for t in f.get("themes", [])]
        g = read_json(groupings_path(L)) or {}
        for x in g.get("groupings", []):
            out += [x["grouping"], x["dimension"]] + [grp["group"] for grp in x["groups"]]
        for n in range(1, n_groupings(L) + 1):
            c = read_json(comparison_path(L, n)) or {}
            out += [t["theme"] for t in c.get("themes", [])]
        for a in acts:
            w = read_json(within_path(L, act_stem(a))) or {}
            out += [t["theme"] for t in w.get("themes", [])]
    return [x for x in out if x]


def word_problem(text, limit, what):
    n = nwords(text)
    return [f"{what} has {n:,} words; keep it to about {limit:,}."] if n > limit * WORDS_SLACK else []


def _memo1(letter, acts):
    st = load_state(letter)
    if os.path.exists(rpath(f"agent_{letter}", "memo_part1.md")):
        say(f"[{letter}] round 5 is already done.")
        return
    if not os.path.exists(final_path(letter)):
        raise AgentStop("round 4 (final themes) is not done for this agent")
    st["stopped"] = None
    s3 = step3(letter)
    tools = folder_tools(build_folder(letter, acts))
    text = fill(PROMPTS[5], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"],
                             "MEMORY": memory_for(letter, acts, "end"), "FINAL": final_block(letter)})
    scope = sorted(all_tids(acts))
    ok = list(run_labels(letter, acts).values()) + theme_names([letter], acts)

    def check(t):
        first = next((ln for ln in t.splitlines() if ln.strip()), "")
        p = [] if PART_RE[1].match(first) else ['Part 1 must start with the heading "Part 1".']
        return p + word_problem(t, MEMO_WORDS[1], "Part 1") + quote_problems(t, scope, ok)
    res, probs = ask_text(st, "memo1", {"fixed": text, "memory": [], "tail": ""}, tools, check,
                          {"round": "memo1", "label": "round 5 (memo part 1)", "log": "memo1"})
    write_text(rpath(f"agent_{letter}", "memo_part1.md"), res["text"].strip() + "\n")
    write_text(rpath(f"agent_{letter}", f"memo_{letter}.md"), res["text"].strip() + "\n")
    flag(st, "memo1", "", probs, res["call_id"])
    save_state(st)
    say(f"[{letter}] round 5 done: memo part 1 has {nwords(res['text']):,} words, after "
        f"{res.get('tool_steps', 0)} file reads.")


def run_memo1(acts=None, agents=None):
    """Round 5: each agent writes part 1 of its Step 5 memo (Prompt 5)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("memo1", sel, agents)
    return run_agents(_memo1, agents, sel)

# ==================================================================================================== round 6: the merge


def _theme_turns(t, keys, lt, act_stem_of):
    """The turns behind a theme: the turns its agent coded with the codes it carries (label keys), in the acts it
    covers (or, if that leaves none, in any act)."""
    stems = {act_stem_of[x] for x in t.get("acts", []) if x in act_stem_of}
    turns = set()
    for k in keys:
        turns |= lt.get(k, set())
    inside = {x for x in turns if CORPUS["stem_of"].get(x) in stems} if stems else turns
    return inside or turns


def overlap_rows(acts):
    act_stem_of = {a["Act"]: act_stem(a) for a in acts}
    sets = {}
    for L in LETTERS:
        lt = label_turns(L, acts)
        for t, listed, carried, _ in final_bases(L, acts):
            sets[(L, t["theme"])] = _theme_turns(t, carried, lt, act_stem_of)
    rows = []
    for (L, name), s in sets.items():
        best = []
        for (M, other), s2 in sets.items():
            if M == L or not s or not s2:
                continue
            shared = len(s & s2)
            if shared:
                best.append((shared, round(shared / len(s | s2), 3), M, other, len(s2)))
        best.sort(key=lambda x: (-x[1], -x[0]))            # closest first: a large catch-all theme shares many turns
                                                           # with everything, so the ranking is by Jaccard; but Jaccard
                                                           # is low whenever two themes differ much in size, so each
                                                           # pair also shows the share of each theme's turns, and each
                                                           # theme the other theme that holds most of its turns
        most = max(best, key=lambda x: (x[0], -x[4]), default=None)
        rows.append({"agent": L, "theme": name, "turns": len(s), "closest": best[:3],
                     "most_within": most if most and most not in best[:3] else None,
                     "share_within": round(most[0] / len(s), 3) if most else 0.0})
    return rows


def overlap_pair(b, n):
    """One pair of themes, for overlap.txt: shared turns as a share of each theme's turns, and the Jaccard index."""
    return (f"Agent {b[2]}'s \"{b[3]}\" ({b[0]} shared: {b[0] / max(n, 1):.0%} of this theme's turns and "
            f"{b[0] / max(b[4], 1):.0%} of that theme's {b[4]} turns; Jaccard {b[1]:.2f})")


CELL_MAX = 32000


def _cell(v):
    if isinstance(v, (list, tuple)):
        v = "; ".join(str(x) for x in v)
    v = "" if v is None else v
    if isinstance(v, str) and len(v) > CELL_MAX:
        v = v[:CELL_MAX] + " … [cut]"
    return v


def write_xlsx(path, sheets):
    """sheets: [(name, header, rows, widths)]."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    wb.remove(wb.active)
    A, AB = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    WRAP, HEAD = Alignment(wrap_text=True, vertical="top"), PatternFill("solid", start_color="E7F1FB")
    for name, header, rows, widths in sheets:
        ws = wb.create_sheet(name[:31])
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill, c.alignment = AB, HEAD, WRAP
        for r in rows:
            ws.append([_cell(v) for v in r])
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font, c.alignment = A, WRAP
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = ws.dimensions
    ensure(os.path.dirname(path))
    wb.save(path)


def write_csv(path, header, rows):
    ensure(os.path.dirname(path))
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            w.writerow([_cell(v) for v in r])


def earlier_index(W_L, C_L):
    """An agent's earlier themes, for following a final theme's "comes from" back to codes: within-case themes by
    (act, theme), column themes by (grouping n, theme), and every theme by its name alone."""
    by_key, by_name = {}, collections.defaultdict(list)
    for act, w in W_L.items():
        for t in w["themes"]:
            by_key[(act_key(act), label_key(t["theme"]))] = t["codes"]
            by_name[label_key(t["theme"])].append(t["codes"])
    for n, c in C_L.items():
        for t in c["themes"]:
            by_key[(f"grouping {n}", label_key(t["theme"]))] = t["codes"]
            by_name[label_key(t["theme"])].append(t["codes"])
    return by_key, by_name


def trace(entry, by_key, by_name):
    """The codes of the earlier theme a "comes from" entry names ("act: theme", "grouping n: theme", or the theme's
    name alone if no other theme has it), or None."""
    m = re.match(r"\s*(.+?)\s*[:–—-]\s+(.+)$", entry)
    if m:
        g = re.match(r"(?i)grouping\s*(\d+)", m.group(1))
        k = (f"grouping {g.group(1)}" if g else act_key(m.group(1)), label_key(m.group(2)))
        if k in by_key:
            return by_key[k]
    for name in [entry] + ([m.group(2)] if m else []):
        hits = by_name.get(label_key(name)) or []
        if len(hits) == 1:
            return hits[0]
    return None


def final_bases(letter, acts, W_L=None, C_L=None):
    """For each of an agent's final themes: the theme, the codes it lists, the codes it carries, as label keys, and its
    "comes from" entries that name no earlier theme. A final theme carries the codes it lists and all the codes of
    each earlier theme that went into it alone. An earlier theme that several final themes come from (split between
    them, or shared) gives each of them only the codes they list, as which of its codes went where is not recorded;
    otherwise every final theme that names a broad earlier theme would take on all its turns."""
    if W_L is None:
        W_L = {a["Act"]: read_json(within_path(letter, act_stem(a))) for a in acts}
    if C_L is None:
        C_L = {n: read_json(comparison_path(letter, n)) for n in range(1, n_groupings(letter) + 1)}
    by_key, by_name = earlier_index({a: w for a, w in W_L.items() if w}, {n: c for n, c in C_L.items() if c})
    finals = (read_json(final_path(letter)) or {}).get("themes", [])
    traced = [[(e, trace(e, by_key, by_name)) for e in t.get("comes_from", [])] for t in finals]
    uses = collections.Counter()               # how many final themes come from each earlier theme (its codes list)
    for tr in traced:
        uses.update({id(codes) for _, codes in tr if codes is not None})
    out = []
    for t, tr in zip(finals, traced):
        listed = {label_key(x) for x in t["codes"]}
        carried, not_found = set(listed), []
        for e, codes in tr:
            if codes is None:
                not_found.append(e)
            elif uses[id(codes)] == 1:
                carried |= {label_key(x) for x in codes}
        out.append((t, listed, carried, not_found))
    return out


def run_merge(acts=None):
    """Round 6: the three agents' work in one workbook (merged/step5_combined.xlsx) and as CSV files in merged/folder,
    without changing or resolving anything; and where their final themes draw on the same turns (merged/overlap.txt)."""
    load_corpus()
    sel = select_acts(acts)
    missing = [L for L in LETTERS if not os.path.exists(final_path(L))]
    if missing:
        print(f"Not merged: round 4 (final themes) is not done for agent(s) {', '.join(missing)}.")
        return
    d, fd = ensure(rpath("merged")), ensure(rpath("merged", "folder"))
    W = {L: {a["Act"]: read_json(within_path(L, act_stem(a))) for a in sel} for L in LETTERS}
    G = {L: read_json(groupings_path(L)) for L in LETTERS}
    C = {L: {n: read_json(comparison_path(L, n)) for n in range(1, len(G[L]["groupings"]) + 1)} for L in LETTERS}
    Fn = {L: read_json(final_path(L)) for L in LETTERS}
    # codes by pass: one row per agent and label, one column per pass
    codes_rows, missing_of = [], {}
    for L in LETTERS:
        s3 = step3(L)
        uses, acts_of = collections.Counter(), collections.defaultdict(set)
        for a in sel:
            for r in act_codes(L, act_stem(a)).values():
                for c in r["codes"]:
                    uses[label_key(c["label"])] += 1
                    acts_of[label_key(c["label"])].add(a["Act"])
        within_of, unplaced_in = collections.defaultdict(list), collections.defaultdict(list)
        for act, w in W[L].items():
            for t in w["themes"]:
                for c in t["codes"]:
                    within_of[label_key(c)].append(f"{act}: {t['theme']}")
            for c in w.get("unplaced_codes", []):
                unplaced_in[label_key(c)].append(f"within ({act})")
        col_of = {n: collections.defaultdict(list) for n in C[L]}
        for n, c in C[L].items():
            for t in c["themes"]:
                for x in t["codes"]:
                    col_of[n][label_key(x)].append(t["theme"])
        final_of, via_of = collections.defaultdict(list), collections.defaultdict(list)
        for t, listed, carried, not_found in final_bases(L, sel, W[L], C[L]):
            for k in listed:
                final_of[k].append(t["theme"])
            for k in carried:
                via_of[k].append(t["theme"])
            missing_of[(L, t["theme"])] = not_found
        for k, n in sorted(uses.items(), key=lambda kv: (-kv[1], kv[0])):
            x = s3["labels"].get(k, {})
            codes_rows.append([L, x.get("label", k), x.get("from", ""), n, sorted(acts_of[k]), x.get("definition", ""),
                               within_of.get(k, [])] + [col_of[g].get(k, []) for g in sorted(col_of)]
                              + [final_of.get(k, []), via_of.get(k, []), unplaced_in.get(k, [])])
    ng = max(len(G[L]["groupings"]) for L in LETTERS)
    for r in codes_rows:                       # pad the grouping columns of an agent with fewer groupings
        while len(r) < 7 + ng + 3:
            r.insert(-3, [])
    codes_head = (["Agent", "Label", "Data or own", "Times used", "Acts", "Definition", "Within-case themes"]
                  + [f"Grouping {n} themes" for n in range(1, ng + 1)]
                  + ["Final themes (listed)", "Final themes (listed, or through an earlier theme that went into it alone)",
                     "Listed as fitting no within-case theme in"])
    final_rows, rev_rows, grp_rows, comp_rows, gth_rows, wth_rows = [], [], [], [], [], []
    for L in LETTERS:
        for t in Fn[L]["themes"]:
            final_rows.append([L, t["theme"], t["definition"], t["boundaries"], t["codes"], t.get("acts", []),
                               t["periods"], t["supporting"], t["contrary"], t.get("too_little_evidence", []),
                               t.get("comes_from", []), missing_of.get((L, t["theme"]), [])])
        for r in Fn[L]["revisions"]:
            rev_rows.append([L, r["change"], r["themes"], r["result"], r["why"]])
        for n, g in enumerate(G[L]["groupings"], 1):
            for grp in g["groups"]:
                grp_rows.append([L, n, g["grouping"], g["dimension"], g["basis"], g["why"], grp["group"],
                                 grp["definition"], grp["acts"]])
        for n, c in C[L].items():
            for x in c["comparisons"]:
                comp_rows.append([L, n, x.get("kind", ""), x["compared"], x["why"], x["found"], x["changed_or_open"]])
            for t in c["themes"]:
                gth_rows.append([L, n, t["theme"], t["definition"], t["boundaries"], t["codes"], t.get("acts", []),
                                 t["periods"], t["supporting"], t["contrary"], t.get("too_little_evidence", []),
                                 t.get("builds_on", [])])
        for act, w in W[L].items():
            for t in w["themes"]:
                wth_rows.append([L, act, t["theme"], t["definition"], t["boundaries"], t["codes"], t["periods"],
                                 t["supporting"], t["contrary"]])
    sum_rows = [[a["Act"]] + [W[L][a["Act"]]["summary"] for L in LETTERS] for a in sel]
    memo_rows = [[L, x["where"], x["change"], x["entry"], x["step"], x["before"], x["after"]]
                 for L in LETTERS for x in running_memo(L, sel, "end")[2]]
    ov = overlap_rows(sel)
    ov_rows = [[r["agent"], r["theme"], r["turns"], b[2], b[3], b[4], b[0], round(b[0] / max(r["turns"], 1), 3),
                round(b[0] / max(b[4], 1), 3), b[1]]
               for r in ov for b in r["closest"] + ([r["most_within"]] if r["most_within"] else [])] + \
              [[r["agent"], r["theme"], r["turns"], "", "(no other agent's theme shares a turn)", 0, 0, 0, 0, 0]
               for r in ov if not r["closest"]]
    th_head = ["Agent", "Theme", "Definition", "Boundaries", "Codes", "Acts", "Periods", "Supporting Turn IDs",
               "Contrary Turn IDs", "Too little evidence in"]
    sheets = [
        ("Codes by pass", codes_head, codes_rows, [7, 30, 9, 8, 30, 40, 40] + [30] * ng + [30, 30, 20]),
        ("Final themes", th_head + ["Comes from", "Comes from: earlier theme not found"], final_rows,
         [7, 30, 50, 40, 40, 30, 12, 24, 20, 20, 30, 24]),
        ("Revisions", ["Agent", "Change", "Themes", "Result", "Why"], rev_rows, [7, 12, 40, 40, 60]),
        ("Groupings", ["Agent", "Grouping", "Name", "Dimension", "Basis", "Why", "Group", "Definition", "Acts"],
         grp_rows, [7, 9, 26, 30, 9, 40, 24, 40, 40]),
        ("Comparisons", ["Agent", "Grouping", "Kind", "Compared", "Why", "Found", "Changed or left open"], comp_rows,
         [7, 9, 14, 36, 40, 50, 40]),
        ("Grouping themes", ["Agent", "Grouping"] + th_head[1:] + ["Builds on"], gth_rows,
         [7, 9, 30, 50, 40, 40, 30, 12, 24, 20, 20, 30]),
        ("Within-case themes", ["Agent", "Act", "Theme", "Definition", "Boundaries", "Codes", "Periods",
                                "Supporting Turn IDs", "Contrary Turn IDs"], wth_rows, [7, 22, 30, 50, 40, 40, 12, 24, 20]),
        ("Summaries", ["Act"] + [f"Agent {L}" for L in LETTERS], sum_rows, [22, 60, 60, 60]),
        ("Overlap", ["Agent", "Final theme", "Turns behind it", "Other agent", "Other agent's theme",
                     "Turns behind the other theme", "Shared turns", "Share of this theme's turns",
                     "Share of the other theme's turns", "Jaccard"], ov_rows, [7, 34, 9, 7, 34, 9, 9, 9, 9, 9]),
        ("Running memo", ["Agent", "Call", "Change", "Entry", "From step", "Before", "After"], memo_rows,
         [7, 34, 9, 8, 8, 60, 60]),
    ]
    write_xlsx(os.path.join(d, "step5_combined.xlsx"), sheets)
    for name, header, rows, _ in sheets:
        write_csv(os.path.join(fd, re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") + ".csv"), header, rows)
    def count(n, what):                        # the counts are given, so that no agent has to count them
        return f"{n} {what}{'' if n == 1 else 's'}"
    write_text(os.path.join(d, "groupings_all.txt"), "\n\n".join(
        f"AGENT {L} ({count(len(G[L]['groupings']), 'grouping')})\n\n"
        + "\n\n".join(grouping_block(g, n) for n, g in enumerate(G[L]["groupings"], 1))
        for L in LETTERS) + "\n")
    write_text(os.path.join(d, "final_themes_all.txt"), "\n\n".join(
        f"AGENT {L} ({count(len(Fn[L]['themes']), 'final theme')})\n\n"
        + "\n\n".join(theme_text(t, True) for t in Fn[L]["themes"]) for L in LETTERS) + "\n")
    lines = []
    for r in ov:
        near = "; ".join(overlap_pair(b, r["turns"]) for b in r["closest"])
        if r["most_within"]:
            near += "; most of its turns: " + overlap_pair(r["most_within"], r["turns"])
        lines.append(f"Agent {r['agent']}'s \"{r['theme']}\" ({r['turns']} turns behind it): "
                     + (near or "no other agent's theme shares a turn"))
    write_text(os.path.join(d, "overlap.txt"),
               "For each final theme: the turns behind it are the turns its agent coded, in the acts it covers, with "
               "the codes the theme lists and all the codes of each earlier theme that went into it alone (an earlier "
               "theme that several final themes come from, split between them or shared, gives each of them only the "
               "codes they list). Its closest themes from the other agents are those with the highest Jaccard "
               "index: the turns both draw on, as a share of the turns either draws on. Jaccard is low whenever two "
               "themes differ much in size, even when one holds most of the other's turns, so each pair also shows "
               "the shared turns as a share of each theme's turns; where another theme holds more of a theme's turns "
               "than its three closest, it is added as \"most of its turns\".\n\n"
               + "\n".join(lines) + "\n")
    write_json(os.path.join(d, "done.json"), {"acts": [act_stem(a) for a in sel], "time": now(), "script": VERSION})
    say(f"Merge done: {os.path.join(d, 'step5_combined.xlsx')} ({len(final_rows)} final themes from three agents).")

# ==================================================================================================== round 7: the lead's comparison


LEAD_HEAD = re.compile(r"^\s*(?:#+\s*)?\**\s*Comparison of the three agents", re.I)


def _lead(letter, acts):
    st = load_state(letter)
    out_path = rpath("merged", "comparison_by_lead.md")
    if os.path.exists(out_path):
        say(f"[{letter}] round 7 is already done.")
        return
    if not os.path.exists(rpath("merged", "done.json")):
        raise AgentStop("round 6 (the merge) has not been run")
    st["stopped"] = None
    tools = folder_tools(rpath("merged", "folder"))
    text = fill(PROMPTS[6], {"GUIDE": load_guide(), "GROUPINGS3": read_text(rpath("merged", "groupings_all.txt")),
                             "FINAL3": read_text(rpath("merged", "final_themes_all.txt")),
                             "OVERLAP": read_text(rpath("merged", "overlap.txt"))})
    scope = sorted(all_tids(acts))
    ok = [x for L in LETTERS for x in run_labels(L, acts).values()] + theme_names(LETTERS, acts)

    def check(t):
        first = next((ln for ln in t.splitlines() if ln.strip()), "")
        p = [] if LEAD_HEAD.match(first) else ["Start with the heading \"Comparison of the three agents' Step 5 themes\"."]
        return p + word_problem(t, LEAD_WORDS, "The comparison") + quote_problems(t, scope, ok)
    res, probs = ask_text(st, "lead", {"fixed": text, "memory": [], "tail": ""}, tools, check,
                          {"round": "lead", "label": "round 7 (lead's comparison)", "log": "lead"})
    write_text(out_path, res["text"].strip() + "\n")
    flag(st, "lead", "", probs, res["call_id"])
    save_state(st)
    say(f"[{letter}] round 7 done: the comparison has {nwords(res['text']):,} words.")


def run_lead(acts=None):
    """Round 7: the lead agent (A) compares the three agents' work (Prompt 6)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("lead", sel, [LEAD])
    return run_agents(_lead, [LEAD], sel)

# ==================================================================================================== round 8: memo parts 2 and 3


def _memo23(letter, acts):
    st = load_state(letter)
    if os.path.exists(rpath(f"agent_{letter}", "memo_parts2and3.md")):
        say(f"[{letter}] round 8 is already done.")
        return
    memo1 = read_text(rpath(f"agent_{letter}", "memo_part1.md"))
    comp = read_text(rpath("merged", "comparison_by_lead.md"))
    if not memo1:
        raise AgentStop("round 5 (memo part 1) is not done for this agent")
    if not comp:
        raise AgentStop("round 7 (the lead's comparison) is not done")
    st["stopped"] = None
    tools = folder_tools(rpath("merged", "folder"))
    text = fill(PROMPTS[7], {"GUIDE": load_guide(), "LETTER": letter, "MEMO1": memo1.strip(), "COMPARISON": comp.strip(),
                             "MEMORY": memory_for(letter, acts, "end"),
                             "FINAL3": read_text(rpath("merged", "final_themes_all.txt"))})
    scope = sorted(all_tids(acts))
    ok = [x for L in LETTERS for x in run_labels(L, acts).values()] + theme_names(LETTERS, acts)

    def check(t):
        p = []
        m2, m3 = PART_RE[2].search(t), PART_RE[3].search(t)
        if not m2 or not m3:
            p.append('Parts 2 and 3 must start with the headings "Part 2" and "Part 3".')
        if PART1_LINE.search(t):
            p.append("Return parts 2 and 3 only; part 1 stays as it is.")
        if m2 and m3 and m3.start() > m2.start():
            p += word_problem(t[m2.start():m3.start()], MEMO_WORDS[2], "Part 2")
            p += word_problem(t[m3.start():], MEMO_WORDS[3], "Part 3")
        return p + quote_problems(t, scope, ok)
    res, probs = ask_text(st, "memo23", {"fixed": text, "memory": [], "tail": ""}, tools, check,
                          {"round": "memo23", "label": "round 8 (memo parts 2 and 3)", "log": "memo23"})
    write_text(rpath(f"agent_{letter}", "memo_parts2and3.md"), res["text"].strip() + "\n")
    write_text(rpath(f"agent_{letter}", f"memo_{letter}.md"), memo1.strip() + "\n\n" + res["text"].strip() + "\n")
    flag(st, "memo23", "", probs, res["call_id"])
    save_state(st)
    say(f"[{letter}] round 8 done: parts 2 and 3 appended to memo_{letter}.md ({nwords(res['text']):,} words).")


def run_memo23(acts=None, agents=None):
    """Round 8: each agent reads the lead's comparison and writes parts 2 and 3 of its memo (Prompt 7)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts(acts)
    check_settings("memo23", sel, agents)
    return run_agents(_memo23, agents, sel)


def run_all(acts=None, agents=None):
    """Every round in order, stopping at the first round where an agent stops."""
    steps = [("round 1 (within)", lambda: run_within(acts, agents)),
             ("round 2 (groupings)", lambda: run_groupings(acts, agents)),
             ("round 3 (comparisons)", lambda: run_comparisons(acts, agents)),
             ("round 4 (final themes)", lambda: run_final(acts, agents)),
             ("round 5 (memo part 1)", lambda: run_memo1(acts, agents)),
             ("round 6 (merge)", lambda: run_merge(acts)),
             ("round 7 (lead's comparison)", lambda: run_lead(acts)),
             ("round 8 (memo parts 2 and 3)", lambda: run_memo23(acts, agents))]
    for name, fn in steps:
        res = fn()
        if name.startswith("round 6"):
            if not os.path.exists(rpath("merged", "done.json")):
                print(f"Stopped at {name}.")
                return
            continue
        if not res:
            print(f"Stopped at {name}: it did not start (see the message above).")
            return
        bad = {L: r for L, r in res.items() if r != "done"}
        if bad:
            print(f"Stopped at {name}: " + "; ".join(f"agent {L}: {r}" for L, r in bad.items()))
            return res
    print("All eight rounds are done. pilot_report() sums them up.")

# ==================================================================================================== checks, dry run, test, reports


def check_inputs():
    """Checks the corpus, Step 3's work, the events, the prompts, the guide, the keys and the SDK versions, without
    calling any model."""
    print(VERSION)
    for mod, want in (("anthropic", "2.3.0"), ("openai", "3.37.0"), ("google.genai", "2.33.0")):
        try:
            v = getattr(__import__(mod, fromlist=["_"]), "__version__", "?")
            print(f"{mod} {v}" + ("" if v == want else f"  (the script was tested with {want}: run the pip line in "
                                                         f"the notes at the top, then Runtime > Restart session)"))
        except Exception as e:
            print(f"{mod}: NOT INSTALLED ({e}); run the pip line in the notes at the top")
    CORPUS["acts"] = None
    c = load_corpus()
    print(f"Corpus: {len(c['acts'])} acts, {len(c['by_id']):,} turns; {len(c['excluded'])} turns left out by Step 3 "
          f"({EXCLUDE_FILE}).")
    sel = select_acts()
    EVENTS["by_act"] = None
    ev = load_events()
    print(f"Events ({EVENTS_FILE}): {sum(len(v) for v in ev.values())} events for {len(ev)} acts."
          + "".join(f"\n  PROBLEM: no events for {a['Act']}" for a in sel if not act_events(a)))
    for L in LETTERS:
        S3.pop(L, None)
        try:
            s3 = step3(L)
            miss = [a["Act"] for a in sel if not os.path.exists(s3path(L, "codes", f"codes_{act_stem(a)}.csv"))]
            print(f"Agent {L}'s Step 3 work: memo {nwords(s3['memo']):,} words, memory {nwords(s3['memory']):,} words, "
                  f"{len(s3['labels']):,} labels" + (f"; PROBLEM: no codes for {len(miss)} act(s), e.g. {miss[0]}"
                                                     if miss else "; codes for every act"))
        except AgentStop as e:
            print(f"Agent {L}'s Step 3 work: PROBLEM: {e}")
    PROMPTS.clear()
    for n in PROMPT_FILES:
        try:
            s = read_prompt(n)
            pr = check_prompt(n, s)
            print(f"Prompt {n} ({PROMPT_FILES[n]}): {nwords(s):,} words" + (f"; PROBLEMS: {pr}" if pr else "; fine"))
        except SystemExit as e:
            print(f"Prompt {n}: {e}")
    GUIDE["text"] = None
    g = load_guide()
    print(f"Guide ({GUIDE_FILE}): {nwords(g):,} words, about {est_tokens(g):,} tokens.")
    for L, cfg in MODELS.items():
        try:
            secret(SECRETS[cfg["provider"]])
            ok = "key found"
        except AgentStop as e:
            ok = f"PROBLEM: {e}"
        print(f"Agent {L}: {cfg['model']} ({cfg['provider']}), effort {cfg.get('effort')}: {ok}")
    print(f"Run {RUN!r} on {len(sel)} act(s) ({ACTS!r}); Step 3 run {STEP3_RUN!r}; the run folder is {rpath()}; "
          f"budget ${BUDGET_USD:,.0f}.")


def dry_run(acts=None):
    """Writes the filled prompts of Agent A to step5_<RUN>/dry_run, so you can read what each agent will get, and
    prints the size of each act's Prompt 1. No model is called; later prompts show where earlier rounds' work goes."""
    load_corpus(), load_prompts(), load_events()
    sel = select_acts(acts)
    d = ensure(rpath("dry_run"))
    fixed, post = within_fixed("A")
    print(f"Prompt 1: the part every act shares (guide, instructions, Step 3 memo and memory) is about "
          f"{est_tokens(fixed):,} tokens, cached after the first act.")
    sizes = []
    for a in sel:
        parts, level = within_parts("A", a, fixed, post)
        n = est_tokens(parts["tail"])
        sizes.append((n, a, level))
        print(f"  {a['Act']}: about {n:,} tokens of its own"
              + ("" if level == 0 else " (uncoded turns on one line)" if level == 1 else " (coded turns without text)"))
    for n, a, level in [sizes[0], max(sizes, key=lambda x: x[0])][:1 if len(sizes) == 1 else 2]:
        parts, _ = within_parts("A", a, fixed, post)
        write_text(os.path.join(d, f"prompt1_agent_A_{act_stem(a)}.txt"), full_text(parts))
    s3 = step3("A")
    stub = {"GUIDE": load_guide(), "LETTER": "A", "MEMO3": s3["memo"], "MEMORY": s3["memory"]
            + "\n\nYour entries from Step 5:\n[the running memo as the earlier calls left it: the entries they added, "
              "with their revisions, and without the entries they dropped]"}
    write_text(os.path.join(d, "prompt2_agent_A.txt"), fill(PROMPTS[2], dict(
        stub, ACTS="[for each act: its name, kind, result and year, turns from articles and video, the agent's "
                   "summary and its within-case themes]")))
    write_text(os.path.join(d, "prompt3_agent_A.txt"), fill(PROMPTS[3], dict(
        stub, GNUM=1, GTOTAL="[the number of groupings]", GNAME="[grouping name]", GDIM="[dimension]", N_ACTS=len(sel),
        GROUPING="[the grouping and its groups]", ACTS="[for each act: its group, summary and within-case themes]")))
    write_text(os.path.join(d, "prompt4_agent_A.txt"), fill(PROMPTS[4], dict(
        stub, N_ACTS=len(sel), COLUMNS="[the agent's columns of themes: within-case, then one for each grouping]")))
    write_text(os.path.join(d, "prompt5_agent_A.txt"), fill(PROMPTS[5], dict(stub, FINAL="[the final themes]")))
    write_text(os.path.join(d, "prompt6_agent_A.txt"), fill(PROMPTS[6], {
        "GUIDE": load_guide(), "GROUPINGS3": "[the three agents' groupings]", "FINAL3": "[the three agents' final themes]",
        "OVERLAP": "[the overlap figures]"}))
    write_text(os.path.join(d, "prompt7_agent_A.txt"), fill(PROMPTS[7], {
        "GUIDE": load_guide(), "LETTER": "A", "MEMO1": "[memo part 1]", "COMPARISON": "[the lead's comparison]",
        "MEMORY": stub["MEMORY"], "FINAL3": "[the three agents' final themes]"}))
    tot = sum(n for n, _, _ in sizes)
    print(f"The filled prompts are in {d}. Round 1 sends about {tot:,} tokens of acts per agent, plus the shared part "
          f"once per act (mostly from the cache).")


def test_models(agents=None):
    """A few small calls per model, with the same settings as the real calls: a JSON reply; the cache (two calls with
    the same long first part, the second reading it from the cache); and the folder tool. Under a dollar in all; these
    calls are not counted in the run's budget."""
    if busy():
        print("Agents from the last run are still finishing their calls; wait a minute and run this again.")
        return
    STOP_WHY[0] = "stopped by hand"
    STOP.clear()
    print(VERSION)
    g = load_guide()
    folder = ensure(rpath("test_folder"))
    with open(os.path.join(folder, "note.txt"), "w", encoding="utf-8") as f:
        f.write("The word to find is: lighthouse\n")
    for L in (agents or LETTERS):
        tools = folder_tools(folder)
        name = f"Agent {L} ({MODELS[L]['model']})"
        try:
            prov = provider(L)
            res = prov.json_call({"fixed": 'This is a test of the connection. Reply in this JSON format: {"memory": '
                                           '["connected"]}', "memory": [], "tail": ""}, SCHEMA_TEST, False)
            obj, _ = parse_json(res["text"])
            print(f"{name}: JSON reply {'OK' if obj and res['ok'] else 'PROBLEM: ' + short(res['text'], 80)}; "
                  f"model version {res.get('model')}")
            text = (g + "\n") * (1 + 40000 // max(len(g), 1))
            fixed = "This is a test of the cache. Here is a long text to cache:\n" + text[:40000]
            rs = [prov.json_call({"fixed": fixed, "memory": [], "tail": f'\nTest {k}. Reply in this JSON format: '
                                  f'{{"memory": ["cached {k}"]}}'}, SCHEMA_TEST, True) for k in (1, 2)]
            us = [sum_usage(r["usages"]) for r in rs]
            if not all(r["ok"] for r in rs):
                verdict = "PROBLEM (a reply stopped early)"
            elif us[1]["cache_read"] > 0:
                verdict = "OK"
            elif MODELS[L]["provider"] == "google":
                verdict = "no read yet, which is fine: Gemini caches by itself, as it sees fit"
            else:
                verdict = "PROBLEM: the second call did not read from the cache"
            print(f"{name}: cache: {verdict}. Read from the cache: "
                  + ", then ".join(f"{u['cache_read']:,} of {u['input_total']:,}" for u in us) + " tokens")
            r3 = prov.tool_call({"fixed": "This is a test of your folder tools. List the files in your folder, read "
                                          "note.txt, and reply with the one word it asks you to find, and nothing "
                                          "else.", "memory": [], "tail": ""}, tools)
            print(f"{name}: folder tool {'OK' if 'lighthouse' in r3['text'].lower() else 'PROBLEM'}: "
                  f"{r3.get('tool_steps', 0)} tool step(s); replied {short(r3['text'], 60)!r}")
            cst = sum(cost_of(MODELS[L]["model"], u) for r in [res] + rs + [r3] for u in r["usages"])
            print(f"{name}: these tests cost about ${cst:.3f}")
        except Exception as e:
            print(f"{name}: PROBLEM: {short(e, 400)}")


def status(acts=None):
    """Where each agent is: the rounds done, its cost and flags, and any stop."""
    load_corpus()
    sel = select_acts(acts)
    for L in LETTERS:
        st = load_state(L)
        nw = sum(os.path.exists(within_path(L, act_stem(a))) for a in sel)
        k = n_groupings(L)
        nc = sum(os.path.exists(comparison_path(L, n)) for n in range(1, k + 1))
        print(f"Agent {L} ({MODELS[L]['model']}): round 1 {nw}/{len(sel)} acts; groupings "
              f"{f'done ({k})' if k else 'not done'}; comparisons {f'{nc}/{k}' if k else 'not started'}; final themes "
              f"{'done' if os.path.exists(final_path(L)) else 'not done'}; memo part 1 "
              f"{'done' if os.path.exists(rpath(f'agent_{L}', 'memo_part1.md')) else 'not done'}, parts 2 and 3 "
              f"{'done' if os.path.exists(rpath(f'agent_{L}', 'memo_parts2and3.md')) else 'not done'}; "
              f"{st['calls']} calls, ${st['cost']:.2f}; {st['flags']} flags"
              + (f"; STOPPED {st['stopped']}" if st.get("stopped") else ""))
    print(f"Merge {'done' if os.path.exists(rpath('merged', 'done.json')) else 'not done'}; lead's comparison "
          f"{'done' if os.path.exists(rpath('merged', 'comparison_by_lead.md')) else 'not done'}.")
    left = busy()
    print(f"Total cost of run {RUN!r}: ${sum(load_state(L)['cost'] for L in LETTERS):.2f} of the "
          f"${BUDGET_USD:,.0f} budget." + (f" Still running: {', '.join(left)}." if left else ""))


def pilot_report(acts=None):
    """What the pilot produced, the checks, and the cost and hours of the full run, from the pilot's calls."""
    load_corpus(), load_prompts(), load_events()
    sel = select_acts(acts)
    full = load_corpus()["acts"]
    k = len(full) / max(len(sel), 1)
    lines = [f"# Step 5 pilot report ({RUN}, {len(sel)} acts: {', '.join(a['Act'] for a in sel)})", ""]
    proj_cost, proj_hours = 0.0, 0.0
    for L in LETTERS:
        calls = [c for c in logged_calls(L)]
        by = collections.defaultdict(list)
        for c in calls:
            by[c.get("kind")].append(c)
        lines += [f"## Agent {L} ({MODELS[L]['model']})", ""]
        ws = [read_json(within_path(L, act_stem(a))) for a in sel]
        ws = [w for w in ws if w]
        if ws:
            nth = [len(w["themes"]) for w in ws]
            cpt = [len(t["codes"]) for w in ws for t in w["themes"]]
            sw = [nwords(w["summary"]) for w in ws]
            left = [len(w.get("labels_left_out", [])) for w in ws]
            n_lab = n_in = n_none = 0              # the acts' labels: in a theme, listed as fitting none, or neither
            for w in ws:
                keys = set(act_labels(L, w["stem"]))
                in_th = {label_key(c) for t in w["themes"] for c in t["codes"]} & keys
                n_lab, n_in = n_lab + len(keys), n_in + len(in_th)
                n_none += len({label_key(c) for c in w.get("unplaced_codes", [])} & keys - in_th)
            pc = (lambda x: f"{100 * x / max(n_lab, 1):.0f}%")
            lines.append(f"Round 1: {len(ws)} act(s); themes per act {min(nth)} to {max(nth)} "
                         f"(mean {sum(nth) / len(nth):.1f}); codes per theme {min(cpt) if cpt else 0} to "
                         f"{max(cpt) if cpt else 0}; summaries {min(sw)} to {max(sw)} words; of the acts' {n_lab:,} "
                         f"labels, {pc(n_in)} are in a theme, {pc(n_none)} are listed as fitting none and "
                         f"{pc(n_lab - n_in - n_none)} are neither ({sum(left)}); acts shown in short form: "
                         f"{sum(1 for w in ws if w.get('level'))}.")
        g = read_json(groupings_path(L))
        if g:
            lines.append("Round 2: " + "; ".join(
                f"\"{x['grouping']}\" ({x['basis']}): " + ", ".join(f"{grp['group']} {len(grp['acts'])}"
                                                                    for grp in x["groups"]) for x in g["groupings"])
                         + (f". Problems left: {len(g.get('problems', []))}" if g.get("problems") else ""))
        for n in range(1, n_groupings(L) + 1):
            c = read_json(comparison_path(L, n))
            if c:
                lines.append(f"Round 3, grouping {n}: {len(c['comparisons'])} comparisons, {len(c['themes'])} themes"
                             + (f"; problems left: {len(c.get('problems', []))}" if c.get("problems") else ""))
        f = read_json(final_path(L))
        if f:
            Wl = {a["Act"]: read_json(within_path(L, act_stem(a))) for a in sel}
            Cl = {n: read_json(comparison_path(L, n)) for n in range(1, n_groupings(L) + 1)}
            by_key, by_name = earlier_index({a: w for a, w in Wl.items() if w}, {n: c for n, c in Cl.items() if c})
            earlier = set(by_key)              # (act or grouping, theme): a name used in two groupings counts twice
            named = set()
            for e in [x for r in f["revisions"] for x in r["themes"]] + [x for t in f["themes"]
                                                                          for x in t.get("comes_from", [])]:
                m = re.match(r"\s*(.+?)\s*[:–—-]\s+(.+)$", e)
                if m:
                    g = re.match(r"(?i)grouping\s*(\d+)", m.group(1))
                    key = (f"grouping {g.group(1)}" if g else act_key(m.group(1)), label_key(m.group(2)))
                    if key in earlier:
                        named.add(key)
                        continue
                names = {label_key(e)} | ({label_key(m.group(2))} if m else set())
                named |= {key for key in earlier if key[1] in names}
            cf = [e for t in f["themes"] for e in t.get("comes_from", [])]
            found = sum(1 for e in cf if trace(e, by_key, by_name) is not None)
            lines.append(f"Round 4: {len(f['themes'])} final themes, {len(f['revisions'])} revision records; "
                         f"{len(earlier & named)} of the {len(earlier)} earlier themes are named in a record or a "
                         f"final theme's \"comes from\"; {found} of {len(cf)} \"comes from\" entries found among the "
                         f"earlier themes; {len(f.get('column_codes_not_listed', []))} codes of the columns are in no "
                         f"final theme's list.")
        if os.path.exists(final_path(L)):
            log = running_memo(L, sel, "end")[2]
            ch = collections.Counter((x["change"], x["step"]) for x in log)
            lines.append(f"Running memo: {ch[('added', 5)]} entries added in Step 5, {ch[('revised', 5)]} of them "
                         f"revised and {ch[('dropped', 5)]} dropped; of Step 3's entries, {ch[('revised', 3)]} revised "
                         f"and {ch[('dropped', 3)]} dropped. It has "
                         f"{est_tokens(running_memo(L, sel, 'end')[0]):,} tokens as Step 5 leaves it.")
        m1, m23 = read_text(rpath(f"agent_{L}", "memo_part1.md")), read_text(rpath(f"agent_{L}", "memo_parts2and3.md"))
        if m1:
            lines.append(f"Memo: part 1 {nwords(m1):,} words" + (f"; parts 2 and 3 {nwords(m23):,} words" if m23 else ""))
        fl = collections.Counter(x.get("round") for x in read_jsonl(rpath(f"agent_{L}", "flags.jsonl")))
        if fl:
            lines.append("Flags: " + ", ".join(f"{r} {n}" for r, n in fl.items()))
        cost = {kd: sum(c.get("cost", 0) or 0 for c in v) for kd, v in by.items()}
        secs = {kd: sum(c.get("seconds", 0) or 0 for c in v) for kd, v in by.items()}
        lines.append("Cost: " + ", ".join(f"{kd} ${v:.2f}" for kd, v in cost.items()) + f" (all ${sum(cost.values()):.2f})")
        again = [c for c in calls if (c.get("attempt") or 1) > 1]
        if again:
            per = collections.Counter(c.get("kind") for c in again)
            lines.append(f"Asked again: {len(again)} call(s) (" + ", ".join(f"{k} {n}" for k, n in per.items())
                         + f"), costing ${sum(c.get('cost', 0) or 0 for c in again):.2f} of the total.")
        steps = [c.get("tool_steps", 0) or 0 for c in calls if c.get("kind") in
                 ("comparison", "final", "memo1", "lead", "memo23")]
        if steps:
            lines.append(f"File reads per call in rounds 3 to 8: mean {sum(steps) / len(steps):.1f}, most {max(steps)}.")
        limit = MODELS[L]["max_output"]
        for kd, what in (("comparison", "round 3"), ("final", "round 4")):
            outs = [c.get("last_output", (c.get("usage") or {}).get("output", 0)) for c in by.get(kd, [])
                    if c.get("ok")]
            if outs:
                warn = max(outs) > 0.25 * limit
                lines.append(f"Largest {what} reply: {max(outs):,} output tokens (thinking included) of the model's "
                             f"{limit:,}." + (" WARNING: replies grow with the number of acts, so with all 52 this one "
                                              "may not fit; check before the full run." if warn else ""))
        w_calls = by.get("within", [])
        if w_calls:
            firsts = [c for c in w_calls if (c.get("attempt") or 1) == 1]
            tok_first = sum(c.get("tokens_counted_before", 0) or 0 for c in firsts)
            fixed, post = within_fixed(L)
            tok_full = 0
            for a in full:
                try:
                    tok_full += est_tokens(fixed) + est_tokens(within_parts(L, a, fixed, post)[0]["tail"])
                except AgentStop:
                    pass
            # round 1: the pilot's cost per token of each act's first call, re-asks included, applied to every act
            pc = cost["within"] * tok_full / max(tok_first, 1)
            ph = secs["within"] / max(len(firsts), 1) * len(full) / max(PARALLEL_ACTS, 1) / 3600
            # rounds 2 to 5: only the part of each prompt that depends on the acts grows; the guide, the Step 3 memo
            # and the memory stay the same
            s3 = step3(L)
            fixed_tok = est_tokens(load_guide()) + est_tokens(s3["memo"]) + est_tokens(s3["memory"]) + 1500
            rest = 0.0
            for c in calls:
                if c.get("kind") in ("groupings", "comparison", "final", "memo1"):
                    i0 = c.get("tokens_counted_before", 0) or 0
                    rest += (c.get("cost", 0) or 0) * (fixed_tok + max(i0 - fixed_tok, 0) * k) / max(i0, 1)
                elif c.get("kind") in ("lead", "memo23"):
                    rest += (c.get("cost", 0) or 0) * 1.5
            rest_h = sum(secs.get(x, 0) for x in ("groupings", "comparison", "final", "memo1", "lead", "memo23")) \
                * 2 / 3600
            proj_cost += pc + rest
            proj_hours = max(proj_hours, ph + rest_h)
            lines.append(f"Full run, roughly: round 1 ${pc:,.0f} and {ph:.1f} hours; the other rounds ${rest:,.0f}.")
        lines.append("")
    lines += [f"Projected cost of the full run, all three agents: about ${proj_cost:,.0f} (BUDGET_USD is "
              f"${BUDGET_USD:,.0f}; set it with some room); about {proj_hours:.1f} hours, as the agents work side by "
              f"side. A rough guide: round 1 scales with the size of each act's material, re-asks included; in "
              f"rounds 2 to 5 only the part of each prompt that lists the acts grows; the lead's comparison and memo "
              f"parts 2 and 3 grow a little. Calls in progress when the budget runs out still finish, so a run can "
              f"pass BUDGET_USD by the cost of those calls (up to 12 in round 1).", ""]
    for L in LETTERS:
        m = MODELS[L]["model"]
        for d in price_changes(m):
            now_p, then_p = price(m), price(m, d)
            lines += [f"Agent {L}'s model ({m}) changes its prices on {nice_date(d)}: input ${now_p['input']:.2f} to "
                      f"${then_p['input']:.2f} and output ${now_p['output']:.2f} to ${then_p['output']:.2f} per million "
                      f"tokens. The projection uses today's prices.", ""]
    if os.path.exists(rpath("merged", "done.json")):
        ov = overlap_rows(sel)
        matched = sum(1 for r in ov if r["closest"] and r["closest"][0][1] >= 0.3)
        half = sum(1 for r in ov if r["share_within"] >= 0.5)
        lines.append(f"Merge: {len(ov)} final themes in all; {half} have at least half of their turns in one theme of "
                     f"another agent, and {matched} have a Jaccard index of 0.3 or more with one (Jaccard is low "
                     f"whenever two themes differ much in size; see overlap.txt).")
    rep = "\n".join(lines)
    write_text(rpath("reports", "pilot_report.md"), rep)
    print(rep)


if __name__ == "__main__":
    print(VERSION, "is loaded. Start with check_inputs(), then dry_run(); see the notes at the top of the script.")
