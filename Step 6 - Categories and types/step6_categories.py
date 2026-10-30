"""step6_categories.py, version 1.4 (30 October 2026)

What it does: runs Step 6, developing categories and classifying the cases (the guide's Step 6), with the three AI
agents of Steps 3 and 5, each starting from its own Step 5 work, in three rounds. In every round the agents work side
by side, each in calls of its own; the script does the pooling and the counting between rounds.
  1. own: each agent, on its own, groups its Step 5 final themes into aggregate dimensions, builds types of act and
     places every act, with its evidence (Prompt 1). It sees no other agent's work and no performance data.
  2. vote: the script pools the three agents' dimensions and types into one list (merged/pool.json), in a neutral
     order. The vote is blind: the list gives each item's name, definition and boundaries, not its author or the themes
     behind it, and this round sees the running memo without round 1's entries (they name the agent's own items; they
     come back in round 3). For every item, each agent names the pooled item closest to it in meaning and marks the
     item keep, drop or the same as another; it maps its own themes to the dimensions it keeps (at least one each, and
     no theme under more than two) and lists the acts in each type it keeps (at least MIN_TYPE_ACTS) (Prompt 2). The
     script counts (merged/shared_scheme.json): two items count as one when at least AGREE agents mark them the same,
     under the name of the item most agents chose; an agent backs an item only with that evidence, and an item stands
     when at least AGREE agents back it; an act stands in a type when at least AGREE agents place it there, an agent's
     acts for an item counting for every item it marks the same, and a type needs at least MIN_TYPE_ACTS such acts. A
     type that leaves fewer than MIN_TYPE_ACTS acts outside it is kept apart as a pattern that nearly every act shows.
     Each agent also says, for every type, whether it rests on what round 3 measures (splits, returns, stops,
     longevity, output, later success); a type that at least AGREE agents say does is set aside, out of the tables,
     since comparing it would be circular. Everything else is kept on record as a disagreement.
  3. careers: the shared scheme is fixed. The script reads the performance data (step6_inputs/<DV_FILE>) and
     tabulates each measure, counted from each act's Best New Artist ceremony, for the acts in each agreed type and
     for the acts outside it (merged/performance_by_type.csv and performance_by_act.csv). Each agent reads the tables
     against the scheme, without changing it, and writes its Step 6 memo (Prompt 3).
  Then the merge: merged/step6_combined.xlsx, with every round's work and the counts; merged/step6_memos.md, the
  three memos in one file to read, after the shared scheme they cite; and reports/step6_report.md.
Rounds 1 and 2 update each agent's running memo, which starts from its memory as Step 5 left it; round 3 and later
steps see it as updated (agent_X/memory.txt).

Agents (see MODELS): A = Claude Opus 5.5 (at low effort), B = Gemini 4 Argon and C = GPT-6.1 Sol (both at medium), as
in Steps 3 and 5.

Reads, in MyDrive/harrison_v7 (it never changes them):
  band_voice_by_act/*.jsonl and harrison_v7_band_voice_acts.json   the corpus (the acts, and the turns quotations are
                                                                     checked against)
  step3_inputs/excluded_turns.csv and step3_inputs/<GUIDE_FILE>      as in Step 5
  step3_<STEP3_RUN>/agent_A, agent_B, agent_C                        each agent's Step 3 memo, memory and labels
  step5_<STEP5_RUN>/agent_A, agent_B, agent_C                        each agent's Step 5 work: final.json, memo_part1.md,
                                                                     memo_parts2and3.md, memory.txt (its running memo as
                                                                     Step 5 left it), within/ (its summaries of the
                                                                     acts) and folder/ (the files it can open, which
                                                                     Step 6 opens too)
  step6_inputs/prompt1.txt, prompt2.txt, prompt3.txt                 the three prompts, as plain text
  step6_inputs/<DV_FILE>                                             the performance data (sheets Acts and Albums)
The acts are those of the Step 5 run (its merged/done.json).
Writes, in MyDrive/harrison_v7/step6_<RUN>/:
  agent_A, agent_B, agent_C   own.json (round 1), votes.json (round 2), memo6.md (round 3), memory.txt (the running
               memo as Step 6 leaves it) and memo_changes.csv, calls/ (every call), prompts/ (the fixed parts),
               flags.jsonl (problems that remained), state.json
  merged/      pool.json, shared_scheme.json and shared_scheme.txt, performance_by_type.csv, performance_by_act.csv,
               step6_combined.xlsx, and step6_memos.md (the three memos to read and send on, with the shared scheme
               they cite and the tables they read)
  reports/     step6_report.md
  run_settings.json, run_log.txt
Everything is saved as it goes: a round that stops (Colab disconnects, you press Stop, the budget runs out) carries on
where it left off when you run it again, and nothing done is done twice.

HOW TO RUN, in Colab, with the same keys in Colab Secrets as Steps 3 and 5:
  0. Upload this script to MyDrive/harrison_v7, and the three prompts and the performance data to
     MyDrive/harrison_v7/step6_inputs.
  1. !pip install -q "anthropic==2.3.0" "openai==3.37.0" "google-genai==2.33.0" tiktoken python-docx openpyxl
  2. from google.colab import drive; drive.mount("/content/drive")
     %run -i "/content/drive/MyDrive/harrison_v7/step6_categories.py"
     STEP5_RUN = "pilot3"; RUN = "pilot3"    a pilot: Step 6 on pilot 3's Step 5 work (its eight acts); earlier
                                             pilots are in step6_pilot (version 1.0) and step6_pilot2 (1.2)
     STEP5_RUN = "full"; RUN = "full"        the real run, on the full Step 5 run (all 52 acts)
  3. check_inputs()      the corpus, Step 3's and Step 5's work, the prompts, the guide, the data, the keys (no calls)
     dry_run()           writes round 1's filled prompts to step6_<RUN>/dry_run, to read before spending anything
     test_models()       a few small calls per model
  4. run_all()           every round in order; or one at a time: run_own(), run_vote(), run_careers(), run_merge()
     report()            what each round produced, the counts and the cost
     status()            where each agent is, its cost and its flags, at any time
  To stop, press Stop once and wait: the calls in progress finish and are saved first.
"""
import collections, csv, datetime, hashlib, json, os, re, shutil, threading, time, traceback, unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed

VERSION = "step6_categories 1.4 (30 Oct 2026)"

if any(t.name.startswith("step6-agent") and t.is_alive() for t in threading.enumerate()):
    raise SystemExit("Agents from the last run are still finishing their calls in this session, so the script was "
                     "not loaded again. Wait a minute (status() shows where they are), then run this cell again.")


# ==================================================================================================== settings
BASE = "/content/drive/MyDrive/harrison_v7"
RUN = "full"                       # the run folder is step6_<RUN>; a new name starts a new run
STEP5_RUN = "full"                 # the Step 5 run whose work Step 6 starts from (step5_<STEP5_RUN>); its acts are
                                   # Step 6's acts
STEP3_RUN = "full"                 # Step 3's run folder is step3_<STEP3_RUN>
STEP3_INPUTS = "step3_inputs"      # the guide and the excluded turns, as Steps 3 and 5 used them
INPUT_DIR = "step6_inputs"         # the three prompts and the performance data
PROMPT_FILES = {n: f"prompt{n}.txt" for n in range(1, 4)}
GUIDE_FILE = "best_practices_guide.docx"
EXCLUDE_FILE = "excluded_turns.csv"
DV_FILE = "dv_performance_data.xlsx"
CEREMONIES = {   # the Best New Artist ceremonies (TVmaze's list of Grammy ceremonies and Wikipedia)
    1980: "1980-02-27", 1981: "1981-02-25", 1982: "1982-02-24", 1983: "1983-02-23", 1984: "1984-02-28",
    1985: "1985-02-26", 1986: "1986-02-25", 1987: "1987-02-24", 1988: "1988-03-02", 1989: "1989-02-22",
    1990: "1990-02-21"}

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
BUDGET_USD = 40.0                  # no new call starts once the three agents' total cost in this run passes this

AGREE = 2                          # agents (of the three) that must back an item, a merge or an act's place in a type
SHUFFLE_SEED = "step6"             # the pooled list's neutral order (the same for every agent, and on every run)
MIN_TYPE_ACTS = 3                  # round 1: a type with fewer acts that fit it clearly is noted; round 2: an agent
                                   # backs a type only if it places this many acts in it, a type stands only with this
                                   # many agreed acts, and one leaving fewer acts outside it is a pattern, not a type
MAX_SUPPORT, MAX_CONTRARY = 3, 2   # round 1: Turn IDs per act placement
MEMO6_WORDS = {1: 400, 2: 700, 3: 300}
WORDS_SLACK = 1.3                  # a memo part this much over its length is asked for again, once
CONTEXT_ROOM = 0.8
OUTPUT_ROOM = 40000
RETRIES_UNUSABLE = 2               # re-asks when a reply cannot be used (not JSON, parts missing)
RETRIES_PROBLEMS = 1               # re-asks when a reply can be used but a check finds problems; then it is flagged
TRANSIENT_RETRIES = 8              # after network or server errors, waits and tries again: about 45 minutes in all
FIRST_WAIT, MAX_WAIT = 30, 600
OTHER_RETRIES = 4
MAX_TOOL_STEPS = 60                # most times an agent may use its folder in one call,
READ_BUDGET = 800000               # and most characters it may read from it in all (about 200,000 tokens)
READ_LINES, READ_CHARS = 300, 60000                    # and in one read
GPT_CHUNK_TOKENS = 4000
GPT_MAX_CHUNKS = 40
MEASURES = [   # the performance measures of round 3: (key, name, column of the Acts sheet or None if computed, lower
               # is better)
    ("albums", "Studio albums after the ceremony", "Studio albums after the ceremony", False),
    ("us_albums", "US studio albums after the ceremony", None, False),
    ("charted", "Albums on the Billboard 200 after the ceremony", "Albums charted after the ceremony", False),
    ("peak", "Best Billboard 200 peak after the ceremony", "Best Billboard 200 peak after the ceremony", True),
    ("top40", "Top-40 albums after the ceremony", "Top-40 albums after the ceremony", False),
    ("riaa", "RIAA-certified albums after the ceremony", "RIAA-certified albums after the ceremony", False),
    ("units", "RIAA units after the ceremony (millions)", "RIAA units, millions, after the ceremony", False),
    ("noms", "Later Grammy nominations", "Later Grammy nominations", False),
    ("wins", "Later Grammy wins", "Later Grammy wins", False),
    ("years", "Years active after the ceremony", "Years active after the ceremony", False),
]
AROUND_MONTHS = 12                 # (as in Step 5: "around" the nomination is within 12 months of it)
NO_CHART = 201                     # a best peak for an act with no album on the Billboard 200 after its ceremony, so
                                   # that it ranks below every act that charted (shown as "none" in the tables)

# ==================================================================================================== basics
_G = globals()                     # with %run -i, loading the script again keeps the same locks and stop signal
PRINT_LOCK = _G.get("PRINT_LOCK") or threading.Lock()
COST_LOCK = _G.get("COST_LOCK") or threading.Lock()
FILE_LOCK = _G.get("FILE_LOCK") or threading.Lock()
STOP = _G.get("STOP") or threading.Event()
STOP_WHY = _G.get("STOP_WHY") or ["stopped by hand"]      # why STOP was set: by hand, or the budget
THREAD_PREFIX = "step6-agent"
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
    return os.path.join(base, f"step6_{run}", *p)


def s5path(letter, *p):
    """A file of the agent's Step 5 work (letter None: of the Step 5 run as a whole)."""
    base = getattr(_LOCAL, "base", None) or BASE
    return os.path.join(base, f"step5_{STEP5_RUN}", *(([f"agent_{letter}"] if letter else []) + list(p)))


def s3path(letter, *p):
    return os.path.join(getattr(_LOCAL, "base", None) or BASE, f"step3_{STEP3_RUN}", f"agent_{letter}", *p)


def say(*a):
    """Prints a line, and keeps it in step6_<RUN>/run_log.txt, so the log survives a lost Colab page."""
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
    """Step 6's acts: those of the Step 5 run it starts from (its merged/done.json), in nomination-year order; or,
    given as names (for tests), those."""
    c = load_corpus()
    if acts is None:
        done = read_json(s5path(None, "merged", "done.json"))
        if not done or not done.get("acts"):
            raise SystemExit(f"STOPPED: {s5path(None, 'merged', 'done.json')} not found: is STEP5_RUN right, and is "
                             f"that Step 5 run finished?")
        stems = set(done["acts"])
        return [a for a in c["acts"] if act_stem(a) in stems]
    by_name = {a["Act"].casefold(): a for a in c["acts"]}
    out = []
    for n in ([acts] if isinstance(acts, str) else list(acts)):
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


def act_key(name):
    """An act's name for matching the events file to the corpus: accents, case, a leading "The" and punctuation aside
    ("The Go-Go's" and "Go-Go’s", "Tone Loc" and "Tone Lōc" match)."""
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode().casefold().strip()
    s = re.sub(r"^the\s+", "", s)
    return re.sub(r"[^a-z0-9]+", "", s)

# ==================================================================================================== prompts and guide


PROMPTS, GUIDE = {}, {"text": None}
PH = re.compile(r"\{(?![\"“])([^{}\n]+)\}")     # a placeholder; a JSON object in a format starts with {" instead
REQUIRED = {
    1: {"GUIDE", "LETTER", "MEMO3", "MEMO5P1", "MEMORY", "FINAL", "ACTS"},
    2: {"GUIDE", "LETTER", "MEMO3", "MEMO5", "MEMORY", "POOLD", "POOLT", "FINALSHORT", "ACTS"},
    3: {"GUIDE", "LETTER", "MEASURES", "MEMO3", "MEMO5", "MEMORY", "OWN", "SHARED", "TABLES"}}
BIG = ("GUIDE", "MEMO3", "MEMO5P1", "MEMO5", "MEMORY", "FINAL", "FINALSHORT", "ACTS", "POOLD", "POOLT", "OWN", "SHARED",
       "TABLES", "MEASURES")


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
    if low.startswith("part 1 of your step 5 memo"):
        return "MEMO5P1"
    if "step 5 memo" in low:
        return "MEMO5"
    if low.startswith("your final themes"):
        return "FINAL" if "boundaries" in low else "FINALSHORT"
    if low.startswith("for each act"):
        return "ACTS"
    if low.startswith("the pooled dimensions"):
        return "POOLD"
    if low.startswith("the pooled types"):
        return "POOLT"
    if low.startswith("your own scheme"):
        return "OWN"
    if low.startswith("the shared scheme"):
        return "SHARED"
    if low.startswith("the performance tables"):
        return "TABLES"
    if low.startswith("the measures"):
        return "MEASURES"
    return None


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
    if n in (1, 2) and "JSON" not in s:
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


ROUNDS = {"own": 1, "vote": 2, "careers": 3}


def dv_sha():
    p = os.path.join(BASE, INPUT_DIR, DV_FILE)
    if not os.path.exists(p):
        return ""
    with open(p, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def round_settings(r, acts):
    s = {"script": VERSION, "acts": [act_stem(a) for a in acts], "step3_run": STEP3_RUN, "step5_run": STEP5_RUN,
         "agree": AGREE}
    if r in ROUNDS:
        load_prompts()
        s.update(models=MODELS, guide_sha=sha(load_guide()), prompt_sha=sha(PROMPTS[ROUNDS[r]]))
    if r == "vote":
        s["shuffle_seed"] = SHUFFLE_SEED
    if r == "careers":
        s["dv_sha"] = dv_sha()
    return s


def round_started(r):
    if r == "merge":
        return os.path.exists(rpath("merged", "done.json"))
    names = {"own": "own.json", "vote": "votes.json", "careers": "memo6.md"}
    return any(os.path.exists(rpath(f"agent_{L}", names[r])) for L in LETTERS) if r in names else False


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
    last } (rounds 1 and 2 may use their folder first and put a sentence before the JSON). A complete object followed
    by a stray closing brace, or by a note with no braces in it, is taken as it is. The error reported is the one from
    the JSON block, or from the first { to the last }, where there is one: the whole reply often starts with ```."""
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
    errs = []
    for c in cands:
        try:
            obj = json.loads(c)
        except Exception as e:
            errs.append(short(e, 200))
            continue
        if isinstance(obj, dict):
            return obj, None
        errs.append("the reply is JSON but not an object")
    if i >= 0:
        try:
            obj, end = json.JSONDecoder().raw_decode(s, i)
        except Exception:
            obj, end = None, 0
        if isinstance(obj, dict) and "{" not in s[end:]:
            return obj, None
    return None, errs[1] if len(errs) > 1 else errs[0]


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


def _obj(props, required=None):
    return {"type": "object", "additionalProperties": False, "required": required or list(props), "properties": props}


_S = {"type": "string"}
_LIST = {"type": "array", "items": _S}
_PERIODS = {"type": "array", "items": {"type": "string", "enum": list(PERIODS)}}
SCHEMA_TEST = _obj({"memory": _LIST})
_MEMO = _obj({"add": _LIST, "revise": {"type": "array", "items": _obj({"entry": _S, "text": _S})}, "drop": _LIST})

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
    "own": "one sentence for each definition and boundary, and one sentence on why for each act",
    "vote": "one sentence on why for each item",
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

import itertools, random

# ==================================================================================================== Step 5's work, as Step 6 reads it


def own_path(letter):
    return rpath(f"agent_{letter}", "own.json")


def votes_path(letter):
    return rpath(f"agent_{letter}", "votes.json")


def memo6_path(letter):
    return rpath(f"agent_{letter}", "memo6.md")


NUMBERED = re.compile(r"^\s*[A-Z]\d{1,3}\s*(?:[.:)]|[\u2014\u2013-])\s*")   # a number the agent put in front of
                                                                         # a theme's name ("F01. ...", "T01 — ...")


def final_themes(letter):
    """The agent's Step 5 final themes, each with its Step 6 ID (F1, F2 ..., in Step 5's order) and its name without a
    number the agent may have put in front of it."""
    f = read_json(s5path(letter, "final.json"))
    if not f:
        raise AgentStop(f"Step 5's final themes are not there ({s5path(letter, 'final.json')}): is STEP5_RUN right?")
    return [dict(t, id=f"F{i}", theme=NUMBERED.sub("", t["theme"]).strip() or t["theme"])
            for i, t in enumerate(f["themes"], 1)]


def theme_block(t, short_form=False):
    lines = [f"ID: {t['id']}", f"Theme: {t['theme']}", f"Definition: {t['definition']}"]
    if short_form:
        return "\n".join(lines)
    if t.get("boundaries"):
        lines.append(f"Boundaries: {t['boundaries']}")
    lines.append("Codes: " + "; ".join(t.get("codes") or []))
    if t.get("acts"):
        lines.append("Acts: " + "; ".join(t["acts"]))
    if t.get("periods"):
        lines.append("Periods: " + ", ".join(t["periods"]))
    lines.append("Supporting Turn IDs: " + (", ".join(t.get("supporting") or []) or "none"))
    if t.get("contrary"):
        lines.append("Contrary Turn IDs: " + ", ".join(t["contrary"]))
    if t.get("too_little_evidence"):
        lines.append("Too little evidence in: " + "; ".join(t["too_little_evidence"]))
    return "\n".join(lines)


def final_block6(letter, short_form=False):
    return "\n\n".join(theme_block(t, short_form) for t in final_themes(letter))


def summary_of(letter, a):
    w = read_json(s5path(letter, "within", f"{act_stem(a)}.json")) or {}
    return w.get("summary") or "(no summary)"


def acts_block6(letter, acts):
    return "\n\n".join(f"Act: {a['Act']} ({kind(a)}); Best New Artist {'won' if won(a) else 'nominated'} at the "
                       f"{a['Nominated']} ceremony\nYour Step 5 summary: {summary_of(letter, a)}" for a in acts)


def memo5(letter, part1_only=False):
    """The agent's Step 5 memo: part 1 only (written before it saw the other agents' work), or all three parts."""
    p1 = (read_text(s5path(letter, "memo_part1.md")) or "").strip()
    if part1_only:
        return p1 or "(none)"
    whole = (read_text(s5path(letter, f"memo_{letter}.md")) or "").strip()
    if whole:
        return whole
    p23 = (read_text(s5path(letter, "memo_parts2and3.md")) or "").strip()
    return (p1 + "\n\n" + p23).strip() or "(none)"


def all_tids(acts):
    return {t["Turn ID"] for a in acts for t in load_corpus()["all_turns"][act_stem(a)]}


def act_tids(acts):
    return {act_key(a["Act"]): {t["Turn ID"] for t in load_corpus()["all_turns"][act_stem(a)]} for a in acts}

# ==================================================================================================== the running memo


MEMO_LINE = re.compile(r"^(M\d+) \(([^)]*)\): (.*)$")
ADD_RE = re.compile(r"^\s*(?:M\d+\s*)?\(([^)]*)\)\s*:?\s*(.*)$", re.S)   # an entry the agent numbered itself is
                                                                         # read without its number


def memo5_base(letter):
    """The running memo as Step 5 left it (its memory.txt): the text, its numbered entries and the next number."""
    text = read_text(s5path(letter, "memory.txt"))
    if not text or not text.strip():
        raise AgentStop(f"Step 5's running memo is not there ({s5path(letter, 'memory.txt')})")
    entries = {}
    for line in text.splitlines():
        m = MEMO_LINE.match(line.strip())
        if m:
            entries[m.group(1)] = {"id": m.group(1), "where": m.group(2), "text": m.group(3), "step": 5}
    return text.rstrip(), entries, max([int(k[1:]) + 1 for k in entries] + [1])


def memo6_updates(letter, upto):
    """The changes the agent made to its running memo in Step 6 that the call at stage `upto` ("own", "vote", "careers"
    or "end") sees, in order. Round 2 sees none: round 1's entries would tell the agent which pooled items are its own,
    so they come back from round 3 on."""
    out = []
    if upto in ("careers", "end"):              # not "vote": round 1's entries name the agent's own dimensions and
        out.append(("making your own scheme", (read_json(own_path(letter)) or {}).get("memo") or {}))  # types
    if upto in ("careers", "end"):
        out.append(("judging the pooled scheme", (read_json(votes_path(letter)) or {}).get("memo") or {}))
    return out


def running_memo6(letter, upto):
    """The running memo as the call at stage `upto` sees it: Step 5's with the Step 6 changes made, in order. Returns
    its text, its entries (dropped ones marked) and the log of changes."""
    base, entries, nxt = memo5_base(letter)
    entries = {k: dict(v) for k, v in entries.items()}
    added, log = [], []
    for where, upd in memo6_updates(letter, upto):
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
                          "text": (m.group(2) if m else a).strip(), "step": 6, "added": where}
            added.append(k)
            log.append({"where": where, "change": "added", "entry": k, "step": 6, "before": "",
                        "after": f"({entries[k]['where']}): {entries[k]['text']}"})
    lines = []
    for line in base.splitlines():
        m = MEMO_LINE.match(line.strip())
        e = entries.get(m.group(1)) if m else None
        if e and e.get("dropped"):
            continue
        if e and e.get("revised"):
            line = f"{e['id']} ({e['where']}): {e['text']} [revised in Step 6, {e['revised']}]"
        lines.append(line)
    groups = collections.OrderedDict()
    for k in added:
        e = entries[k]
        if not e.get("dropped"):
            groups.setdefault(e["added"], []).append(
                f"{k} ({e['where']}): {e['text']}" + (f" [revised, {e['revised']}]" if e.get("revised") else ""))
    text = "\n".join(lines)
    text += "\n\nYour entries from Step 6:\n" + ("\n\n".join(f"From {w}:\n" + "\n".join(v) for w, v in groups.items())
                                               if groups else "none yet.")
    return text, entries, log


def memo_known6(letter, stage):
    return {k for k, e in running_memo6(letter, stage)[1].items() if not e.get("dropped")}


def save_memory6(letter):
    """After round 2: the running memo as Step 6 leaves it (memory.txt, for later steps), and its changes."""
    text, entries, log = running_memo6(letter, "end")
    write_text(rpath(f"agent_{letter}", "memory.txt"), text + "\n")
    write_csv(rpath(f"agent_{letter}", "memo_changes.csv"), ["Call", "Change", "Entry", "From step", "Before", "After"],
              [[x["where"], x["change"], x["entry"], x["step"], x["before"], x["after"]] for x in log])
    return log

# ==================================================================================================== checks of the replies


FIT = {"clear": "clear", "clearly": "clear", "partial": "partial", "partly": "partial", "partially": "partial"}
DECISIONS = {"keep": "keep", "kept": "keep", "keeps": "keep", "drop": "drop", "dropped": "drop", "same": "same",
             "same as": "same", "merge": "same", "merged": "same"}
F_ID = re.compile(r"^\s*(F\d+)\b", re.I)


def _fit(v):
    return FIT.get(str(v or "").strip().lower().rstrip("."), "partial") if v else "clear"


def _theme_ids(v, ids):
    """Theme IDs from a reply ("F3", "F3: name", or the theme's name): the IDs found, and the entries that name none."""
    by_name = {label_key(n): k for k, n in ids.items()}
    found, unknown = [], []
    for x in _strs(v):
        m = F_ID.match(x)
        k = f"F{int(m.group(1)[1:])}" if m else None   # "F03" is F3
        name = label_key(re.sub(r"^\s*F\d+\s*[.:)\-]?\s*", "", x))
        if k in ids:
            found.append(k)
        elif name in by_name:
            found.append(by_name[name])
        else:
            unknown.append(x)
    return list(dict.fromkeys(found)), unknown


def _act(name, canon):
    return canon.get(act_key(_text(name)))


def check_own(obj, ids, act_names, tids_by_act, known, scope):
    """Round 1's reply: dimensions (each at least two of the agent's final themes, by ID), the themes that fit none
    (every theme in one or the other), types (with the dimensions and themes that define them), every act placed (in
    the types it fits, or none, with Turn IDs of that act), misfits and empty cells, and the memo. Returns what to keep,
    problems that make it unusable, problems to ask again about, and notes."""
    if not isinstance(obj, dict) or not all(isinstance(obj.get(k), list) for k in ("dimensions", "types", "acts")):
        return None, ['The reply needs the lists "dimensions", "types" and "acts".'], [], []
    soft, notes = [], []
    canon = {act_key(n): n for n in act_names}
    dims, seen = [], set()
    for d in obj["dimensions"]:
        if not isinstance(d, dict):
            continue
        name = _text(d.get("dimension"))
        if not name:
            soft.append("A dimension has no name.")
            continue
        if label_key(name) in seen:
            soft.append(f'Two dimensions are called "{short(name, 60)}": give each its own name.')
            continue
        seen.add(label_key(name))
        th, unk = _theme_ids(d.get("themes"), ids)
        if unk:
            soft.append(f'Dimension "{short(name, 60)}": these are not IDs of your final themes: '
                        f'{", ".join(short(x, 30) for x in unk[:6])}.')
        if len(th) < 2:
            soft.append(f'Dimension "{short(name, 60)}" holds {len(th)} of your final themes; a dimension holds at '
                        f'least two.')
        dims.append({"dimension": name, "definition": _text(d.get("definition")),
                     "boundaries": _text(d.get("boundaries")), "themes": th})
    unplaced = []
    for u in obj.get("unplaced_themes") or []:
        th, unk = _theme_ids([u.get("theme") if isinstance(u, dict) else u], ids)
        unplaced += [{"theme": x, "why": _text(u.get("why")) if isinstance(u, dict) else ""} for x in th]
        if unk:
            notes.append(f'"unplaced_themes" names a theme that is not one of yours: {short(unk[0], 40)}.')
    placed = {x for d in dims for x in d["themes"]} | {u["theme"] for u in unplaced}
    missing = [k for k in ids if k not in placed]
    if missing:
        soft.append(f"These final themes are in no dimension and not listed as fitting none: {', '.join(missing)}. Put "
                    f"each in the dimension it belongs to, or list it under \"unplaced_themes\" with the reason.")
    dim_names = {label_key(d["dimension"]): d["dimension"] for d in dims}
    types, tseen = [], set()
    for ty in obj["types"]:
        if not isinstance(ty, dict):
            continue
        name = _text(ty.get("type"))
        if not name:
            soft.append("A type has no name.")
            continue
        if label_key(name) in tseen:
            soft.append(f'Two types are called "{short(name, 60)}": give each its own name.')
            continue
        tseen.add(label_key(name))
        th, unk = _theme_ids(ty.get("themes"), ids)
        if unk:
            soft.append(f'Type "{short(name, 60)}": these are not IDs of your final themes: '
                        f'{", ".join(short(x, 30) for x in unk[:6])}.')
        dn = []
        for x in _strs(ty.get("dimensions")):
            if label_key(x) in dim_names:
                dn.append(dim_names[label_key(x)])
            else:
                notes.append(f'Type "{short(name, 60)}": "{short(x, 50)}" is not one of your dimensions.')
        types.append({"type": name, "definition": _text(ty.get("definition")),
                      "boundaries": _text(ty.get("boundaries")), "dimensions": list(dict.fromkeys(dn)), "themes": th})
    if not types:
        soft.append("The reply has no types of act: build them from your dimensions, as the prompt asks.")
    type_names = {label_key(t["type"]): t["type"] for t in types}
    placements = {}
    for x in obj["acts"]:
        if not isinstance(x, dict):
            continue
        nm = _act(x.get("act"), canon)
        if not nm:
            soft.append(f'"{short(x.get("act"), 60)}" is not one of the acts: write the act names exactly as given.')
            continue
        if nm in placements:
            notes.append(f"{nm} is placed twice; the first placement is kept.")
            continue
        tys = []
        for y in (x.get("types") or []):
            tn = _text(y.get("type")) if isinstance(y, dict) else _text(y)
            if not tn or label_key(tn) in ("none", "no type", "fits none"):
                continue
            if label_key(tn) not in type_names:
                soft.append(f'{nm}: "{short(tn, 60)}" is not one of your types.')
                continue
            tys.append({"type": type_names[label_key(tn)], "fit": _fit(y.get("fit") if isinstance(y, dict) else None)})
        sup, con = _ids(x.get("supporting")), _ids(x.get("contrary"))
        bad = [i for i in sup + con if i not in tids_by_act.get(act_key(nm), set())]
        if bad:
            soft.append(f"{nm}: these Turn IDs are not turns of this act: {', '.join(bad[:5])}.")
        placements[nm] = {"act": nm, "types": list({y["type"]: y for y in tys}.values()), "why": _text(x.get("why")),
                          "supporting": [i for i in sup if i not in bad][:MAX_SUPPORT],
                          "contrary": [i for i in con if i not in bad][:MAX_CONTRARY]}
    left = [n for n in act_names if n not in placements]
    if left:
        soft.append(f"These acts are not placed: {', '.join(left)}. Place every act, one at a time, in the types it "
                    f"fits, or in none.")
    for ty in types:
        n_clear = sum(1 for p in placements.values() for y in p["types"] if y["type"] == ty["type"] and
                      y["fit"] == "clear")
        if n_clear < MIN_TYPE_ACTS:
            notes.append(f'Type "{short(ty["type"], 60)}" holds {n_clear} act(s) that fit it clearly (a type should '
                         f'hold at least {MIN_TYPE_ACTS}).')
    misfits = []
    for m in obj.get("misfits") or []:
        if isinstance(m, dict):
            nm = _act(m.get("act"), canon) or _text(m.get("act"))
            kind_ = _text(m.get("kind")).lower()
            misfits.append({"act": nm, "type": _text(m.get("type")), "kind": "core" if "core" in kind_ else "detail",
                            "what": _text(m.get("what"))})
    empty = [{"cell": _text(c.get("cell")), "possible": _text(c.get("possible")).lower(), "why": _text(c.get("why"))}
             for c in (obj.get("empty_cells") or []) if isinstance(c, dict)]
    memo, mnotes = check_memo(obj, known, scope)
    hard = [] if dims else ["The reply has no dimensions."]
    out = {"dimensions": dims, "unplaced_themes": unplaced, "types": types, "empty_cells": empty,
           "acts": [placements[n] for n in act_names if n in placements], "misfits": misfits, "memo": memo}
    return out, hard, soft, notes + mnotes


POOL_ID = re.compile(r"\b([DT])\s*0*(\d+)\b", re.I)
YES_NO = {"yes": True, "y": True, "true": True, "no": False, "n": False, "false": False}


def _pool_id(v):
    """A pooled item's ID as the reply gives it ("D3", "d 3", "D03", "D3: name"), or None."""
    m = POOL_ID.search(str(v or ""))
    return f"{m.group(1).upper()}{int(m.group(2))}" if m else None


def check_votes(obj, pool, ids, act_names, known, scope):
    """Round 2's reply: one entry for every pooled item, with the item closest to it in meaning and keep, drop or same
    (as another item of the same kind); for a dimension kept, the agent's own themes under it (at least one, and no
    theme under more than two of the dimensions it keeps); for a type kept, its acts (at least MIN_TYPE_ACTS, clearly
    or partly). Items the agent marks the same count as one throughout. Returns as check_own does."""
    if not isinstance(obj, dict) or not isinstance(obj.get("items"), list):
        return None, ['The reply needs the list "items", with one entry for every pooled item.'], [], []
    kinds = dict([(x["id"], "D") for x in pool["dimensions"]] + [(x["id"], "T") for x in pool["types"]])
    canon = {act_key(n): n for n in act_names}
    got, soft, notes = {}, [], []
    for it in obj["items"]:
        if not isinstance(it, dict):
            continue
        iid = _pool_id(it.get("id")) or ""
        if iid not in kinds:
            soft.append(f'"{short(it.get("id"), 20)}" is not the ID of a pooled item.')
            continue
        if iid in got:
            notes.append(f"{iid} is judged twice; the first is kept.")
            continue
        dec = DECISIONS.get(_text(it.get("decision")).lower())
        if not dec and re.match(r"(?i)\s*same\b", _text(it.get("decision"))):
            dec = "same"
        if not dec:
            soft.append(f'{iid}: the decision must be keep, drop or same, not "{short(it.get("decision"), 30)}".')
            continue
        same = _pool_id(it.get("same_as"))
        if dec == "same" and not same:             # "same as D3" written in the decision itself
            same = _pool_id(re.sub(r"(?i)^\s*same\s*(as)?", "", _text(it.get("decision"))))
        if dec == "same":
            if not same or same not in kinds or same == iid:
                soft.append(f'{iid} is marked the same as another item, but "same_as" does not give the ID of another '
                            f'pooled item.')
                continue
            if kinds[same] != kinds[iid]:
                soft.append(f"{iid} is marked the same as {same}, but one is a dimension and the other a type.")
                continue
        else:
            same = None
        close = _pool_id(it.get("closest"))
        if not close or close not in kinds or close == iid or kinds[close] != kinds[iid]:
            kind_ = "dimension" if kinds[iid] == "D" else "type"
            soft.append(f'{iid}: "closest" must give the ID of the other pooled {kind_} closest to it in meaning, '
                        f'whether or not the two are the same.')
            close = None
        rec = {"id": iid, "decision": dec, "same_as": same, "closest": close, "why": _text(it.get("why"))}
        if kinds[iid] == "D":
            th, unk = _theme_ids(it.get("themes"), ids)
            if unk:
                soft.append(f"{iid}: these are not IDs of your final themes: {', '.join(short(x, 30) for x in unk[:6])}.")
            rec["themes"] = th if dec != "drop" else []
        else:
            acts_l = {}
            for y in (it.get("acts") or []):
                raw = y.get("act") if isinstance(y, dict) else y
                nm = _act(raw, canon)
                if not nm:
                    soft.append(f'{iid}: "{short(raw, 60)}" is not one of the acts: write the act names exactly as '
                                f'given.')
                    continue
                acts_l.setdefault(nm, {"act": nm, "fit": _fit(y.get("fit") if isinstance(y, dict) else None)})
            rec["acts"] = list(acts_l.values()) if dec != "drop" else []
            out_ = YES_NO.get(_text(it.get("rests_on_outcomes")).lower().rstrip("."))
            if out_ is None:
                soft.append(f'{iid}: "rests_on_outcomes" must be yes or no: whether the type rests on whether or when '
                            f'an act split, came back or stopped, how long it lasted or how much it released, or its '
                            f'later success.')
            rec["outcomes"] = out_
        got[iid] = rec
    left = [k for k in kinds if k not in got]
    if left:
        soft.append(f"These pooled items are not judged: {', '.join(left)}. Give every pooled item one entry: keep, "
                    f"drop or same.")
    for r in got.values():
        if r["decision"] == "same" and got.get(r["same_as"], {}).get("decision") == "drop":
            soft.append(f"{r['id']} is marked the same as {r['same_as']}, but {r['same_as']} is dropped: keep "
                        f"{r['same_as']}, drop both, or mark {r['id']} the same as an item you keep.")
    under = collections.defaultdict(list)          # own theme -> the dimensions kept that list it
    for grp in own_groups(got):                    # what the agent keeps, items it marks the same counting as one
        grp = [i for i in grp if i in got]         # (an item named in a mark but not judged is asked about above)
        if not grp or all(got[i]["decision"] == "drop" for i in grp):
            continue
        lab = grp[0] if len(grp) == 1 else " and ".join(grp) + " (which you mark the same)"
        ev = group_evidence(got, grp)
        if kinds[grp[0]] == "D":
            if not ev:
                soft.append(f"{lab}: kept, but none of your final themes is listed under it. List the themes of yours "
                            f"that belong under it; if none does, drop it, or mark it the same as a dimension you keep.")
            for th in ev:
                under[th].append("/".join(grp))
        elif len(ev) < MIN_TYPE_ACTS:
            soft.append(f"{lab}: kept, but your reply places only {len(ev)} act(s) in it. Keep a type only if at least "
                        f"{MIN_TYPE_ACTS} acts fit it, clearly or partly, testing each act as in round 1; otherwise drop "
                        f"it, or mark it the same as a type you keep.")
    for th, where in under.items():
        if len(where) > 2:
            soft.append(f"Your theme {th} is listed under {len(where)} of the dimensions you keep ({', '.join(where)}): "
                        f"list each of your themes under the dimension that fits it best, and under a second only when "
                        f"it clearly belongs to both.")
    memo, mnotes = check_memo(obj, known, scope)
    return {"items": [got[k] for k in kinds if k in got], "memo": memo}, [], soft, notes + mnotes


def own_groups(items):
    """One agent's votes in groups: each item with the items its marks put it together with, directly or through a
    chain of marks, in ID order. `items` maps each pooled ID to the agent's vote on it."""
    u = _Union()
    for v in items.values():
        u.find(v["id"])
        if v["decision"] == "same" and v.get("same_as"):
            u.join(v["id"], v["same_as"])
    g = collections.defaultdict(list)
    for i in list(u.up):
        g[u.find(i)].append(i)
    return [sorted(x, key=_order) for x in sorted(g.values(), key=lambda x: min(_order(i) for i in x))]


def group_evidence(items, grp):
    """What an agent gives for a group of items it counts as one: its own themes under them (dimensions), or the acts
    it places in them, clearly or partly (types). `items` maps each pooled ID to the agent's vote on it."""
    if grp[0].startswith("D"):
        return list(dict.fromkeys(th for i in grp for th in (items.get(i) or {}).get("themes", [])))
    return list(dict.fromkeys(y["act"] for i in grp for y in (items.get(i) or {}).get("acts", [])))


def has_evidence(items, grp):
    return len(group_evidence(items, grp)) >= (1 if grp[0].startswith("D") else MIN_TYPE_ACTS)


def backed_items(items):
    """The pooled items an agent backs: each item it keeps or marks the same as another, and each item a mark of it
    names, provided it gives evidence for the group they make (at least one of its own themes for a dimension, at least
    MIN_TYPE_ACTS acts that fit, clearly or partly, for a type). `items` maps each pooled ID to the agent's vote."""
    out = set()
    for grp in own_groups(items):
        if all((items.get(i) or {}).get("decision", "drop") == "drop" for i in grp):
            continue
        if has_evidence(items, grp):
            out |= set(grp)
    return out


def backing(letter, pool, items):
    """How many of its own pooled items, and of the other agents', an agent backs (as backed_items says). Returns
    (own backed, own, others backed, others)."""
    author = {x["id"]: x["author"] for x in pool["dimensions"] + pool["types"]}
    b = backed_items({v["id"]: v for v in items})
    own = [i for i in author if author[i] == letter]
    other = [i for i in author if author[i] != letter]
    return sum(i in b for i in own), len(own), sum(i in b for i in other), len(other)

# ==================================================================================================== round 1: own scheme


def _own(letter, acts):
    st = load_state(letter)
    if os.path.exists(own_path(letter)):
        say(f"[{letter}] round 1 is already done.")
        return
    folder = s5path(letter, "folder")
    if not os.path.isdir(folder):
        raise AgentStop(f"Step 5's folder for this agent is not there ({folder})")
    st["stopped"] = None
    s3 = step3(letter)
    ids = {t["id"]: t["theme"] for t in final_themes(letter)}
    names = [a["Act"] for a in acts]
    known = memo_known6(letter, "own")
    text = fill(PROMPTS[1], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"],
                             "MEMO5P1": memo5(letter, True), "MEMORY": running_memo6(letter, "own")[0],
                             "FINAL": final_block6(letter), "ACTS": acts_block6(letter, acts)})
    out, hard, soft, notes, res = ask_until_usable(
        st, "own", {"fixed": text, "memory": [], "tail": ""}, None,
        lambda o: check_own(o, ids, names, act_tids(acts), known, all_tids(acts)),
        {"label": "round 1 (own scheme)", "log": "own", "round": "own"}, "in the JSON format given above",
        tools=folder_tools(folder))
    write_json(own_path(letter), dict(out, call=res["call_id"], problems=soft, notes=notes))
    flag(st, "own", "", soft + notes, res["call_id"])
    save_state(st)
    n_in = sum(1 for x in out["acts"] if x["types"])
    say(f"[{letter}] round 1 done: {len(out['dimensions'])} dimensions, {len(out['types'])} types; {n_in} of "
        f"{len(acts)} acts in at least one type.")


def run_own(agents=None):
    """Round 1: each agent, on its own, builds its dimensions and types and places every act (Prompt 1)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts()
    check_settings("own", sel, agents)
    return run_agents(_own, agents, sel)

# ==================================================================================================== round 2: the pooled scheme and the vote


def build_pool():
    """The three agents' dimensions and types in one list, D1 ... and T1 ..., in a neutral order (the same on every
    run of the same Step 5 work) and without the agents' letters, which pool.json keeps for the record only."""
    path = rpath("merged", "pool.json")
    if os.path.exists(path):
        return read_json(path)
    left = [L for L in LETTERS if not os.path.exists(own_path(L))]
    if left:
        raise SystemExit(f"STOPPED: round 1 (own scheme) is not done for agent(s) {', '.join(left)}; run run_own() "
                         f"again first.")
    dims, types = [], []
    for L in LETTERS:
        o = read_json(own_path(L))
        th = {t["id"]: t for t in final_themes(L)}
        for d in o["dimensions"]:
            dims.append({"author": L, "name": d["dimension"], "definition": d["definition"],
                         "boundaries": d["boundaries"],
                         "themes": [{"theme": th[x]["theme"], "definition": th[x]["definition"]}
                                    for x in d["themes"] if x in th]})
        for ty in o["types"]:
            dims_names = ty["dimensions"]
            types.append({"author": L, "name": ty["type"], "definition": ty["definition"],
                          "boundaries": ty["boundaries"], "dimensions": dims_names,
                          "themes": [th[x]["theme"] for x in ty["themes"] if x in th],
                          "acts": [{"act": p["act"], "fit": y["fit"]} for p in o["acts"] for y in p["types"]
                                   if y["type"] == ty["type"]]})
    rng = random.Random(f"{SHUFFLE_SEED}:{STEP5_RUN}")
    rng.shuffle(dims)
    rng.shuffle(types)
    for i, x in enumerate(dims, 1):
        x["id"] = f"D{i}"
    for i, x in enumerate(types, 1):
        x["id"] = f"T{i}"
    pool = {"dimensions": dims, "types": types, "seed": f"{SHUFFLE_SEED}:{STEP5_RUN}", "time": now()}
    write_json(path, pool)
    say(f"Pooled {len(dims)} dimensions and {len(types)} types from the three agents: {path}")
    return pool


def pool_block_d(pool):
    """The pooled dimensions as round 2 shows them: name, definition and boundaries only. The themes behind each are
    left out, because a voter would know its own themes by name and so know which items are its own."""
    return "\n\n".join(f"{x['id']}: {x['name']}\nDefinition: {x['definition']}\nBoundaries: "
                       f"{x['boundaries'] or 'none given'}" for x in pool["dimensions"]) or "none"


def pool_block_t(pool):
    """The pooled types as round 2 shows them, without the themes behind them (as for the dimensions)."""
    out = []
    for x in pool["types"]:
        acts = "; ".join(f"{y['act']} ({y['fit']})" for y in x["acts"]) or "none"
        out.append(f"{x['id']}: {x['name']}\nDefinition: {x['definition']}\nBoundaries: {x['boundaries'] or 'none given'}"
                   f"\nDefined by the dimensions: {'; '.join(x['dimensions']) or 'none given'}"
                   f"\nActs its author placed in it: {acts}")
    return "\n\n".join(out) or "none"


def _vote(letter, acts, pool):
    st = load_state(letter)
    if os.path.exists(votes_path(letter)):
        say(f"[{letter}] round 2 is already done.")
        return
    if not os.path.exists(own_path(letter)):
        raise AgentStop("round 1 (own scheme) is not done for this agent")
    st["stopped"] = None
    s3 = step3(letter)
    ids = {t["id"]: t["theme"] for t in final_themes(letter)}
    names = [a["Act"] for a in acts]
    known = memo_known6(letter, "vote")
    text = fill(PROMPTS[2], {"GUIDE": load_guide(), "LETTER": letter, "MEMO3": s3["memo"], "MEMO5": memo5(letter),
                             "MEMORY": running_memo6(letter, "vote")[0], "POOLD": pool_block_d(pool),
                             "POOLT": pool_block_t(pool), "FINALSHORT": final_block6(letter, True),
                             "ACTS": acts_block6(letter, acts)})
    out, hard, soft, notes, res = ask_until_usable(
        st, "vote", {"fixed": text, "memory": [], "tail": ""}, None,
        lambda o: check_votes(o, pool, ids, names, known, all_tids(acts)),
        {"label": "round 2 (shared scheme)", "log": "vote", "round": "vote"}, "in the JSON format given above",
        tools=folder_tools(s5path(letter, "folder")))
    write_json(votes_path(letter), dict(out, call=res["call_id"], problems=soft, notes=notes))
    flag(st, "vote", "", soft + notes, res["call_id"])
    save_state(st)
    log = save_memory6(letter)
    c = collections.Counter(x["decision"] for x in out["items"])
    ob, on, xb, xn = backing(letter, pool, out["items"])
    say(f"[{letter}] round 2 done: keep {c['keep']}, same {c['same']}, drop {c['drop']}; it backs {ob} of its own "
        f"{on} items and {xb} of the other agents' {xn}; running memo: {sum(x['change'] == 'added' for x in log)} "
        f"entries added in Step 6, saved in agent_{letter}/memory.txt.")


def run_vote(agents=None):
    """Round 2: the script pools the agents' dimensions and types; each agent judges every item (Prompt 2)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts()
    check_settings("vote", sel, agents)
    pool = build_pool()
    return run_agents(_vote, agents, sel, pool)

# ==================================================================================================== the count


class _Union:
    def __init__(self):
        self.up = {}

    def find(self, x):
        self.up.setdefault(x, x)
        while self.up[x] != x:
            self.up[x] = self.up[self.up[x]]
            x = self.up[x]
        return x

    def join(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.up[max(ra, rb, key=_order)] = min(ra, rb, key=_order)


def _order(i):
    return (i[0], int(i[1:]))


def shared_scheme():
    """Counts round 2's votes. Two items count as one when at least AGREE agents mark them the same (directly, or
    through their own chain of marks); an agent backs an item it keeps or marks the same, and the item another of its
    marks names; a concept (one item or several counted as one) stands when at least AGREE agents back it, under the
    name of the item most agents chose (kept, or named in a mark as the one that puts the concept better), then of the
    one with most backers; an act stands in an agreed type when at least AGREE of the agents that back the type place
    it there. An agent's themes and acts count for every item it puts together with the one it lists them under, as
    Prompt 2 tells it ("give the themes or acts for the two once, under either of them"). An agent backs an item only
    with evidence (backed_items): a vote to keep without it is recorded but not counted. A type also needs at least
    MIN_TYPE_ACTS agreed acts, or it falls away; one that at least AGREE agents say rests on what round 3 measures
    (splits, returns, stops, longevity, output or later success) is set aside, since comparing it would be circular;
    one that leaves fewer than MIN_TYPE_ACTS acts outside it is kept as a pattern that nearly every act shows, apart
    from the types, since it sets no acts apart to compare."""
    path = rpath("merged", "shared_scheme.json")
    if os.path.exists(path):
        s = read_json(path)
        if s.get("agree") != AGREE:
            raise SystemExit(f"STOPPED: the shared scheme of run {RUN!r} was counted with AGREE = {s.get('agree')}; put "
                             f"that back, or set RUN to a new name to start a new run.")
        return s
    pool = read_json(rpath("merged", "pool.json"))
    left = [L for L in LETTERS if not os.path.exists(votes_path(L))]
    if not pool or left:
        raise SystemExit(f"STOPPED: round 2 is not done for agent(s) {', '.join(left) or '?'}; run run_vote() again.")
    item = {x["id"]: x for x in pool["dimensions"] + pool["types"]}
    ids = sorted(item, key=_order)
    votes = {L: {v["id"]: v for v in read_json(votes_path(L))["items"]} for L in LETTERS}
    group = {L: {i: g for g in own_groups(votes[L]) for i in g} for L in LETTERS}     # each agent's own groups
    pair = collections.Counter()
    backs = {L: backed_items(votes[L]) for L in LETTERS}
    unproven = collections.defaultdict(set)        # item -> agents that kept it without evidence (not counted)
    chosen = collections.defaultdict(set)          # item -> the agents that keep it or name it in a mark
    for L in LETTERS:
        for v in votes[L].values():
            if v["decision"] == "keep":
                chosen[v["id"]].add(L)
            if v["decision"] == "same":
                chosen[v["same_as"]].add(L)
            if v["decision"] != "drop" and v["id"] not in backs[L]:
                unproven[v["id"]].add(L)
        for g in {tuple(g) for g in group[L].values()}:
            for a, b in itertools.combinations(g, 2):
                pair[(a, b)] += 1
    u = _Union()
    merges = []
    for (a, b), n in sorted(pair.items(), key=lambda kv: (_order(kv[0][0]), _order(kv[0][1]))):
        if n >= AGREE:
            u.join(a, b)
            merges.append({"items": [a, b], "agents": sorted(L for L in LETTERS if b in group[L].get(a, []))})
    concepts = collections.defaultdict(list)
    for i in ids:
        concepts[u.find(i)].append(i)

    def listed(L, members, key):
        """What agent L lists under these items, or under any item it puts together with one of them."""
        out_ = []
        for i in members:
            for j in group[L].get(i, [i]):
                out_ += (votes[L].get(j) or {}).get(key, [])
        return out_

    n_acts = len(select_acts())
    out = {"dimensions": [], "types": [], "patterns": [], "outcome_types": [], "fell_away": [], "merges": merges,
           "agree": AGREE,
           "min_type_acts": MIN_TYPE_ACTS, "script": VERSION, "time": now()}
    for members in sorted(concepts.values(), key=lambda m: _order(m[0])):
        backers = sorted(L for L in LETTERS if backs[L] & set(members))
        rep = max(members, key=lambda i: (len(chosen[i]), sum(i in backs[L] for L in LETTERS),
                                          [-x for x in _order(i)[1:]]))
        x = item[rep]
        rec = {"items": members, "name": x["name"], "definition": x["definition"], "boundaries": x["boundaries"],
               "backers": backers, "kind": "dimension" if rep.startswith("D") else "type", "named_after": rep,
               "members": [{"id": i, "name": item[i]["name"], "chosen_by": sorted(chosen[i])} for i in members],
               "kept_without_evidence": sorted(set().union(*(unproven[i] for i in members)) - set(backers))}
        if len(backers) < AGREE:
            rec["why"] = f"backed by {len(backers) or 'no'} agent{'' if len(backers) == 1 else 's'}"
            out["fell_away"].append(rec)
            continue
        if rep.startswith("D"):
            rec["themes"] = {L: list(dict.fromkeys(listed(L, members, "themes"))) for L in backers}
            out["dimensions"].append(rec)
            continue
        placed = collections.defaultdict(dict)
        for L in backers:
            for y in listed(L, members, "acts"):
                f = placed[y["act"]]
                f[L] = "clear" if "clear" in (f.get(L), y["fit"]) else y["fit"]
        rec["acts"] = [{"act": a, "agents": sorted(f), "fits": f} for a, f in placed.items() if len(f) >= AGREE]
        rec["one_agent"] = [{"act": a, "agent": next(iter(f)), "fit": next(iter(f.values()))}
                            for a, f in placed.items() if len(f) < AGREE]
        rec["outcome_flags"] = sorted(L for L in LETTERS if any((votes[L].get(i) or {}).get("outcomes") for i in members))
        if len(rec["acts"]) < MIN_TYPE_ACTS:       # its backers agree on too few of its acts to share the type
            rec["why"] = (f"backed by {len(backers)} agents, who agree on only {len(rec['acts'])} of its acts (a type "
                          f"needs {MIN_TYPE_ACTS})")
            out["fell_away"].append(rec)
        elif len(rec["outcome_flags"]) >= AGREE:   # it rests on what round 3 measures: comparing would be circular
            out["outcome_types"].append(rec)
        elif n_acts - len(rec["acts"]) < MIN_TYPE_ACTS:   # it sets (almost) no act apart
            out["patterns"].append(rec)
        else:
            out["types"].append(rec)
    for k, key in (("dimensions", "SD"), ("types", "ST"), ("patterns", "SP"), ("outcome_types", "SO")):
        for n, rec in enumerate(out[k], 1):
            rec["id"] = f"{key}{n}"
    order = {a["Act"]: i for i, a in enumerate(select_acts())}
    for t in out["types"] + out["patterns"] + out["outcome_types"] + [x for x in out["fell_away"] if "acts" in x]:
        t["acts"].sort(key=lambda x: order.get(x["act"], 999))
        t["one_agent"].sort(key=lambda x: order.get(x["act"], 999))
    write_json(path, out)
    write_text(rpath("merged", "shared_scheme.txt"), shared_text(out) + "\n")
    say(f"Shared scheme: {len(out['dimensions'])} dimensions and {len(out['types'])} types stand (backed by at least "
        f"{AGREE} agents" + (f"; {len(out['patterns'])} agreed patterns that nearly every act shows are kept apart from "
                             f"the types" if out["patterns"] else "") +
        (f"; {len(out['outcome_types'])} agreed types are set aside because they rest on what round 3 measures"
         if out["outcome_types"] else "") +
        f"); {len(out['fell_away'])} items fell away; {len(merges)} pairs counted as one. {path}")
    return out


def _pooled(rec):
    """Where an agreed item came from in the pool: its pooled items and, where several count as one, the names of
    those it is not named after."""
    others = [m for m in rec.get("members", []) if m["id"] != rec.get("named_after")]
    s = f"pooled as {', '.join(rec['items'])}"
    if others:
        s += (f"; named after {rec['named_after']}, with " + "; ".join(f"{m['id']} ({m['name']})" for m in others)
              + " counted as the same")
    return s


def shared_text(s):
    """The shared scheme as round 3 shows it."""
    names = {L: {t["id"]: t["theme"] for t in final_themes(L)} for L in LETTERS}
    lines = ["AGREED DIMENSIONS"]
    for d in s["dimensions"]:
        lines.append(f"\n{d['id']}: {d['name']} (backed by {len(d['backers'])} of the three agents; {_pooled(d)})"
                     f"\nDefinition: {d['definition']}\nBoundaries: {d['boundaries'] or 'none given'}")
        for L, th in d["themes"].items():
            lines.append(f"Agent {L}'s themes under it: " + ("; ".join(names[L].get(t, t) for t in th) or "none"))
    def type_lines(t):
        acts = "; ".join(f"{x['act']} (placed there by {len(x['agents'])})" for x in t["acts"]) or "none"
        one = "; ".join(x["act"] for x in t["one_agent"]) or "none"
        thin = f" (fewer than {MIN_TYPE_ACTS} acts agreed: too few to compare)" if t.get("thin") else ""
        return (f"\n{t['id']}: {t['name']} (backed by {len(t['backers'])} of the three agents; {_pooled(t)})"
                f"\nDefinition: {t['definition']}\nBoundaries: {t['boundaries'] or 'none given'}"
                f"\nActs in it (placed there by at least {s['agree']} agents): {acts}{thin}"
                f"\nActs placed there by one agent only: {one}")
    lines.append("\nAGREED TYPES")
    lines += [type_lines(t) for t in s["types"]] or ["none"]
    if s.get("outcome_types"):
        lines.append(f"\nAGREED TYPES SET ASIDE (agreed as types, but at least {s['agree']} agents judged that they rest "
                     "on what the tables measure: whether an act split, came back or stopped, how long it lasted, how "
                     "much it released, or its later success. Comparing them on those measures would be circular, so "
                     "the tables leave them out)")
        lines += [type_lines(t) for t in s["outcome_types"]]
    if s.get("patterns"):
        lines.append("\nAGREED PATTERNS THAT NEARLY EVERY ACT SHOWS (agreed as types, but they leave fewer than "
                     f"{s.get('min_type_acts', MIN_TYPE_ACTS)} acts outside them, so they set no acts apart and the "
                     "tables leave them out)")
        lines += [type_lines(t) for t in s["patterns"]]
    lines.append(f"\nITEMS THAT FELL AWAY (backed by fewer than {s['agree']} agents, or, for a type, with fewer than "
                 f"{s.get('min_type_acts', MIN_TYPE_ACTS)} acts that its backers agree on)")
    lines += [f"- {x['kind']} {', '.join(x['items'])}: {x['name']} ("
              + (x.get("why") or f"backed by {len(x['backers']) or 'no'} agent{'' if len(x['backers']) == 1 else 's'}")
              + ")" for x in s["fell_away"]] or ["- none"]
    return "\n".join(lines)

# ==================================================================================================== the performance data


def load_dv():
    """The performance data, act by act (by act_key): each measure of MEASURES, from the Acts sheet, and the US
    studio albums after the ceremony counted from the Albums sheet (released after it, and released in the US)."""
    from openpyxl import load_workbook
    p = os.path.join(BASE, INPUT_DIR, DV_FILE)
    if not os.path.exists(p):
        raise SystemExit(f"STOPPED: the performance data {p} not found.")
    wb = load_workbook(p, read_only=True, data_only=True)
    if "Acts" not in wb.sheetnames:
        raise SystemExit(f"STOPPED: {DV_FILE} has no sheet called Acts.")
    rows = list(wb["Acts"].iter_rows(values_only=True))
    col = {str(h or "").strip(): i for i, h in enumerate(rows[0])}
    need = ["Act"] + [m[2] for m in MEASURES if m[2]]
    miss = [h for h in need if h not in col]
    if miss:
        raise SystemExit(f"STOPPED: {DV_FILE}, sheet Acts, has no column {miss}.")
    us = collections.Counter()
    if "Albums" in wb.sheetnames:
        ar = list(wb["Albums"].iter_rows(values_only=True))
        ah = {str(h or "").strip(): i for i, h in enumerate(ar[0])}
        if all(k in ah for k in ("Act", "Released after the ceremony?", "Released in the US?")):
            for r in ar[1:]:
                if (r[ah["Act"]] and str(r[ah["Released after the ceremony?"]] or "").strip().lower() == "yes"
                        and str(r[ah["Released in the US?"]] or "").strip().lower() == "yes"):
                    us[act_key(r[ah["Act"]])] += 1
    out = {}
    for r in rows[1:]:
        if not r or not r[col["Act"]]:
            continue
        k = act_key(r[col["Act"]])
        rec = {"act": str(r[col["Act"]]).strip()}
        for key, name, c, low in MEASURES:
            if c is None:
                v = float(us.get(k, 0))
            else:
                v = r[col[c]]
                try:
                    v = None if v in (None, "") else float(v)
                except (TypeError, ValueError):
                    v = None
            if key == "peak" and v is None:
                v = float(NO_CHART)            # no album on the Billboard 200 after the ceremony
            rec[key] = v
        out[k] = rec
    return out


def _median(xs, low=False):
    """The median; with low, the lower of the two middle values when there is an even number of them, so that the
    median best peak is a real chart position (or no chart), never an average with an act that did not chart."""
    xs = sorted(xs)
    n = len(xs)
    if not n:
        return None
    if n % 2 or low:
        return xs[(n - 1) // 2]
    return (xs[n // 2 - 1] + xs[n // 2]) / 2


def _fmt(v, key):
    if v is None:
        return "n/a"
    if key == "peak" and v >= NO_CHART:
        return "none"
    return f"{v:.1f}".rstrip("0").rstrip(".") if v != int(v) else str(int(v))


def _summ(vals, key):
    vals = [v for v in vals if v is not None]
    if not vals:
        return "n/a"
    s = f"{_fmt(_median(vals, key == 'peak'), key)} ({_fmt(min(vals), key)} to {_fmt(max(vals), key)})"
    if key == "peak":
        s += f", {sum(v < NO_CHART for v in vals)} of {len(vals)} on the chart"
    return s


def measures_text():
    return ("The measures, all counted from the act's Best New Artist ceremony: studio albums released after it, all of "
            "them and those that the sources show were released in the US (Harrison et al. counted only US albums); "
            "albums that reached the Billboard "
            "200 after it, and the best peak any of them reached (1 is the top of the chart; \"none\" means that no "
            "album reached the chart); top-40 albums after it; RIAA-certified albums after it, and their RIAA units in "
            "millions (Gold is 0.5, Platinum 1, n-times Platinum n, Diamond 10); Grammy nominations and wins credited "
            "to the act at later ceremonies; and years active after it, from the ceremony year to the year of the "
            "act's last studio album, counting both, less the years between a breakup and the reunion that ended it.")


def performance_tables(acts=None):
    """Round 3's tables, by agreed type and by act (merged/performance_by_type.csv, performance_by_act.csv and
    performance_tables.txt, which round 3 shows)."""
    sel = acts or select_acts()
    s = shared_scheme()
    dv = load_dv()
    missing = [a["Act"] for a in sel if act_key(a["Act"]) not in dv]
    if missing:
        raise SystemExit(f"STOPPED: {DV_FILE} has no row for: {', '.join(missing)}.")
    types_of = collections.defaultdict(list)
    for t in s["types"]:
        for x in t["acts"]:
            types_of[x["act"]].append(t["id"])
    by_type, lines = [], []
    lines.append(f"BY AGREED TYPE (median, and in brackets the lowest and highest value; {len(sel)} acts in all)")
    for t in s["types"]:
        inside = {x["act"] for x in t["acts"]}
        ins = [dv[act_key(a["Act"])] for a in sel if a["Act"] in inside]
        outs = [dv[act_key(a["Act"])] for a in sel if a["Act"] not in inside]
        few = min(len(ins), len(outs))
        note = (" No act is agreed in it, so there is nothing to compare." if not ins else
                " Every act is in it, so there is nothing to compare it with." if not outs else
                f" With only {few} act{'' if few == 1 else 's'} on one side, a median says little." if few < MIN_TYPE_ACTS
                else "")
        lines.append(f"\n{t['id']}: {t['name']}: {len(ins)} acts in the type, {len(outs)} outside it.{note}")
        for key, name, c, low in MEASURES:
            a_, b_ = _summ([r[key] for r in ins], key), _summ([r[key] for r in outs], key)
            if ins and outs:
                lines.append(f"- {name}{' (1 is the top; the median is the lower middle value)' if low else ''}: in "
                             f"the type {a_}; outside it {b_}")
            by_type.append([t["id"], t["name"], name, len(ins), a_, len(outs), b_])
    none = [a["Act"] for a in sel if not types_of.get(a["Act"])]
    lines.append("\nActs in no agreed type: " + ("; ".join(none) or "none"))
    lines.append("\nBY ACT")
    head = ["Act", "Group or solo", "Ceremony", "Best New Artist", "Agreed types"] + [m[1] for m in MEASURES]
    by_act = []
    for a in sel:
        r = dv[act_key(a["Act"])]
        row = [a["Act"], kind(a), a["Nominated"], "won" if won(a) else "nominated",
               "; ".join(types_of.get(a["Act"], [])) or "none"] + [_fmt(r[m[0]], m[0]) for m in MEASURES]
        by_act.append(row)
        lines.append(f"- {a['Act']} ({kind(a)}, {a['Nominated']} ceremony, {'won' if won(a) else 'nominated'}; types: "
                     f"{row[4]}): " + "; ".join(f"{m[1]} {v}" for m, v in zip(MEASURES, row[5:])))
    write_csv(rpath("merged", "performance_by_type.csv"),
              ["Type", "Name", "Measure", "Acts in it", "Median (range) in it", "Acts outside it",
               "Median (range) outside it"], by_type)
    write_csv(rpath("merged", "performance_by_act.csv"), head, by_act)
    write_text(rpath("merged", "performance_tables.txt"), "\n".join(lines) + "\n")
    return "\n".join(lines), by_type, head, by_act

# ==================================================================================================== round 3: later careers and the Step 6 memo


def own_block(letter):
    o = read_json(own_path(letter))
    names = {t["id"]: t["theme"] for t in final_themes(letter)}
    lines = ["YOUR DIMENSIONS"]
    for d in o["dimensions"]:
        lines.append(f"- {d['dimension']}: {d['definition']} Themes: " + "; ".join(names.get(x, x) for x in d["themes"]))
    lines.append("\nYOUR TYPES")
    for t in o["types"]:
        lines.append(f"- {t['type']}: {t['definition']}")
    lines.append("\nYOUR PLACEMENTS")
    for p in o["acts"]:
        tys = "; ".join(f"{y['type']} ({y['fit']})" for y in p["types"]) or "no type"
        lines.append(f"- {p['act']}: {tys}. {p['why']}")
    if o.get("misfits"):
        lines.append("\nYOUR MISFITS")
        lines += [f"- {m['act']} ({m['type']}, {m['kind']}): {m['what']}" for m in o["misfits"]]
    return "\n".join(lines)


def step5_names(acts):
    """Names a Step 6 memo may give in quotation marks without being a quotation: the agents' Step 5 themes (final,
    column and within-case), groupings, groups and dimensions, and Step 6's dimensions and types."""
    out = []
    for L in LETTERS:
        f = read_json(s5path(L, "final.json")) or {}
        out += [t["theme"] for t in f.get("themes", [])] + [t["theme"] for t in final_themes(L)]
        g = read_json(s5path(L, "groupings.json")) or {}
        for x in g.get("groupings", []):
            out += [x["grouping"], x["dimension"]] + [grp["group"] for grp in x["groups"]]
        for n in range(1, len(g.get("groupings", [])) + 1):
            out += [t["theme"] for t in (read_json(s5path(L, f"comparison_{n}.json")) or {}).get("themes", [])]
        for a in acts:
            out += [t["theme"] for t in (read_json(s5path(L, "within", f"{act_stem(a)}.json")) or {}).get("themes", [])]
        o = read_json(own_path(L)) or {}
        out += [d["dimension"] for d in o.get("dimensions", [])] + [t["type"] for t in o.get("types", [])]
    out += [a["Act"] for a in acts]
    return [x for x in out if x]


def word_problem(text, limit, what):
    n = nwords(text)
    return [f"{what} has {n:,} words; keep it to about {limit:,}."] if n > limit * WORDS_SLACK else []


PACKED = re.compile(r"(?<![\w./-])\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)+(?![\w/-]|\.\d)")


NOT_PACKED = re.compile(r"24/7|(?:19|20)\d\d/(?:\d\d|(?:19|20)\d\d)|\d\d?/\d\d?/(?:\d\d|\d{4})")
# not packed figures: "24/7", year spans such as "1985/86", and dates such as "10/30/2026"


def packed_problems(text):
    """Figures packed together with slashes ("4.5/3.5 versus 3/2.5"), which a reader cannot follow."""
    hits = [m.group(0) for m in PACKED.finditer(text) if not NOT_PACKED.fullmatch(m.group(0))]
    if len(hits) < 3:
        return []
    return [f"The memo packs figures together with slashes ({', '.join(hits[:4])}{' …' if len(hits) > 4 else ''}): "
            f"write each figure in a plain sentence with the name of its measure, and keep to the measures that matter "
            f"for the point."]


SCHEME_ID = re.compile(r"\bS([DTPO])\s?(\d+)\b")
ID_NAME = [re.compile(r"\*\s*(S[DTPO]\s?\d+)\s*[:.]\s*([^*\n]{3,150}?)\s*\*"),     # *SD3: name*
           re.compile(r"(?<!\*)\b(S[DTPO]\s?\d+)\s*[:,]?\s*\*([^*\n]{3,150}?)\*"),  # SD3 *name*, SD3: *name*
           re.compile(r"\*([^*\n]{3,150}?)\*\s*\((S[DTPO]\s?\d+)\)")]              # *name* (SD3)


def _name_fits(given, names):
    g = label_key(given)
    return any(g == n or n.startswith(g) or g.startswith(n) or (len(g) >= 12 and g in n) for n in names)


def id_problems(text, s):
    """The memo's IDs of agreed dimensions and types: each must exist, and where the memo gives a name with it, the
    name must not be another agreed item's (a name that fits no agreed item, such as a shortened one, passes)."""
    items = {x["id"]: x for x in s["dimensions"] + s["types"] + s.get("patterns", []) + s.get("outcome_types", [])}
    names = {i: [label_key(m["name"]) for m in x.get("members", [])] + [label_key(x["name"])] for i, x in items.items()}
    used = sorted({f"S{m.group(1)}{int(m.group(2))}" for m in SCHEME_ID.finditer(text)}, key=lambda i: (i[1], int(i[2:])))
    has = f"SD1 to SD{len(s['dimensions'])} and ST1 to ST{len(s['types'])}" + (
        f", and SP1 to SP{len(s['patterns'])}" if s.get("patterns") else "") + (
        f", and SO1 to SO{len(s['outcome_types'])}" if s.get("outcome_types") else "")
    p = [f"{i} is not the ID of an agreed dimension, type or pattern: the shared scheme has {has}." for i in used
         if i not in items]
    seen = set()
    for k, rx in enumerate(ID_NAME):
        for m in rx.finditer(text):
            i, given = (m.group(2), m.group(1)) if k == 2 else (m.group(1), m.group(2))
            i = re.sub(r"\s", "", i)
            given = given.strip(" .,;:")
            if i not in items or (i, label_key(given)) in seen or _name_fits(given, names[i]):
                continue
            seen.add((i, label_key(given)))
            other = [j for j in items if j != i and _name_fits(given, names[j])]
            if other:
                p.append(f"{i} is {items[i]['name']}, not {short(given, 80)}, which is {other[0]}: give each agreed "
                         f"item its own ID, as the shared scheme does.")
    return p


def _careers(letter, acts):
    st = load_state(letter)
    if os.path.exists(memo6_path(letter)):
        say(f"[{letter}] round 3 is already done.")
        return
    if not os.path.exists(votes_path(letter)):
        raise AgentStop("round 2 (shared scheme) is not done for this agent")
    st["stopped"] = None
    s3 = step3(letter)
    text = fill(PROMPTS[3], {"GUIDE": load_guide(), "LETTER": letter, "MEASURES": measures_text(), "MEMO3": s3["memo"],
                             "MEMO5": memo5(letter), "MEMORY": running_memo6(letter, "careers")[0],
                             "OWN": own_block(letter), "SHARED": read_text(rpath("merged", "shared_scheme.txt")),
                             "TABLES": read_text(rpath("merged", "performance_tables.txt"))})
    scope = sorted(all_tids(acts))
    ok = [x for L in LETTERS for x in run_labels(L, acts).values()] + step5_names(acts)
    ok += [x["name"] for x in (read_json(rpath("merged", "pool.json")) or {}).get("dimensions", [])]
    ok += [x["name"] for x in (read_json(rpath("merged", "pool.json")) or {}).get("types", [])]
    shared = shared_scheme()

    def check(t):
        p = []
        m = [PART_RE[n].search(t) for n in (1, 2, 3)]
        if not all(m):
            p.append('The memo must have its three parts, starting with the headings "Part 1", "Part 2" and "Part 3".')
        elif not m[0].start() < m[1].start() < m[2].start():
            p.append("The parts must come in order: Part 1, then Part 2, then Part 3.")
        else:
            p += word_problem(t[m[0].start():m[1].start()], MEMO6_WORDS[1], "Part 1")
            p += word_problem(t[m[1].start():m[2].start()], MEMO6_WORDS[2], "Part 2")
            p += word_problem(t[m[2].start():], MEMO6_WORDS[3], "Part 3")
        return p + quote_problems(t, scope, ok) + id_problems(t, shared) + packed_problems(t)
    res, probs = ask_text(st, "careers", {"fixed": text, "memory": [], "tail": ""}, folder_tools(s5path(letter, "folder")),
                          check, {"round": "careers", "label": "round 3 (later careers and memo)", "log": "careers"})
    write_text(memo6_path(letter), res["text"].strip() + "\n")
    flag(st, "careers", "", probs, res["call_id"])
    save_state(st)
    say(f"[{letter}] round 3 done: the Step 6 memo has {nwords(res['text']):,} words.")


def run_careers(agents=None):
    """Round 3: the script counts round 2's votes and tabulates the performance data by agreed type; each agent reads
    the tables against the fixed scheme and writes its Step 6 memo (Prompt 3)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts()
    check_settings("careers", sel, agents)
    shared_scheme()
    performance_tables(sel)
    return run_agents(_careers, agents, sel)

# ==================================================================================================== the merge


CELL_MAX = 32000


def _cell(v):
    if isinstance(v, (list, tuple)):
        v = "; ".join(str(x) for x in v)
    if isinstance(v, dict):
        v = "; ".join(f"{k}: {x}" for k, x in v.items())
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


def run_merge():
    """Every round's work and the counts in one workbook (merged/step6_combined.xlsx) and as CSV files in
    merged/folder, without changing anything."""
    load_corpus()
    sel = select_acts()
    left = [L for L in LETTERS if not os.path.exists(memo6_path(L))]
    if left:
        print(f"Not merged: round 3 is not done for agent(s) {', '.join(left)}.")
        return
    pool, s = read_json(rpath("merged", "pool.json")), shared_scheme()
    names = {L: {t["id"]: t["theme"] for t in final_themes(L)} for L in LETTERS}
    own_d, own_t, own_p, mis, votes, memo = [], [], [], [], [], []
    for L in LETTERS:
        o = read_json(own_path(L))
        for d in o["dimensions"]:
            own_d.append([L, d["dimension"], d["definition"], d["boundaries"], [names[L].get(x, x) for x in d["themes"]]])
        for u in o["unplaced_themes"]:
            own_d.append([L, "(fits no dimension)", u["why"], "", [names[L].get(u["theme"], u["theme"])]])
        for t in o["types"]:
            clear = [p["act"] for p in o["acts"] for y in p["types"] if y["type"] == t["type"] and y["fit"] == "clear"]
            part = [p["act"] for p in o["acts"] for y in p["types"] if y["type"] == t["type"] and y["fit"] == "partial"]
            own_t.append([L, t["type"], t["definition"], t["boundaries"], t["dimensions"],
                          [names[L].get(x, x) for x in t["themes"]], clear, part])
        for p in o["acts"]:
            own_p.append([L, p["act"], [f"{y['type']} ({y['fit']})" for y in p["types"]] or "none", p["why"],
                          p["supporting"], p["contrary"]])
        for m in o["misfits"]:
            mis.append([L, m["act"], m["type"], m["kind"], m["what"]])
        v = read_json(votes_path(L))
        nm = {x["id"]: x["name"] for x in pool["dimensions"] + pool["types"]}
        for it in v["items"]:
            votes.append([L, it["id"], nm.get(it["id"], ""), it.get("closest") or "", it["decision"],
                          it.get("same_as") or "", {True: "yes", False: "no"}.get(it.get("outcomes"), ""), it["why"],
                          [names[L].get(x, x) for x in it.get("themes", [])],
                          [f"{y['act']} ({y['fit']})" for y in it.get("acts", [])]])
        for x in running_memo6(L, "end")[2]:
            memo.append([L, x["where"], x["change"], x["entry"], x["step"], x["before"], x["after"]])
    pool_rows = [[x["id"], "dimension", x["author"], x["name"], x["definition"], x["boundaries"],
                  [t["theme"] for t in x["themes"]]] for x in pool["dimensions"]] + \
                [[x["id"], "type", x["author"], x["name"], x["definition"], x["boundaries"],
                  [f"{y['act']} ({y['fit']})" for y in x["acts"]]] for x in pool["types"]]
    sd = [[d["id"], d["name"], d["definition"], d["boundaries"], _pooled(d), d["backers"]]
          + [[names[L].get(t, t) for t in d["themes"].get(L, [])] for L in LETTERS] for d in s["dimensions"]]
    status = {"types": "type", "patterns": "pattern nearly every act shows",
              "outcome_types": "set aside: rests on what round 3 measures"}
    stt = [[t["id"], status[k], t["name"], t["definition"], t["boundaries"], _pooled(t), t["backers"],
            [f"{x['act']} ({', '.join(f'{L} {f}' for L, f in x['fits'].items())})" for x in t["acts"]],
            [f"{x['act']} ({x['agent']})" for x in t["one_agent"]], t.get("outcome_flags", [])]
           for k in status for t in s.get(k, [])]
    fell = [[", ".join(x["items"]), x["kind"], x["name"], x["backers"], x.get("why", "")] for x in s["fell_away"]]
    place = []
    vote_acts = {L: collections.defaultdict(list) for L in LETTERS}
    for L in LETTERS:
        for it in read_json(votes_path(L))["items"]:
            for y in it.get("acts", []):
                vote_acts[L][y["act"]].append(it["id"])
    for a in sel:
        agreed = [t["id"] for t in s["types"] if a["Act"] in {x["act"] for x in t["acts"]}]
        place.append([a["Act"], agreed or "none"] + [vote_acts[L].get(a["Act"], []) or "none" for L in LETTERS])
    _, by_type, head, by_act = performance_tables(sel)
    sheets = [
        ("Own dimensions", ["Agent", "Dimension", "Definition", "Boundaries", "Themes"], own_d, [7, 30, 50, 40, 60]),
        ("Own types", ["Agent", "Type", "Definition", "Boundaries", "Dimensions", "Themes", "Acts (clear)",
                       "Acts (partly)"], own_t, [7, 30, 50, 40, 30, 40, 40, 30]),
        ("Own placements", ["Agent", "Act", "Types", "Why", "Supporting Turn IDs", "Contrary Turn IDs"], own_p,
         [7, 22, 40, 60, 24, 20]),
        ("Misfits", ["Agent", "Act", "Type", "Kind", "What"], mis, [7, 22, 30, 8, 60]),
        ("Pool", ["ID", "Kind", "Author", "Name", "Definition", "Boundaries", "Themes or acts"], pool_rows,
         [6, 10, 7, 30, 50, 40, 50]),
        ("Votes", ["Agent", "ID", "Name", "Closest", "Decision", "Same as", "Rests on round 3's measures", "Why",
                   "Own themes", "Acts"], votes, [7, 6, 30, 8, 9, 8, 10, 50, 40, 40]),
        ("Shared dimensions", ["ID", "Name", "Definition", "Boundaries", "Pooled items", "Backed by"]
         + [f"Agent {L}'s themes" for L in LETTERS], sd, [6, 30, 50, 40, 12, 10, 40, 40, 40]),
        ("Shared types", ["ID", "Status", "Name", "Definition", "Boundaries", "Pooled items", "Backed by",
                          "Acts (agreed)", "Acts placed by one agent", "Rests on round 3's measures, for"], stt,
         [6, 14, 30, 50, 40, 24, 10, 50, 40, 14]),
        ("Fell away", ["Pooled items", "Kind", "Name", "Backed by", "Why"], fell, [12, 10, 40, 10, 40]),
        ("Act placements", ["Act", "Agreed types"] + [f"Agent {L}'s pooled types" for L in LETTERS], place,
         [22, 20, 24, 24, 24]),
        ("Performance by type", ["Type", "Name", "Measure", "Acts in it", "Median (range) in it", "Acts outside it",
                                 "Median (range) outside it"], by_type, [6, 30, 34, 9, 18, 9, 18]),
        ("Performance by act", head, by_act, [22, 10, 9, 10, 14] + [12] * len(MEASURES)),
        ("Running memo", ["Agent", "Call", "Change", "Entry", "From step", "Before", "After"], memo,
         [7, 30, 9, 8, 8, 60, 60]),
    ]
    d, fd = ensure(rpath("merged")), ensure(rpath("merged", "folder"))
    write_text(os.path.join(d, "step6_memos.md"), memos_for_readers(s))
    write_xlsx(os.path.join(d, "step6_combined.xlsx"), sheets)
    for name, header, rows, _ in sheets:
        write_csv(os.path.join(fd, re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") + ".csv"), header, rows)
    write_json(os.path.join(d, "done.json"), {"acts": [act_stem(a) for a in sel], "time": now(), "script": VERSION,
                                             "step5_run": STEP5_RUN})
    say(f"Merge done: {os.path.join(d, 'step6_combined.xlsx')}.")


def demote_headings(text, by=2):
    """Markdown headings in a memo moved down by two levels, so that "# Part 1" sits under the memo's own heading in
    the reading copy; lines inside code blocks are left as they are."""
    out, code = [], False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            code = not code
        m = None if code else re.match(r"(#{1,6})(\s.*)$", line)
        out.append(("#" * min(6, len(m.group(1)) + by) + m.group(2)) if m else line)
    return "\n".join(out)


def memos_for_readers(s):
    """The three Step 6 memos as one file to read and send on (merged/step6_memos.md): first the shared scheme they
    cite, each item with its ID, definition and acts, then the memos, then the performance tables they read."""
    parts = [("dimensions", "Agreed dimensions", "the agreed dimensions (SD)"),
             ("types", "Agreed types", "the agreed types (ST)"),
             ("patterns", "Agreed patterns that nearly every act shows (not compared in the tables)",
              "the patterns that nearly every act shows (SP)"),
             ("outcome_types", "Agreed types set aside because they rest on what the tables measure (not compared in "
                               "the tables)", "the types set aside because they rest on what the tables measure (SO)")]
    have = [p for p in parts if s.get(p[0])]
    cited = [p[2] for p in have]
    cited = ", ".join(cited[:-1]) + " and " + cited[-1] if len(cited) > 1 else "".join(cited)
    out = [f"# Step 6 memos ({len(select_acts())} acts)", "",
           f"The three agents' Step 6 memos. They cite {cited} by the IDs defined below, with the acts placed in each "
           f"type. Turn IDs in brackets point to the turns of the corpus. The performance tables the memos read come "
           f"at the end.", "", "## The shared scheme", ""]
    for k, head, _ in have:
        out += [f"### {head}", ""]
        for x in s[k]:
            acts = "; ".join(a["act"] for a in x.get("acts", []))
            out.append(f"- **{x['id']} {x['name']}.** {x['definition']}" + (f" *Acts:* {acts}." if acts else ""))
        out.append("")
    for L in LETTERS:
        out += [f"## Agent {L}'s memo ({MODELS[L]['model']})", "", demote_headings(read_text(memo6_path(L)).strip()),
                ""]
    out += ["## The performance tables the memos read", "", "```", read_text(rpath("merged",
                                                                                    "performance_tables.txt")).strip(),
            "```", ""]
    return "\n".join(out)


def run_all(agents=None):
    """Every round in order; each skips what is done. Stops, with a message, where a round is not finished."""
    for name, fn in (("round 1", run_own), ("round 2", run_vote), ("round 3", run_careers)):
        res = fn(agents)
        if STOP.is_set():
            say(f"Stopped in {name}: {STOP_WHY[0]}. Run run_all() again to carry on.")
            return res
        bad = {L: r for L, r in (res or {}).items() if r != "done"}
        if bad or not res:
            say(f"Not every agent finished {name} ({'; '.join(f'{L}: {r}' for L, r in bad.items()) or 'not started'}). "
                f"Run run_all() again to carry on.")
            return res
    run_merge()
    say("All three rounds are done.")
    return {"done": True}

# ==================================================================================================== checks, tests and reports


def check_inputs():
    """Checks everything Step 6 reads, without any model call."""
    print(VERSION)
    for mod, want in (("anthropic", "2.3.0"), ("openai", "3.37.0"), ("google.genai", "2.33.0")):
        try:
            m = __import__(mod, fromlist=["x"])
            v = getattr(m, "__version__", "?")
            print(f"{mod} {v}" + ("" if v == want else f"  (the script was tested with {want}: run the pip line in the "
                                                       "notes at the top)"))
        except Exception as e:
            print(f"{mod}: NOT INSTALLED ({e}); run the pip line in the notes at the top")
    c = load_corpus()
    print(f"Corpus: {len(c['acts'])} acts, {len(c['by_id']):,} turns.")
    sel = select_acts()
    print(f"Step 5 run '{STEP5_RUN}': {len(sel)} acts.")
    for L in LETTERS:
        try:
            s3 = step3(L)
            th = final_themes(L)
            mem = memo5_base(L)
            summ = sum(1 for a in sel if summary_of(L, a) != "(no summary)")
            ok = os.path.isdir(s5path(L, "folder"))
            print(f"Agent {L}: Step 3 memo {nwords(s3['memo']):,} words; Step 5: {len(th)} final themes, memo part 1 "
                  f"{nwords(memo5(L, True)):,} words, all parts {nwords(memo5(L)):,} words, running memo "
                  f"{len(mem[1]):,} entries, summaries for {summ} of {len(sel)} acts, folder "
                  f"{'there' if ok else 'MISSING'}")
        except (Exception, SystemExit) as e:
            print(f"Agent {L}: PROBLEM: {e}")
    for n in PROMPT_FILES:
        try:
            s = read_prompt(n)
            pr = check_prompt(n, s)
            print(f"Prompt {n} ({PROMPT_FILES[n]}): {nwords(s):,} words" + (f"; PROBLEMS: {pr}" if pr else "; fine"))
        except (Exception, SystemExit) as e:
            print(f"Prompt {n}: {e}")
    g = load_guide()
    print(f"Guide ({GUIDE_FILE}): {nwords(g):,} words, about {est_tokens(g):,} tokens.")
    try:
        dv = load_dv()
        miss = [a["Act"] for a in sel if act_key(a["Act"]) not in dv]
        print(f"Performance data ({DV_FILE}): {len(dv)} acts" + (f"; NO ROW FOR: {', '.join(miss)}" if miss else
                                                                  ", every act of the run included"))
    except (Exception, SystemExit) as e:
        print(f"Performance data: PROBLEM: {e}")
    for L in LETTERS:
        try:
            secret(SECRETS[MODELS[L]["provider"]])
            print(f"Agent {L}: {MODELS[L]['model']} ({MODELS[L]['provider']}), effort {MODELS[L].get('effort')}: key "
                  f"found")
        except Exception as e:
            print(f"Agent {L}: PROBLEM: {e}")
    print(f"Run '{RUN}' on Step 5 run '{STEP5_RUN}'; the run folder is {rpath()}; budget ${BUDGET_USD:,.0f}.")


def dry_run():
    """Writes each agent's round 1 prompt, filled, to step6_<RUN>/dry_run, with its size (no model call)."""
    load_corpus(), load_prompts(), load_guide()
    sel = select_acts()
    d = ensure(rpath("dry_run"))
    for L in LETTERS:
        s3 = step3(L)
        text = fill(PROMPTS[1], {"GUIDE": load_guide(), "LETTER": L, "MEMO3": s3["memo"], "MEMO5P1": memo5(L, True),
                                 "MEMORY": running_memo6(L, "own")[0], "FINAL": final_block6(L),
                                 "ACTS": acts_block6(L, sel)})
        write_text(os.path.join(d, f"prompt1_agent_{L}.txt"), text)
        n = own_tokens(L, est_tokens(text))
        lim = CONTEXT_ROOM * CONTEXT[MODELS[L]["model"]] - OUTPUT_ROOM
        print(f"Agent {L}: round 1 prompt about {n:,} tokens of its model ({'fits' if n <= lim else 'TOO LONG'}: "
              f"{lim:,.0f} fit), {len(final_themes(L))} final themes, {len(sel)} acts")
    print(f"The filled prompts are in {d}.")


def test_models(agents=None):
    """A few small calls per model: a JSON reply and the folder tool. These are not counted in the run's budget."""
    if busy():
        print("Agents from the last run are still finishing their calls; wait a minute and run this again.")
        return
    STOP_WHY[0] = "stopped by hand"
    STOP.clear()
    print(VERSION)
    folder = ensure(rpath("test_folder"))
    with open(os.path.join(folder, "note.txt"), "w", encoding="utf-8") as f:
        f.write("The word to find is: lighthouse\n")
    for L in (agents or LETTERS):
        name = f"Agent {L} ({MODELS[L]['model']})"
        try:
            prov = provider(L)
            res = prov.json_call({"fixed": 'This is a test of the connection. Reply in this JSON format: {"memory": '
                                           '["connected"]}', "memory": [], "tail": ""}, SCHEMA_TEST, False)
            obj, _ = parse_json(res["text"])
            print(f"{name}: JSON reply {'OK' if obj and res['ok'] else 'PROBLEM: ' + short(res['text'], 80)}")
            r3 = prov.tool_call({"fixed": "This is a test of your folder tools. List the files in your folder, read "
                                          "note.txt, and reply with the one word it asks you to find, and nothing "
                                          "else.", "memory": [], "tail": ""}, folder_tools(folder))
            print(f"{name}: folder tool {'OK' if 'lighthouse' in r3['text'].lower() else 'PROBLEM'}: "
                  f"{r3.get('tool_steps', 0)} tool step(s); replied {short(r3['text'], 60)!r}")
        except Exception as e:
            print(f"{name}: PROBLEM: {short(e, 400)}")


def status():
    """Where each agent is, its cost and its flags."""
    for L in LETTERS:
        st = load_state(L)
        done = [r for r, p in (("round 1", own_path(L)), ("round 2", votes_path(L)), ("round 3", memo6_path(L)))
                if os.path.exists(p)]
        print(f"Agent {L}: {', '.join(done) or 'nothing yet'} done; ${st['cost']:.2f} in {st['calls']} calls; "
              f"{st['flags']} flags" + (f"; stopped: {st['stopped']}" if st.get("stopped") else ""))
    print(f"All agents: ${total_cost():.2f} of the budget of ${BUDGET_USD:,.0f}." + (f" Running: {', '.join(busy())}"
                                                                                   if busy() else ""))


def report():
    """What each round produced, the counts and the cost (reports/step6_report.md)."""
    sel = select_acts()
    lines = [f"# Step 6 report ({RUN}, on Step 5 run {STEP5_RUN}, {len(sel)} acts)", ""]
    for L in LETTERS:
        st = load_state(L)
        lines.append(f"## Agent {L} ({MODELS[L]['model']})\n")
        o = read_json(own_path(L))
        if o:
            n_th = len(final_themes(L))
            in_dim = len({x for d in o["dimensions"] for x in d["themes"]})
            lines.append(f"Round 1: {len(o['dimensions'])} dimensions holding {in_dim} of its {n_th} final themes "
                         f"({len(o['unplaced_themes'])} fit none); {len(o['types'])} types; "
                         f"{sum(1 for p in o['acts'] if p['types'])} of {len(sel)} acts in at least one type, "
                         f"{sum(1 for p in o['acts'] if not p['types'])} in none; {len(o['misfits'])} misfits "
                         f"({sum(1 for m in o['misfits'] if m['kind'] == 'core')} against a type's core); "
                         f"{len(o['empty_cells'])} empty cells.")
        v = read_json(votes_path(L))
        pool = read_json(rpath("merged", "pool.json"))
        if v and pool:
            c = collections.Counter(x["decision"] for x in v["items"])
            ob, on, xb, xn = backing(L, pool, v["items"])
            lines.append(f"Round 2: keep {c['keep']}, same {c['same']}, drop {c['drop']} of {len(v['items'])} pooled "
                         f"items; backs {ob} of its own {on} items and {xb} of the other agents' {xn} (with evidence).")
            if xn >= 5 and xb >= 0.9 * xn:
                lines.append(f"WARNING: Agent {L} backs nearly all the other agents' items, so its votes do little to "
                             f"test agreement: an item it backs stands as soon as its author backs it too.")
            if xn >= 5 and xb <= 0.1 * xn:
                lines.append(f"WARNING: Agent {L} backs almost none of the other agents' items, so the shared scheme "
                             f"rests on the other two agents alone.")
        if os.path.exists(memo6_path(L)):
            lines.append(f"Round 3: the Step 6 memo has {nwords(read_text(memo6_path(L))):,} words.")
        calls = logged_calls(L)
        by = collections.defaultdict(float)
        for x in calls:
            by[x.get("log") or x.get("kind")] += x.get("cost") or 0
        again = sum(1 for x in calls if int(str(x.get("attempt") or 1)) > 1)
        lines.append("Cost: " + ", ".join(f"{k} ${c:.2f}" for k, c in sorted(by.items())) + f" (all ${st['cost']:.2f}); "
                     f"asked again {again} time(s); {st['flags']} flags.\n")
    votes = {L: read_json(votes_path(L)) for L in LETTERS}
    if all(votes.values()):
        b = {L: backed_items({x["id"]: x for x in votes[L]["items"]}) for L in LETTERS}
        ids = [x["id"] for x in votes[LETTERS[0]]["items"]]
        lines.append("## Agreement in round 2\n")
        for L1, L2 in itertools.combinations(LETTERS, 2):
            same = sum((i in b[L1]) == (i in b[L2]) for i in ids)
            lines.append(f"- Agents {L1} and {L2} agree on {same} of the {len(ids)} pooled items (both back it, or "
                         f"neither does).")
        lines.append("")
    s = read_json(rpath("merged", "shared_scheme.json"))
    if s:
        lines.append("## Shared scheme\n")
        lines.append(f"{len(s['dimensions'])} dimensions and {len(s['types'])} types stand (backed by at least "
                     f"{s['agree']} of the three agents)" + (f", with {len(s['patterns'])} patterns that nearly every act "
                                                              f"shows kept apart" if s.get("patterns") else "") +
                     f"; {len(s['fell_away'])} items fell away; {len(s['merges'])} pairs of items counted as one.")
        if s.get("outcome_types"):
            lines.append(f"{len(s['outcome_types'])} agreed types are set aside because at least {s['agree']} agents "
                         f"judged that they rest on what round 3 measures.")
        for k in ("dimensions", "types", "patterns", "outcome_types"):
            for t in s.get(k, []):
                n_acts = (f"{len(t['acts'])} acts agreed{' (too few to compare)' if t.get('thin') else ''}, "
                          f"{len(t['one_agent'])} placed by one agent only; " if k != "dimensions" else "")
                ign = (f"; kept without evidence (not counted) by {', '.join(t['kept_without_evidence'])}"
                       if t.get("kept_without_evidence") else "")
                if t.get("outcome_flags"):
                    ign += f"; judged to rest on what round 3 measures by {', '.join(t['outcome_flags'])}"
                lines.append(f"- {t['id']} {t['name']}: {n_acts}backed by {', '.join(t['backers'])}; {_pooled(t)}{ign}")
        for x in s["fell_away"]:
            lines.append(f"- fell away: {x['kind']} {', '.join(x['items'])} {x['name']}" +
                         (f" ({x['why']})" if x.get("why") else ""))
        types_of = {x["act"] for t in s["types"] for x in t["acts"]}            # (patterns set no act apart)
        lines.append(f"Acts in no agreed type: {sum(1 for a in sel if a['Act'] not in types_of)} of {len(sel)}.")
    lines.append(f"\nAll agents: ${total_cost():.2f}.")
    write_text(rpath("reports", "step6_report.md"), "\n".join(lines) + "\n")
    print("\n".join(lines))


print(VERSION, "is loaded. Start with check_inputs(), then dry_run(); see the notes at the top of the script.")
