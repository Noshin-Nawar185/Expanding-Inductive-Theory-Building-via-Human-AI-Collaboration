"""transcribe_videos.py: transcribe the chosen videos with WhisperX.
Step 1 (videos). Version 1.3, 16 October 2026.

Transcribes the chosen YouTube videos (the To transcribe sheet of harrison_v7_VIDEO_SELECTION.xlsx) with WhisperX.
For each video it records:
  - the words (Whisper large-v3, with the act's and members' names given as a prompt, since names are where
    transcription goes wrong most);
  - when each word was said (wav2vec2 alignment);
  - who spoke when (pyannote speaker-diarization-community-1; its "exclusive" version, without overlaps, assigns
    each word to one speaker, and the full version marks where two people talk at once);
  - which parts sound sung, or are music with little speech (an AudioSet sound classifier, AST, on 10-second
    windows; music with little speech counts as sung, since 1980s TV vocals are seldom recognised as singing),
    so that lyrics can be left out of the turns.
Speakers are numbered (SPEAKER_00, SPEAKER_01 ...). Naming them is the next step.

Put this script, harrison_v7_VIDEO_SELECTION.xlsx and the audio files (named <YouTube ID>.m4a or .webm, loose or
in any folder inside) in one Google Drive folder, for example MyDrive/harrison_v7/Video for Harrison, where
the later video scripts look for them.

Output, in that folder's video_transcripts:
  V1234.json   everything above, word by word, with the settings and timings
  V1234.txt    the transcript to read or listen along to: [time] SPEAKER_01: words
  transcription_log.csv   one row per video; missing_audio.txt   chosen videos with no audio file yet

In Colab (Runtime > Change runtime type > A100 GPU, or L4 if A100 is not offered):
  Cell 1:  !pip install -q "whisperx==3.8.6" "transformers>=4.48,<5"
           then Runtime > Restart session (needed once per session, after the install)
  Cell 2:  from google.colab import drive; drive.mount('/content/drive')
           %run "/content/drive/MyDrive/harrison_v7/Video for Harrison/transcribe_videos.py" --act "Big Country"
           (the path of the script in your Drive: Files panel > drive > MyDrive > the script > Copy path)
Other runs:  --ids V1319,V971 (just these videos, from any act), --all (every act),
             --priority 1 (priority 1 only; with --all or --act), --max 5 (first five),
             --redo (transcribe again), --model large-v3-turbo (faster, slightly less accurate), --no-music.
It needs HF_TOKEN in Colab Secrets, from a Hugging Face account that has accepted the user agreement of
pyannote/speaker-diarization-community-1. Safe to run again: a video already done is skipped.
"""
import bisect, csv, dataclasses, datetime, gc, glob, importlib.metadata as im, json, math, os, re, sys, time, zlib

VERSION = "transcribe_videos 1.3 (16 Oct 2026)"

# ------------------------------------------------------------------ settings
def script_folder(path=None):
    """The Drive folder this script is in (when run from Drive with %run); otherwise MyDrive/harrison_v7."""
    if path is None:
        path = globals().get("__file__")
    if path and os.path.abspath(path).startswith("/content/drive/"):
        return os.path.dirname(os.path.abspath(path))
    return "/content/drive/MyDrive/harrison_v7"


BASE = script_folder()                                  # the folder with this script, the workbook and the audio
SELECTION = f"{BASE}/harrison_v7_VIDEO_SELECTION.xlsx"   # its To transcribe sheet is the list of videos
AUDIO_DIR = BASE                                        # audio named <YouTube ID>.<ext>, here or in any folder inside
OUT_DIR = f"{BASE}/video_transcripts"
ONLY_ACTS = ["Big Country"]    # the pilot; set to [] for every act
ONLY_IDS = []                  # for example ["V1319", "V971"]: just these videos, from any act; [] for all
ONLY_PRIORITIES = []           # for example [1] for priority 1 only; [] for all
MAX_VIDEOS = None              # stop after this many videos (None: no limit)
REDO = False                   # True: transcribe again the videos already done
MODEL = "large-v3"             # "large-v3-turbo" is about twice as fast and slightly less accurate
LANGUAGE = "en"
MUSIC_CHECK = True             # False skips the sung-parts check

DIAR_MODEL = "pyannote/speaker-diarization-community-1"
MUSIC_MODEL = "MIT/ast-finetuned-audioset-10-10-0.4593"
MUSIC_WINDOW = 10.0            # seconds per window of the sung-parts check
SUNG_MIN = 0.5                 # a window sounds sung when singing >= this and singing > speech,
MUSIC_ONLY = 0.5               # or when music >= this and speech < SPEECH_LOW (in the five-video test, sung
SPEECH_LOW = 0.3               # parts scored singing 0.12 to 0.24 at most, but music 0.56 to 0.95 and speech 0.01 to 0.24)
LOW_LOGPROB = -1.0             # Whisper's average log probability below this marks low confidence
REPETITIVE = 2.4               # text compression ratio above this marks a likely repetition loop
SR = 16000

# The act's own name and its nomination-era members, given to Whisper as a prompt (from the name lists).
ACT_NAMES = {'Dire Straits': ['Dire Straits', 'David Knopfler', 'John Illsley', 'Mark Knopfler', 'Pick Withers'],
 'Rickie Lee Jones': ['Rickie Lee Jones'],
 'The Knack': ['The Knack', 'Berton Averre', 'Bruce Gary', 'Doug Fieger', 'Prescott Niles'],
 'Amy Holland': ['Amy Holland'],
 'Christopher Cross': ['Christopher Cross'],
 'Irene Cara': ['Irene Cara'],
 'Robbie Dupree': ['Robbie Dupree'],
 'The Pretenders': ['The Pretenders', 'Chrissie Hynde', 'James Honeyman-Scott', 'Martin Chambers', 'Pete Farndon'],
 'Adam and the Ants': ['Adam and the Ants',
                       'Adam Ant',
                       'Chris Hughes',
                       'Gary Tibbs',
                       'Marco Pirroni',
                       'Terry Lee Miall'],
 'James Ingram': ['James Ingram'],
 'Luther Vandross': ['Luther Vandross'],
 'Sheena Easton': ['Sheena Easton'],
 "The Go-Go's": ["The Go-Go's",
                 'Belinda Carlisle',
                 'Charlotte Caffey',
                 'Gina Schock',
                 'Jane Wiedlin',
                 'Kathy Valentine'],
 'Asia': ['Asia', 'Carl Palmer', 'Geoff Downes', 'John Wetton', 'Steve Howe'],
 'Jennifer Holliday': ['Jennifer Holliday'],
 'Men at Work': ['Men at Work', 'Colin Hay', 'Greg Ham', 'Jerry Speiser', 'John Rees', 'Ron Strykert'],
 'Stray Cats': ['Stray Cats', 'Brian Setzer', 'Lee Rocker', 'Slim Jim Phantom'],
 'The Human League': ['The Human League',
                      'Ian Burden',
                      'Jo Callis',
                      'Joanne Catherall',
                      'Philip Adrian Wright',
                      'Philip Oakey',
                      'Susan Ann Sulley'],
 'Big Country': ['Big Country', 'Bruce Watson', 'Mark Brzezicki', 'Stuart Adamson', 'Tony Butler'],
 'Culture Club': ['Culture Club', 'Boy George', 'Jon Moss', 'Mikey Craig', 'Roy Hay'],
 'Eurythmics': ['Eurythmics', 'Annie Lennox', 'Dave Stewart'],
 'Men Without Hats': ['Men Without Hats', 'Allan McCarthy', 'Colin Doroschuk', 'Ivan Doroschuk', 'Stefan Doroschuk'],
 'Musical Youth': ['Musical Youth',
                   'Dennis Seaton',
                   'Frederick Waite',
                   'Kelvin Grant',
                   'Michael Grant',
                   'Patrick Waite'],
 'Corey Hart': ['Corey Hart'],
 'Cyndi Lauper': ['Cyndi Lauper'],
 'Frankie Goes to Hollywood': ['Frankie Goes to Hollywood',
                               'Brian Nash',
                               'Holly Johnson',
                               "Mark O'Toole",
                               'Paul Rutherford',
                               'Peter Gill'],
 'Sheila E.': ['Sheila E.'],
 'The Judds': ['The Judds', 'Naomi Judd', 'Wynonna Judd'],
 'a-ha': ['a-ha', 'Magne Furuholmen', 'Morten Harket', 'Paul Waaktaar-Savoy'],
 'Freddie Jackson': ['Freddie Jackson'],
 'Julian Lennon': ['Julian Lennon'],
 'Katrina and the Waves': ['Katrina and the Waves',
                           'Alex Cooper',
                           'Katrina Leskanich',
                           'Kimberley Rew',
                           'Vince de la Cruz'],
 'Sade': ['Sade', 'Andrew Hale', 'Paul S. Denman', 'Sade Adu', 'Stuart Matthewman'],
 'Bruce Hornsby and the Range': ['Bruce Hornsby and the Range',
                                 'Bruce Hornsby',
                                 'David Mansfield',
                                 'George Marinelli',
                                 'Joe Puerta',
                                 'John Molo',
                                 'Peter Harris'],
 'Glass Tiger': ['Glass Tiger', 'Al Connelly', 'Alan Frew', 'Michael Hanson', 'Sam Reid', 'Wayne Parker'],
 'Nu Shooz': ['Nu Shooz', 'John Smith', 'Valerie Day'],
 'Simply Red': ['Simply Red',
                'Chris Joyce',
                'Fritz McIntyre',
                'Ian Kirkham',
                'Mick Hucknall',
                'Sylvan Richardson',
                'Tim Kellett',
                'Tony Bowers'],
 'Timbuk3': ['Timbuk3', 'Barbara K. MacDonald', 'Pat MacDonald'],
 'Breakfast Club': ['Breakfast Club',
                    'Dan Gilroy',
                    'Ed Gilroy',
                    'Gary Burke',
                    'Paul Kauk',
                    'Randy Jackson',
                    'Stephen Bray'],
 'Cutting Crew': ['Cutting Crew', 'Colin Farley', 'Kevin MacMichael', 'Martin Beedle', 'Nick Van Eede', 'Tony Moore'],
 'Jody Watley': ['Jody Watley'],
 'Swing Out Sister': ['Swing Out Sister', 'Andy Connell', 'Corinne Drewery', 'Martin Jackson'],
 "Terence Trent D'Arby": ["Terence Trent D'Arby"],
 'Rick Astley': ['Rick Astley'],
 'Take 6': ['Take 6', 'Alvin Chea', 'Cedric Dent', 'Claude McKnight', 'David Thomas', 'Mark Kibble', 'Mervyn Warren'],
 'Toni Childs': ['Toni Childs'],
 'Tracy Chapman': ['Tracy Chapman'],
 'Vanessa Williams': ['Vanessa Williams'],
 'Indigo Girls': ['Indigo Girls', 'Amy Ray', 'Emily Saliers'],
 'Neneh Cherry': ['Neneh Cherry'],
 'Soul II Soul': ['Soul II Soul',
                  'Aitch B',
                  'Caron Wheeler',
                  'Daddae Harvey',
                  'Jazzie B',
                  'Nellee Hooper',
                  'Simon Law'],
 'Tone Loc': ['Tone Loc']}

SPEECH_LABELS = ["Speech", "Male speech, man speaking", "Female speech, woman speaking",
                 "Child speech, kid speaking", "Conversation", "Narration, monologue"]
SINGING_LABELS = ["Singing", "Male singing", "Female singing", "Child singing", "Synthetic singing", "Choir",
                  "Chant", "Yodeling", "Rapping"]


class Stop(Exception):
    pass


def stop(msg):
    print("\n" + "=" * 78 + "\nSTOPPED: " + msg + "\n" + "=" * 78)
    raise Stop(msg)


# ------------------------------------------------------------------ pure helpers (tested without a GPU)
def prompt_for(acts):
    """'Big Country. Bruce Watson, Mark Brzezicki, Stuart Adamson, Tony Butler.' for one or more acts."""
    parts = []
    for act in acts:
        names = [n for n in ACT_NAMES.get(act, []) if n != act]
        parts.append(f"{act}. " + (", ".join(names) + "." if names else ""))
    return " ".join(p.strip() for p in parts).strip() or None


def hms(t):
    t = max(0, int(round(t or 0)))
    return f"{t // 3600:02d}:{t % 3600 // 60:02d}:{t % 60:02d}"


def compression_ratio(text):
    b = text.encode("utf-8")
    return len(b) / max(1, len(zlib.compress(b))) if b else 0.0


AUDIO_EXTS = (".m4a", ".webm", ".mp3", ".mp4", ".opus", ".ogg", ".oga", ".wav", ".aac", ".flac", ".mka", ".mkv")


def audio_index(folder, skip=()):
    """YouTube ID -> audio file, from this folder and every folder inside it. Partial downloads, other files and
    the folders in skip (the transcripts) are ignored; if the same audio is in two places, the first is used."""
    out = {}
    if not os.path.isdir(folder):
        return out
    skip = {os.path.abspath(s) for s in skip}
    for root, dirs, names in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not d.startswith(".") and os.path.abspath(os.path.join(root, d)) not in skip)
        for name in sorted(names):
            stem, ext = os.path.splitext(name)
            if name.startswith(".") or ext.lower() not in AUDIO_EXTS:
                continue
            p = os.path.join(root, name)
            if os.path.getsize(p) == 0:
                continue
            out.setdefault(stem, p)
    return out


def overlap_len(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def speaker_at(w0, w1, turns, ends, near=1.0):
    """The speaker of the exclusive turn that overlaps [w0, w1] most; else the nearest turn within `near` s.
    `turns` are (start, end, speaker), sorted and never overlapping, so their ends are sorted too."""
    lo = bisect.bisect_right(ends, w0)          # first turn ending after the word starts
    best, best_ov = None, 0.0
    j = lo
    while j < len(turns) and turns[j][0] < w1:
        ov = overlap_len(w0, w1, turns[j][0], turns[j][1])
        if ov > best_ov:
            best, best_ov = turns[j][2], ov
        j += 1
    if best is not None:
        return best
    mid, best_d = (w0 + w1) / 2, near
    for t0, t1, spk in turns[max(0, lo - 1):lo + 1]:
        d = 0.0 if t0 <= mid <= t1 else min(abs(mid - t0), abs(mid - t1))
        if d < best_d:
            best, best_d = spk, d
    return best


def _num(x):
    """A float, or None for missing or NaN values."""
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(x) else x


def in_spans(t, spans):
    return any(s0 <= t <= s1 for s0, s1 in spans)


def sung_spans(windows):
    """10-second windows that sound sung, or that are music with little speech: words there are usually lyrics."""
    return [(w["start"], w["end"]) for w in windows
            if (w["singing"] >= SUNG_MIN and w["singing"] > w["speech"])
            or (w["music"] >= MUSIC_ONLY and w["speech"] < SPEECH_LOW)]


def label_words(segments, excl_turns, overlaps, sung, chunks):
    """Give every word a speaker, an overlap flag, a sung flag and its chunk's confidence. Words the aligner could
    not time (numbers, for instance) take the time and speaker of the word before them, or after them."""
    turns = sorted(excl_turns)
    ends = [t[1] for t in turns]
    for seg in segments:
        seg["start"], seg["end"] = _num(seg.get("start")), _num(seg.get("end"))
        words = seg.get("words") or []
        if not words:   # alignment failed for this segment: keep it as one timed piece
            words = [{"word": seg.get("text", "").strip(), "start": seg.get("start"), "end": seg.get("end"),
                      "unaligned": True}]
            seg["words"] = words
        for w in words:
            w["start"], w["end"] = _num(w.get("start")), _num(w.get("end"))
            if w["start"] is None or w["end"] is None:
                w["timed"] = False
                continue
            w["timed"] = True
            w["speaker"] = speaker_at(w["start"], w["end"], turns, ends)
            mid = (w["start"] + w["end"]) / 2
            w["overlap"] = in_spans(mid, overlaps)
            w["sung"] = in_spans(mid, sung)
        # untimed words: borrow from the neighbours
        for i, w in enumerate(words):
            if w["timed"]:
                continue
            prev = next((x for x in reversed(words[:i]) if x["timed"]), None)
            nxt = next((x for x in words[i + 1:] if x["timed"]), None)
            src = prev or nxt
            w["start"] = src["end"] if prev else (src["start"] if src else seg.get("start"))
            w["end"] = w["start"]
            w["speaker"] = src.get("speaker") if src else None
            w["overlap"] = src.get("overlap", False) if src else False
            w["sung"] = src.get("sung", False) if src else False
        # words with no speaker at all: the segment's main speaker, else UNKNOWN
        dur = {}
        for w in words:
            if w.get("speaker"):
                dur[w["speaker"]] = dur.get(w["speaker"], 0.0) + max(0.01, (w["end"] or 0) - (w["start"] or 0))
        main = max(dur.items(), key=lambda x: x[1])[0] if dur else "UNKNOWN"
        for w in words:
            if not w.get("speaker"):
                w["speaker"] = main
        seg["speaker"] = main
        lp = [c["avg_logprob"] for c in chunks
              if overlap_len(seg.get("start") or 0, seg.get("end") or 0, c["start"], c["end"]) > 0
              and c.get("avg_logprob") is not None]
        seg["avg_logprob"] = min(lp) if lp else None
        for w in words:
            w["avg_logprob"] = seg["avg_logprob"]
    return segments


def build_lines(segments, gap=10.0):
    """Speaker turns for reading: consecutive words by one speaker, split where the speaker changes or after a
    pause longer than `gap` seconds."""
    words = [w for seg in segments for w in seg["words"] if (w.get("word") or "").strip()]
    words.sort(key=lambda w: (w["start"] if w["start"] is not None else 0))
    lines, cur = [], None
    for w in words:
        if cur and w["speaker"] == cur["speaker"] and (w["start"] or 0) - (cur["end"] or 0) <= gap:
            cur["words"].append(w)
            cur["end"] = max(cur["end"], w["end"] or cur["end"])
        else:
            cur = {"speaker": w["speaker"], "start": w["start"], "end": w["end"], "words": [w]}
            lines.append(cur)
    out = []
    for ln in lines:
        ws = ln["words"]
        text = re.sub(r"\s+", " ", " ".join(w["word"].strip() for w in ws)).strip()
        n = len(ws)
        lp = [w["avg_logprob"] for w in ws if w.get("avg_logprob") is not None]
        flags = []
        sung = sum(1 for w in ws if w.get("sung")) / n
        over = sum(1 for w in ws if w.get("overlap")) / n
        if sung >= 0.5:
            flags.append("sung")
        if over >= 0.3:
            flags.append("overlap")
        if lp and min(lp) < LOW_LOGPROB:
            flags.append("low confidence")
        if len(text) >= 50 and compression_ratio(text) > REPETITIVE:
            flags.append("repetitive")
        out.append({"start": round(ln["start"] or 0, 2), "end": round(ln["end"] or 0, 2), "speaker": ln["speaker"],
                    "text": text, "words": n, "sung_share": round(sung, 2), "overlap_share": round(over, 2),
                    "min_avg_logprob": round(min(lp), 3) if lp else None, "flags": flags})
    return out


def prompt_echo(lines, prompt):
    """Mark lines that look like the prompt being repeated back (Whisper sometimes does this over music)."""
    if not prompt:
        return
    names = [n.strip().lower() for n in re.split(r"[.,]", prompt) if len(n.strip()) > 2]
    for ln in lines:
        low = ln["text"].lower()
        hits = sum(1 for n in names if n in low)
        if len(names) >= 2 and hits >= max(2, math.ceil(0.75 * len(names))) and len(low) < 1.5 * len(prompt):
            ln["flags"].append("prompt echo")


def speaker_summary(lines, embeddings):
    out = {}
    for ln in lines:
        s = out.setdefault(ln["speaker"], {"seconds": 0.0, "words": 0, "first_seen": ln["start"], "lines": 0})
        s["seconds"] += max(0.0, ln["end"] - ln["start"])
        s["words"] += ln["words"]
        s["lines"] += 1
    for spk, s in out.items():
        s["seconds"] = round(s["seconds"], 1)
        emb = (embeddings or {}).get(spk)
        s["embedding"] = emb
    return dict(sorted(out.items(), key=lambda kv: -kv[1]["seconds"]))


def write_txt(path, meta, lines, speakers):
    head = [f"{meta['id']} | {meta['act']} | priority {meta['priority']} | {meta['title']}",
            meta["url"],
            f"{meta['audio_seconds'] / 60:.1f} minutes; {len(speakers)} speakers: "
            + ", ".join(f"{k} ({v['seconds'] / 60:.1f} min)" for k, v in speakers.items()),
            "Marks: [sung] sung, or music with little speech; [overlap] two people at once; [low confidence]; [repetitive] probably a "
            "transcription loop; [prompt echo] the names prompt repeated back.", ""]
    body = []
    for ln in lines:
        mark = "".join(f" [{f}]" for f in ln["flags"])
        body.append(f"[{hms(ln['start'])}] {ln['speaker']}:{mark} {ln['text']}")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(head + body) + "\n")
    os.replace(tmp, path)


def write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)
    os.replace(tmp, path)


def already_done(path):
    """The JSON is written in one step, and only when a video is finished, so its presence means done."""
    return os.path.exists(path) and os.path.getsize(path) > 0


LOG_COLS = ["ID", "Act", "Priority", "Audio minutes", "Status", "Seconds", "Speakers", "Words", "Sung minutes",
            "Flagged lines", "Note", "Finished", "GPU"]


def log_row(row):
    path = os.path.join(OUT_DIR, "transcription_log.csv")
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_COLS)
        if new:
            w.writeheader()
        w.writerow({k: row.get(k, "") for k in LOG_COLS})


def read_selection():
    import pandas as pd
    if not os.path.exists(SELECTION):
        stop(f"The selection workbook is not at {SELECTION}. Upload harrison_v7_VIDEO_SELECTION.xlsx there.")
    T = pd.read_excel(SELECTION, "To transcribe")
    need = {"ID", "Act", "Priority", "Title", "URL", "YouTube ID", "Minutes"}
    if need - set(T.columns):
        stop(f"The To transcribe sheet lacks the columns {sorted(need - set(T.columns))}.")
    T["_order"] = range(len(T))
    T["_acts"] = T["Act"].astype(str).str.split(";").apply(lambda xs: [x.strip() for x in xs if x.strip()])
    if ONLY_IDS:
        unknown = [i for i in ONLY_IDS if i not in set(T["ID"].astype(str))]
        if unknown:
            stop(f"These IDs are not in the To transcribe sheet: {unknown}. Use the IDs as in the selection, for example V1319.")
        T = T[T["ID"].astype(str).isin(ONLY_IDS)]
    if ONLY_ACTS:
        T = T[T["_acts"].apply(lambda xs: any(a in ONLY_ACTS for a in xs))]
    if ONLY_PRIORITIES:
        T = T[T["Priority"].isin(ONLY_PRIORITIES)]
    T = T.sort_values(["Priority", "_order"])
    if MAX_VIDEOS:
        T = T.head(MAX_VIDEOS)
    return T


# ------------------------------------------------------------------ the models (need the GPU)
class Engines:
    def __init__(self):
        os.environ.setdefault("PYANNOTE_METRICS_ENABLED", "false")   # no usage reports from pyannote
        os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
        import torch
        self.torch = torch
        if not torch.cuda.is_available():
            stop("There is no GPU. Choose Runtime > Change runtime type > A100 (or L4) GPU, then run again.")
        self.gpu = torch.cuda.get_device_name(0)
        mem = torch.cuda.get_device_properties(0).total_memory / 1e9
        self.batch = 32 if mem >= 30 else (16 if mem >= 20 else 8)
        _preload_cuda_libraries()
        self.versions = {p: _version(p) for p in ("whisperx", "faster-whisper", "ctranslate2", "torch",
                                                     "pyannote.audio", "transformers")}
        if not str(self.versions["torch"]).startswith("2.8"):
            stop(f"torch is {self.versions['torch']}, not 2.8. Run the install cell, then Runtime > Restart "
                 f"session, then this cell again.")
        if self.versions["whisperx"] != "3.8.6":
            print(f"Note: whisperx is {self.versions['whisperx']}; this script was written for 3.8.6.")
        try:
            from google.colab import userdata
            token = userdata.get("HF_TOKEN")
        except Exception as e:
            stop(f"Could not read HF_TOKEN from Colab Secrets ({type(e).__name__}). Add it in the key icon on "
                 f"the left, with notebook access switched on.")
        import whisperx
        self.whisperx = whisperx
        print(f"GPU: {self.gpu} ({mem:.0f} GB), batch size {self.batch}. Loading the models (a few minutes the "
              f"first time) ...")
        t = time.time()
        self.asr = whisperx.load_model(MODEL, "cuda", compute_type="float16", language=LANGUAGE)
        self.align_model, self.align_meta = whisperx.load_align_model(language_code=LANGUAGE, device="cuda")
        from pyannote.audio import Pipeline
        try:
            self.diar = Pipeline.from_pretrained(DIAR_MODEL, token=token)
        except Exception as e:
            stop(f"Could not load {DIAR_MODEL} ({type(e).__name__}: {str(e)[:200]}). Check that HF_TOKEN is a "
                 f"Read token and that the same Hugging Face account accepted the user agreement at "
                 f"https://huggingface.co/{DIAR_MODEL}")
        if self.diar is None:
            stop(f"Hugging Face refused {DIAR_MODEL}. Accept its user agreement with the account that made HF_TOKEN.")
        self.diar.to(torch.device("cuda"))
        self.music = None
        if MUSIC_CHECK:
            try:
                from transformers import ASTFeatureExtractor, ASTForAudioClassification
                self.fe = ASTFeatureExtractor.from_pretrained(MUSIC_MODEL)
                self.music = ASTForAudioClassification.from_pretrained(MUSIC_MODEL).to("cuda").eval()
                lab = {v: int(k) for k, v in self.music.config.id2label.items()}
                self.i_speech = [lab[x] for x in SPEECH_LABELS if x in lab]
                self.i_sing = [lab[x] for x in SINGING_LABELS if x in lab]
                self.i_music = [lab[x] for x in ["Music"] if x in lab]
                if not (self.i_speech and self.i_sing and self.i_music):
                    print("Note: the music model's labels were not as expected; the sung-parts check is off.")
                    self.music = None
            except Exception as e:
                print(f"Note: the sung-parts check could not start ({type(e).__name__}: {str(e)[:150]}); "
                      f"transcribing without it.")
                self.music = None
        print(f"Models ready in {time.time() - t:.0f} s.")

    def transcribe(self, audio, prompt):
        self.asr.options = dataclasses.replace(self.asr.options, initial_prompt=prompt)
        bs = self.batch
        while True:
            try:
                return self.asr.transcribe(audio, batch_size=bs, language=LANGUAGE)
            except Exception as e:
                if "out of memory" in str(e).lower() and bs > 1:
                    bs //= 2
                    self.torch.cuda.empty_cache()
                    print(f"   (GPU memory full; retrying with batch size {bs})")
                    self.batch = bs
                    continue
                raise

    def align(self, segments, audio):
        if not segments:
            return []
        out = self.whisperx.align(segments, self.align_model, self.align_meta, audio, "cuda",
                                  return_char_alignments=False)
        return out["segments"]

    def diarize(self, audio):
        torch = self.torch
        output = self.diar({"waveform": torch.from_numpy(audio[None, :]), "sample_rate": SR})
        reg, exc = output.speaker_diarization, output.exclusive_speaker_diarization
        regular = [(round(s.start, 3), round(s.end, 3), spk) for s, _, spk in reg.itertracks(yield_label=True)]
        exclusive = [(round(s.start, 3), round(s.end, 3), spk) for s, _, spk in exc.itertracks(yield_label=True)]
        overlaps = [(round(s.start, 3), round(s.end, 3)) for s in reg.get_overlap()]
        emb = {}
        if output.speaker_embeddings is not None:
            for i, spk in enumerate(reg.labels()):
                if i < len(output.speaker_embeddings):
                    v = output.speaker_embeddings[i]
                    emb[spk] = None if any(math.isnan(float(x)) for x in v) else [round(float(x), 5) for x in v]
        return regular, exclusive, overlaps, emb

    def music_windows(self, audio):
        if self.music is None:
            return None
        torch = self.torch
        step = int(MUSIC_WINDOW * SR)
        starts = list(range(0, len(audio), step))
        out = []
        for b in range(0, len(starts), 16):
            batch = [audio[s:s + step] for s in starts[b:b + 16]]
            keep = [i for i, x in enumerate(batch) if len(x) >= SR]
            if not keep:
                continue
            feats = self.fe([batch[i] for i in keep], sampling_rate=SR, return_tensors="pt")
            with torch.no_grad(), torch.autocast("cuda", dtype=torch.float16):
                p = torch.sigmoid(self.music(feats["input_values"].to("cuda")).logits.float()).cpu().numpy()
            for j, i in enumerate(keep):
                s0 = starts[b + i] / SR
                out.append({"start": round(s0, 2), "end": round(s0 + len(batch[i]) / SR, 2),
                            "speech": round(float(p[j, self.i_speech].max()), 3),
                            "singing": round(float(p[j, self.i_sing].max()), 3),
                            "music": round(float(p[j, self.i_music].max()), 3)})
        return out

    def free(self):
        gc.collect()
        self.torch.cuda.empty_cache()


def _version(pkg):
    try:
        return im.version(pkg)
    except Exception:
        return None


def _preload_cuda_libraries():
    """CTranslate2 (Whisper) needs cuDNN 9 and cuBLAS; load the copies that came with torch, so it finds them."""
    import ctypes, site
    roots = site.getsitepackages() + [site.getusersitepackages()]
    for root in roots:
        for pat in ("nvidia/cublas/lib/libcublas*.so*", "nvidia/cudnn/lib/libcudnn*.so*"):
            for lib in sorted(glob.glob(os.path.join(root, pat))):
                try:
                    ctypes.CDLL(lib, mode=ctypes.RTLD_GLOBAL)
                except OSError:
                    pass


# ------------------------------------------------------------------ one video
def process(E, row, path):
    times = {}
    t0 = time.time()
    audio = E.whisperx.load_audio(path)
    secs = len(audio) / SR
    times["load"] = time.time() - t0
    if secs < 1:
        raise ValueError("the audio file is shorter than a second")
    acts = row["_acts"]
    prompt = prompt_for(acts)

    t = time.time()
    asr = E.transcribe(audio, prompt)
    chunks = [{"start": round(s["start"], 3), "end": round(s["end"], 3), "text": s["text"].strip(),
               "avg_logprob": None if s.get("avg_logprob") is None else round(float(s["avg_logprob"]), 4)}
              for s in asr["segments"]]
    times["transcribe"] = time.time() - t

    t = time.time()
    segments = E.align([{"start": c["start"], "end": c["end"], "text": c["text"]} for c in chunks if c["text"]], audio)
    times["align"] = time.time() - t

    t = time.time()
    regular, exclusive, overlaps, emb = E.diarize(audio)
    times["speakers"] = time.time() - t

    t = time.time()
    windows = E.music_windows(audio)
    times["music"] = time.time() - t

    sung = sung_spans(windows) if windows else []
    label_words(segments, exclusive, overlaps, sung, chunks)
    lines = build_lines(segments)
    prompt_echo(lines, prompt)
    speakers = speaker_summary(lines, emb)
    times["total"] = time.time() - t0

    meta = {"id": row["ID"], "act": "; ".join(acts), "priority": int(row["Priority"]), "title": str(row["Title"]),
            "url": str(row["URL"]), "youtube_id": str(row["YouTube ID"]), "minutes_listed": float(row["Minutes"]),
            "audio_file": os.path.basename(path), "audio_seconds": round(secs, 2)}
    doc = {"status": "done", "version": VERSION, **meta,
           "settings": {"model": MODEL, "language": LANGUAGE, "batch_size": E.batch, "prompt": prompt,
                        "speaker_model": DIAR_MODEL, "music_model": MUSIC_MODEL if windows is not None else None,
                        "music_window_seconds": MUSIC_WINDOW, "sung_min": SUNG_MIN, "music_only": MUSIC_ONLY, "speech_low": SPEECH_LOW,
                        "low_logprob": LOW_LOGPROB, "repetitive": REPETITIVE},
           "run": {"gpu": E.gpu, "finished": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
                   "seconds": {k: round(v, 1) for k, v in times.items()}, "versions": E.versions},
           "speakers": speakers, "lines": lines,
           "segments": [{"start": _r(s.get("start")), "end": _r(s.get("end")), "text": s.get("text", "").strip(),
                         "speaker": s.get("speaker"), "avg_logprob": s.get("avg_logprob"),
                         "words": [{"word": w.get("word"), "start": _r(w.get("start")), "end": _r(w.get("end")),
                                    "score": _r(w.get("score")), "speaker": w.get("speaker"),
                                    "timed": w.get("timed"), "overlap": w.get("overlap"), "sung": w.get("sung")}
                                   for w in s["words"]]} for s in segments],
           "asr_chunks": chunks,
           "diarization": [{"start": a, "end": b, "speaker": c} for a, b, c in regular],
           "exclusive_diarization": [{"start": a, "end": b, "speaker": c} for a, b, c in exclusive],
           "overlaps": [list(x) for x in overlaps],
           "music_windows": windows}
    return doc, meta, lines, speakers, secs, times


def _r(x, n=3):
    x = _num(x)
    return None if x is None else round(x, n)


# ------------------------------------------------------------------ main
def parse_args(argv):
    """Command-line choices (after %run ... in Colab). Without --all or --act, the pilot act is used."""
    import argparse
    global ONLY_ACTS, ONLY_IDS, ONLY_PRIORITIES, MAX_VIDEOS, REDO, MODEL, MUSIC_CHECK
    ap = argparse.ArgumentParser(prog="transcribe_videos.py")
    ap.add_argument("--ids", default="", help="just these videos, by ID, separated by commas (for example V1319,V971)")
    ap.add_argument("--act", action="append", default=[], help="an act to transcribe (can be repeated)")
    ap.add_argument("--all", action="store_true", help="every act")
    ap.add_argument("--priority", action="append", type=int, default=[], help="only this priority (can be repeated)")
    ap.add_argument("--max", type=int, default=None, help="stop after this many videos")
    ap.add_argument("--redo", action="store_true", help="transcribe again videos already done")
    ap.add_argument("--model", default=MODEL, help="Whisper model (default large-v3)")
    ap.add_argument("--no-music", action="store_true", help="skip the sung-parts check")
    a = ap.parse_args(argv)
    unknown = [x for x in a.act if x not in ACT_NAMES]
    if unknown:
        stop(f"Unknown act name(s): {unknown}. Use the names as in the selection, for example \"Big Country\".")
    ONLY_IDS = [x.strip().upper() for x in re.split(r"[,\s]+", a.ids) if x.strip()]
    ONLY_ACTS = [] if (a.all or (ONLY_IDS and not a.act)) else (a.act or ONLY_ACTS)
    ONLY_PRIORITIES, MAX_VIDEOS, REDO, MODEL = a.priority, a.max, a.redo, a.model
    MUSIC_CHECK = MUSIC_CHECK and not a.no_music


def main():
    print(VERSION)
    print(("Videos: " + ", ".join(ONLY_IDS) + ("; acts: " + ", ".join(ONLY_ACTS) if ONLY_ACTS else "") if ONLY_IDS
           else "Acts: " + (", ".join(ONLY_ACTS) if ONLY_ACTS else "all")) +
          (f"; priorities {ONLY_PRIORITIES}" if ONLY_PRIORITIES else "") + f"; model {MODEL}")
    os.makedirs(OUT_DIR, exist_ok=True)
    T = read_selection()
    files = audio_index(AUDIO_DIR, skip=[OUT_DIR])
    if not files:
        stop(f"No audio files found in {AUDIO_DIR} or the folders inside it. Upload the audio there "
             "(files named <YouTube ID>.m4a or .webm, as get_audio.sh saved them).")
    todo, missing, done = [], [], 0
    for _, r in T.iterrows():
        out = os.path.join(OUT_DIR, f"{r['ID']}.json")
        if not REDO and already_done(out):
            done += 1
            continue
        p = files.get(str(r["YouTube ID"]))
        (todo if p else missing).append((r, p))
    with open(os.path.join(OUT_DIR, "missing_audio.txt"), "w", encoding="utf-8") as f:
        f.write("Chosen videos with no audio file in video_audio yet (ID, YouTube ID, act, title):\n")
        for r, _ in missing:
            f.write(f"{r['ID']}\t{r['YouTube ID']}\t{'; '.join(r['_acts'])}\t{r['Title']}\n")
    mins = sum(float(r["Minutes"]) for r, _ in todo)
    print(f"Videos in this run: {len(T)}. Already done: {done}. To do now: {len(todo)} ({mins / 60:.1f} hours). "
          f"No audio yet: {len(missing)} (listed in missing_audio.txt).")
    if not todo:
        print("Nothing to transcribe.")
        return
    E = Engines()
    t_start, audio_done, n_ok, n_bad = time.time(), 0.0, 0, 0
    for k, (r, p) in enumerate(todo, 1):
        label = f"[{k}/{len(todo)}] {r['ID']} {'; '.join(r['_acts'])} ({float(r['Minutes']):.1f} min)"
        try:
            doc, meta, lines, speakers, secs, times = process(E, r, p)
            write_json(os.path.join(OUT_DIR, f"{r['ID']}.json"), doc)
            write_txt(os.path.join(OUT_DIR, f"{r['ID']}.txt"), meta, lines, speakers)
            sung_min = sum(ln["end"] - ln["start"] for ln in lines if "sung" in ln["flags"]) / 60
            flagged = sum(1 for ln in lines if ln["flags"])
            note = "" if abs(secs / 60 - float(r["Minutes"])) <= max(0.5, 0.1 * float(r["Minutes"])) else \
                f"audio is {secs / 60:.1f} min, the list says {float(r['Minutes']):.1f}"
            log_row({"ID": r["ID"], "Act": "; ".join(r["_acts"]), "Priority": int(r["Priority"]),
                     "Audio minutes": round(secs / 60, 2), "Status": "done", "Seconds": round(times["total"], 1),
                     "Speakers": len(speakers), "Words": sum(s["words"] for s in speakers.values()),
                     "Sung minutes": round(sung_min, 2), "Flagged lines": flagged, "Note": note,
                     "Finished": datetime.datetime.now().isoformat(timespec="seconds"), "GPU": E.gpu})
            audio_done += secs
            n_ok += 1
            rate = (time.time() - t_start) / max(1.0, audio_done)
            left = sum(float(x["Minutes"]) for x, _ in todo[k:]) * 60 * rate
            print(f"{label}: {times['total']:.0f} s, {len(speakers)} speakers"
                  + (f", {sung_min:.1f} min sung" if sung_min >= 0.1 else "")
                  + (f"  NOTE: {note}" if note else "") + f"  | about {left / 3600:.1f} h left")
        except Stop:
            raise
        except Exception as e:
            n_bad += 1
            msg = f"{type(e).__name__}: {str(e)[:300]}"
            print(f"{label}: FAILED ({msg})")
            log_row({"ID": r["ID"], "Act": "; ".join(r["_acts"]), "Priority": int(r["Priority"]), "Status": "failed",
                     "Note": msg, "Finished": datetime.datetime.now().isoformat(timespec="seconds"), "GPU": E.gpu})
        finally:
            E.free()
    took = time.time() - t_start
    print(f"\nFinished: {n_ok} done, {n_bad} failed, in {took / 60:.0f} min for {audio_done / 3600:.1f} h of audio "
          f"({audio_done / max(1.0, took):.0f} times real time). Files are in {OUT_DIR}.")


if __name__ == "__main__":
    try:
        parse_args(sys.argv[1:])
        main()
    except Stop:
        pass
