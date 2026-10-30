"""
===============================================================================
turns_v7_stage4to9.py: STAGES 4 TO 9 OF THE TURN BUILD
QUOTATIONS, SPEAKERS, TURNS, INTERVIEW DATES, CONFIDENCE
Step 2 (articles), stages 4 to 9. Version 1.3, 12 Oct 2026.
Reads version 3 of the stage 3 file (turns_v7_stage3.py, 12 Oct).
===============================================================================
Runs after turns_v7_stage3.py. The stages are those of Step 2 (articles) in
the technical appendix.
Reads   harrison_v7_CLEAN7.json            (never changed)
        harrison_v7_TURNS_stage1.json      the cleaned texts (turns_v7.py)
        harrison_v7_TURNS_stage2.json      the routes and segments (turns_v7.py)
        harrison_v7_TURNS_stage3.json      the Q&A and first-person speakers
                                           (turns_v7_stage3.py)
Writes  harrison_v7_turns_final.json       THE TURNS TABLE: one row per turn
        harrison_v7_spans_final.json       the spans table: every character of
                                           every text in exactly one span
        harrison_v7_records_final.json     the records table: route, quotation
                                           convention and quality checks
        harrison_v7_turns_final.xlsx       the turns table, for browsing, with a
                                           Read me that defines every column
        harrison_v7_dashboard_final.xlsx   the dashboard (by act, publication,
                                           route and year), the records, the
                                           second-hand turns, the stated
                                           interview dates, and an optional
                                           check of 66 turns
The two large JSON tables are in pandas' "split" layout:
    pd.read_json("harrison_v7_turns_final.json", orient="split")

The final name lists are built into this script. Forum threads (49) are not
processed, and record 11698 (machine-generated text) is skipped; both stay in
the records table.

WHAT THE STAGES DO
 4. Quotations in the prose (prose records, and the prose parts of Q&A and
    first-person records): the quotation marks of the record's convention
    (double, or single in British style) are paired paragraph by paragraph;
    paragraphs whose marks cannot be paired are "unresolved" and kept out of
    the turns. Each quotation is speech (a speech verb or a name label beside
    it, it continues the speech before it, or it is a sentence in the first or
    second person), non-speech (a title, lyrics, a scare quote) or undecided.
    Passages inside a Q&A that stage 2 left as prose (mostly answer
    paragraphs that report a conversation) are read as part of the
    interview: question lines go to the interviewer, and the rest continues
    the answer before it (rule names "Q&A passage: ...").
 5. Who says each speech quotation, in this order:
      1  a member's name beside a speech verb                        high
      2  someone else named beside a speech verb                     high
         (one name, or a first name not on the list: low; an unnamed
         speaker such as "a spokesperson": low)
      3  a name several people share ("Knopfler", "Lennon"): the person
         the title or introduction is about, else the nearest earlier
         full name                                                   medium
         (one of several members, none named: "band member,
         unspecified", medium; otherwise ambiguous, low)
      4  the band named as speaker ("the band said")                 medium
      5  a pronoun or "the singer" beside a speech verb, or a quotation
         that continues the current speaker                          medium
         (low when the gender could not be checked)
      5b a pronoun with no current speaker: the person the narration
         is about (members medium, anyone else low)
      5c no current speaker in a piece about one member: that member  low
      6  nothing applied                                             unknown
    Confidence drops one level when the member was named by a surname or
    first name that a relative or namesake named in the record shares, and
    a turn is never surer than the speaker it continues.
    "The band said" (rule 4) is the act only when it can be: never in a solo
    act's record, only with "the", a size word that fits the line-up ("the
    duo"), and the act or a member named just before, with no other group or
    person the subject since; otherwise it is an unnamed speaker (low). A
    member named beside the band ("the duo's Dave Stewart") is the speaker,
    and what the band is given must be something said, not a title, lyric
    or fragment. The same holds for second-hand turns. "Sade says" is Sade
    Adu, as in stage 3. An alias whose first and last names make the name of
    someone on the act's non-member list is not used ("John Charles Julian
    Lennon" would make "John Lennon" read as Julian Lennon).
 6. The current speaker is replaced by rules 1 to 4, and cleared after more
    than two paragraphs without a quotation, or when the narration names
    someone else (as the subject of a sentence; or, for "he"/"she", anyone
    the pronoun can stand for). A second-hand sentence ("Setzer says he
    isn't bothered.") makes its member the current speaker.
 7. Turns: speech spans by the same speaker join across attribution clauses
    ("...," she says. "...") and line breaks; narration (with non-speech and
    undecided quotations) is the journalist's (the Author, else the outlet).
    A narration sentence in which a member is the subject of a speech verb,
    with no quotation marks ("Adam says they have been practising"), is a
    second-hand turn: the speaker is the journalist, the member is the
    reported member, and it is never band voice. Unresolved spans, sung lines
    and stage directions belong to no turn.
 8. Interview date: a year bound to a speaking verb in the attribution's
    sentence ("told NME in 1983", "in a 1985 interview"), or, for a Q&A, in
    the title or introduction ("This interview took place on January 29,
    1981", "The 1982 Stray Cats Interview"); otherwise the publication date.
 9. Confidence: as above. The main set is high and medium; Main set is
    filled for band-voice turns only (blank for the rest).
 The Q&A and first-person blocks keep the speakers stage 3 gave them.

THE OUTPUTS
 Repeat of points to the earliest turn with the same text (an article
 published more than once). Quotation convention is blank on the turns of a
 record with no prose to read (mostly first-person accounts), as in the
 records table. The dashboard groups outlets without "www." and trailing
 full stops. The Read me sheets define every column of the turns table and
 say what each dashboard sheet is for; the dashboard's Read me also gives the
 blind checks by AI agents and the known limits.

THE CHECKS THE SCRIPT MAKES (it stops without writing if one fails)
   - CLEAN7 is the expected file (record count, turn set, fingerprint);
   - the stage 1, 2 and 3 files belong to it and are the versions this script
     was checked with (fingerprints of the routes, segments and speakers);
   - every record's spans cover its text exactly once (the round trip); a
     record that fails is flagged, not dropped.

COLAB (one cell)
  from google.colab import drive
  drive.mount('/content/drive')
  %run /content/drive/MyDrive/harrison_v7/turns_v7_stage4to9.py

It takes about 6 to 10 minutes. Set DRY_RUN = True to run every check and print
the summary without writing.
===============================================================================
"""

import base64
import hashlib
import json
import math
import os
import random
import re
import sys
import time
import types
import zlib
from collections import Counter, defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "turns_v7 stages 4 to 9, version 1.3 (12 Oct 2026; reads stage 3 version 3)"
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN7.json")
IN_STAGE1 = os.path.join(DIR, "harrison_v7_TURNS_stage1.json")
IN_STAGE2 = os.path.join(DIR, "harrison_v7_TURNS_stage2.json")
IN_STAGE3 = os.path.join(DIR, "harrison_v7_TURNS_stage3.json")
OUT_TURNS = os.path.join(DIR, "harrison_v7_turns_final.json")
OUT_SPANS = os.path.join(DIR, "harrison_v7_spans_final.json")
OUT_RECORDS = os.path.join(DIR, "harrison_v7_records_final.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_turns_final.xlsx")
OUT_DASH = os.path.join(DIR, "harrison_v7_dashboard_final.xlsx")
DRY_RUN = False
REPEAT_MIN_WORDS = 8

EXPECTED_RECORDS = 13993
EXPECTED_TURN_SET = 9316
# MD5 of "ID|URL" for every turn-set record in ID order, from CLEAN7 as written on 4 Oct 2026
TURN_SET_FINGERPRINT = "e07271c9aacf5776a0d593bbd64cf7c0"
# the routes version 2 of turns_v7.py gave (5 Oct 2026), and fingerprints of the stage 2 and stage 3 files
STAGE2_COUNTS = {"labelled Q&A": 631, "unlabelled Q&A": 652, "first-person": 170, "prose": 7814, "forum thread": 49}
STAGE2_FINGERPRINT = "3b320e4f99f02b9b51aae0d1d017cf09"
STAGE3_FINGERPRINT = "7ca3d5913b22fdae5f7d13bbd967b30a"
SKIP = {11698: "machine-generated text"}
FORUM = "forum thread"
ROUTES = ["prose", "labelled Q&A", "unlabelled Q&A", "first-person"]
SAMPLE_SEED = 20261006
EXCEL_CELL_CAP = 32000
# counts this script produced on the same files when it was checked on 12 Oct 2026, and a fingerprint of every turn's
# speaker, confidence, position and interview date; a run on those files should give exactly these
CHECK_COUNTS = {"records processed": 9266, "turns": {"narration": 44065, "second-hand": 1900, "speech": 86870}, "spans": 337356, "band-voice turns by route": {"prose": 8774, "labelled Q&A": 6229, "unlabelled Q&A": 4534, "first-person": 176}, "band-voice turns by confidence": {"high": 8060, "medium": 8810, "low": 2843}, "records with band-voice turns": 2447, "unknown speech turns": 8716, "stated interview dates": 1789, "round-trip failures": 0, "unresolved spans": 2682, "Q&A passages read as interview": 582, "band-voice turns from Q&A passages": 307, "turns that repeat an earlier turn": 9662}
CHECK_FINGERPRINT = "07bfc7cdfc5a06b48cf409a069ce3ccd"
# the sources and relative acts watched closely on the dashboard
WATCH_SOURCES = ["npr.org", "theguardian.com", "recordcollectormag.com", "musicmusingsandsuch.com",
                 "electricityclub.co.uk"]


# ═════════════════════════════════════════════════════════════════════════════
# THE RULES: stage 2 (version 2, for its name and label tests), stage 3, stage 4, stages 5 to 9, and the first-name
# genders used for the pronoun check. Each is loaded as its own module, exactly as written.
# ═════════════════════════════════════════════════════════════════════════════
SOURCES = {}
SOURCES["stage2"] = r'''"""Stage 2: route each record (and each part of it) by how speakers are marked.

Routes (four, plus forum threads, which are not processed further):
  labelled Q&A    speaker labels mark each turn: "LAUPER:", "Q:", "NVE:", "[Host]", "PHIL OAKEY –",
                  a label on its own line, labels inside paragraphs, lecture transcripts
  unlabelled Q&A  questions followed by first-person answers, with no labels; the question may
                  stand alone or run straight into the answer in the same paragraph
  first-person    one person speaking throughout, or one person per section under a heading
  prose           everything else: the journalist's narration with quotations
  forum thread    posts by several users (from Page Type)
Every record gets segments [(start, end, kind)] that cover the text exactly once, in order:
a prose introduction, the block, and a prose tail. Quotations in the introduction and tail
go through the prose route like any other prose."""
import base64
import re
import zlib
from collections import Counter, defaultdict

# first names: the 1990 US Census first-name lists (public domain) plus names from the corpus's member tables
# and British, Irish and European names common in music writing; used to tell a person's name from a heading
_FIRST_NAMES_Z = (
    "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")
FIRST_NAMES = set(zlib.decompress(base64.b64decode(_FIRST_NAMES_Z)).decode().split("|"))
# first names that are also everyday words: not enough on their own to make a heading a name
WORDLIKE_NAMES = set("Will May June April August Major King Royal Golden Young Long Love Sunny Rich Star Chance Summer Autumn Winter Spring Hope Joy Dawn Crystal Iris Violet Daisy Olive Ginger Ivy Eve Gay Art Sandy Misty Salt Boy Chase Page Lane Clay Hunter Reed Bell Ward Dale Mason Carter Taylor Walker Parker Porter Cole Wade Pearl Heather Hazel".split())
FIRST_NAMES_STRICT = FIRST_NAMES - WORDLIKE_NAMES



SPEECH_VERB = (r"(?:says?|said|asks?|asked|adds?|added|recalls?|recalled|explains?|explained|told|tells|laughs?|"
               r"laughed|notes?|noted|admits?|admitted|continues?|continued|replies|replied|insists?|insisted|"
               r"remembers?|remembered|wrote|writes|joked|jokes|confesses|confessed|reflects?|reflected|muses?|"
               r"mused|observes?|observed|concedes?|conceded|agrees?|agreed|declares?|declared|states?|stated|"
               r"argues?|argued)")

# ── speaker labels ───────────────────────────────────────────────────────────
# "Name:", "Name Surname:", "Ms. FIONA:", "SADE (singer-songwriter):", "Q:", "A.", "Q –"
LABEL_LINE = re.compile(
    r"^[ \t]*(?:[-–•*>][ \t]*)?(?:(?:Mr|Ms|Mrs|Dr|Miss)\.?[ \t]+)?"
    r"(?:(?P<lab>[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ.'’&-]*(?:[ \t]+(?:[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ.'’&-]*|de|van|von|la|le|of))?"
    r"(?:[ \t]+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ.'’&-]*){0,2})(?:[ \t]*(?:\([^()\n]{1,40}\)|\[[^\[\]\n]{1,40}\]))?[ \t]*:"
    r"|(?P<qa>Q|A)[ \t]*(?::|[.)](?=[ \t])|[ \t]+[-–—](?=[ \t])))"
    r"[ \t]*(?P<val>.*)$")
# radio transcripts introduce each speaker with name and role: "TONYA MOSLEY, HOST:", "KEN TUCKER, BYLINE:"
NPR_LABEL = re.compile(r"^[ \t]*(?:(?:Mr|Ms|Mrs|Dr|Miss)\.?[ \t]+)?(?P<lab>[A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ.'’-]*(?:[ \t]+[A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ.'’-]+){0,3}),"
                       r"[ \t]*[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ .'’&/()-]{1,50}:[ \t]*(?P<val>.*)$")
# the same in ordinary case: "Joe Elliott, lead singer/co-founder:", "Cliff Burnstein, co-manager:" (the name must start
# with a known first name; checked in label_lines)
ROLE_LABEL = re.compile(r"^[ \t]*(?:(?:Mr|Ms|Mrs|Dr|Miss)\.?[ \t]+)?(?P<lab>[A-ZÀ-ÖØ-Þ][a-zß-öø-ÿ'’.-]+(?:[ \t]+(?:[A-ZÀ-ÖØ-Þ][\w'’.-]+|de|van|von|la|le|da|di|del)){1,3}),"
                        r"[ \t]*(?P<role>[A-Za-zÀ-ÖØ-öø-ÿ][^:\n.!?\"]{1,60}):[ \t]+(?P<val>\S.*)$")
# capital-letter names before a dash: "PHIL OAKEY – We didn't...", "MF – ...", "Jody Watley — ..."
DASH_LABEL = re.compile(r"^[ \t]*(?:(?P<lab>[A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ.'’-]*(?:[ \t]+[A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ.'’-]*){0,3})[ \t]+[–—-][ \t]+(?P<val>\S.*)"
                        r"|(?P<lab2>[A-Z]{2,4})[–—-](?P<val2>[\"'A-Z].*)"
                        r"|(?P<lab3>[A-ZÀ-ÖØ-Þ][a-zß-öø-ÿ'’.-]+(?:[ \t]+[A-ZÀ-ÖØ-Þ][\w'’.-]+){0,2})[ \t]+[–—][ \t]+(?P<val3>\S.*))$")
# initials and a space, no colon: "CL There are drag queens...", "MC Well, a drag icon." (kept only when a name in the
# record has those initials)
SPACE_INITIALS = re.compile(r"^[ \t]*(?P<lab>[A-Z]{2,4})[ \t]+(?P<val>[A-Z\"'(\[][^\n]{10,})$")
# "[Host] ...", "- (00:12) [Christopher] ..."
BRACKET_LABEL = re.compile(r"^[ \t]*(?:-[ \t]*)?(?:\([^)\n]*\)[ \t]*-?[ \t]*)?\[(?P<lab>[A-Z][\w .'’-]{0,30})\][ \t]*(?P<val>.*)$")
# a label inside a paragraph, after the end of a sentence or a closing bracket: "...questions. D.T.: Rehearsals..."
INLINE_LABEL = re.compile(
    r"(?:(?<=[.!?\"'…)\]])|(?<=[.!?\"'…)\]][ \t])|(?<=[.!?\"'…)\]][ \t]{2}))"
    r"(?P<lab>(?:[A-Z]\.){1,3}[A-Z]?\.?|[A-Z][a-z]+(?:[ \t][A-Z][a-z'’-]+){0,2}|[A-Z]{2,}(?:[ \t][A-Z]{2,}){0,2})"
    r":[ \t]+(?=[A-Z\"'(\[])")
DASH_TURN = re.compile(r"^[ \t]*-[ \t]+\S")

# metadata labels seen in this corpus (28 Sep list, 4 Oct scan, 5 Oct scan); compared in lower case
METADATA = {x.lower() for x in """
IN THIS PHOTO|FEATURE|PHOTO CREDIT|IMAGE CREDIT|Photo|PHOTOS|Photos|Pic|Release Date|Release date|Album|Disc|Label|Labels|
Publisher|Producers|Producer|Released|Song|Peak Date|Type|Review|Standout Tracks|Standout Track|Choice Cut|Verdict|
Key track|Key Cut|Killer track|Featured Song|Album Review|Buy|Browse by Topic|Stream|Twitter|Facebook|Instagram|View|
More from|From|YouTube|Clip|Read more|READ MORE|Official|Spotify|Website|Related Artists|Hear it|Video|Browse category|
Listen|Related|RELATED|SoundCloud|Tickets|Description|Source|When|Artist|INTERVIEW|Currency|Interview|Note|Dates|Role|Date|
Q&A|Director|Where|Starring|By|Result|Hometown|Cities|Author|Posted by|Synopsis|ALSO|Words|Topic|Links|
LIVE ACTION|GRADE|CD|Genre|Who|Publication Date|BONUS BEATS|BONUS BONUS BEATS|Information|POP MUSIC|
Opera Recording|Flashback|Format|Formats|Catalogue|Catalog|Cat No|Price|Rating|Runtime|Length|Duration|Venue|Location|
Time|Tags|Tag|Category|Categories|Credits|Credit|Mixed by|Written by|Recorded|Engineer|Artwork|Track|Tracks|Tracklist|
Tracklisting|Track listing|Bonus tracks|Editor|Edited by|Photography|Photographs|Image|Images|Caption|Copyright|Watch|
Subscribe|Share|Email|Print|Comments|Comment|Reply|Updated|Published|Posted|Filed under|Filed Under|See also|Further reading|
Also read|Read|More|Next|Previous|Prev|Home|Menu|Search|Login|Sign up|Contact|About|Awards|Award|Nominee|Nominees|
Winner|Winners|Category|Ceremony|Notes|Sources|References|Web|Site|URL|Born|Died|Origin|Years active|Occupation|
Occupations|Instruments|Instrument|Members|Past members|Associated acts|Spouse|Spouses|Children|Height|Education|
Genres|Recorded at|Studio|Length|Sales|Certification|Chart|Charts|Peak|Weeks|Edition|Volume|Issue|Pages|Page|
Rank|Position|Points|Score|Total|Year|Years|Season|Network|Air date|Aired|Channel|
Station|Program|Programme|Format|Contributor|Contributors|Writer|Writers|Composer|Lyricist|Arranger|Conductor|Orchestra|
Featuring|Feat|Ft|Vocals|Guitar|Bass|Drums|Keyboards|Piano|Synthesizer|Saxophone|Percussion|Backing vocals|
Cover|Sleeve|Design|Designer|Management|Manager|Booking|Press|PR|Distribution|Distributor|Imprint|ISBN|UPC|EAN|Barcode|
Download|Downloads|Purchase|Order|Pre-order|Available|Availability|Out now|Out Now|Out|Now|Today|Tonight|Tomorrow|
Update|UPDATE|Editor's note|Editor’s note|NOTE|Warning|Disclaimer|Advertisement|Ad|Sponsored|Promoted|PLUS|Plus|
Also|Watch|WATCH|LISTEN|Listen now|SEE ALSO|See Also|MORE|PHOTO|Getty Images|Getty|AP|Reuters|Credit|Pictured|
Above|Below|Top|Left|Right|Main image|Main picture|Picture|Pictures|Illustration|Graphic|Chart|Map|Table|Figure|Fig|
Photo Credits|Featured Videos|Featured Video|Retracted|Collapse|Expand|Show more|Show less|Transcript|TRANSCRIPT|
Extract|Excerpt|Highlights|My Take|Original Chart Peak|Should Be Covered By|Number of Tracks|Classic Track|Advert|
Read Our Review|Read Our Feature|Albums|Singles|Name|Profession|Age|Earnings|First Contact|Previous Signings|
Recent Signings|Reasons|Nomination|Debates|Film|Why It Was Shelved|Closed|Cast|Theater|Theatre|Changes|Opened|
Exclusive|Film Review|Book Review|Sample|Loop Time|End Run|Monitoring|Microphones|Signal Processors|Computers|Software|
Lyrics|Chorus|Verse|Bridge|Intro|Outro|Hook|Refrain|Pre-chorus|Spoken|Man|Woman|Both|All|Together|Choir|Personnel|
Line-up|Lineup|Setlist|Set list|Musicians|Band members|Recorded live|Mastered by|Mastering|Session musicians|
Laughter|Laughs|Laughing|Chuckles|Applause|Music|Inaudible|Crosstalk|Silence|Pause|Sighs|Fade Out|Plays|
Most Telling Lyric|Male Vocal|Female Vocal|Best New Artist|Record of the Year|Album of the Year|Song of the Year
""".replace("\n", "").split("|") if x}
ATTRIB_LABEL = re.compile(r"^(?:he|she|they|we|it|i|you|[A-Z][\w'’-]*)\s+" + SPEECH_VERB + r"$", re.I)
PRONOUN_LABEL = re.compile(r"^(?:he|she|they|we|it|i|you|this|that|these|those|there|here|but|and|so|then|now|note|"
                           r"answer|question|yes|no|well|ok|okay|update|example|warning|tip|fact|bonus|p\.?s|"
                           r"edit|edited|addendum|correction|clarification|spoiler|trivia|fun fact)$", re.I)
SUNG = re.compile(r"\((?:sing|soundbite|vocaliz|rapping|chanting|music|instrumental)", re.I)

# words and marks that show a value is someone speaking
# ── pronouns, in English and in the other languages of the turn set ──────────
# route() sets the record's language; the patterns below then also match that language's pronouns, so a
# French or Spanish interview is read the same way as an English one. English records are unaffected.
_LANG = ["en"]
FP_OTHER = {
    "fr": r"je|j'|moi|me|m'|nous|mon|ma|mes|notre|nos",
    "es": r"yo|me|mi|mis|conmigo|nosotros|nosotras|nos|nuestro|nuestra|nuestros|nuestras",
    "de": r"ich|mich|mir|mein|meine|meinen|meinem|meiner|wir|uns|unser|unsere|unseren",
    "it": r"io|mi|me|mio|mia|miei|mie|noi|ci|nostro|nostra|nostri|nostre",
    "da": r"jeg|mig|min|mit|mine|vi|os|vores",
}
FPS_OTHER = {
    "fr": r"je|j'|moi|me|m'|mon|ma|mes",
    "es": r"yo|me|mi|mis|conmigo",
    "de": r"ich|mich|mir|mein|meine|meinen|meinem|meiner",
    "it": r"io|mi|me|mio|mia|miei|mie",
    "da": r"jeg|mig|min|mit|mine",
}
YOU_OTHER = {
    "fr": r"vous|tu|toi|te|t'|votre|vos|ton|ta|tes",
    "es": r"tú|tu|tus|usted|ustedes|te|ti|contigo|vosotros|vosotras|os|vuestro|vuestra",
    "de": r"du|dich|dir|dein|deine|deinen|Sie|Ihnen|Ihr|Ihre|euch|euer|eure",
    "it": r"tu|te|ti|tuo|tua|tuoi|voi|vi|vostro|vostra|lei",
    "da": r"du|dig|din|dit|dine|De|Dem|jer|jeres",
}


class _ByLanguage:
    """A compiled English pattern plus, for the record's language, the same pattern with that language's
    pronouns added (matched without regard to case)."""

    def __init__(self, pattern, flags=0, extra=None):
        self.en = re.compile(pattern, flags)
        self.by = {}
        for lang, words in (extra or {}).items():
            whole = "|".join(w for w in words.split("|") if not w.endswith("'"))
            elided = "|".join(w for w in words.split("|") if w.endswith("'"))      # French j', m', t' + verb
            alt = r"(?i:(?<![\w'’])(?:" + whole + r")(?![\w'’]))"
            if elided:
                alt += r"|(?i:(?<![\w'’])(?:" + elided + r")(?=\w))"
            self.by[lang] = re.compile(r"(?:" + pattern + r")|" + alt, flags)

    def _rx(self):
        return self.by.get(_LANG[0], self.en)

    def search(self, s, *args):
        return self._rx().search(s, *args)

    def findall(self, s, *args):
        return self._rx().findall(s, *args)

    def finditer(self, s, *args):
        return self._rx().finditer(s, *args)

    def match(self, s, *args):
        return self._rx().match(s, *args)


SPEECH_MARK = _ByLanguage(r"\b(?:I|I'm|I've|I'd|I'll|me|my|mine|we|we're|we've|we'd|our|us|you|your|you're|you've)\b|\?|"
                          r"\((?:laughter|laughs|laughing)\)|\[(?:laughter|laughs|laughing)\]", re.I,
                          {k: FP_OTHER[k] + "|" + YOU_OTHER[k] for k in FP_OTHER})
# radio and podcast transcripts
TRANSCRIPT_CUE = re.compile(r"\((?:LAUGHTER|SOUNDBITE|CROSSTALK|APPLAUSE|CHEERING|MUSIC|BEGIN|END)\b|"
                            r"\b(?:HOST|BYLINE|ANCHOR|CORRESPONDENT|GUEST)[ \t]*:|\[(?:laughter|applause|crosstalk)\]", re.I)
GENERIC_SPEAKER = {"q", "a", "question", "answer", "interviewer", "interviewee", "host", "co-host", "guest", "caller",
                   "announcer", "audience", "audience member", "narrator", "moderator", "reporter", "presenter", "dj",
                   "mc", "panelist", "panellist", "speaker", "participant", "student", "fan"}
# capitalised words that sit next to a name without being part of it
FUNC = {"The", "A", "An", "And", "But", "Or", "Of", "In", "On", "At", "To", "For", "With", "By", "From", "As", "If", "So",
        "Then", "When", "While", "This", "That", "These", "Those", "It", "He", "She", "They", "We", "I", "You", "His",
        "Her", "Their", "Our", "My", "Your", "Its", "Says", "Said", "Asked", "Told", "Tells", "Like", "After", "Before",
        "Since", "Although", "Though", "Because", "Meanwhile", "Now", "Today", "Also", "Even", "Only", "Just", "Here",
        "There", "Yes", "No", "Well", "Oh", "Sir", "Mr", "Mrs", "Ms", "Dr", "Best", "New", "Top", "Read", "Watch",
        "Listen", "Grammy", "Album", "Song", "Record", "Award", "Awards", "Review", "Live", "Official", "Video", "Music",
        "Radio", "TV", "Show", "Classic", "Exclusive", "Interview", "Feature", "Update", "Episode", "Part", "Chapter",
        "Day", "Night", "Year", "January", "February", "March", "April", "May", "June", "July", "August", "September",
        "October", "November", "December", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}


def _clean_label(lab):
    lab = re.sub(r"\s+", " ", lab).strip().rstrip(".") if not re.fullmatch(r"(?:[A-Z]\.){1,4}", lab.strip()) else lab.strip()
    lab = re.sub(r"\s*(?:\([^()]*\)|\[[^\[\]]*\])$", "", lab)
    return re.sub(r"^(?:Mr|Ms|Mrs|Dr|Miss)\.?\s+", "", lab)


def _keep_label(lab, val):
    return not (len(lab) > 40 or lab.lower() in METADATA or ATTRIB_LABEL.match(lab) or PRONOUN_LABEL.match(lab)
                or SUNG.match(val))


def label_lines(paras, context=""):
    """(paragraph index, label, value, style) for every paragraph that starts with a speaker label.
    style: colon, dash, bracket, space (initials and a space), or alone (the label stands on its own line and
    the words follow)."""
    out = []
    whole = None
    for i, p in enumerate(paras):
        b = BRACKET_LABEL.match(p)
        if b:
            lab, val = b.group("lab").strip(), b.group("val").strip()
            if _keep_label(lab, val):
                out.append([i, lab, val, "bracket"])
            continue
        m = NPR_LABEL.match(p)
        style = "npr"
        if not m:
            m = ROLE_LABEL.match(p)
            style = "role"
            if m and m.group("lab").split()[0].title() not in FIRST_NAMES:
                m = None
        if not m:
            m = LABEL_LINE.match(p)
            style = "colon"
        if not m:
            m = DASH_LABEL.match(p)
            style = "dash"
            if not m:
                m = SPACE_INITIALS.match(p)
                style = "space"
                if not m:
                    continue
                if whole is None:
                    whole = context + "\n" + "\n".join(paras)
                rx = _initials_rx(m.group("lab"))
                if not (rx and rx.search(whole)):
                    continue
        g = m.groupdict()
        lab = _clean_label(g.get("lab") or g.get("lab2") or g.get("lab3") or g.get("qa") or "")
        val = (g.get("val") or g.get("val2") or g.get("val3") or "").strip()
        if not lab or not _keep_label(lab, val):
            continue
        out.append([i, lab, val, style])
    # a bare name on its own line that comes back two or more times ("Teddy Riley" ... "Teddy Riley") is a label
    # without a colon; a name that appears once is more likely a heading
    bare = [(i, m.group("name")) for i, p in enumerate(paras) if len(p.strip()) <= 40
            for m in [BARE_NAME.match(p)] if m and m.group("name").split()[0].title() in FIRST_NAMES]
    times = Counter(nm for _, nm in bare)
    if sum(1 for k in times.values() if k >= 2) >= 2:
        taken = {x[0] for x in out}
        out += [[i, nm, "", "alone"] for i, nm in bare if times[nm] >= 2 and i not in taken]
        out.sort(key=lambda x: x[0])
    # a label alone on its line: the words are in the next paragraph(s)
    starts = {x[0] for x in out}
    for x in out:
        if len(x[2]) <= 2:
            j = x[0] + 1
            while j < len(paras) and not paras[j].strip():
                j += 1
            if j < len(paras) and j not in starts:
                x[2] = paras[j].strip()
                x[3] = "npr" if x[3] == "npr" else "alone"
    return [tuple(x) for x in out]


def inline_labels(paras):
    """(paragraph index, label, value, 'inline', offset in paragraph) for labels inside paragraphs."""
    out = []
    for i, p in enumerate(paras):
        ms = [m for m in INLINE_LABEL.finditer(p) if m.start() > 0]
        for k, m in enumerate(ms):
            lab = _clean_label(m.group("lab"))
            end = ms[k + 1].start() if k + 1 < len(ms) else len(p)
            val = p[m.end():end].strip()
            if _keep_label(lab, val):
                out.append((i, lab, val, "inline", m.start()))
    return out


def _initials_rx(lab):
    letters = re.sub(r"[^A-Z]", "", lab)
    if not 2 <= len(letters) <= 4:
        return None
    parts = [letters[0] + r"[a-zà-ÿA-Z'’]+"] + [r"(?:[A-Z]\.?[ \t]+)?" + c + r"(?:[\w'’-]+|\.)" for c in letters[1:]]
    return re.compile(r"(?<![\w-])" + r"[ \t]+".join(parts) + r"(?![\w-])")


def person_like(lab, text_wo_labels):
    """Does this label name a person? Generic speaker words count; a first name counts, alone or followed by
    capitalised words ('Ann Powers', 'Geoff'); initials count when a name in the record has those initials
    ('PO' and Phil Oakey); a surname counts when the record names someone with it ('Collen' and Phil Collen)."""
    low = lab.lower()
    if low in GENERIC_SPEAKER or low.startswith("unidentified"):
        return True
    letters = re.sub(r"[^A-Z]", "", lab)
    if re.fullmatch(r"(?:[A-Z]\.?){2,4}", lab.replace(" ", "")):
        rx = _initials_rx(lab)
        if rx and rx.search(text_wo_labels):
            return True
        if len(letters) <= 2 or "." in lab:
            return False          # two letters, or initials with full stops: nothing else to try
    words = [w if not w.isupper() or len(w) <= 1 else w[0] + w[1:].lower() for w in lab.split()]
    if words[0] in FIRST_NAMES and all(w[0].isupper() and w not in FUNC for w in words[1:]):
        return True
    last = words[-1]
    if len(last) < 2 or last in FUNC or not (len(words) == 1 or (len(words) == 2 and len(words[0].strip(".")) == 1)):
        return False
    word = last[0] + "(?i:" + re.escape(last[1:]) + ")"
    for m in re.finditer(r"(?<![\w-])([A-Z][A-Za-zà-ÿ'’]+)[ \t]+(?:(?:[A-Z]\.|\"[^\"\n]{1,20}\"|'[^'\n]{1,20}')[ \t]+)?" + word +
                         r"(?![\w-])(?![ \t]*:)", text_wo_labels):
        if m.group(1).title() in FIRST_NAMES:
            return True
    return False


def merge_labels(lines):
    """One person, several labels: 'John Wetton' then 'JW', 'TONYA MOSLEY, HOST' then 'MOSLEY', 'Geoff Downes'
    then 'Geoff'. Initials, a surname or a first name that fits exactly one full-name label become that label."""
    labels = {x[1] for x in lines}
    full = [l for l in labels if len(l.split()) >= 2 and not re.search(r"(?i)\b(?:and|with)\b|[&/,+]", l)]
    mapping = {}
    for l in labels:
        if l in full:
            continue
        if re.fullmatch(r"(?:[A-Z]\.?){2,4}", l.replace(" ", "")):
            letters = re.sub(r"[^A-Z]", "", l)
            cands = [f for f in full if "".join(w[0] for w in f.split() if w[0].isupper()) == letters]
        else:
            cands = [f for f in full if l.lower() in (f.split()[-1].lower(), f.split()[0].lower())]
        if len(cands) == 1:
            mapping[l] = cands[0]
    return [(x[0], mapping.get(x[1], x[1])) + tuple(x[2:]) for x in lines]


def assess(lines, paras, text, title="", singletons=False):
    """Measure a set of label lines: counts, alternation, value lengths, how speech-like the values
    are, and how many labels name people. Labels used once are dropped unless singletons is True."""
    lines = merge_labels(lines)
    cnt = Counter(x[1] for x in lines)
    keep = {l for l, k in cnt.items() if k >= 2 or singletons}
    if "Q" in cnt and "A" in cnt:
        keep |= {"Q", "A"}
    lines = [x for x in lines if x[1] in keep]
    n = len(lines)
    if n < 4:
        return None
    labs = Counter(x[1] for x in lines)
    switches = sum(1 for a, b in zip(lines, lines[1:]) if a[1] != b[1])
    by = defaultdict(list)
    for x in lines:
        by[x[1]].append(len(x[2]))
    lmed = max(sorted(v)[len(v) // 2] for v in by.values())
    med_all = sorted(len(x[2]) for x in lines)[n // 2]
    top2 = sum(k for _, k in labs.most_common(2)) / n
    speech = sum(1 for x in lines if SPEECH_MARK.search(x[2])) / n
    # the record's text without the labels themselves, to look for the names elsewhere
    label_paras = {x[0] for x in lines}
    wo = title + "\n" + "\n".join(p if i not in label_paras else re.sub(r"^\W*[A-Z][\w.'’ &-]{0,40}[:–—-]", " ", p)
                                    for i, p in enumerate(paras))
    pl = {l: person_like(l, wo) for l in labs}
    person = sum(k for l, k in labs.items() if pl[l]) / n
    return dict(lines=lines, n=n, labs=labs, switches=switches, lmed=lmed, med_all=med_all, top2=top2,
                speech=speech, person=person, person_labels=pl, cue=bool(TRANSCRIPT_CUE.search(text)))


def dense_enough(paras, lines):
    """The labelled stretch is a conversation, not an article with a few labelled lines in it: at most half of
    its characters may be narration (quotations attributed in the third person, long paragraphs with no
    first person)."""
    a, b = min(x[0] for x in lines), max(x[0] for x in lines)
    labelled = {x[0] for x in lines}
    total = narr = 0
    for i in range(a, b + 1):
        p = paras[i].strip()
        if not p:
            continue
        total += len(p)
        if i not in labelled and narration(p):
            narr += len(p)
    return total == 0 or narr <= 0.5 * total


def labelled_verdict(paras, text, lines=None, title=""):
    """Is this (part of the) record a labelled Q&A? Returns (yes, kind, lines, measures)."""
    ok, kind, kept, meas = _labelled_verdict(paras, text, lines, title)
    if ok and not dense_enough(paras, kept):
        return False, "", kept, meas
    return ok, kind, kept, meas


def _labelled_verdict(paras, text, lines=None, title=""):
    lines = label_lines(paras, title) if lines is None else lines
    a = assess(lines, paras, text, title)
    # oral histories where most speakers say one thing each, and radio transcripts that name each speaker once
    b = assess(lines, paras, text, title, singletons=True)
    if b and (not a or a["n"] < 6):
        caps = sum(1 for x in b["lines"] if x[1].isupper()) / b["n"]
        quoted = sum(1 for x in b["lines"] if x[2][:1] in "\"'") / b["n"]
        if b["n"] >= 6 and len(b["labs"]) >= 4 and b["person"] >= 0.75 and b["med_all"] >= 80 and b["speech"] >= 0.6 and (
                TITLE_SECTIONS.search(title) or re.search(r"\bexcerpt\b", title, re.I) or quoted >= 0.5):
            return True, "labelled turns, many speakers", b["lines"], b
        if b["cue"] and b["n"] >= 3 and len(b["labs"]) >= 2 and caps >= 0.6 and b["person"] >= 0.5 and \
                any(x[3] == "npr" for x in b["lines"]) and dense_enough(paras, b["lines"]):
            return True, "radio transcript", b["lines"], b
    if not a:
        return False, "", [], None
    n, labs, sw, lmed = a["n"], a["labs"], a["switches"], a["lmed"]
    sp, pe, cue = a["speech"], a["person"], a["cue"]
    if n >= 6 and len(labs) >= 2 and sw >= 2 and a["top2"] >= 0.6 and lmed >= 40 and (
            sp >= 0.4 or (pe >= 0.6 and (sp >= 0.2 or cue)) or (cue and pe >= 0.5)):
        return True, "labelled turns", a["lines"], a
    if n >= 8 and len(labs) >= 3 and sw >= 0.5 * n and a["med_all"] >= 30 and (
            (sp >= 0.55 and pe >= 0.5) or (pe >= 0.6 and sp >= 0.3) or (cue and pe >= 0.4)):
        return True, "labelled turns, many speakers", a["lines"], a
    if n >= 4 and len(labs) >= 2 and sw >= 3 and a["top2"] >= 0.8 and lmed >= 40 and (sp >= 0.5 or (pe >= 0.6 and sp >= 0.2)) and any(
            l in ("Q", "A") or re.fullmatch(r"[A-Z]{2,4}", l) or l.isupper() for l in labs):
        return True, "labelled turns, short", a["lines"], a
    top, k = labs.most_common(1)[0]
    own = [x for x in a["lines"] if x[1] == top]
    # one-sided, the interviewer's questions carry the label ("Michael Christopher: ...?")
    asks = [x for x in own if x[2].rstrip("\"' )").endswith("?")]
    if k >= 4 and len(asks) >= 0.6 * k:
        answered = 0
        for x in own:
            j = x[0] + 1
            while j < len(paras) and not paras[j].strip():
                j += 1
            if j < len(paras) and not LABEL_LINE.match(paras[j]) and is_answer(paras[j]):
                answered += 1
        if answered >= 0.6 * k:
            return True, "questions labelled", own, a
    # one-sided: only the answers carry a label ("SH: ..."), the questions are plain paragraphs
    meds = sorted(len(x[2]) for x in own)
    own_speech = sum(1 for x in own if SPEECH_MARK.search(x[2])) / k
    if k >= 6 and meds[len(meds) // 2] >= 60 and own_speech >= 0.4:
        idx = [x[0] for x in own]
        qs = sum(1 for p, q in zip(idx, idx[1:]) for j in range(p + 1, q) if is_question(paras[j]))
        if qs >= 3:
            return True, "answers labelled", own, a
    # one speaker labelled throughout ("Ms. FIONA: I love the instrumentation...")
    if k >= 4 and meds[len(meds) // 2] >= 60 and own_speech >= 0.6 and a["person_labels"][top] and len(labs) <= 3 \
            and all(x[3] != "inline" for x in own):
        return True, "one speaker labelled", own, a
    return False, "", a["lines"], a


def lecture_transcript(text, title):
    """Red Bull Music Academy lectures: 'Hosted by  Emma Warren  We're very lucky... Martyn Ware  Thank you.'
    The names stand before each turn with no colon. Returns (host, guest, offsets) or None."""
    m = re.search(r"Hosted by[ \t]+((?:[A-Z][\w.'’-]*|\"[^\"\n]{1,20}\")(?:[ \t]+(?:[A-Z][\w.'’-]*|\"[^\"\n]{1,20}\")){1,3})",
                  text)
    guest = (title or "").strip()
    if not m or not re.fullmatch(r"[A-Z][\w.'’-]*(?:[ \t]+[A-Z][\w.'’-]*){0,3}", guest):
        return None
    toks = m.group(1).split()

    def occ(name):
        rx = re.compile(r"(?<![\w-])" + r"[ \t]+".join(re.escape(t) for t in name.split()) +
                        r"(?=[ \t]+[A-Z\"'(\[…])", re.I)
        return [x.start() for x in rx.finditer(text)]

    host = None
    for k in range(len(toks), 1, -1):
        cand = " ".join(toks[:k])
        if len(occ(cand)) >= 4:
            host = cand
            break
    if not host:
        return None
    g, h = occ(guest), occ(host)
    if len(g) < 4:
        return None
    return host, guest, sorted(set(g + h))


# ── questions and answers without labels ─────────────────────────────────────
FIRST_PERSON = _ByLanguage(r"\b(?:I|I'm|I've|I'd|I'll|me|my|mine|myself|we|we're|we've|we'd|we'll|us|our|ours|ourselves)\b",
                           0, FP_OTHER)
FP_SINGULAR = _ByLanguage(r"\b(?:I|I'm|I've|I'd|I'll|me|my|mine|myself)\b", 0, FPS_OTHER)
SECOND_PERSON = _ByLanguage(r"\b(?:you|your|you're|you've|you'd|you'll|yourself|yourselves)\b", re.I, YOU_OTHER)
QUOTE_START = re.compile(r"^\s*[\"']")
THIRD_ATTR_ANY = re.compile(r"[,.!?][\"']\s*(?:he|she|they|[A-Z][\w'-]+(?:\s+[A-Z][\w'-]+)?)\s+" + SPEECH_VERB + r"\b"
                            r"|\b(?:he|she|they)\s+" + SPEECH_VERB + r"\b", re.I)


LIST_ITEM = re.compile(r"^\s*(?:\d{1,4}[.)]\s*)?[^?\n]{1,60}\s[–—-]\s[\"']?[A-Z0-9]")


# set by route() for each record: readers' questions signed with a handle ("...? Integrity411"), and pages whose
# title says they are a Q&A ("Q&A: Martin Rushent", "FAN QUESTIONS!")
_HANDLES = [False]
_TITLE_LINE = [-1]


def nonempty_title_line(paras, title):
    first = next((p for p in paras if p.strip()), "")
    return bool(title) and first.strip().lower()[:60] == title.strip().lower()[:60]
_QA_TITLE = [False]
HANDLE_END = re.compile(r"\?[ \t]+[\w.@-]{3,30}[ \t]*$")
QA_TITLE = re.compile(r"\bQ[ \t]?&[ \t]?A\b|\bfan questions?\b|\byou asked\b|\breaders?['’]?[ \t]+(?:questions|ask)\b|"
                      r"\bwebchat\b|\bquestions? (?:and|&) answers?\b|\bask (?:me|us) anything\b", re.I)
# a reader's letter in a Q&A page: "KJ writes, ..."
READER_Q = re.compile(r"^\s*[A-Z][\w.'’-]*(?:\s+[A-Z][\w.'’-]*){0,2}\s+(?:writes|asks|wrote|asked)\s*[,:]")
# a paragraph opening with a short quoted title is not a quotation: '"One Lost Day" is your newest release...'
GREETING = re.compile(r"^\s*(?:hello|hi|hey|dear|greetings)\b[^\n]{0,30}[,!.]?\s*$", re.I)
QUOTED_TITLE_START = re.compile(r"^\s*[\"'][^\"'\n]{1,80}[\"'][\s,:;]+\S")


def is_question(p):
    s = re.sub(r"\s*\([^()]{1,60}\)\s*$", "", p.strip()).rstrip("\"' ")
    if _HANDLES[0] and HANDLE_END.search(s):
        s = s[:HANDLE_END.search(s).start() + 1]
    if _QA_TITLE[0] and READER_Q.match(s) and SECOND_PERSON.search(s) and 8 <= len(s) <= 600:
        return True
    letters = [c for c in s if c.isalpha()]
    caps = sum(1 for c in letters if c.isupper()) / max(1, len(letters))
    quoted = QUOTE_START.match(p) and not (QUOTED_TITLE_START.match(p) and len(s) <= 300 and
                                           SECOND_PERSON.search(re.sub(r"^\s*[\"'][^\"'\n]{1,80}[\"']", " ", p)))
    return s.endswith("?") and 8 <= len(s) <= 600 and not quoted and not (
        len(s) < 120 and LIST_ITEM.match(s)) and caps < 0.7


def is_answer(p, lenient=False):
    s = p.strip()
    if not s or is_question(p) or len(s) < 15:
        return False
    # third-person narration ('..." she says.') is not an answer; once a record is known to be a Q&A, an answer can
    # report others ("whatever they said")
    return not THIRD_ATTR_ANY.search(s) or (lenient and bool(FP_SINGULAR.search(_unquote_para(s)[:300])))


def qa_pairs(paras):
    """(question paragraph, answer paragraph) for questions followed directly by an answer."""
    idx = [i for i, p in enumerate(paras) if p.strip()]
    return [(a, b) for a, b in zip(idx, idx[1:]) if a != _TITLE_LINE[0] and is_question(paras[a]) and is_answer(paras[b])]


# a paragraph that opens with the interviewer's question(s) and runs straight into the answer
RUN_IN = re.compile(r"^[ \t]*(?P<q>(?:[^\n?\"]{3,300}?[ \t]?\?[ \t]+){1,4})(?P<a>[A-Z][^\n]{29,})$")


def run_in_pairs(paras):
    """Paragraphs that start with a question addressed to the interviewee ('you') and continue with a
    first-person answer, with no quotation marks around either."""
    out = []
    for i, p in enumerate(paras):
        if len(p) > 20000 or QUOTE_START.match(p):
            continue
        m = RUN_IN.match(p)
        if not m:
            continue
        q, a = m.group("q"), m.group("a")
        head = a[:300]
        if SECOND_PERSON.search(q) and FIRST_PERSON.search(re.sub(r'"[^"]*"', " ", head)) and \
                not THIRD_ATTR_ANY.search(head) and '"' not in q:
            out.append((i, m.start("a")))
    return out


ADDRESSED = _ByLanguage(r"\b(?:you|your|you're|you've|you'd|you'll|yourself|yourselves|he|she|his|her|him)\b", re.I,
                        YOU_OTHER)
# a quotation attributed in the third person to a named person ('..." says Hucknall.', '..." Flower says')
THIRD_ATTR_NAMED = re.compile(r"[,.!?…-][\"']\s*(?:" + SPEECH_VERB + r"\s+(?P<who2>[A-Z][\w'’-]+(?:\s+[A-Z][\w'’-]+)?)|"
                              r"(?P<who>[A-Z][\w'’-]+(?:\s+[A-Z][\w'’-]+)?)\s+" + SPEECH_VERB + r")\b")
# a quotation attributed in the third person ('..." says Hucknall.'): the journalist is narrating again
THIRD_ATTR_QUOTE = re.compile(r"[,.!?…-][\"']\s*(?:he|she|they|[A-Z][\w'’-]+(?:\s+[A-Z][\w'’-]+)?)\s+" + SPEECH_VERB + r"\b")


def interviewer_statement(p):
    """A short paragraph addressed to the interviewee that is not a question ('Tell me about the new album.')."""
    s = p.strip()
    return bool(25 <= len(s) <= 300 and len(s.split()) >= 5 and s[-1] in ".…!" and not s[0].isdigit()
                and SECOND_PERSON.search(s) and not FP_SINGULAR.search(_unquote_para(s))
                and not QUOTE_START.match(s) and not THIRD_ATTR_QUOTE.search(s))


# a questionnaire prompt: a short line put to the interviewee without a question mark ("Your aesthetic hero",
# "A book you've found inspiring", "What gets you up in the morning")
def questionnaire_prompt(p):
    s = p.strip()
    words = s.split()
    caps = sum(1 for w in words if w[:1].isupper()) / max(1, len(words))
    return bool(6 <= len(s) <= 70 and len(words) >= 2 and s[-1] not in ".!:;,\"'" and SECOND_PERSON.search(s)
                and not FP_SINGULAR.search(s) and not QUOTE_START.match(s) and not re.match(r"\s*[\d(•*-]", s)
                and caps < 0.6)            # not a numbered track list or a song title ("8. Within You Without You")


# an opening prompt from the interviewer ("Thanks for sitting down to talk with us today! Let's start with a little
# bit about your background...", "Tell us about some cool memories you have from being on tour!")
OPENING_PROMPT = re.compile(r"^\s*(?:thanks?(?: you)? (?:so much )?for|let'?s (?:start|begin|talk)|tell (?:us|me)\b|talk (?:us|me) "
                            r"(?:through|about)|describe|explain|walk (?:us|me) through|give (?:us|me)|share (?:with )?(?:us|me)|"
                            r"congratulations)", re.I)


# the interviewer reading out a prompt that is not a question ("The first card is the EMS Polysynthi…") or reacting
# between answers ("That's amazing!")
def interviewer_prompt(p):
    s = p.strip()
    if not s or FP_SINGULAR.search(_unquote_para(s)) or QUOTE_START.match(s) or THIRD_ATTR_QUOTE.search(s):
        return False
    return (8 <= len(s) <= 120 and (s.endswith("…") or s.endswith("...")) and len(s.split()) >= 3) or \
        (len(s) <= 60 and s.endswith("!") and len(s.split()) >= 2)


def narration(p):
    """A paragraph that is the journalist narrating, not an answer: a quotation attributed in the third person,
    page furniture, or a long paragraph with no first person at all outside quotation marks."""
    s = p.strip()
    if THIRD_ATTR_QUOTE.search(s) or FURNITURE.match(s):
        return True
    body = s if QUOTE_START.match(s) else _unquote_para(s)     # some print every answer inside quotation marks
    return len(s) >= 400 and not FIRST_PERSON.search(body)


def qa_units(paras, min_answer=50, prompts=False, lenient=None):
    """Interview units: a question (on its own line, or run into its answer) and the answer paragraphs after it,
    up to the next question, a paragraph of narration, or page furniture. A short statement addressed to the
    interviewee after an answer starts a new unit, as interviewers do not always ask with a question mark.
    Returns [(first paragraph, last paragraph, answer opens in the first person?, question is put to someone?)]
    for units with an answer. A question is put to someone when it says 'you', or asks about 'he' or 'she' (some
    magazines print questions in the third person: 'How does he cope with producing other guitarists?')."""
    idx = [i for i, p in enumerate(paras) if p.strip()]
    run_at = dict(run_in_pairs(paras))
    runs = set(run_at)
    starts = {i for i in idx if i != _TITLE_LINE[0] and is_question(paras[i])} | runs
    # a questionnaire (at least four short prompts each answered in the first person)
    qp = [i for k, i in enumerate(idx[:-1]) if i not in starts and questionnaire_prompt(paras[i]) and
          idx[k + 1] not in starts and len(paras[idx[k + 1]].strip()) >= 40 and
          paras[idx[k + 1]].strip()[-1] in ".!?…\"'”)" and          # an answer is a sentence, not a lyric line
          FP_SINGULAR.search(_unquote_para(paras[idx[k + 1]])[:300])]
    if len(qp) >= 4:
        starts |= set(qp)
    if len(starts) >= 2:
        # prompts and opening statements count as the interviewer's turns when an answer in the first person follows
        first_q = min(starts)
        for k, i in enumerate(idx[:-1]):
            if i in starts:
                continue
            nxt = paras[idx[k + 1]]
            if idx[k + 1] in starts or not FIRST_PERSON.search(_unquote_para(nxt)[:400]):
                continue
            if paras[i].rstrip().endswith(("…", "...")) and re.match(r"\s*[a-z]", nxt):
                continue                  # "HER album, RED, arrives…" / "on 31st August, so I have been..." is one sentence
            # a prompt or reaction never comes straight after a question: "In one word: great!" answers one
            after_q = k > 0 and idx[k - 1] in starts
            if (i < first_q and interviewer_statement(paras[i]) and OPENING_PROMPT.match(paras[i])) or \
                    (prompts and not after_q and interviewer_prompt(paras[i])):
                starts.add(i)
    units = []
    k = 0
    while k < len(idx):
        i = idx[k]
        if i not in starts:
            k += 1
            continue
        last, answer = i, ""
        run_start = run_at.get(i, 0)
        if i in runs:
            answer = paras[i][run_start:]
        j = k + 1
        while j < len(idx):
            q = idx[j]
            p = paras[q]
            if q in starts:
                break
            if narration(p):
                nxt = next((x for x in idx if x > q), None)
                # the speaker quoting someone inside their own account: first person outside the quotation marks
                # in this paragraph and the next ('"Well," he says, "We could be New Shoes!" and we should spell it')
                if not (prompts and answer and len(p.strip()) <= 250 and nxt is not None and nxt not in starts and
                        FIRST_PERSON.search(_unquote_para(p)) and not QUOTE_START.match(paras[nxt]) and
                        FIRST_PERSON.search(_unquote_para(paras[nxt])[:300])):
                    break
            if answer and interviewer_statement(p) and j + 1 < len(idx) and idx[j + 1] not in starts and \
                    FIRST_PERSON.search(_unquote_para(paras[idx[j + 1]])):
                starts.add(q)
                break
            if not answer and GREETING.match(p):
                last = q                   # "Hello KJ!" before the answer to a reader's letter
                j += 1
                continue
            if not answer and not is_answer(p, lenient=prompts if lenient is None else lenient):
                break
            answer += " " + p
            last = q
            j += 1
        n_ans = sum(1 for x in idx if i < x <= last) + (1 if i in runs else 0)
        if answer.strip() and len(answer.strip()) >= min_answer * max(1, n_ans) and len(paras[i].strip()) >= 12:
            first_answer = paras[i][run_start:] if i in runs else paras[next(x for x in idx if x > i)]
            q_text = paras[i][:run_start] if i in runs else paras[i]
            units.append((i, last, bool(FIRST_PERSON.search(first_answer[:400])),
                          2 if SECOND_PERSON.search(q_text) else 1 if ADDRESSED.search(q_text) else 0))
        k = j if j > k else k + 1
    return units


def first_person_answers(paras, pairs):
    """Share of answers (with their continuation paragraphs) that speak in the first person."""
    if not pairs:
        return 0.0
    qs = {a for a, _ in pairs}
    hits = 0
    for a, b in pairs:
        j, txt = b, ""
        while j < len(paras) and j not in qs and len(txt) < 1500:
            txt += " " + paras[j]
            j += 1
        hits += bool(FIRST_PERSON.search(txt))
    return hits / len(pairs)


# ── quotation convention (needed to read first person outside quotes) ────────
ATTR_AFTER = r"\s*(?:,\s*)?(?:he|she|they|[A-Z][\w'-]+(?:\s+[A-Z][\w'-]+)?)\s+" + SPEECH_VERB + r"\b"
D_ATTR = re.compile(r"[,.!?]\"" + ATTR_AFTER + r"|" + SPEECH_VERB + r"[,:]?\s+\"[A-Z]")
S_ATTR = re.compile(r"[,.!?]'" + ATTR_AFTER + r"|" + SPEECH_VERB + r"[,:]?\s+'[A-Z]")


def convention(text):
    d = len(D_ATTR.findall(text))
    s = len(S_ATTR.findall(text))
    return ("single" if s > d and s >= 2 else "double"), d, s


def _unquote_para(p):
    """A paragraph with its double-quoted passages blanked out. A quotation that opens and does not close in
    the paragraph (British magazines open every paragraph of a long quotation and close only the last) is
    blanked to the end of the paragraph."""
    p = re.sub(r'"[^"\n]{0,2000}"', " ", p)
    k = p.find('"')
    return p if k < 0 else p[:k]


def strip_quoted(text, conv):
    """Text with quoted passages removed (rough: used only to measure first person outside quotes)."""
    if conv == "double":
        return "\n".join(_unquote_para(p) for p in text.split("\n"))
    return re.sub(r"(?:(?<=^)|(?<=[\s(\[—-]))'[^'\n]{1,2000}?'(?=[\s,.;:!?)\]—-]|$)", " ", text)


THIRD_ATTR = re.compile(r"(?:[,.!?][\"']\s*(?:he|she|they|[A-Z][\w'-]+)\s+" + SPEECH_VERB + r"\b)|"
                        r"(?:\b(?:he|she|they)\s+" + SPEECH_VERB + r"\b)", re.I)
WORD = re.compile(r"\b\w+\b")


def first_person_profile(text, conv):
    """First person per 1,000 words outside quotation marks (singular, and singular plus plural),
    third-person attributions and quotation marks per 1,000 words."""
    outside = strip_quoted(text, conv)
    words = max(1, len(WORD.findall(outside)))
    allw = max(1, len(WORD.findall(text)))
    qmarks = text.count('"') if conv == "double" else len(re.findall(r"(?:(?<=^)|(?<=\s))'|'(?=\s|[,.;:!?]|$)", text))
    return dict(words=words, fp_per_k=1000 * len(FP_SINGULAR.findall(outside)) / words,
                fp_all_per_k=1000 * len(FIRST_PERSON.findall(outside)) / words,
                third_per_k=1000 * len(THIRD_ATTR.findall(text)) / allw, quote_per_k=1000 * qmarks / allw)


# titles that say the named person is the one speaking
TITLE_FP = re.compile(r"\b(?:my life|in my own words|as told to|how i (?:made|wrote)|soundtrack of (?:my|his|her|their|our) li(?:fe|ves)|my favou?rite|"
                      r"my top \d+|diary|letter to|in (?:his|her|their) own words|i wrote|what i've learned|"
                      r"the songs? that changed my life|me and my|my first|i was there|my \d+ songs|"
                      r"on (?:his|her|their) favou?rite|(?:his|her|their) favou?rite|(?:picks|chooses|selects) (?:his|her|their)|"
                      r"(?:his|her|their|our) life in (?:music|songs|\d+ songs)|records that changed|music that made me|"
                      r"desert island discs|first (?:and|&) last|as chosen by|chosen by|picked by|on my radar|"
                      r"my cultural life|cultural highlights|this much i know|the books? (?:of|that made|that changed) (?:my|me)|"
                      r"my perfect|my favou?rite things|songs i wish i'?d written|my (?:week|weekend|day) in|"
                      r"what i(?:'m| am) (?:listening|into)|the record that changed|my style|i'm a (?:fan|believer)|"
                      r"letter to my (?:younger|teenage) self|(?:his|her|their) (?:top|best) \d+ (?:songs|albums|records|tracks|picks|favou?rites))\b", re.I)
# "Jazzie B's 10 best 80s anthems", "Rick Astley's top 10 favourite rock songs": the named person chooses
TITLE_PICKS = re.compile(r"\b[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){0,3}['’]s\s+(?:top\s+\d+|\d+\s+(?:best|favou?rite|greatest)|"
                         r"favou?rite|best)\b")
# "Jane Wiedlin on the birth of LA punk", "Erasure's Andy Bell on ...": a person on a subject, often told in their words
TITLE_ON = re.compile(r"^(?:(?:The\s+)?[\w'’.&-]+['’]s?\s+)?[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){1,3}\s+on\s+\S")
# titles of pieces told by several people, each under their own name
TITLE_SECTIONS = re.compile(r"oral history|in their own words|remember(?:s|ing)?\b|tribute|toast|recall|how we (?:made|wrote)|"
                            r"how [A-Z][\w'’&. -]{1,40} made|the making of|behind the (?:scenes|song)|as told by|"
                            r"we were there|stories behind|collaborators|friends|\brip\b|r\.i\.p|in memoriam|tributes? (?:pour|flow)|"
                            r"pay tribute|remembered by|reacts? to", re.I)
NAME = r"[A-Z][\w'’.-]+(?:\s+(?:[A-Z][\w'’.-]*|de|van|von|la|le|da|di|del|'[^'\n]{1,25}'|\"[^\"\n]{1,25}\")){0,4}"
ROLE_WORD = (r"(?i:singer|songwriter|guitarist|guitar|bassist|bass|drummer|drums|keyboardist|keyboards|keys|vocalist|vocals|"
             r"producer|co-producer|engineer|manager|musician|director|dj|rapper|mc|writer|co-writer|composer|arranger|"
             r"founder|co-founder|member|frontman|frontwoman|saxophonist|sax|trumpeter|percussionist|dancer|backing|session|"
             r"label|a&r|promoter|photographer|stylist|choreographer|actor|actress|presenter|journalist|editor|executive|"
             r"designer|artist|programmer|mixer|remixer|lyricist|pianist|violinist|cellist|horn|choir|bandleader|"
             r"collaborator|friend|tour|roadie|publicist|agent|trumpet|trombone|saxophone|horns|flute|violin|cello|"
             r"harmonica|organ|piano|synths?|synthesi[sz]ers?|programming|percussion|strings|lawyer|attorney|"
             r"president|partner|associate|chairman|boss|owner|head|assistant|teacher|mentor|brother|sister|"
             r"father|mother|son|daughter|wife|husband|widow|bandmate)")
# "Nick Van Eede, singer, songwriter", "Robert Hodgens, AKA Bobby Bluebell, songwriter", "Herb Alpert: 'He put...'",
# "Stuart Matthewman (Sade co-founder)", "Youth, songwriter and producer"
ROLE_HEADING = re.compile(r"^\s*(?P<name>" + NAME + r")\s*(?:,\s*(?:(?:aka|AKA|a\.k\.a\.)\s+[^,\n]{2,40},\s*)?"
                          r"(?=[^\n]{0,60}\b" + ROLE_WORD + r"\b)[A-Za-z][\w ,'’&/()-]{2,100}"
                          r"|\((?=[^)\n]{0,40}\b" + ROLE_WORD + r"\b)[^)\n]{2,60}\)"
                          r"|\s[–—-]\s+(?=[^\n]{0,60}\b" + ROLE_WORD + r"\b)[^\n]{2,100}|:\s*[\"'].{3,200})\s*$")
# a short biography that introduces the speaker of the paragraphs after it (the Independent's "How We Met")
BIO_INTRO = re.compile(r"^\s*(?P<name>[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]*){1,3}),\s*\d{1,3},\s*(?:is|was|has|lives|works|"
                       r"runs|founded|co-founded|plays|sings|became|joined|formed)\b")
# a bare name on its own line: "Trevor Horn", "JIM KERR"
BARE_NAME = re.compile(r"^\s*(?P<name>(?:[A-Z][a-zà-ÿ'’.-]+|[A-Z]{2,})(?:\s+(?:[A-Z][\w'’.-]+|de|van|von|la|le|da|di|del)){1,2})\s*$")
# a short line naming a song or record: "Barbra Streisand – The Love Inside", "1. Song – Artist", or any short title line
TOPIC_HEADING = re.compile(r"^\s*(?:\d{1,2}[.)]\s*)?[\"']?[A-Z0-9][^\n]{1,80}?\s[–—-]\s[\"']?[A-Z0-9(][^\n]{0,80}$")
SHORT_LINE = re.compile(r"^\s*(?:\d{1,2}[.)]\s*)?[\"']?[A-Z0-9][^\n.!?]{1,70}[\"')]?\s*$")
NOT_NAME = {"The", "A", "An", "This", "That", "Our", "My", "Your", "His", "Her", "Their", "In", "On", "At", "When",
            "How", "What", "Why", "Where", "Who", "Read", "Watch", "Listen", "Related", "More", "Photo", "Video", "From",
            "By", "For", "With", "Disc", "Bonus", "Track", "Tracks", "Side", "Part", "Chapter", "Episode", "Every", "Latest",
            "Suggested", "Further", "Featured", "Today", "Search", "Name", "Copyright", "Album", "Albums", "Single"}


def fp_outside_quotes(p, limit=None, singular=False):
    s = _unquote_para(p)
    s = s if limit is None else s[:limit]
    rx = FP_SINGULAR if singular else FIRST_PERSON
    return bool(rx.search(s))


def _mentions(name, p, title=""):
    """Does the paragraph talk about this person by name (first name or surname) near its start? Names that
    are also in the title (the person the whole piece is about) do not count."""
    toks = [t for t in re.findall(r"[A-Z][\w'’-]+", name) if len(t) > 2]
    head = p[:300]
    return any(re.search(r"(?<![\w-])" + re.escape(t) + r"(?![\w-])", head) for t in toks)


COMMON_NOUNS = {"Mother", "Father", "Mum", "Mom", "Dad", "Brother", "Sister", "Son", "Daughter", "Wife", "Husband",
                "Friend", "Manager", "Producer", "Engineer", "Interviewer", "Narrator", "Editor", "Author", "Writer",
                "Band", "Group", "Label", "Album", "Single", "Song", "Track", "Everyone", "Everybody", "Nobody"}


def headed_sections(paras, title="", single_names=False):
    """Headings followed by first person outside quotation marks.
    Returns (role headings [(index, name)], bare-name headings [(index, name)], topic headings [index])."""
    roles, bare, topics = [], [], []
    idx = [i for i, p in enumerate(paras) if p.strip()]
    for k, a in enumerate(idx[:-1]):
        h = paras[a].strip()
        bio = BIO_INTRO.match(h)
        if bio and bio.group("name").split()[0] in FIRST_NAMES and \
                fp_outside_quotes(paras[idx[k + 1]], 400, True) and not _mentions(bio.group("name"), paras[idx[k + 1]]):
            roles.append((a, bio.group("name")))
            continue
        if len(h) > 240 or h.endswith("?"):
            continue
        nxt = []
        for b in idx[k + 1:k + 3]:
            q = paras[b].strip()
            if ROLE_HEADING.match(q) or (len(q) <= 120 and (BARE_NAME.match(q) or TOPIC_HEADING.match(q))):
                break
            nxt.append(b)
        if not nxt:
            continue
        m = ROLE_HEADING.match(h)
        nm = m.group("name").split() if m else []
        if m and h[0].isupper() and " – " not in m.group("name") and (nm[0] in FIRST_NAMES or (
                single_names and len(nm) == 1 and nm[0] not in COMMON_NOUNS and "," in h)):
            # the account is the first paragraph in the first person; a short third-person introduction under the
            # heading ("A founding member of Toto, Lukather is ...") may come before it
            acc = next((b for b in nxt if fp_outside_quotes(paras[b], 400, True)), None)
            if acc is not None and not _mentions(m.group("name"), paras[acc], title):
                roles.append((a, m.group("name")))
            continue
        if len(h) > 120 or h.endswith("."):
            continue
        m = BARE_NAME.match(h)
        if m and m.group("name").split()[0].title() in FIRST_NAMES_STRICT and fp_outside_quotes(paras[nxt[0]], 150, True) \
                and not _mentions(m.group("name"), paras[nxt[0]], title):
            bare.append((a, m.group("name")))
            continue
        if (TOPIC_HEADING.match(h) or SHORT_LINE.match(h)) and fp_outside_quotes(paras[nxt[0]], 300):
            topics.append(a)
    return roles, bare, topics


FURNITURE = re.compile(r"^\s*(?:[-•*]\s*)?(?:sign up|subscribe|related|read more|more from|you may also like|you might also like|share this|"
                       r"more details on copyright|the current copyright owner|(?:feature|interview|review|article|words|text) by\b|"
                       r"you'?ve reached your limit|reached your limit of free|keep reading for|subscribe for [£$€]|"
                       r"browse by topic|browse category|more from these topics|more with these topics|related articles|"
                       r"also featuring gear in this article|featuring related gear|if you value this resource|"
                       r"(?:previous|next) article in this issue|unlimited access|subscriber-only|already a subscriber|"
                       r"award-winning (?:email )?newsletters?|"
                       r"new issues that have been donated|all donations and support|do you have any of these magazine|"
                       r"if so, and you can donate|mu:zines|this article also has google|\{\{cite|<ref name|publisher:|"
                       r"follow @|follow \w+ on (?:twitter|instagram|facebook)|previous article|next article|"
                       r"about the author|join our commenting|log in to comment|you must be logged in|leave a reply|"
                       r"your email address will not|tags:|posted in|categories:|share on|tweet this|"
                       r"advertisement|©|copyright|all rights reserved|tags?:|filed under|comments?\b|leave a comment|"
                       r"follow us|click here|this article|this interview was|this story|originally published|"
                       r"a version of this|photo:|photograph:|getty|image:|top image|main image|watch:|listen:)", re.I)

# page furniture that publishers drop into the middle of a piece ("Sign up below to get the latest from Classic
# Rock...", "You might also like", "RELATED:"); a stricter list than FURNITURE, so an answer is never cut by it
INBLOCK_FURNITURE = re.compile(r"^\s*(?:sign up (?:below|to|for|now|here)\b|subscribe (?:to|for|now|today|here)\b|"
                               r"(?:you (?:may|might) also like|also read|see also|read more|"
                               r"related(?: articles?| stories| posts| content| reading)?|more (?:from|on) [^\n]{0,60})\s*(?::|»|›|$)|"
                               r"advertisement\s*$|share this (?:article|story|post)\b|follow us on\b|click here (?:to|for)\b)", re.I)


def inblock_furniture(p):
    return len(p.strip()) <= 200 and bool(INBLOCK_FURNITURE.match(p))


# the interviewer signing off after the last answer ("I am honored that Franne was able to...", "Thanks to Jody for...")
CLOSING = re.compile(r"^\s*(?:[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+)?,\s*)?(?:i(?: am|'m) (?:very |so |truly |really )?(?:honou?red|grateful|thankful|thrilled|delighted|pleased)|"
                     r"i want to (?:take this (?:occasion|opportunity) to )?(?:again )?thank|i'?d like to thank|thanks? (?:so much |again )?(?:to|for)|"
                     r"thanks?(?: you)?\b[^.\n]{0,40}\bfor (?:taking|sharing|talking|speaking|chatting|your time|the time|"
                     r"joining|being|going)|pick up a copy|(?:buy|order|grab|get) (?:it|a copy|the (?:book|album|record))|"
                     r"thank you (?:so much |again )?(?:to|for)|special thanks|many thanks|a big thank|"
                     r"(?:the )?(?:album|single|ep|book|record|box set|tour)\b[^.]{0,80}\b(?:is|are) (?:out|available)|"
                     r"(?:tickets|tour dates|details)\b)", re.I)


# NPR pages that print the interviewee's answers under topic headings
HIGHLIGHTS = re.compile(r"^\s*(?:interview highlights|highlights from (?:the|this|our) (?:interview|conversation))\s*:?\s*$", re.I)
# a writer's biography at the foot of a piece ("Joe Bloggs is a writer based in...", "About the author")
BYLINE_BIO = re.compile(r"^\s*(?:about the author|[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){1,2} (?:is|was) (?:a|an|the) "
                        r"(?:[\w-]+ ){0,3}(?:writer|journalist|editor|critic|author|contributor|freelance|music writer))")


def lyrics_like(paras):
    """Song lyrics and verse: many short lines and hardly any paragraph of prose."""
    lens = [len(p.strip()) for p in paras if p.strip()]
    if len(lens) < 12:
        return False
    return sum(1 for x in lens if x < 60) / len(lens) >= 0.7 and sum(1 for x in lens if x >= 200) / len(lens) < 0.1


# lines of page clutter at the end of a piece: dates, credits, plugs, links, listings boxes, sign-off names
CLUTTER = re.compile(
    r"^\s*(?:(?:where|when|category|categories|tags?|photo(?:graph)?s? credits?|image credits?|credits?|website|official website|"
    r"facebook|twitter|instagram|youtube|tiktok|spotify|bandcamp|soundcloud|visit|tickets?|more [\w ]{0,20}interviews?|"
    r"get more at|follow|find (?:him|her|them|out more)|for more|listen|watch|stream|buy|pre-?order|out now|"
    r"transcript provided by|editor'?s note)\b"
    r"|(?:[A-Z][a-z]+ \d{1,2}(?:st|nd|rd|th)?, \d{4}|\d{1,2}(?:st|nd|rd|th)? [A-Z][a-z]+,? \d{4}|[A-Z][a-z]+ \d{4})\b[^\n]{0,80}$"
    r"|photos?\b[^.\n]{0,30}\bby\b|[-–—]\s*[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){0,2}\s*$|\|"
    r"|[_\-=*~·•]{3,}\s*$|(?:thank you|thanks)\b[!. ]*$)", re.I)
# the same, anywhere in a short closing line: plugs, credits, pointers elsewhere
CLUTTER_ANY = re.compile(r"\b(?:for tickets|tickets (?:are )?(?:on sale|available)|on sale now|visit:|(?:is|are) out (?:now|on|\d)|"
                         r"(?:is|are) released on|out \d{1,2}/\d{1,2}|produced and edited|watch the video|pre-?order|"
                         r"available (?:now|to stream|to buy|from|on (?:cd|vinyl|digital))|learn more|find out more|"
                         r"if you'?d like to know more|for more information|(?:feature|interview|article) by [A-Z]|"
                         r"you can order|order it here|here are our interviews|our interviews? with|"
                         r"for more (?:exclusive )?news|gives its (?:warmest |sincere |grateful )?thanks|"
                         r"tours? (?:begins|starts|kicks off|runs|continues)|this interview (?:originally )?(?:appeared|was)|"
                         r"interview was conducted|(?:has|have) been (?:edited|condensed)|edited (?:and condensed )?for (?:length|clarity)|"
                         r"^note:|contact page|got a question for|send (?:us )?your questions|get in touch)", re.I)
SOCIAL = re.compile(r"\b(?:website|facebook|instagram|twitter|youtube|tiktok|spotify|bandcamp)\b", re.I)
NPR_TAIL = re.compile(r"\s*(?:Transcript provided by NPR, Copyright NPR\.|Copyright © \d{4} NPR\.)[^\n]*\s*$")


def clutter(p):
    """A short line of page clutter (as opposed to the last words of an answer)."""
    s = p.strip()
    if len(s) > 220:
        return False
    if CLUTTER.match(s) or FURNITURE.match(s) or BYLINE_BIO.match(s) or CLUTTER_ANY.search(s) or CLOSING.match(s):
        return True
    if s.endswith(":") or (len(s) <= 60 and re.fullmatch(r"[\"'][^\"'\n]+[\"']", s)):
        return True                       # a quoted title on its own, or a line leading into a list of links
    if len(SOCIAL.findall(s)) >= 2 or s.count(" | ") >= 2 or re.search(r"https?://|www\.", s):
        return True
    if re.match(r"^[A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){1,2} is (?:a|an|the) ", s) and not FIRST_PERSON.search(s):
        return True                       # "Tig Notaro is a comedian whose documentary ..."
    return len(s) <= 40 and bool(BARE_NAME.match(s))     # the interviewer's name signing off


# ── routing ──────────────────────────────────────────────────────────────────
def para_spans(text):
    """(start, end) of every line, end exclusive, without the newline."""
    spans, pos = [], 0
    for p in text.split("\n"):
        spans.append((pos, pos + len(p)))
        pos += len(p) + 1
    return spans


def _labels_text(lines):
    return ", ".join(f"{l} ({k})" for l, k in Counter(x[1] for x in lines).most_common(6))


def route(text, title="", page_type="", lang="en"):
    """Return dict(route, segments=[(start, end, kind)], evidence). lang is the record's Language code."""
    _LANG[0] = (lang or "en").strip().lower()[:2]
    paras = text.split("\n")
    _HANDLES[0] = sum(1 for p in paras if len(p) >= 30 and HANDLE_END.search(p.strip())) >= 3
    _QA_TITLE[0] = bool(QA_TITLE.search(title or ""))
    _TITLE_LINE[0] = next((i for i, p in enumerate(paras) if p.strip()), -1) if nonempty_title_line(paras, title) else -1
    spans = para_spans(text)
    nonempty = [i for i, p in enumerate(paras) if p.strip()]
    ev = {}
    if page_type == "forum thread":
        ev["detail"] = "posts by several users (not processed)"
        return dict(route="forum thread", segments=[(0, len(text), "forum thread")], evidence=ev)
    conv, _, _ = convention(text)
    ev["quote convention"] = conv
    kind, block, cstart = "prose", None, None     # block = (first para, last para); cstart = char offset inside first
    blocks = []                                  # several blocks (unlabelled Q&A units): [(first para, last para, None)]

    # 1. speaker labels at the start of paragraphs
    lab_ok, lab_kind, lab_lines, meas = labelled_verdict(paras, text, title=title or "")
    styles = Counter(x[3] for x in lab_lines)
    # 2. labels inside paragraphs (scraped pages that lost their line breaks)
    if not lab_ok:
        inl = inline_labels(paras)
        if len(inl) >= 4:
            merged = sorted(list(label_lines(paras, title or "")) + [x[:4] for x in inl], key=lambda x: (x[0], 0))
            ok2, kind2, lines2, meas2 = labelled_verdict(paras, text, merged, title or "")
            n_inline = sum(1 for x in lines2 if x[3] == "inline")
            if ok2 and n_inline >= 0.3 * len(lines2):
                lab_ok, lab_kind, lab_lines, meas = ok2, kind2 + " (labels inside paragraphs)", lines2, meas2
                styles = Counter(x[3] for x in lab_lines)
                first_inline = min((x for x in inl if x[0] == min(y[0] for y in lab_lines)), key=lambda x: x[4],
                                   default=None)
                if first_inline and not any(x[0] == first_inline[0] and x[3] != "inline" for x in lab_lines):
                    cstart = spans[first_inline[0]][0] + first_inline[4]
    # 2b. quotations labelled with the speaker's name, one or more per person ('John Illsley [bass]: "We spent..."')
    if not lab_ok:
        raw = [x for x in label_lines(paras, title or "") if x[2][:1] in "\"'" and x[3] == "colon"]
        if len(raw) >= 4 and len({x[1] for x in raw}) >= 2:
            wo = (title or "") + "\n" + "\n".join(re.sub(r"^\W*[A-Z][\w.'’ &-]{0,40}(?:\[[^\]]*\]|\([^)]*\))?[:–—-]", " ", p) for p in paras)
            people = [x for x in raw if person_like(x[1], wo)]
            if len(people) >= 4 and len(people) >= 0.75 * len(raw) and 2 * len(people) >= 0.2 * len(nonempty):
                lab_ok, lab_kind, lab_lines, meas = True, "quotes labelled with the speaker's name", people, None
                styles = Counter(x[3] for x in people)
    if lab_ok:
        kind = "labelled Q&A"
        # a heading with a colon used once ("Music Loses Two Giants: Vandross and Benson") does not start the block
        _cnt = Counter(x[1] for x in lab_lines)
        _wo = (title or "") + "\n" + text
        _starts = [x[0] for x in lab_lines if x[0] != _TITLE_LINE[0] and
                   (_cnt[x[1]] >= 2 or x[3] in ("npr", "role") or person_like(x[1], _wo))]
        first = min(_starts or [x[0] for x in lab_lines])
        if lab_kind == "answers labelled":       # the question just before the first labelled answer
            j = first - 1
            while j >= 0 and not paras[j].strip():
                j -= 1
            if j >= 0 and is_question(paras[j]):
                first = j
        last = max(x[0] for x in lab_lines)
        if lab_kind == "questions labelled":     # the answer after the last question
            j = last + 1
            while j < len(paras) and not paras[j].strip():
                j += 1
            last = min(j, len(paras) - 1)
        if not any(x[3] == "inline" for x in lab_lines):
            known = {x[1] for x in lab_lines if x[0] in set(_starts)} or {x[1] for x in lab_lines}
            wo_t = (title or "") + "\n" + text
            surnames = {k.split()[-1].upper() for k in known}
            all_lines = merge_labels(label_lines(paras, title or ""))
            any_label = {x[0] for x in all_lines
                         if x[1] in known or person_like(x[1], wo_t) or NPR_LABEL.match(paras[x[0]])}
            same = {x[0] for x in all_lines if x[1].split()[-1].upper() in surnames}
            j, gap, seen = first - 1, 0, 0
            while j >= 0:
                if paras[j].strip():
                    seen += 1
                    if j in any_label and j != _TITLE_LINE[0] and (gap <= 2 or j in same):
                        first, gap, seen = j, 0, 0
                    else:
                        gap += 1
                        # past two unlabelled paragraphs, keep looking only for the same speaker's earlier label,
                        # over paragraphs that read as their answer (not the journalist's narration)
                        if gap > 2 and (narration(paras[j]) or seen > 15 or not same):
                            break
                j -= 1
        j = first - 1
        while j >= 0 and not paras[j].strip():
            j -= 1
        if j >= 0 and is_question(paras[j]):
            first = j
        block = (first, last)
        ev["detail"] = lab_kind
        ev["label lines"] = len(lab_lines)
        ev["labels"] = _labels_text(lab_lines)
        ev["label style"] = ", ".join(f"{s} ({k})" for s, k in styles.most_common())
        if meas:
            ev["speech share"] = round(meas["speech"], 2)
            ev["labels naming people"] = round(meas["person"], 2)
    # 3. lecture transcripts (names before each turn, no colon)
    if kind == "prose":
        lec = lecture_transcript(text, title)
        if lec:
            host, guest, offs = lec
            kind = "labelled Q&A"
            ev["detail"] = "lecture transcript (speaker names before each turn, no colons)"
            ev["labels"] = f"{guest}, {host}"
            first = next(i for i, (s, e) in enumerate(spans) if s <= offs[0] <= e)
            last = next(i for i, (s, e) in enumerate(spans) if s <= offs[-1] <= e)
            block, cstart = (first, last), offs[0]
    # 4. caption transcripts (a dash starts each new speaker)
    if kind == "prose" and sum(1 for p in paras if DASH_TURN.match(p)) >= 10 and \
            sum(1 for p in paras if BRACKET_LABEL.match(p)) >= 3:
        kind = "labelled Q&A"
        idx = [i for i, p in enumerate(paras) if DASH_TURN.match(p) or BRACKET_LABEL.match(p)]
        block = (idx[0], idx[-1])
        ev["detail"] = "caption transcript (a dash starts each new speaker)"
    # 5. questions and answers without labels
    if kind == "prose":
        pairs = qa_pairs(paras)
        runs = run_in_pairs(paras)
        covered = 2 * (len(pairs) + len(runs)) / max(1, len(nonempty))   # a run-in paragraph holds both
        fpa = first_person_answers(paras, pairs) if pairs else 1.0
        allu = qa_units(paras)
        fpu = [u for u in allu if u[2]]                      # units whose answer opens in the first person
        you_share = sum(1 for u in allu if u[3]) / len(allu) if allu else 0.0           # 'you', 'he' or 'she'
        you2_share = sum(1 for u in allu if u[3] == 2) / len(allu) if allu else 0.0      # 'you' only
        unit_cov = sum(1 for i in nonempty if any(a <= i <= b for a, b, _, _ in allu)) / max(1, len(nonempty))
        by_pairs = (len(pairs) >= 4 or len(runs) >= 4 or (len(pairs) >= 2 and len(runs) >= 2)) and \
            covered >= 0.25 and fpa >= 0.5
        by_units = (you2_share >= 0.34 and (len(fpu) >= 8 or (len(fpu) >= 4 and unit_cov >= 0.4))) or \
            (you2_share >= 0.67 and len(fpu) >= 3 and unit_cov >= 0.4) or \
            (you_share >= 0.6 and len(fpu) >= 5 and unit_cov >= 0.4)
        # the title says it is a Q&A ("Q&A: Martin Rushent", "FAN QUESTIONS!"): fewer exchanges are enough
        by_title = _QA_TITLE[0] and len(fpu) >= 1 and len(fpu) >= 0.5 * len(allu) and you_share >= 0.5 and unit_cov >= 0.3
        # pages that repeat their own text (a list printed twice) count each question once
        distinct = len({paras[a].strip() for a, _ in pairs}) + len({paras[i].strip()[:200] for i, _ in runs})
        by_many = distinct >= 8 and fpa >= 0.6 and len(fpu) >= 0.5 * len(allu) and you2_share >= 0.34
        by_lenient = False
        if not (by_pairs or by_units or by_many or by_title):
            # a short interview whose answers mention what others said ("If there were cell phones back then...
            # they weren't that friendly"): every answer in the first person, short questions, mostly put to 'you'
            lu = qa_units(paras, lenient=True)
            if len(lu) >= 4 and all(u[2] for u in lu):
                lq = sorted(len(paras[u[0]].strip()) for u in lu)
                l_you2 = sum(1 for u in lu if u[3] == 2) / len(lu)
                l_cov = sum(1 for i in nonempty if any(a <= i <= b for a, b, _, _ in lu)) / max(1, len(nonempty))
                by_lenient = l_you2 >= 0.5 and l_cov >= 0.4 and lq[len(lq) // 2] <= 120
                if by_lenient:
                    allu = lu
                    fpu = list(lu)
                    you_share = sum(1 for u in lu if u[3]) / len(lu)
                    unit_cov = l_cov
        if (by_pairs or by_units or by_many or by_title or by_lenient) and allu:
            kind = "unlabelled Q&A"
            run_set = {i for i, _ in runs}
            n_runs = sum(1 for a, _, _, _ in allu if a in run_set)
            ev["detail"] = ("questions on their own lines" if n_runs * 2 < len(allu) else "questions run into the answers")
            ev["question-answer units"] = len(allu)
            ev["answers opening in the first person"] = len(fpu)
            ev["run-in questions"] = n_runs
            ev["questions put to someone"] = round(you_share, 2)
            ev["share of paragraphs in units"] = round(unit_cov, 2)
            if by_title and not (by_pairs or by_units or by_many):
                ev["title cue"] = "Q&A"
            # for the blocks, short answers count too ("Fear, anxiety and panic attacks."); a question that is not
            # put to anyone and gets no answer in the first person is a standfirst, not an exchange
            def _standfirst(u):
                ans = " ".join(paras[x] for x in range(u[0] + 1, u[1] + 1))
                body = ans if QUOTE_START.match(ans.strip()) else _unquote_para(ans)   # answers printed in quotes
                return not u[3] and len(ans) >= 400 and not FIRST_PERSON.search(body)
            blk_units = [u for u in qa_units(paras, min_answer=8, prompts=True) if not _standfirst(u)]
            # consecutive units form one block; anything between blocks is prose
            for a, b, _, _ in (blk_units or allu):
                if blocks and not any(paras[x].strip() for x in range(blocks[-1][1] + 1, a)):
                    blocks[-1] = (blocks[-1][0], b, None)
                else:
                    blocks.append((a, b, None))
    # 6. first-person accounts
    if kind == "prose":
        fp = first_person_profile(text, conv)
        ev.update({k: round(v, 1) for k, v in fp.items()})
        bare_title = re.sub(r"\"[^\"]*\"|(?:(?<=^)|(?<=[\s(]))'[^']*'(?=[\s:,.;!?)]|$)", " ", title or "")
        title_cue = bool(TITLE_FP.search(bare_title))
        picks = bool(TITLE_PICKS.search(bare_title))
        sections_cue = bool(TITLE_SECTIONS.search(bare_title))
        roles, bare, topics = headed_sections(paras, title or "", single_names=sections_cue)
        ev["title cue"] = "one speaker" if title_cue else "one speaker chooses" if picks else \
            "several speakers" if sections_cue else ""
        words, fps, fpa_, third, quote = fp["words"], fp["fp_per_k"], fp["fp_all_per_k"], fp["third_per_k"], fp["quote_per_k"]
        detail = None
        bare_names = set(nm for _, nm in bare)
        named = roles + (bare if sections_cue or (len(bare_names) >= 3 and not (title_cue or picks)) else [])
        distinct = list(dict.fromkeys(nm for _, nm in sorted(named)))
        if len(distinct) >= 2:
            # 'How we made' features, tributes, oral histories by section: each heading names the speaker
            a = min(i for i, _ in named)
            sp = first_person_profile("\n".join(paras[a:]), conv)
            if sp["words"] >= 150 and sp["fp_all_per_k"] >= 15 and sp["third_per_k"] <= 4:
                detail, block = "sections headed by the speaker's name", (a, nonempty[-1])
                ev["headings"] = "; ".join(distinct[:8])
        heads = sorted(topics + [i for i, _ in bare])
        if not detail and len(heads) >= 3 and (title_cue or picks):
            # 'Andy Bell on his favourite diva songs': one speaker, a heading for each song or artist
            a = heads[0]
            sp = first_person_profile("\n".join(paras[a:]), conv)
            if sp["words"] >= 150 and sp["fp_per_k"] >= (20 if title_cue else 30) and sp["third_per_k"] <= (3 if title_cue else 2):
                detail, block = "one speaker, sections headed by topic", (a, nonempty[-1])
        marker = next((i for i in nonempty if HIGHLIGHTS.match(paras[i])), None)
        if not detail and marker is not None:
            sp = first_person_profile("\n".join(paras[marker + 1:]), conv)
            if sp["words"] >= 100 and sp["fp_per_k"] >= 20:
                detail, block = "interview highlights (the interviewee's answers, questions left out)", (marker + 1, nonempty[-1])
        if not detail and words >= 150 and third <= 1.5 and quote <= 8 and (fps >= 35 or fpa_ >= 45) and \
                not lyrics_like(paras):
            detail = "one speaker throughout"
        if not detail and words >= 150 and TITLE_ON.match(bare_title.strip()) and fps >= 35 and third <= 1 and quote <= 15 \
                and not lyrics_like(paras):
            detail = "one speaker throughout (title names the person)"
        if not detail and title_cue:
            allw = max(1, len(WORD.findall(text)))
            fp_in = 1000 * len(FIRST_PERSON.findall(text)) / allw
            fps_in = 1000 * len(FP_SINGULAR.findall(text)) / allw
            if allw >= 150 and fp_in >= 25 and fps_in >= 15 and third <= 3:
                detail = "one speaker throughout (title says so; quoted answers count)"
                qheads = [a for k, a in enumerate(nonempty[:-1]) if len(paras[a].strip()) <= 120 and
                          (TOPIC_HEADING.match(paras[a]) or BARE_NAME.match(paras[a]) or
                           re.match(r"^\s*[A-Z][^.!?\n]{1,60}\([A-Z][\w .'’-]{1,30}\)\s*$", paras[a])) and
                          FIRST_PERSON.search(paras[nonempty[k + 1]][:300])]
                not_title = [i for i in nonempty if paras[i].strip()[:40].lower() != (title or "").strip()[:40].lower()]
                if len(qheads) >= 3:
                    block = (qheads[0], nonempty[-1])
                else:
                    fpidx = [i for i in not_title if FIRST_PERSON.search(paras[i][:150])]
                    block = (fpidx[0], nonempty[-1]) if fpidx else (nonempty[0], nonempty[-1])
        if detail and block is None:
            qheads = [a for k, a in enumerate(nonempty[:-1]) if len(paras[a].strip()) <= 120 and
                      (TOPIC_HEADING.match(paras[a]) or
                       re.match(r"^\s*[A-Z][^.!?\n]{1,60}\([A-Z][\w .'’-]{1,30}\)\s*$", paras[a])) and
                      FIRST_PERSON.search(paras[nonempty[k + 1]][:300])]
            if len(qheads) >= 3:
                block = (qheads[0], nonempty[-1])
        if detail:
            kind = "first-person"
            ev["detail"] = detail
            if block is None:
                # skip a third-person introduction (a standfirst) before the first first-person paragraph
                fpidx = [i for i in nonempty if fp_outside_quotes(paras[i])]
                block = (fpidx[0], nonempty[-1]) if fpidx else (nonempty[0], nonempty[-1])
    segments = []
    if block is not None and kind == "first-person" and block[0] > 0:
        # the heading of the first section belongs to the account ("When I discovered dub" above the first song)
        prev = max((i for i in nonempty if i < block[0]), default=None)
        if prev is not None and prev == block[0] - 1:
            hp = paras[prev].strip()
            if len(hp) <= 80 and SHORT_LINE.match(hp) and not HIGHLIGHTS.match(hp) and not FURNITURE.match(hp) and \
                    not CLUTTER_ANY.search(hp) and not re.search(r"https?://|www\.|\|", hp) and \
                    hp.lower() != (title or "").strip().lower() and not hp.endswith(":") and \
                    any(k > block[0] and len(paras[k].strip()) <= 80 and SHORT_LINE.match(paras[k]) for k in nonempty):
                block = (prev, block[1])
    if block is not None:
        a, b = block
        if kind == "first-person":
            # the account ends before the first piece of page furniture after it starts
            def _ends(j):
                if BYLINE_BIO.match(paras[j]):
                    return True
                if not FURNITURE.match(paras[j]):
                    return False
                if inblock_furniture(paras[j]):
                    nxt = [k for k in nonempty if k > j][:2]      # a section heading may come first
                    return not any(fp_outside_quotes(paras[k]) for k in nxt)
                return True
            stop = next((j for j in nonempty if a < j <= b and _ends(j)), None)
            last = max(i for i in nonempty if i < stop) if stop is not None else b
            last = max(last, a)
        else:
            # extend the block over continuation paragraphs after the last label or answer (a long last answer),
            # up to page furniture or the interviewer's sign-off, and for at most eight paragraphs
            last, extra = b, 0
            for j in range(b + 1, len(paras)):
                if not paras[j].strip():
                    continue
                if FURNITURE.match(paras[j]) or CLOSING.match(paras[j]) or narration(paras[j]) or extra >= 8:
                    break
                last, extra = j, extra + 1
        blocks = [(a, last, cstart)]
        if kind == "labelled Q&A" and cstart is None and lab_lines and not lab_kind.startswith(("lecture", "caption")):
            # a labelled exchange set inside an article: where the journalist quotes one of the labelled speakers
            # in the third person ('"I was in at half past 10," says Rocco') and writes no first person of their
            # own, the paragraph is the journalist's narration and stays prose. Pronouns ("he said") are not
            # enough: speakers tell stories with reported speech in them.
            labelled = {x[0] for x in lab_lines}
            names = {w for x in lab_lines for w in re.findall(r"[A-Z][\w'’-]{2,}", x[1].title())}
            idx = [j for j in range(a, last + 1) if paras[j].strip()]
            brk = set()
            for j in idx:
                if j in labelled:
                    continue
                for m in THIRD_ATTR_NAMED.finditer(paras[j]):
                    who = m.group("who") or m.group("who2")
                    if who.split()[-1].title() in names and not FIRST_PERSON.search(_unquote_para(paras[j])):
                        brk.add(j)
                        break
            if brk:
                blocks, cur = [], None
                for j in idx:
                    if j in brk:
                        if cur:
                            blocks.append((cur[0], cur[1], None))
                        cur = None
                    else:
                        cur = (cur[0], j) if cur else (j, j)
                if cur:
                    blocks.append((cur[0], cur[1], None))
                # a stretch with no label line in it is narration too
                blocks = [blk for blk in blocks if any(blk[0] <= x <= blk[1] for x in labelled)] or [(a, last, None)]
    if blocks:
        a, b, cs = blocks[-1]
        while b > a and (not paras[b].strip() or clutter(paras[b])):
            b -= 1
        blocks[-1] = (a, b, cs)
        if len(blocks) > 1 and blocks[-1][0] == blocks[-1][1] and is_question(paras[blocks[-1][0]]):
            blocks = blocks[:-1]          # "Got a Question for The Shooz?" with nothing after it
        split = []
        for a, b, cs in blocks:
            cur = None
            for j in range(a, b + 1):
                if paras[j].strip() and j > a and inblock_furniture(paras[j]):
                    if cur:
                        split.append(cur)
                    cur = None
                    continue
                if cur is None:
                    if not paras[j].strip():
                        continue
                    cur = (j, j, cs if j == a else None)
                else:
                    cur = (cur[0], j, cur[2])
            if cur:
                split.append(cur)
        blocks = split
        if kind == "unlabelled Q&A":
            # "...brought Tears for Fears their first chart success. Did its achievements come as a surprise?": the
            # block starts at the question, and the narration before it in the paragraph stays prose
            fixed = []
            for a, b, cs in blocks:
                pa = paras[a]
                if cs is None and len(pa) > 200 and is_question(pa):
                    m = re.search(r"(?:(?<=[.!…\"'])|(?<=[.!…\"'][)\]]))\s+(?=[A-Z][^.!?\n]{3,300}\?\s*$)", pa)
                    pre = pa[:m.start()] if m else ""
                    # only when what comes before the question is the journalist's narration, not the interviewer
                    # talking to the interviewee ("When Hate For Sale came out, the first thing that struck me...")
                    if m and m.end() >= 250 and not SECOND_PERSON.search(pre) and not FIRST_PERSON.search(_unquote_para(pre)):
                        cs = spans[a][0] + m.end()
                fixed.append((a, b, cs))
            blocks = fixed
    pos = 0
    for a, b, cs in blocks:
        start = spans[a][0] if cs is None else cs
        end = spans[b][1]
        tail = NPR_TAIL.search(text[start:end])
        if tail and end - tail.start() - start < end - start:
            end = start + tail.start()
        if start > pos:
            segments.append((pos, start, "prose"))
        segments.append((start, end, kind))
        pos = end
    if pos < len(text):
        segments.append((pos, len(text), "prose"))
    ev["block share"] = round(sum(e - s for s, e, k in segments if k != "prose") / max(1, len(text)), 2)
    return dict(route=kind, segments=segments, evidence=ev)
'''
SOURCES["stage3"] = r'''"""Stage 3 of the turn build: who speaks in the Q&A and first-person blocks (technical appendix, Step 2, stage 3).

For every record that stage 2 routed as a labelled Q&A, an unlabelled Q&A or a first-person account, the blocks are
cut into spans that cover them exactly once:
    turn         words spoken by one speaker (a labelled line, a question, an answer, an account)
    sung         a labelled line that is sung, not spoken ("BLEACHERS: (Singing) ...")
    heading      a section heading inside the block ("The first song I remember hearing")
    label        a speaker label standing on its own line before the words
    narration    the journalist's own words inside a block (a third-person line under a heading)
    non-speech   stage directions and page furniture ("(SOUNDBITE OF MUSIC)", "Sign up below...")
Each turn gets a speaker: a member of the focal act (nomination-era or other-era), several members without saying
which, the interviewer, another person, or unknown; with a confidence and the rule that decided it.
The prose segments of stage 2 are left for stage 4.
"""
import functools
import re
from collections import Counter, defaultdict

import stage2 as S

# ── speaker types ────────────────────────────────────────────────────────────
MEMBER = "band member"                         # nomination-era member: band voice
MEMBERS_UNSPEC = "band member, unspecified"     # several nomination-era members, not which one: band voice
OTHER_ERA = "other-era member"
OTHER_ERA_UNSPEC = "other-era member, unspecified"
INTERVIEWER = "interviewer"
OTHER = "other person"
UNKNOWN = "unknown"
AMBIGUOUS = "ambiguous"
BAND_VOICE = {MEMBER, MEMBERS_UNSPEC}
LEVELS = ["unknown", "low", "medium", "high"]


def lower_level(c):
    return LEVELS[max(1, LEVELS.index(c) - 1)] if c in LEVELS[1:] else c


def fold(s):
    return re.sub(r"\s+", " ", s.replace("’", "'")).strip().casefold()


# ── the name lists ───────────────────────────────────────────────────────────
def _accent_classes():
    """For each plain letter, the letter and its accented forms ("o" also matches "ō", "ö", "ó")."""
    import unicodedata
    groups = defaultdict(set)
    for code in list(range(0xC0, 0x250)):
        ch = chr(code)
        base = "".join(c for c in unicodedata.normalize("NFKD", ch) if not unicodedata.combining(c))
        if len(base) == 1 and base.isalpha() and base.isascii():
            groups[base.lower()].add(ch.lower())
            groups[base.lower()].add(ch.upper())
    return {b: "[" + re.escape(b) + re.escape(b.upper()) + "".join(sorted(v)) + "]" for b, v in groups.items()}


ACCENTS = _accent_classes()


def loose(word):
    """A regex for a word that tolerates accents either way ("Lōc" for "Loc", "Dorper" for "Dörper")."""
    import unicodedata
    out = []
    for ch in word:
        base = "".join(c for c in unicodedata.normalize("NFKD", ch) if not unicodedata.combining(c)).lower()
        out.append(ACCENTS.get(base, re.escape(ch)) if len(base) == 1 else re.escape(ch))
    return "".join(out)


def _forms(name):
    """A name as written in the lists, plus its plain variants: "(Helen) Folasade Adu, Helen Adu" gives
    "Helen Folasade Adu", "Folasade Adu" and "Helen Adu"; "Chris 'Merrick' Hughes" gives "Chris Hughes" too."""
    out = set()
    for part in [x.strip() for x in name.split(",")] if name.count(",") and "(" in name else [name]:
        if not part:
            continue
        variants = {part}
        if "(" in part:
            variants = {re.sub(r"\s+", " ", part.replace("(", "").replace(")", "")).strip(),
                        re.sub(r"\s+", " ", re.sub(r"\([^)]*\)", " ", part)).strip()}
        for v in list(variants):
            nick = re.sub(r"\s+", " ", re.sub(r"\s['\"‘“][^'\"’”]{1,20}['\"’”]\s", " ", " " + v + " ")).strip()
            variants.add(nick)
        out |= {v for v in variants if len(v) >= 2}
    out |= {v[:-1] for v in out if v.endswith(".") and len(v) > 4}       # "Sheila E." is also written "Sheila E"
    return out


def _tokens(name):
    return [t for t in re.sub(r"['\"‘“][^'\"’”]{1,20}['\"’”]", " ", name).replace(",", " ").split() if t]


SUFFIX = {"sr.", "sr", "jr.", "jr", "ii", "iii", "iv"}


class ActNames:
    """Every way the act's members can be named, and the decisions from the name-list review."""

    def __init__(self, act, members, nonmembers, labels, no_bare, duplicates):
        self.act = act
        self.members = {}
        self.form = defaultdict(set)          # folded full form -> members
        self.first = defaultdict(set)         # folded first name -> members
        self.surname = defaultdict(set)       # folded surname -> members
        self.initials = defaultdict(set)      # "MW" -> members
        for m in members:
            name = m["member"]
            forms = {name} | {f for a in m["aliases"] for f in _forms(a)}
            for d in duplicates:
                if name in d["one_person"]:
                    forms |= set(d["one_person"])
            forms = {f for x in forms for f in _forms(x)}
            toks = [t for t in _tokens(name) if t.casefold() not in SUFFIX]
            self.members[name] = dict(name=name, era=m["era"], forms=forms, pronoun=m.get("pronoun"))
            for f in forms:
                self.form[fold(f)].add(name)
            if len(toks) >= 2:
                if len(toks[0].strip(".")) >= 2:
                    self.first[fold(toks[0])].add(name)
                if len(toks[-1].strip(".")) >= 3:
                    self.surname[fold(toks[-1])].add(name)
                caps = [t[0] for t in toks if t[0].isupper()]
                if 2 <= len(caps) <= 4:
                    self.initials["".join(caps)].add(name)
                    if len(caps) == 3:                 # "Paul S. Denman" is also "PD"
                        self.initials[caps[0] + caps[2]].add(name)
            # aliases' surnames ("Pete Harris" for Peter Harris) and first names ("Andy Hale")
            for f in forms:
                ft = [t for t in _tokens(f) if t.casefold() not in SUFFIX]
                if len(ft) >= 2 and len(ft[-1].strip(".")) >= 3:
                    self.surname[fold(ft[-1])].add(name)
                if len(ft) >= 2 and len(ft[0].strip(".")) >= 2 and ft[0][0].isupper():
                    self.first[fold(ft[0])].add(name)
        # the act's name is also its singer's name ("SADE:" is Sade Adu)
        for name in self.members:
            if fold(name.split()[0]) == fold(act) and fold(act) not in self.form:
                self.form[fold(act)].add(name)
        self.no_bare = {fold(x["surname"]) for x in no_bare if x["act"] == act}
        self.nonmembers = [x["name"] for x in nonmembers if x["act"] == act]
        self.labels = {}                        # folded label -> list of candidate members ([] = not a member)
        for x in labels:
            if x["act"] == act:
                self.labels[fold(x["label"])] = list(x.get("candidates") or ([x["member"]] if x["member"] else []))
        self.rx = {}
        for name, m in self.members.items():
            alts = sorted(m["forms"], key=len, reverse=True)
            pat = "|".join(r"[\s\-‐]+".join(loose(w) for w in f.replace("-", " ").split()) for f in alts)
            self.rx[name] = re.compile(r"(?<![\w'’])(?:" + pat + r")(?![\w])", re.I)

    def era(self, name):
        return self.members[name]["era"]

    def mentions(self, s):
        """{member: first position} for members named in full (by a full form or alias) in s."""
        out = {}
        for name, rx in self.rx.items():
            m = rx.search(s)
            if m:
                out[name] = m.start()
        return out


def load_names(d):
    acts = sorted({m["act"] for m in d["members"]})
    return {a: ActNames(a, [m for m in d["members"] if m["act"] == a], d["nonmembers"], d["member_labels"],
                        d["no_bare_surname"], d["duplicates"]) for a in acts}, set(d["stoplist"])


# ── people named in a stretch of text ────────────────────────────────────────
PARTICLES = ("de", "van", "von", "la", "le", "da", "di", "del", "der", "den", "du", "dos")
ROLE_WORDS = (r"singer|guitarist|bassist|drummer|frontman|frontwoman|keyboardist|keyboard player|vocalist|star|legend|icon|"
              r"founder|co-founder|member|producer|songwriter|singer-songwriter|rapper|dj|pianist|saxophonist|leader|"
              r"mainman|maestro|artist|musician|bass player|guitar hero|percussionist|composer|engineer|manager|director")
# capitalised words in titles that end a name rather than continue it ("Joe Elliott Reflects on...")
NAME_BREAK = set(("talks talk speaks reveals recalls reflects remembers looks opens explains discusses tells says picks "
                  "chooses selects shares answers breaks walks guides chats names lists ranks gets goes dishes warming "
                  "returns announces releases celebrates confirms covers drops debuts joins leaves plays performs sings "
                  "writes is was has on in at and with about interview interviews interviewed q&a meets meet talking "
                  "chatting exclusive live tops hits wins scores earns makes takes gives sets leads launches unveils "
                  "who what where when why how which whose eyes looks article girl girls boy boys band woman women man "
                  "men brothers sisters").split()) | set(ROLE_WORDS.replace("-", " ").split("|"))
# first names that are mostly words in this corpus ("Sunshine Walkers", "Prince Charming", "Tiny Desk")
NONNAME_FIRST = set("See Man Numbers Hang Brain Carry Tiny Son Sunshine Cherry Genesis Forest Merry Bee Prince Karma "
                    "Liberty Windy China Belle Ja Valentine".split())
# a last word that makes the run a place or an organisation ("El Rey Theatre", "Erika Heet Photography")
ORG_WORDS = set(("theatre theater arena center centre stadium club college university school records recordings "
                 "music studios studio festival clinic hotel avenue times magazine media radio news billboard "
                 "productions orchestra choir show awards award photography photos pictures films tv support "
                 "foundation institute museum cathedral academy society publishing entertainment management "
                 "agency").split())
NOT_PERSONS = {"cherry red", "santa barbara", "santa monica", "santa cruz", "santa fe", "santa ana", "santa clara",
               "victoria bc", "rose bowl", "iron maiden", "jack swing"}


def _strip_word(w):
    return re.sub(r"^[\"'‘“(\[]+|[\"'’”),.:;!?\]…]+$", "", w)


def person_spans(s, names, stop, act=None, lines=False):
    """People named in s, in order: [(start, end, name, member or None)]. Members by any full form; other people as a
    capitalised first name (from the first-name list) followed by one to three capitalised words ("Singer Paulette
    McWilliams", "Tracy Chapman Producer David Kershenbaum"). The act's own name is not a person. "Bruce and Jamie
    Watson" names Bruce Watson too."""
    out, taken = [], []
    for name, rx in names.rx.items():
        for m in rx.finditer(s):
            out.append((m.start(), m.end(), name, name))
            taken.append((m.start(), m.end()))
    masked = list(s)
    for a, b in taken:
        masked[a:b] = "#" * (b - a)
    masked = "".join(masked)
    if act:
        masked = _act_rx(act).sub(lambda m: "#" * len(m.group(0)), masked)
    toks = [(m.start(), m.end(), m.group(0)) for m in re.finditer(r"[^\s]+", masked)]
    k = 0
    while k < len(toks):
        a, b, w = toks[k]
        word = _strip_word(w)
        cap = word[:1].isupper() and "#" not in word
        caps_run = word.isupper() and len(word) > 2
        first = word.title() if caps_run else word
        if cap and first in S.FIRST_NAMES_STRICT and first not in S.FUNC and first not in stop and \
                first not in NONNAME_FIRST and not re.search(r"['’]s$", word) and not re.search(r"[,.:;!?…)]$", w):
            parts, ends, j = [first], [b], k + 1
            while j < len(toks) and len(parts) < 4:
                w2 = toks[j][2]
                if re.search(r"^[\"'‘“(]", w2) or "#" in w2 or (lines and "\n" in s[toks[j - 1][1]:toks[j][0]]):
                    break
                core = re.sub(r"[\"'’”),.:;!?\]…]+$", "", w2)
                poss = bool(re.search(r"['’]s?$", core)) and len(core) > 2
                core = re.sub(r"['’]s?$", "", core)
                initial = bool(re.fullmatch(r"[A-Z]\.", w2))
                if not core or not (core[:1].isupper() or core in PARTICLES) or core.title() in S.FUNC or \
                        core.lower() in NAME_BREAK or (core.isupper() and len(core) >= 2 and not caps_run):
                    break
                parts.append(core.title() if core.isupper() and len(core) > 2 else core)
                ends.append(toks[j][0] + len(core) + (1 if initial else 0))
                j += 1
                if poss or (re.search(r"[,.:;!?)…]$", w2) and not initial):
                    break
            while len(parts) > 1 and (parts[-1] in PARTICLES or parts[-1].lower() in ORG_WORDS):
                parts.pop()
                ends.pop()
            start = a + (len(w) - len(w.lstrip("\"'‘“([")))
            if len(parts) >= 2 and " ".join(parts).lower() not in NOT_PERSONS and \
                    not (parts[0] == "Santa" or (len(parts) == 2 and parts[1].isupper() and len(parts[1]) == 2)):
                out.append((start, ends[-1], " ".join(parts), None))
                k = j
                continue
            # "Bruce and Jamie Watson": Bruce Watson
            m2 = re.match(r"\s+(?:and|&)\s+([A-Z][\w'’.-]+)\s+([A-Z][\w'’-]+)", s[b:b + 60])
            if len(parts) == 1 and m2 and m2.group(1) in S.FIRST_NAMES and m2.group(2) not in S.FUNC:
                full = first + " " + m2.group(2)
                mem = next((n for n in names.members if fold(full) in {fold(f) for f in names.members[n]["forms"]}),
                           None)
                out.append((start, b, mem or full, mem))
        k += 1
    return sorted(out)


def persons_in(s, names, stop, act=None):
    """[(position, name, member or None)]."""
    return [(a, n, m) for a, b, n, m in person_spans(s, names, stop, act)]


def quoted_spans(s):
    """Character ranges inside quotation marks (headline quotes: "'I love Sade but...'"), including a quotation that
    opens a title and contains apostrophes ("'I've followed Radiohead from the start...': Rickie Lee Jones's...")."""
    spans = []
    m = re.match(r"\s*['‘\"“].{6,300}?['’\"”](?=\s*[:–—-]\s)", s)
    if m:
        spans.append((m.start(), m.end()))
    for m in re.finditer(r"\"[^\"\n]{2,300}\"|(?<![\w])'[^'\n]{2,300}'(?![\w])|“[^”\n]{2,300}”|‘[^’\n]{2,300}’", s):
        spans.append((m.start(), m.end()))
    return spans


def outside_quotes(pos, spans):
    return not any(a <= pos < b for a, b in spans)


def unquoted(s, spans):
    out, pos = [], 0
    for a, b in spans:
        if a >= pos:
            out.append(s[pos:a])
            pos = b
    out.append(s[pos:])
    return "".join(out)


# ── whose interview is it? ───────────────────────────────────────────────────
TITLE_CUE_BEFORE = re.compile(r"(?i)(?:(?:q\s?&\s?a|interview(?:ed|s)?|spotlight(?: on)?|profile|soundtrack of (?:my|his|her|their) "
                              r"life|this much i know|on my radar|five questions|questions|in (?:his|her|their) own words|"
                              r"exclusive)\s*(?:with|:|—|–|-|\||\.{3}|…)?\s*…?|(?:in conversation|conversation|chat(?:s|ting)?|"
                              r"talk(?:s|ing)?|meets?|meeting|catching up|catch(?:es)? up|sits? down|one-on-one|q\s?&\s?a)"
                              r"\s+(?:with|to)\s*…?)\s*$")
# what may stand between the name and the verb: ", the Go-Go's bassist,", " of Big Country", " (Dire Straits)"
APPOSITIVE = (r"(?:\s*,[^,\n]{2,120},|\s+(?:of|from)\s+(?:the\s+)?[A-Z][\w'’.&+-]*(?:\s+(?:[A-Z][\w'’.&+-]*|and|the|&)){0,4}"
              r"|\s*\([^)\n]{1,60}\))?")
TITLE_CUE_AFTER = re.compile(r"(?i)^" + APPOSITIVE + r"\s*(?::|\s(?:talks|talk|speaks|reveals|recalls|reflects|"
                             r"remembers|looks back|looking back|opens up|explains|discusses|tells|says|picks|chooses|"
                             r"selects|shares|on\b|in conversation|interview|q\s?&\s?a|answers|breaks down|walks us|"
                             r"guides us|chats|names|lists|ranks|gets candid|gets personal|goes deep|dishes|interviewed))")
# what may come before the speaker's name at the start of a title: nothing, "Interview:", "Culture Club's",
# "Big Country Redux guitarist", "Exclusive |"
LEAD_OK = re.compile(r"(?i)(?:^|[:—–\-|(]|['’]s|s['’]|\b(?:" + ROLE_WORDS + r"))\s*$")
# "How Robbie McIntosh Became...", "Why Jennifer Holliday Decided...", '"...": How Marco Pirroni teamed...'
HOW_LEAD = re.compile(r"(?i)(?:^|[:.!?|—–]\s*|[\"”’']\s*:?\s*)(?:how|why|when|where)\s*$")
# a name in the title after these is the topic, not the speaker ("My Ö1 Interview on Tracy Chapman")
TOPIC_BEFORE = re.compile(r"(?i)\b(?:on|about|by|feat\.?|featuring|ft\.?|vs\.?|versus|like|than|for|to|re:?)\s*$")
# names after these are the journalist's, not the interviewee's
BY_BEFORE = re.compile(r"(?i)(?:\bby|\bwords:?|\bphotos?:?|\bphotography:?|\binterviewer:?|\bhost(?:ed)? by|"
                       r"\bpresented by|\bwritten by|\bwith host|\bhost|\bpresenter|\bcorrespondent|\breporter|"
                       r"\bour writer|\bjournalist)\s*$")
# "X interviews Y": X is the journalist
ASKER_AFTER = re.compile(r"(?i)^\s+(?:interviews|interviewed|quizzes|quizzed|grills|grilled|asks|asked|questions|"
                         r"questioned|puts (?:the |some |a few )?questions to|put (?:the |some |a few )?questions to)"
                         r"\s+(?:the\s+)?[A-Z]")
# "X talks to Y": either could be the interviewee; the member is
MUTUAL_AFTER = re.compile(r"(?i)^\s+(?:meets|met|catches up with|caught up with|sits down with|sat down with|"
                          r"speaks (?:to|with)|spoke (?:to|with)|talks (?:to|with)|talked (?:to|with)|chats (?:to|with)|"
                          r"chatted (?:to|with)|in conversation with|in discussion with|goes one-on-one with|"
                          r"went one-on-one with)\s+")
INTRO_CUE_BEFORE = re.compile(r"(?i)(?:spoke|speak|speaks|speaking|talked|talk|talks|talking|chatted|chat|chats|chatting|"
                              r"caught up|catch up|catches up|sat down|sit down|met|meet|interviewed|interview|interviewing|"
                              r"conversation|asked|asking|joined by|joins us|welcome|welcomed|questions? (?:to|for)|q&a|"
                              r"checked in|check in)"
                              r"(?:\s+(?:to|with|exclusively (?:to|with)|by|in|for))?\s+"
                              r"(?:(?:the\s+)?[\w'’.-]+(?:\s+[\w'’.-]+){0,4}?\s+)?$")
INTRO_CUE_AFTER = re.compile(r"(?i)^" + APPOSITIVE + r"\s+(?:(?:has|had|have|was|is|recently|also|very|kindly|graciously|"
                             r"generously|gladly|finally|now|then|later|exclusively|happily)\s+){0,3}"
                             r"(?:talks|talked|tells|told|explains|explained|discusses|discussed|"
                             r"reveals|revealed|recalls|recalled|reflects|reflected|answered|answers|answering|kindly|chatted|"
                             r"chats|spoke|speaks|gives|gave|shares|shared|takes us|took us|looks back|looked back|"
                             r"opens up|opened up|joins|joined|was kind enough|took (?:some )?time|takes (?:some )?time|"
                             r"taken (?:some |the )?time|sat down|sits down|caught up|has been speaking|"
                             r"has been talking|dishes|met|meets|agrees|agreed|checked in|checks in|said|says|"
                             r"admits|admitted|insists|insisted)\b")
# a quotation attributed in the introduction: "...," says Ray
SAYS_BEFORE = re.compile(r"(?i)[,.!?…][\"”’']?\s*(?:says|said|explains|explained|recalls|recalled|adds|added|admits|"
                         r"admitted|laughs|laughed|insists|insisted|continues|continued|notes|noted|tells us|told us)\s+$")


def _short_names(names, name):
    """A member's first name or surname when it names only them in the act and is not an ordinary word."""
    toks = [t for t in _tokens(name) if t.casefold() not in SUFFIX]
    out = []
    if len(toks) >= 2:
        for t, idx, minlen in ((toks[0], names.first, 4), (toks[-1], names.surname, 3)):
            if idx is names.surname and fold(t) in names.no_bare:
                continue
            if len(t) >= minlen and t[0].isupper() and len(idx.get(fold(t), ())) == 1 and t not in S.FUNC and \
                    t not in S.WORDLIKE_NAMES:
                out.append(t)
    return out


def _shared_short(names, name, named):
    """A surname (or first name) the member shares with other members, when only they are named in full here
    ("Knopfler" when the introduction names Mark Knopfler and not David)."""
    toks = [t for t in _tokens(name) if t.casefold() not in SUFFIX]
    out = []
    if len(toks) >= 2:
        for t, idx, minlen in ((toks[0], names.first, 4), (toks[-1], names.surname, 3)):
            sharers = idx.get(fold(t), set())
            if len(sharers) > 1 and not (sharers - {name}) & set(named) and len(t) >= minlen and t[0].isupper() and \
                    t not in S.FUNC and t not in S.WORDLIKE_NAMES and not (idx is names.surname and fold(t) in names.no_bare):
                out.append(t)
    return out


def _edits(a, b):
    """Levenshtein distance."""
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _misspelt(tn, tm):
    """Same first name and a surname one letter off (two for long surnames)."""
    return len(tn) == len(tm) == 2 and tn[0] == tm[0] and len(tm[1]) >= 5 and tn[1] != tm[1] and \
        _edits(tn[1], tm[1]) <= (1 if len(tm[1]) < 8 else 2)


def _accentless(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def _author_words(author):
    a = fold(_accentless(author or ""))
    return {w for w in re.findall(r"[a-z0-9'’-]+", a) if len(w) > 1}


# ── names written slightly differently (stage 3 only; version 2) ─────────────
def _plain(s):
    """A name for a loose comparison: case, accents, hyphens, full stops and spacing ignored, so "Pål Waaktaar Savoy"
    is "Pål Waaktaar-Savoy" and "Paul S Denman" is "Paul S. Denman"."""
    return re.sub(r"[\s\-‐.]+", " ", fold(_accentless(s or ""))).strip()


_PLAIN_FORMS = {}


def plain_forms(names):
    """{plain form: {members}} for an act's members (every full form and alias)."""
    key = id(names)
    if key not in _PLAIN_FORMS:
        d = defaultdict(set)
        for n, m in names.members.items():
            for f in set(m["forms"]) | {n}:
                if len(_plain(f).split()) >= 2:
                    d[_plain(f)].add(n)
        _PLAIN_FORMS[key] = (names, d)
    return _PLAIN_FORMS[key][1]


def leading_run(names, lab):
    """Members whose name begins with the label's words, when the label has two words or more and is shorter than
    the name ("Slim Jim" for Slim Jim Phantom)."""
    lt = _plain(lab).split()
    if len(lt) < 2:
        return set()
    out = set()
    for n, m in names.members.items():
        for f in set(m["forms"]) | {n}:
            ft = _plain(f).split()
            if len(ft) > len(lt) and ft[:len(lt)] == lt:
                out.add(n)
    return out


@functools.lru_cache(maxsize=None)
def _short_rx(short):
    return re.compile(r"(?<![\w'’\-])" + loose(short) + r"(?![\w\-])")


def _short_hits(s, short, spans_skip, quotes):
    """Positions of a one-word name in s, not inside a longer name, a quotation or a hyphenated word."""
    out = []
    for mm in _short_rx(short).finditer(s):
        if any(a <= mm.start() < b for a, b in spans_skip) or not outside_quotes(mm.start(), quotes):
            continue
        before = s[max(0, mm.start() - 25):mm.start()]
        if re.search(r"(?<![\w'’])[A-Z][\w'’.-]+\s+$", before) and not re.search(
                r"(?:^|[.!?:;\"“”—–-]\s+|\b(?:" + "|".join(sorted(S.FUNC)) + r")\s+)$", before):
            continue                               # part of someone else's name ("Kirsten Adamson")
        out.append((mm.start(), mm.end()))
    return out


# titles that present a piece as someone's own words; with the person at the head of the title, they are the speaker
TITLE_VOICE = re.compile(r"(?i)\b(?:interview(?:ed|s)?|q\s?&\s?a|in conversation|conversation with|honest playlist|"
                         r"playlist|track[- ]by[- ]track|excerpt|memoir|that's me in the picture|message from|"
                         r"(?:a )?letter (?:from|to)|a note from|favou?rite|picks|chooses|selects|in (?:his|her|their|my) own "
                         r"words|looks back|remembers|recalls|reflects)\b")
# a title made of the speaker's own quotation and their name: "'I don't play music when I have sex': Michelle Visage"
QUOTE_TITLE = re.compile(r"^\s*[\"'‘“].{8,}[\"'’”]\s*[:–—-]\s*$")


def interviewee(names, stop, title, intro, act="", author="", hosts=(), url=""):
    """The person (or people) whose words the interview or account gives, from the title and the introduction.
    Evidence comes in tiers: (1) the title's subject; (2) a person mentioned in the title whom the introduction
    presents as speaking; (3) a person the introduction presents as speaking; then weaker fallbacks.
    Returns dict(kind=member|members|act|other|mixed|others|unknown, who=[names], how=..., confidence=...)."""
    t = title or ""
    intro = intro or ""
    authors = _author_words(author)
    solo = len([n for n in names.members if names.era(n) == "Nomination-era"]) == 1
    tq, iq = quoted_spans(t), quoted_spans(intro)
    subj, cue, ment, member_of, journalists = set(), set(), set(), {}, set()
    marked = set()          # title subjects marked by a cue, not merely standing at the head of the title
    voice = bool(S.TITLE_FP.search(t) or TITLE_VOICE.search(t) or S.TITLE_PICKS.search(t) or
                 re.search(r"(?i)/[^/]*(?:interview|q-?and-?a|q-?a)[^/]*/", re.sub(r"^https?://[^/]+", "", url or "")))
    mentions = Counter()

    # 1. the title
    tp = [x for x in person_spans(t, names, stop, act) if outside_quotes(x[0], tq)]
    subj_next, first_free, last_subject, prev_end = False, True, None, 0
    for k, (p, end, n, m) in enumerate(tp):
        if m:
            member_of[n] = m
        mentions[n] += 1
        lead, after = t[:p], t[end:end + 140]
        lead_u = unquoted(lead, [x for x in tq if x[1] <= p])
        nxt = tp[k + 1] if k + 1 < len(tp) else None
        between, prev_end = t[prev_end:p], end
        if BY_BEFORE.search(lead[-30:]) or ASKER_AFTER.match(after):
            if not m:
                journalists.add(fold(n))
            subj_next = bool(ASKER_AFTER.match(after))
            continue
        mu = MUTUAL_AFTER.match(after)
        if mu and not m and nxt and nxt[3] and nxt[0] - (end + mu.end()) <= 60:   # "Pete Feenstra talks to John Illsley"
            journalists.add(fold(n))
            subj_next = True
            continue
        # a name modifying the next one: "Tracy Chapman Producer David Kershenbaum"
        if nxt and re.fullmatch(r"(?i)(?:['’]s?)?\s+(?:(?:" + ROLE_WORDS + r")\s+)+", t[end:nxt[0]]):
            ment.add(n)
            continue
        prev_and = last_subject == k - 1 and k > 0 and bool(re.fullmatch(r"(?i)\s*(?:and|&|,)\s*", between))
        asked = subj_next                         # the person after "X interviews"
        at_head = subj_next or prev_and or (first_free and (
            (len(lead_u) <= 70 and bool(LEAD_OK.search(lead_u.strip()))) or bool(HOW_LEAD.search(lead_u))))
        subj_next, first_free = False, False
        topic = bool(TOPIC_BEFORE.search(lead[-12:]))
        later = last_subject is not None and not prev_and and not at_head   # "X - Interview — ANTHONY R. CARDNO"
        cued = bool(TITLE_CUE_BEFORE.search(lead[-60:]) or TITLE_CUE_AFTER.match(after) or mu or asked)
        if not topic and not later and (cued or at_head):
            subj.add(n)
            if cued or (prev_and and tp[k - 1][2] in marked) or voice or QUOTE_TITLE.match(lead):
                marked.add(n)
            last_subject = k
        else:
            ment.add(n)
    full_t = [(p, e) for p, e, n, m in tp]
    t_masked = _act_rx(act).sub(lambda mm: "#" * len(mm.group(0)), t) if act else t
    # members named in the title by one name ("Interview with Katrina", "THE HOUSE OF STEWART")
    for name in names.members:
        if name in member_of:
            continue
        for short in _short_names(names, name):
            hits = _short_hits(t_masked, short, full_t, tq) + \
                (_short_hits(t_masked, short.upper(), full_t, tq) if len(short) > 3 else [])
            for a, b in hits:
                after, lead = t[b:b + 140], t[:a]
                member_of[name] = name
                mentions[name] += 1
                if TOPIC_BEFORE.search(lead[-12:]) or BY_BEFORE.search(lead[-30:]):
                    continue
                if TITLE_CUE_AFTER.match(after) or TITLE_CUE_BEFORE.search(lead[-60:]) or HOW_LEAD.search(lead):
                    subj.add(name)
                    if not HOW_LEAD.search(lead):
                        marked.add(name)
                else:
                    ment.add(name)
    # the act itself as the title's subject ("Music Q&A: Swing Out Sister", "Pretenders interview: ...")
    act_subj = bool(act) and not solo and act_subject(t, act, tq)

    # 2. the introduction
    ip = [x for x in person_spans(intro, names, stop, act) if outside_quotes(x[0], iq)]
    for k, (p, end, n, m) in enumerate(ip):
        if m:
            member_of[n] = m
        mentions[n] += 1
        after = intro[end:end + 200]
        nxt = ip[k + 1] if k + 1 < len(ip) else None
        asker = ASKER_AFTER.match(after)
        if BY_BEFORE.search(intro[max(0, p - 30):p]) or asker:
            if not m:
                journalists.add(fold(n))
            if asker and nxt and nxt[0] - end <= 80:
                cue.add(nxt[2])
            continue
        mu = MUTUAL_AFTER.match(after)
        if mu and not m and nxt and nxt[3] and nxt[0] - (end + mu.end()) <= 60:
            journalists.add(fold(n))
            cue.add(nxt[2])             # "Paul Gilby talks to their drummer, Ped (Peter Gill)"
            continue
        if INTRO_CUE_BEFORE.search(intro[max(0, p - 90):p]) or INTRO_CUE_AFTER.match(after) or \
                SAYS_BEFORE.search(intro[max(0, p - 30):p]):
            cue.add(n)
    # people named in full, then by one name ("Knopfler met RC", "as Trevor explains", "says Ray")
    owners = defaultdict(set)
    named_members = {m for p, e, n, m in tp + ip if m}
    for p, e, n, m in tp + ip:
        if m:
            for short in _short_names(names, m) + _shared_short(names, m, named_members):
                owners[short].add(n)
        else:
            toks = n.split()
            for tk in (toks[0], toks[-1]):
                if len(tk) >= 3 and tk not in S.FUNC and tk.title() not in stop and not re.fullmatch(r"[A-Z]\.?", tk):
                    owners[tk].add(n)
    full_i = [(p, e) for p, e, n, m in ip]
    i_masked = _act_rx(act).sub(lambda mm: "#" * len(mm.group(0)), intro) if act else intro
    for short, who in owners.items():
        if len(who) != 1:
            continue                                 # "Paul" for Paul Rutherford and Paul Morley
        n = next(iter(who))
        for a, b in _short_hits(i_masked, short, full_i, iq):
            mentions[n] += 1
            after = re.sub(r"^['’]s\b", "", intro[b:b + 200])
            if INTRO_CUE_BEFORE.search(intro[max(0, a - 90):a]) or INTRO_CUE_AFTER.match(after) or \
                    SAYS_BEFORE.search(intro[max(0, a - 30):a]):
                cue.add(n)

    # 3. one key per person: spellings ("MARTYN WARE"), short forms ("Slim Jim" for Slim Jim Phantom) and longer runs
    # ("Joe Elliott Reflects" for Joe Elliott)
    every = sorted(subj | cue | ment | set(mentions), key=lambda x: (x not in member_of, len(x), x))
    key = {}
    for n in every:
        tn = fold(n).split()
        k2 = None
        for k in list(dict.fromkeys(key.values())):
            tk = fold(k).split()
            if tk == tn or (len(tn) > len(tk) and tn[:len(tk)] == tk and n not in member_of) or \
                    (len(tn) < len(tk) and tk[:len(tn)] == tn and k in member_of):
                k2 = k
                break
        if k2 is None and n not in member_of:
            for name in sorted(set(member_of.values())):
                if any(fold(f).split()[:len(tn)] == tn for f in names.members[name]["forms"]) or \
                        _misspelt(tn, fold(name).split()):          # "Mick Hucknell" for Mick Hucknall
                    k2 = name
                    break
        key[n] = k2 or n
    canon = lambda xs: {key[x] for x in xs}
    subj, cue, ment, marked = canon(subj), canon(cue), canon(ment), canon(marked)
    cnt = Counter()
    for n, v in mentions.items():
        cnt[key[n]] += v
    mentions = cnt
    member_of = {key[n]: m for n, m in member_of.items()}

    def journalist(n):
        """Authors, hosts and named journalists are not interviewees; an author who is the title's subject is the
        narrator ("Dalton Ross on Boy George's..."), and a member never counts as a journalist."""
        if n in member_of:
            return False
        f = fold(_accentless(n))
        if f in hosts or (f in journalists and n not in subj):
            return True
        return bool(authors) and {w for w in f.split() if len(w) > 1} <= authors and n not in subj

    def verdict(group, how, conf):
        group = sorted(x for x in group if not journalist(x))
        if not group:
            return None
        mk = all(x in marked or x in cue for x in group)
        mem = [x for x in group if x in member_of]
        if len(group) == 1:
            return dict(kind="member" if mem else "other", who=group, how=how, confidence=conf, marked=mk)
        if len(mem) == len(group):
            return dict(kind="members", who=[member_of[x] for x in mem], how=how + " (several)", confidence=conf,
                        marked=mk)
        return dict(kind="mixed" if mem else "others", who=group, how=how + " (several)", confidence="low", marked=mk)

    def dominant(pool, ratio=1.5):
        """The member named far more often than anyone else in the title and introduction."""
        memc = sorted(((v, n) for n, v in mentions.items() if n in member_of and n in pool), reverse=True)
        oth = max([v for n, v in mentions.items() if n not in member_of and not journalist(n)] or [0])
        if memc and memc[0][0] >= 3 and memc[0][0] > oth and (len(memc) == 1 or memc[0][0] >= ratio * memc[1][0]):
            return memc[0][1]
        return None

    if subj or act_subj:
        persons = verdict(subj, "title", "medium")
        if persons:
            return persons
        if act_subj:
            # the band is the subject: a member the introduction presents as speaking, or the member it is about
            cued = [n for n in cue if n in member_of and not journalist(n)]
            if len(cued) == 1:
                return dict(kind="member", who=cued, how="the act in the title; this member in the introduction",
                            confidence="medium")
            d = dominant(set(member_of), ratio=2)
            if d:
                return dict(kind="member", who=[d], how="the act in the title; the member the introduction is about",
                            confidence="low")
            return dict(kind="act", who=[], how="the act is named in the title; which member answers is not said",
                        confidence="low")
    res = verdict(ment & cue, "title and introduction", "medium") or verdict(cue, "introduction", "medium")
    if res:
        return res
    # 4. weaker: one member mentioned in the title and no one else; the member the introduction is about
    tm = [n for n in ment if not journalist(n)]
    oth = max([v for n, v in mentions.items() if n not in member_of and not journalist(n)] or [0])
    if len(tm) == 1 and tm[0] in member_of and (mentions[tm[0]] > oth or oth <= 1):
        return dict(kind="member", who=tm, how="named in the title", confidence="low")
    if not [n for n in tm if n not in member_of]:
        d = dominant(set(member_of))
        if d:
            return dict(kind="member", who=[d], how="the member the introduction is about", confidence="low")
    return dict(kind="unknown", who=[], how="", confidence="unknown")


def _act_rx(act):
    words = act.split()
    if words and words[0].lower() == "the":
        words = words[1:]
    return re.compile(r"(?<![\w])(?:the\s+)?" + r"[\s\-]+".join(re.escape(w) for w in words) + r"(?![\w])", re.I)


def act_in(title, act):
    return bool(_act_rx(act).search(title or ""))


def act_subject(t, act, tq=()):
    """The act's name in a subject position of the title."""
    for mm in _act_rx(act).finditer(t):
        if not outside_quotes(mm.start(), tq):
            continue
        lead, after = t[:mm.start()], t[mm.end():mm.end() + 140]
        if BY_BEFORE.search(lead[-30:]) or re.match(r"\s*['’]s\b", after):
            continue
        if TITLE_CUE_BEFORE.search(lead[-60:]) or TITLE_CUE_AFTER.match(after) or HOW_LEAD.search(lead) or \
                (len(lead) <= 70 and LEAD_OK.search(lead.strip()) and
                 re.match(r"(?i)\s*(?:$|[:|—–-]|interview|q\s?&\s?a|talk|on\b|in conversation|answer)", after)):
            return True
    return False


def _site_label(url):
    """The site's own name: 'straycats' for https://www.straycats.com/..., 'rickieleejones' for
    rickieleejones.substack.com; None for .org sites, which are fan or charity sites here."""
    host = re.sub(r"^https?://(?:www\.)?", "", (url or "").lower()).split("/")[0]
    parts = host.split(".")
    if len(parts) < 2 or "org" in parts[1:]:
        return None
    return re.sub(r"[^a-z0-9]", "", parts[0])


def _is_name_site(label, name):
    key = re.sub(r"[^a-z0-9]", "", _accentless(name).lower().replace("&", "and"))
    key = re.sub(r"^the", "", key) if len(key) > 7 else key
    return len(key) >= 4 and label is not None and (label == key or label == "the" + key or label == "official" + key or
                                                   label in (key + "music", key + "official", key + "band",
                                                             key + "hq"))


def official_site(outlet, url, act):
    """The act's own website (not a fan site: 'direstraitsblog.com', 'about-tracy-chapman.net')."""
    return _is_name_site(_site_label(url), act)


def member_site(url, names):
    """The member whose own website this is ('colinhay.com', 'geoffdownes.com')."""
    label = _site_label(url)
    return next((n for n in names.members if _is_name_site(label, n)), None)


def person_from_author(author):
    """'Sweeting; Adam' -> 'Adam Sweeting'; outlet-like or empty authors give None."""
    a = (author or "").strip()
    if not a:
        return None
    if ";" in a:
        parts = [x.strip() for x in a.split(";")]
        if len(parts) == 2 and all(re.fullmatch(r"[A-Z][\w'’.-]+(?: [A-Z][\w'’.-]+)*", x) for x in parts):
            a = parts[1] + " " + parts[0]
        else:
            a = parts[0]
    toks = a.split()
    if len(toks) >= 2 and toks[0].title() in S.FIRST_NAMES and all(t[0].isupper() for t in toks):
        return a
    return None


def who_rule(how):
    """The rule text for a turn given to the interviewee: "interviewee named in the title", "interviewee: the member
    the introduction is about" (version 2; version 1 wrote "interviewee named in the named in the title")."""
    how = how or "title"
    if how.startswith(("title", "introduction")):
        return "interviewee named in the " + how
    if how.startswith("named in"):
        return "interviewee " + how
    return "interviewee: " + how


# ── one record ───────────────────────────────────────────────────────────────
class Record:
    """What stage 3 needs to know about a record: its text, the members it names in full, and its interviewee."""

    def __init__(self, rid, text, title, act, route, detail, segments, names, stop, author="", outlet="", url=""):
        self.rid, self.text, self.title, self.act, self.route, self.detail = rid, text, title or "", act, route, detail
        self.names, self.stop = names, stop
        self.segments = [tuple(x) for x in segments]
        self.blocks = [(a, b) for a, b, k in self.segments if k != "prose"]
        whole = self.title + "\n" + text
        self.full = names.mentions(whole)                    # members named in full anywhere
        self.pspans = S.para_spans(text)
        first_block = self.blocks[0][0] if self.blocks else len(text)
        self.intro = text[:first_block][-2500:]
        self.head = self.title + "\n" + self.intro            # title and introduction
        self.head_members = set(names.mentions(self.head))
        self.hosts = {fold(m.group(1)) for m in re.finditer(
            r"(?mi)^[ \t]*(?:(?:Mr|Ms|Mrs|Dr|Miss)\.?[ \t]+)?([^,:\n]{2,60}),[ \t]*(?:host|co-host|guest host|byline|anchor|"
            r"correspondent|reporter|presenter)[ \t]*:", text)}
        self.who = interviewee(names, stop, self.title, self.intro, act, author, self.hosts, url)
        # version 2: a caption transcript whose title names one member besides the programme ("Overheard with Evan
        # Smith | Kathy Valentine | Season 10 | Episode 25") is that member's interview
        if (detail or "").startswith("caption") and self.who["kind"] in ("unknown", "mixed", "others"):
            tm = names.mentions(self.title)
            if len(tm) == 1:
                self.who = dict(kind="member", who=sorted(tm), how="named in the title of the programme's episode",
                                confidence="medium", marked=True)
        self.author = person_from_author(author)
        self.author_raw = author or ""
        self.outlet = outlet or ""
        self.url = url or ""
        self.official = official_site(outlet, url, act)
        self.site_member = member_site(url, names)
        # other people sharing a member's surname or first name, named in full in the record. A member's full name
        # written slightly differently ("Paul S Denman") is the member, not a namesake (version 2).
        self.namesakes = defaultdict(set)
        pf = plain_forms(names)
        for p, n, m in persons_in(whole, names, stop):
            if m is None:
                same = pf.get(_plain(n), set())
                if len(same) == 1:
                    mem = next(iter(same))
                    self.full[mem] = min(self.full.get(mem, p), p)
                    continue
                toks = n.split()
                self.namesakes[fold(toks[-1])].add(n)
                self.namesakes[fold(toks[0])].add(n)
        for p, n, m in persons_in(self.head, names, stop):
            same = pf.get(_plain(n), set()) if m is None else set()
            if len(same) == 1:
                self.head_members.add(next(iter(same)))
        self._interviewer = self.author or (self.outlet if self.outlet else None)
        # NPR-style "CHRIS (Caller):" labels
        self.callers = {fold(m.group(1)) for m in re.finditer(
            r"(?mi)^[ \t]*([^\n:()]{2,40}?)[ \t]*\((?:caller|audience member|audience|listener|student|fan)\)[ \t]*:", text)}
        self._first_label = {}
        self.field_paras = set()

    def label_start(self, lab):
        """Where a label is first used, as a position in title + text."""
        if lab not in self._first_label:
            m = re.search(r"(?mi)^[ \t]*" + re.escape(lab) + r"\b", self.text)
            self._first_label[lab] = (m.start() + len(self.title) + 1) if m else len(self.title) + 1 + len(self.text)
        return self._first_label[lab]

    # a member's speaker type
    def member_type(self, name):
        return MEMBER if self.names.era(name) == "Nomination-era" else OTHER_ERA

    def speaker_for_who(self, conf_cap=None):
        """The interviewee as a speaker: (type, speaker, confidence, rule)."""
        w = self.who
        if w["kind"] == "member":
            n = w["who"][0]
            return self.member_type(n), n, w["confidence"], who_rule(w["how"])
        if w["kind"] == "members":
            eras = {self.names.era(n) for n in w["who"]}
            typ = MEMBERS_UNSPEC if eras == {"Nomination-era"} else OTHER_ERA_UNSPEC if eras == {"Other-era"} else AMBIGUOUS
            conf = "medium" if typ != AMBIGUOUS else "low"
            return typ, " / ".join(w["who"]), conf, "several members named as interviewees; answers carry no names"
        if w["kind"] == "act":
            return MEMBERS_UNSPEC, None, "low", w["how"]
        if w["kind"] == "other":
            return OTHER, w["who"][0], w["confidence"], who_rule(w["how"])
        if w["kind"] in ("mixed", "others"):
            return UNKNOWN, None, "unknown", "several people named as interviewees, not all members"
        return UNKNOWN, None, "unknown", "no interviewee named in the title or introduction"

    def narrator(self):
        """The speaker of a first-person account with one speaker. The weak fallbacks (a member only mentioned in the
        title, the member the introduction is about, the band as the title's subject) do not name a narrator: fans and
        critics write first-person pieces about the act."""
        if self.site_member:
            n = self.site_member
            return self.member_type(n), n, "medium", "the member's own website"
        # the account names its own narrator: "My name is Aurélie, I'm in my late forties..."
        if self.blocks:
            a0 = self.blocks[0][0]
            m = re.search(r"(?:^|[.!?]\s+)([Mm]y name is|[Mm]y name's|I am|I'm|This is)\s+"
                          r"([A-Z][\w'’.-]+(?:\s+[A-Z][\w'’.-]+){0,3})", self.text[a0:a0 + 300])
            # a member's name here may be a namesake's ("My name is Corey Hart" on a student's fundraising page)
            if m and not self.names.mentions(m.group(2)) and (m.group(1).lower().startswith("my name") or
                                                             _accentless(m.group(2).split()[0]) in S.FIRST_NAMES):
                return OTHER, m.group(2).rstrip(".,"), "high", "the narrator gives their name"
        typ, who, conf, rule = self.speaker_for_who()
        weak = self.who["kind"] == "act" or self.who["how"] in ("named in the title", "the member the introduction is about")
        # a person at the head of the title, with no cue that they speak, is the subject of someone else's piece
        # (an obituary, a fan's tribute, a columnist's essay)
        if typ != UNKNOWN and not weak and self.who.get("marked"):
            return typ, who, conf, rule.replace("interviewee", "narrator")
        if self.official:
            ms = [n for n in self.names.members if self.names.era(n) == "Nomination-era"]
            if len(ms) == 1:
                return MEMBER, ms[0], "medium", "the act's own website; a solo artist"
            return MEMBERS_UNSPEC, None, "low", "the act's own website; which member writes is not said"
        if self.author:
            return OTHER, self.author, "medium", "no narrator named; the author writing in the first person"
        return UNKNOWN, None, "unknown", "no narrator named"

    # ── speaker labels ───────────────────────────────────────────────────────
    def resolve_label(self, label):
        """(type, speaker, confidence, rule) for a speaker label."""
        names = self.names
        lab = re.sub(r"^(?:Mr|Ms|Mrs|Dr|Miss|Sir)\.?\s+", "", label.strip(), flags=re.I).strip(" .:")
        low = fold(lab)
        if low in ("q", "question", "questions", "interviewer", "int", "host", "co-host", "presenter", "moderator",
                   "reporter", "announcer", "narrator", "dj"):
            return INTERVIEWER, self._interviewer, "high", "the interviewer's label"
        if low in self.hosts and not names.mentions(lab):
            return INTERVIEWER, lab.title() if lab.isupper() else lab, "high", "a host or reporter (NPR role)"
        if low in ("a", "answer", "answers", "interviewee", "guest", "respondent"):
            typ, who, conf, rule = self.speaker_for_who()
            return typ, who, ("medium" if conf == "high" else conf), "answer label; " + rule
        if low in ("audience", "audience member", "caller", "fan", "student", "participant", "crowd", "voice",
                   "man", "woman") or low.startswith("unidentified"):
            return OTHER, lab, "high", "a generic speaker"
        if low in self.callers:
            return OTHER, lab.title() if lab.isupper() else lab, "high", "a caller or audience member"
        # the decisions from the name-list review come first
        if low in names.labels:
            cands = names.labels[low]
            if not cands:
                return OTHER, lab.title() if lab.isupper() else lab, "high", "name lists: this label is not a member"
            return self._choose(lab, sorted(cands), "name lists", full=low in names.form or len(lab.split()) >= 2)
        cands, rule = set(), ""
        if low in names.form:
            cands, rule = set(names.form[low]), "full name"
            if len(lab.split()) == 1 and not any(low in {fold(f) for f in names.members[c]["forms"]} for c in cands):
                # "Sade:" is Sade Adu only in a record about the act ("Kit and Sade for Queer Supply" is someone else)
                if act_in(self.title, self.act) or self.head_members:
                    return self._choose(lab, sorted(cands), "the act's name for its singer", full=True)
                return OTHER, lab.title() if lab.isupper() else lab, "medium", \
                    "the act's name for its singer, in a record not about the act: someone else"
            return self._choose(lab, sorted(cands), rule, full=True)
        # a full name written with a space for a hyphen, without an accent or a full stop ("Pål Waaktaar Savoy")
        pf = plain_forms(names).get(_plain(lab))
        if pf:
            return self._choose(lab, sorted(pf), "full name (written slightly differently)", full=True)
        toks = lab.split()
        if len(toks) >= 2:
            # first and last word: "MICK 'RED' HUCKNALL", "Mark Freuder Knopfler"
            f, l = fold(toks[0]), fold(toks[-1])
            both = names.first.get(f, set()) & names.surname.get(l, set())
            if both:
                return self._choose(lab, sorted(both), "full name", full=True)
            # the first words of a member's name ("Slim Jim" for Slim Jim Phantom), by the full-name rule
            lead = leading_run(names, lab)
            if lead and lab.title() not in self.stop:
                return self._choose(lab, sorted(lead), "the first words of a member's name")
            return self._other(lab)
        compact = lab.replace(".", "").replace(" ", "")
        if re.fullmatch(r"[A-Z]{2,4}", compact) and compact in names.initials:
            cands, rule = set(names.initials[compact]), "initials"
        elif low in names.surname and low not in names.no_bare:
            cands, rule = set(names.surname[low]), "surname"
        elif low in names.first:
            cands, rule = set(names.first[low]), "first name"
        if not cands or lab.title() in self.stop:
            return self._other(lab)
        return self._choose(lab, sorted(cands), rule)

    def _other(self, lab):
        disp = lab.title() if lab.isupper() and len(lab) > 3 else lab
        return OTHER, disp, "high" if len(lab.split()) >= 2 else "medium", "a label naming nobody on the line-up"

    def _choose(self, lab, cands, rule, full=False):
        """Settle a label that could be one of several members (or a namesake) by the full-name rule."""
        names = self.names
        if not full:
            # the member must be named in full in the title or introduction, or before the label is first used
            first = self.label_start(lab)
            named = [c for c in cands if c in self.head_members or (c in self.full and self.full[c] < first)]
            if not named:
                low = fold(lab)
                same = sorted({n for n in self.namesakes.get(low, set()) if not names.form.get(fold(n))})
                if len(same) == 1:
                    return OTHER, same[0], "medium", rule + " of someone this record names in full who is not a member"
                if len(cands) == 1 and not same and rule in ("first name", "surname"):
                    c = cands[0]
                    toks = [t for t in _tokens(c) if t.casefold() not in SUFFIX]
                    part = toks[-1] if rule == "first name" else toks[0]
                    conf = "medium" if not self._relative(c) else "low"
                    # "Lee:" when the introduction says "Phantom and Rocker were the rhythm behind the Stray Cats"
                    if len(toks) >= 2 and len(part) >= 3 and \
                            re.search(r"(?<![\w'’])" + re.escape(part) + r"(?![\w])", self.head):
                        return self.member_type(c), c, conf, rule + "; the member's " + \
                            ("surname" if rule == "first name" else "first name") + " is in the title or introduction"
                    if len(lab) >= 4 and (act_in(self.title, self.act) or act_in(self.intro, self.act)):
                        return self.member_type(c), c, conf, rule + " of the only member with it; the record is about " \
                            "the act but does not give the full name first"
                return OTHER, lab.title() if lab.isupper() else lab, "medium", \
                    rule + " of a member not named in full in the title, the introduction or before this label: " \
                    "someone else"
            cands = named
        # a namesake who is not a member, named in full in the record ("Paul Harris" for "Harris:")
        low = fold(lab)
        same = {n for n in self.namesakes.get(low, set()) if not names.form.get(fold(n))} if not full else set()
        if len(cands) == 1:
            c = cands[0]
            conf, why = "high", rule
            if same:
                in_head = [n for n in same if re.search(re.escape(n), self.head, re.I)]
                if in_head and c not in self.head_members:
                    return OTHER, sorted(in_head)[0], "medium", rule + "; the record's interviewee " \
                        "is a namesake who is not a member"
                if c not in self.head_members:
                    return AMBIGUOUS, None, "low", rule + "; a namesake who is not a member is also named"
                why += "; a namesake is named too, but the member is named in the title or introduction"
                conf = "medium"
            # a relative lowers the confidence once more, unless the relative is the namesake just counted (version 2)
            if not full and self._relative(c, exclude=same):
                conf = lower_level(conf)
                why += "; a relative with the same surname is named"
            return self.member_type(c), c, conf, why
        # several members could own the label
        in_head = [c for c in cands if c in self.head_members]
        if len(in_head) == 1:
            c = in_head[0]
            return self.member_type(c), c, "medium", rule + "; several members share it, this one is the interviewee"
        eras = {names.era(c) for c in cands}
        if eras == {"Nomination-era"}:
            return MEMBERS_UNSPEC, " / ".join(sorted(cands)), "medium", rule + " shared by members"
        if eras == {"Other-era"}:
            return OTHER_ERA_UNSPEC, " / ".join(sorted(cands)), "medium", rule + " shared by other-era members"
        return AMBIGUOUS, " / ".join(sorted(cands)), "low", rule + " shared by members of different eras"

    def _relative(self, member, exclude=()):
        toks = [t for t in _tokens(member) if t.casefold() not in SUFFIX]
        if len(toks) < 2:
            return False
        sur = fold(toks[-1])
        return any(fold(n.split()[-1]) == sur and not self.names.form.get(fold(n)) and n not in exclude
                   for n in self.namesakes.get(sur, ()))


# ── cutting the blocks into turns ────────────────────────────────────────────
STAGE_DIRECTION = re.compile(r"^\s*[\(\[](?:soundbite|laughter|laughs|laughing|applause|music|crosstalk|cheering|begin|end|"
                             r"singing|inaudible|pause|break|commercial|musical|audio|video|clip|unintelligible|sound|"
                             r"bleep|technical|archival|music up|instrumental|song|end of|beginning of)\b[^\)\]\n]*[\)\]]\s*$",
                             re.I)
END_PUNCT = ".!?…\"'”’)]"


def heading_like(p):
    s = p.strip()
    return bool(s) and len(s) <= 80 and s[-1] not in END_PUNCT and not s.endswith(":") and \
        bool(S.SHORT_LINE.match(s) or S.TOPIC_HEADING.match(s) or S.BARE_NAME.match(s)) and len(s.split()) <= 12


def furniture_line(p):
    s = p.strip()
    return bool(s) and len(s) <= 220 and bool(S.inblock_furniture(s) or S.FURNITURE.match(s) or S.CLUTTER_ANY.search(s)
                                                or re.search(r"https?://|www\.", s))


def embedded_quotes(s):
    """Quotations inside a turn of at least three words (a speaker reporting someone else's words)."""
    body = s.strip()
    if len(body) > 2 and body[0] in "\"'“‘" and body[-1] in "\"'”’" and body.count(body[0]) <= 2:
        return 0                                  # the whole answer printed in quotation marks
    return sum(1 for m in re.finditer(r"\"([^\"\n]{3,})\"|“([^”\n]{3,})”", s) if len((m.group(1) or m.group(2)).split()) >= 3)


def whole_quoted(s):
    b = s.strip()
    return len(b) > 2 and b[0] in "\"'“‘" and b[-1] in "\"'”’.!?"


class Builder:
    """Collects the spans of one block and merges paragraphs that continue the same turn."""

    def __init__(self, rec):
        self.rec, self.spans = rec, []

    def add(self, start, end, kind, who=None, label=None, speech_start=None, new_turn=False, question=False):
        """who = (type, speaker, confidence, rule)."""
        if start >= end:
            return
        text = self.rec.text
        if not text[start:end].strip():
            if self.spans:
                self.spans[-1]["end"] = max(self.spans[-1]["end"], end)
                return
            kind, who = "non-speech", None
        last = self.spans[-1] if self.spans else None
        typ, spk, conf, rule = who if who else (None, None, None, None)
        if last and not new_turn and last["kind"] == kind and kind in ("turn", "narration", "non-speech", "sung") and \
                last["speaker_type"] == typ and last["speaker"] == spk and last["end"] == start and \
                last["question"] == question:
            last["end"] = end
            return
        member = spk if typ in (MEMBER, OTHER_ERA) else None
        self.spans.append(dict(start=start, end=end, kind=kind, speaker_type=typ, speaker=spk, member=member,
                               era=(self.rec.names.era(member) if member else None),
                               band_voice=bool(kind == "turn" and typ in BAND_VOICE), confidence=conf, rule=rule,
                               label=label, speech_start=speech_start if speech_start is not None else start,
                               question=question))

    def finish(self):
        text = self.rec.text
        for x in self.spans:
            words = text[x["speech_start"]:x["end"]]
            x["quoted"] = x["kind"] == "turn" and whole_quoted(words)
            x["embedded_quotes"] = embedded_quotes(words) if x["kind"] == "turn" else 0
        return self.spans


def _pieces(rec, a, b):
    """(paragraph index, start, end) for the paragraphs of a block, each with its newline."""
    out = []
    for i, (s, e) in enumerate(rec.pspans):
        lo, hi = max(s, a), min(e + 1, b)
        if lo < hi:
            out.append((i, lo, hi))
    return out


# field names in list articles and listings; a line starting with one is page furniture, not speech
FIELD_LABELS = {"songs", "song", "tracks", "track", "singles", "single", "key tracks", "key track", "essential tracks",
                "listen", "hear", "watch", "tickets", "venue", "dates", "where", "when",
                # version 2: a tour listing's fields ("Address:", "Details:") and a studio feature's equipment lists
                "address", "details", "recording", "monitors", "effects", "compressors"}
SUNG_LABEL = re.compile(r"^\s*([A-Z][A-Za-z .'’-]{1,40}):\s*(?=.{0,60}\((?:[Ss]inging|[Rr]apping|[Hh]umming|[Cc]hanting))")


def _label_index(rec, paras):
    """Speaker labels at the start of paragraphs: {paragraph: (label, value offset in paragraph, style)}.
    In labelled Q&As every label stage 2 found counts; elsewhere (a labelled line inside an unlabelled Q&A or a
    first-person account) only a label that names a person and is followed by speech."""
    raw = S.label_lines(paras, rec.title)
    lines = S.merge_labels(raw)
    merged_of = {r[1]: m[1] for r, m in zip(raw, lines)}
    have = {x[0] for x in lines}
    for i, p in enumerate(paras):
        m = SUNG_LABEL.match(p)
        if i not in have and m:
            lines.append((i, m.group(1).strip(), p[m.end(1) + 1:].strip(), "sung"))
    lines.sort()
    cnt = Counter(x[1] for x in lines)
    wo = rec.title + "\n" + rec.text
    out = {}
    for i, lab, val, style in lines:
        clean = S._clean_label(lab)
        if clean.lower() in FIELD_LABELS:
            rec.field_paras.add(i)                # 'Songs: "Pink Houses"' in a list article
            continue
        if clean.lower() in S.METADATA or STAGE_DIRECTION.match("(" + clean + ")"):
            continue
        member = bool(rec.names.mentions(clean)) or fold(clean) in rec.names.labels
        generic = fold(clean) in ("q", "a", "question", "answer", "interviewer")
        if rec.route == "labelled Q&A":
            ok = cnt[lab] >= 2 or style in ("npr", "role", "sung") or generic or S.person_like(lab, wo) or member
            if style == "bracket" and not rec.detail.startswith("caption") and not (cnt[lab] >= 2 and S.person_like(lab, wo)):
                ok = False
        else:
            ok = style == "sung" or style in ("colon", "npr", "role", "space") and (generic or member or (
                S.person_like(lab, wo) and bool(S.SPEECH_MARK.search(val or ""))))
        if not ok:
            continue
        p = paras[i]
        off = p.find(val) if val and style not in ("alone",) else len(p)
        out[i] = (lab, max(off, 0), style)
    if rec.route == "labelled Q&A":
        _recover_labels(rec, paras, out, raw, merged_of)
    return out


def _recover_labels(rec, paras, out, raw, merged_of):
    """Version 2: labels in a labelled Q&A that stage 2 does not read as labels.
    1. Initials stage 2 keeps out as metadata ("CD:", "PR:"), when someone named in full in the record has them and
       the words read as speech: Corinne Drewery's "CD:" between the interviewer's "QH:" and Andy Connell's "AC:".
    2. A label the record uses, followed by a semicolon: "Knopfler; Oh yeah, probably."
    3. Initials the record uses as a label two or more times, run into the first word: "JLYeah.", "JLI was living..."
    The paragraph must be inside a block."""
    used = Counter(r[1] for r in raw if r[0] in out)
    whole = rec.title + "\n" + rec.text
    full = sorted({lab for lab, off, style in out.values() if len(lab.split()) >= 2})
    for i, p in enumerate(paras):
        if i in out or not p.strip() or not any(a <= rec.pspans[i][0] < b for a, b in rec.blocks):
            continue
        m = re.match(r"[ \t]*[A-Z][\w&.-]*(?:[ \t][A-Z][\w&.-]*){0,2}['’]s[ \t]+([A-Z][a-z'’-]+(?:[ \t][A-Z][a-z'’-]+){1,2})"
                     r"[ \t]*:[ \t]+(?=\S)", p)
        if m and _author_words(rec.author_raw) and \
                {w for w in re.findall(r"[a-z'’-]+", fold(m.group(1))) if len(w) > 1} <= _author_words(rec.author_raw):
            out[i] = (m.group(1), m.end(), "colon")       # "RAMzine's Sean Lewis: The pay off, for me..." (the author)
            continue
        m = re.match(r"[ \t]*([A-Z]{2,4}):[ \t]+(?=[A-Z\"'“‘(\[])", p)
        if m and m.group(1).lower() in S.METADATA:
            rx = S._initials_rx(m.group(1))
            if rx and rx.search(whole) and S.SPEECH_MARK.search(p[m.end():]):
                # the record's full-name label with these initials, as stage 2 merges "AC" into "Andy Connell"
                same = [f for f in full if "".join(w[0] for w in f.split() if w[0].isupper()) == m.group(1)]
                out[i] = (same[0] if len(same) == 1 else m.group(1), m.end(), "colon")
                continue
        for lab in sorted(used, key=lambda x: (-len(x), x)):
            m = re.match(r"[ \t]*" + re.escape(lab) + r"[ \t]*;[ \t]+(?=\S)", p)
            if not m and re.fullmatch(r"[A-Z]{2,4}", lab) and used[lab] >= 2:
                m = re.match(r"[ \t]*" + re.escape(lab) + r"(?=[A-Z][a-z’']|I\b|A\b)", p)
            if not m and V3["label_variants"] and used[lab] >= 2:
                # version 3: a label the record uses two or more times, after stray characters or with another
                # separator: '¯¿¼MP-"Of course I like...', "CL– It's a story about an outsider.", "PWS. Yes."
                m = re.match(r"[^A-Za-z0-9\"“'‘(\[]{0,4}" + re.escape(lab) +
                             r"[ \t]*(?:[–—-]|:(?=\S)|\.(?=[ \t]))[ \t]*(?=[A-Z\"“'‘(\[])", p)
            if m and S.SPEECH_MARK.search(p[m.end():]):
                out[i] = (merged_of.get(lab, lab), m.end(), "colon")
                break


def interviewer(rec):
    return INTERVIEWER, rec._interviewer, "high", "the interviewer's question"


def _min_level(a, b):
    return LEVELS[min(LEVELS.index(a), LEVELS.index(b))] if a in LEVELS and b in LEVELS else b


def _qa_roles(rec, paras, labels):
    """Version 2. Who asks and who answers in a labelled Q&A, from its labels: the askers are the labels that are the
    interviewer's (the interviewer's own label, or what _journalists will decide after the build: a label naming the
    author or the outlet, or one whose lines are mostly questions); the answerers are the people the other labels
    name. Returns (askers, answerers): askers as (type, speaker), answerers as {(type, speaker): resolved label}."""
    if getattr(rec, "_roles", None) is not None:
        return rec._roles
    by = defaultdict(list)
    for i, (lab, off, style) in labels.items():
        if style != "sung" and any(a <= rec.pspans[i][0] < b for a, b in rec.blocks):
            by[lab].append((i, off))
    askers, answerers = set(), {}
    for lab, items in sorted(by.items()):
        r = rec.resolve_label(lab)
        if r[0] == INTERVIEWER:
            askers.add((r[0], r[1]))
            continue
        if r[0] == OTHER and r[3] in JOURNALIST_RULES:
            q = sum(1 for i, off in items if S.is_question(paras[i][off:]))
            why = journalist_label(rec, lab)
            if why and any(fold(lab) in fold(n) for n in rec.who["who"]):
                why = None
            if why or (len(items) >= 3 and q >= 0.5 * len(items)):
                askers.add((r[0], r[1]))
                continue
        if fold(lab) in ("a", "answer", "answers", "interviewee", "guest", "respondent"):
            r = rec.speaker_for_who()
        if r[0] in (MEMBER, OTHER_ERA, OTHER, MEMBERS_UNSPEC, OTHER_ERA_UNSPEC):
            answerers.setdefault((r[0], r[1]), r)
    rec._roles = (askers, answerers)
    return rec._roles


ANSWER_RULE = "an unlabelled answer after the interviewer's question"


def _answerer(rec, answerers, who_a):
    """Version 2. The speaker of an unlabelled answer in a labelled Q&A: the one person the other labels name; with no
    other labels, the interviewee; several members, not said which; several people, unknown."""
    if len(answerers) == 1:
        r = next(iter(answerers.values()))
        return r[0], r[1], _min_level(r[2], "medium"), ANSWER_RULE + "; the interview's other answers are this " \
                                                                       "speaker's"
    if not answerers:
        if who_a[0] in (UNKNOWN, AMBIGUOUS):
            return None
        return who_a[0], who_a[1], _min_level(who_a[2], "medium"), ANSWER_RULE + "; " + who_a[3]
    if all(t in (MEMBER, MEMBERS_UNSPEC) for t, n in answerers):
        names = sorted({x for t, n in answerers for x in (n or "").split(" / ") if x})
        return MEMBERS_UNSPEC, " / ".join(names) or None, "low", ANSWER_RULE + "; several members answer, not " \
                                                                              "said which"
    return UNKNOWN, None, "unknown", ANSWER_RULE + "; several people answer, not said which"


def _names_speaker(rec, p, who):
    """Does the paragraph name this speaker (a full name, the first name or the surname)? An answer seldom names its
    own speaker; a caption or the journalist's note does ("Jody Watley at the entrance to the Soul Train tunnel")."""
    if not who or not who[1]:
        return False
    lab = re.match(r"[^:\n]{1,60}:[ \t]", p)
    if lab:
        p = p[lab.end():]             # a label-like start ("Katrina from Katrina & The Waves: Yes, I think so.")
    for n in who[1].split(" / "):
        forms = set(rec.names.members[n]["forms"]) if n in rec.names.members else {n}
        toks = [t for t in _tokens(n) if t.casefold() not in SUFFIX]
        forms |= {t for t in ((toks[0], toks[-1]) if len(toks) >= 2 else ()) if len(t) >= 3}
        if any(re.search(r"(?<![\w'’])" + loose(f) + r"(?![\w'’])", p) for f in forms):
            return True
    return False


def _asks(p):
    """A question, or a prompt that waits for an answer ("The first card is the EMS Polysynthi…"); not a reaction or a
    teaser ending with an exclamation mark."""
    s = p.strip()
    return bool(S.is_question(p) or (S.interviewer_prompt(p) and s.endswith(("…", "..."))))


# how an answer to a statement (not a question) opens: "Absolutely, he always ends up...", "Definitely."
ANSWER_OPENER = re.compile(r"(?i)^[ \t]*[\"“]?(?:yes|yeah|yep|yup|no|nope|absolutely|definitely|exactly|sure|of course|"
                           r"certainly|totally|indeed|correct|thank you|thanks|i'm glad|i am glad|that's right|"
                           r"that's true|not really|not at all|initially|good question|great question|haha|ha ha)\b")


# ── version 3 (12 Oct 2026): a guest's words that a question mark or a missing label gave to the interviewer ──────
# Each rule can be switched off here (they are measured one by one).
V3 = {"goes_on": True, "requests": True, "inline_answer": True, "label_variants": True}
GENERIC_YOU = re.compile(r"(?i)\b(?:you know|y'know|ya know|you see|you guys|thank you|know what i mean)\b")
YOU_ANY = re.compile(r"(?i)\b(?:you|your|you're|you've|you'd|you'll|yours|yourself|yourselves|y'all)\b")
# the interviewer asking for an answer without a question mark ("Tell me a little about that song.", "I'd like to know
# how you decided to do it and where it was shot.")
REQUEST = re.compile(r"(?i)(?:^|[.!?…]\s+)(?:so,?\s+|and\s+|now,?\s+|ok(?:ay)?,?\s+|well,?\s+)?(?:please\s+)?"
                     r"(?:tell (?:me|us)|talk (?:to (?:me|us) )?about|describe|explain|walk (?:me|us) through|"
                     r"take (?:me|us) (?:through|back)|share with (?:me|us)|let'?s talk|"
                     r"i'?d (?:like|love) to (?:know|hear|ask)|i want(?:ed)? to (?:know|ask)|"
                     r"i (?:was going|wanted|meant) to ask you|let me ask you)\b")


def _last_sentence(p):
    parts = [x for x in re.split(r"(?<=[.!?…])[\"'”’)\]]*\s+", p.strip()) if x.strip()]
    return parts[-1] if parts else p


def _says_you(p):
    """The paragraph's last sentence says 'you' outside quotation marks ('you know' and the like do not count)."""
    return bool(YOU_ANY.search(GENERIC_YOU.sub(" ", S._unquote_para(_last_sentence(p)))))


def _quoted_question(p):
    """The paragraph ends with a question inside quotation marks that open after its start: the speaker reporting a
    question ('...I would just look at the Beatles on the wall and think, "How do I become part of this...?"')."""
    s = p.rstrip()
    m = re.search(r"\?([\"”'’])\s*$", s)
    if not m:
        return False
    opener = "\"“" if m.group(1) in "\"”" else "'‘"
    k = max(s.rfind(c, 0, m.start()) for c in opener)
    return k > 0 and bool(s[:k].strip())


def _guest_goes_on(rec, p, cur, nxt_p, nxt_labelled):
    """Version 3. In a Q&A whose answers carry the guest's label and whose questions carry none, a paragraph of the
    guest's answer that has a question mark ('...the start of the year, wasn't it?') is the guest going on, not the
    interviewer's next question, when the guest is speaking, the paragraph is in the first person, says 'you' to no
    one and does not name the guest (the writer's own comment between answers does: "...the appropriate amount of
    Prince talk in a Sheila E interview"), and either the question is one the guest reports in quotation marks or the
    next paragraph is more of the answer (no label, not a question). A real question is followed by the guest's
    label."""
    if not cur or cur[0] in (INTERVIEWER, UNKNOWN, AMBIGUOUS):
        return False
    body = S._unquote_para(p)
    if not S.FIRST_PERSON.search(body) or YOU_ANY.search(GENERIC_YOU.sub(" ", body)) or _names_speaker(rec, p, cur):
        return False
    if _quoted_question(p):
        return True
    return bool(nxt_p) and not nxt_labelled and not (S.is_question(nxt_p) or S.interviewer_statement(nxt_p) or
                                                     S.interviewer_prompt(nxt_p)) and \
        not heading_like(nxt_p) and not furniture_line(nxt_p) and not STAGE_DIRECTION.match(nxt_p)


def _invites(p):
    """Version 3. The interviewer's labelled line asks for an answer without a question mark: a request ('Tell me a
    little about that song.') or a short statement put to the guest ('By my count the only bands not represented on
    your new record are Family and Wishbone Ash.'). A long statement is more likely the host talking, whose next
    paragraph is more of the same."""
    s = p.strip()
    if not s or s.endswith("?") or len(s) > 600:
        return False
    if REQUEST.search(s):
        return True
    return len(s) <= 250 and s[-1] in ".…" and bool(YOU_ANY.search(GENERIC_YOU.sub(" ", S._unquote_para(s))))


def _split_turns(rec, bld, s, e, cuts, cur):
    """Cut s..e at label positions [(position, label, speech start)]: text before the first label continues the current
    speaker; each label starts a new turn. Returns the speaker after the last label."""
    cuts = sorted(c for c in cuts if s <= c[0] < e)
    pos = s
    for k, (o, lab, v) in enumerate(cuts):
        if o > pos:
            bld.add(pos, o, "turn", cur or (UNKNOWN, None, "unknown", "text before the first label"))
        nxt = cuts[k + 1][0] if k + 1 < len(cuts) else e
        cur = rec.resolve_label(lab)
        sung = bool(S.SUNG.match(rec.text[v:nxt]))
        bld.add(o, nxt, "sung" if sung else "turn", cur, label=lab, speech_start=min(max(v, o), nxt), new_turn=True)
        pos = nxt
    if pos < e:
        bld.add(pos, e, "turn", cur or (UNKNOWN, None, "unknown", "text before the first label"))
    return cur


def build_labelled(rec, a, b, bld, paras, labels, inline):
    detail = rec.detail
    cur = None
    who_a = rec.speaker_for_who()
    lec = S.lecture_transcript(rec.text, rec.title) if detail.startswith("lecture") else None
    pieces = _pieces(rec, a, b)
    order = [i for i, s, e in pieces if paras[i].strip()]
    nxt_of = {order[k]: order[k + 1] for k in range(len(order) - 1)}
    clip = False                      # inside a song clip ("(SOUNDBITE OF SONG, ...)") until the next label
    # version 2: who asks and who answers, so that an answer printed without its label is not left in the
    # interviewer's turn. Not in radio transcripts (every change of speaker carries a label there; a host's lines
    # run over several paragraphs) nor in formats with rules of their own.
    askers, answerers = _qa_roles(rec, paras, labels)

    def is_asker(r):
        return r is not None and (r[0] == INTERVIEWER or (r[0], r[1]) in askers)

    answer_rule = not detail.startswith(("answers labelled", "questions labelled", "lecture", "caption", "radio",
                                         "quotes labelled")) and not rec.hosts and \
        not any(style == "npr" for lab, off, style in labels.values())
    answerer = _answerer(rec, answerers, who_a) if answer_rule else None
    # version 3: the answerers' labels as the record writes them at the start of paragraphs, to find one that the page
    # ran into the interviewer's line ("JR: Is there a contagious factor? NJ: Not for my type.")
    ans_forms = {}                    # the label as written ("NJ") -> the label as read ("Naomi Judd")
    for j, (lab, off, style) in labels.items():
        if style in ("sung", "alone") or not any(a <= rec.pspans[j][0] < b for a, b in rec.blocks):
            continue
        r = rec.resolve_label(lab)
        if is_asker(r) or r[0] not in (MEMBER, MEMBERS_UNSPEC, OTHER_ERA, OTHER_ERA_UNSPEC):
            continue
        written = re.sub(r"[ \t]*[:;][ \t]*$", "", paras[j][:off]).strip()
        if 2 <= len(written) <= 40 and re.fullmatch(r"[A-Za-z][\w .'’-]*", written):
            ans_forms.setdefault(written, lab)
    inline_rx = re.compile(r"(?<=[.?!…\"”'’)\]])[ \t]+(" + "|".join(re.escape(x) for x in
                                                               sorted(ans_forms, key=lambda x: (-len(x), x))) +
                           r")[ \t]*:[ \t]+(?=\S)") if answer_rule and ans_forms and V3["inline_answer"] else None
    last_q = False                    # the last paragraph was a question or a prompt
    last_lab = False                  # the last paragraph carried the asker's label
    answered = False                  # someone other than the asker has answered with a label in this block
    last_asker = None                 # the asker's label, as resolved
    a1_on = False                     # the current turn is an answer given by the rule below
    for i, s, e in pieces:
        p = paras[i]
        ps = rec.pspans[i][0]
        if not p.strip() or i in rec.field_paras:
            bld.add(s, e, "non-speech")
            continue
        if STAGE_DIRECTION.match(p):
            bld.add(s, e, "non-speech")
            clip = bool(re.search(r"(?i)soundbite of (?:song|music)|^\s*[(\[](?:music|singing|song)\b", p))
            continue
        if i in labels or i in inline:
            clip = False
        elif clip:
            bld.add(s, e, "sung")
            continue
        if lec:
            host, guest, offs = lec
            cuts = []
            for o in offs:
                if s <= o < e:
                    name = guest if re.match(re.escape(guest), rec.text[o:], re.I) else host
                    cuts.append((o, name, o + len(name)))
            cur = _split_turns(rec, bld, s, e, cuts, cur)
            continue
        if i in inline:
            cuts = []
            for lab, off in inline[i]:
                o = ps + off
                m = re.match(r"[^:\n]{1,60}:\s*", rec.text[o:e])
                cuts.append((o, lab, o + (m.end() if m else 0)))
            if i in labels:
                lab, off, style = labels[i]
                cuts.append((s, lab, max(s, ps + off)))
            cur = _split_turns(rec, bld, s, e, cuts, cur)
            continue
        if i in labels:
            lab, off, style = labels[i]
            cur = rec.resolve_label(lab)
            if style == "alone":
                bld.add(s, e, "label", cur, label=lab, new_turn=True)
                continue
            v = max(s, ps + off)
            sung = bool(S.SUNG.match(rec.text[v:e]))
            m_in = inline_rx.search(rec.text[v:e]) if inline_rx and not sung and is_asker(cur) else None
            if m_in:
                # version 3: the guest's label run into the interviewer's line starts the guest's answer
                cut = v + m_in.start(1)
                cur = _split_turns(rec, bld, s, e, [(s, lab, v), (cut, ans_forms[m_in.group(1)], v + m_in.end())],
                                   cur)
                last_q, last_lab, answered, a1_on = False, False, True, False
                last_asker = rec.resolve_label(lab)
                continue
            bld.add(s, e, "sung" if sung else "turn", cur, label=lab, speech_start=v, new_turn=True)
            last_q = _asks(p[off:]) or bool(V3["requests"] and is_asker(cur) and _invites(p[off:]))
            last_lab = is_asker(cur)
            last_asker = cur if is_asker(cur) else last_asker
            answered = answered or (not is_asker(cur) and not sung)
            a1_on = False
            continue
        # a paragraph without a label
        if furniture_line(p):
            bld.add(s, e, "non-speech")
            continue
        q = S.is_question(p) or S.interviewer_statement(p) or S.interviewer_prompt(p)
        prev_q, last_q = last_q, _asks(p)
        prev_lab, last_lab = last_lab, False
        if detail.startswith("answers labelled") and q:
            nj = nxt_of.get(i)
            if V3["goes_on"] and _guest_goes_on(rec, p, cur, paras[nj] if nj is not None else None,
                                                nj is not None and (nj in labels or nj in inline)):
                bld.add(s, e, "turn", cur)          # version 3: the guest going on (see _guest_goes_on)
                last_q = False
                continue
            cur = interviewer(rec)
            bld.add(s, e, "turn", cur, new_turn=True, question=True)
            continue
        # an unlabelled question between labelled answers ("Neil: ...", question, "Finlay: ...") is the interviewer's.
        # Version 2: not when the next label is the interviewer's own ("Songfacts:" asking after the guest asked
        # "Do you know what The Onion is?"); version 1 could not tell before the build.
        if q and S.is_question(p) and cur and cur[0] != INTERVIEWER and nxt_of.get(i) in labels and \
                not is_asker(rec.resolve_label(labels[nxt_of[i]][0])):
            cur = interviewer(rec)
            bld.add(s, e, "turn", cur, new_turn=True, question=True)
            continue
        if detail.startswith("questions labelled") and cur and cur[0] == INTERVIEWER:
            cur = who_a if who_a[0] != UNKNOWN else (UNKNOWN, None, "unknown", "answer to a labelled question; "
                                                     "no interviewee named")
            bld.add(s, e, "turn", cur, new_turn=True)
            continue
        # version 2: inside an answer printed without its label, a follow-up question printed without its label is
        # the interviewer's again ("Was it surreal getting one of your favourite artists to collaborate with you?")
        if a1_on and q and S.is_question(p) and last_asker and not _names_speaker(rec, p, last_asker) and \
                S.SECOND_PERSON.search(p):
            cur = last_asker
            bld.add(s, e, "turn", cur, new_turn=True, question=True)
            a1_on = False
            continue
        # version 2: an answer printed without its label after the interviewer's question (or after a statement,
        # when it opens like an answer: "Absolutely, ...") is the answerer's, not the interviewer's
        if answerer and is_asker(cur) and not q and (prev_lab or answered) and \
                (prev_q or (prev_lab and ANSWER_OPENER.match(p))) and not heading_like(p) and \
                (S.SPEECH_MARK.search(p) or ANSWER_OPENER.match(p)) and not _names_speaker(rec, p, answerer):
            cur = answerer
            bld.add(s, e, "turn", cur, new_turn=True)
            a1_on = True
            continue
        if heading_like(p) and not (cur and cur[0] in BAND_VOICE | {OTHER_ERA, OTHER} and len(p.split()) > 6):
            bld.add(s, e, "heading")
            continue
        if cur is None:
            if not q and not S.fp_outside_quotes(p):
                bld.add(s, e, "narration", (OTHER, rec.author or rec.outlet or None, "medium",
                                            "the journalist's words before the first label"))
                continue
            cur = interviewer(rec) if q else (UNKNOWN, None, "unknown", "text before the first label")
            bld.add(s, e, "turn", cur, new_turn=True, question=q)
            continue
        bld.add(s, e, "turn", cur, question=bool(cur[0] == INTERVIEWER and q))


# caption transcripts (version 2): a dash at the start of a line, or after an interruption ("tour buses and-- - So
# say something", "So, in your mind- - I was like"), the end of a sentence ("...a series... - Yeah.") or a stage
# direction ("(audience laughing) - Well, actually"), starts a new speaker
CAPTION_MIDLINE = re.compile(r"(?:-|\.\.\.|…|[)\].?!,])(?=[ \t]+-[ \t]+\S)")
CAPTION_LABEL = re.compile(r"[ \t]*(?:-[ \t]*)?\[(?P<lab>[A-Z][\w .'’-]{0,30})\][ \t]*")
CAPTION_DASH = re.compile(r"[ \t]*(?:-[ \t]*)?")
HOST_INTRO = re.compile(r"\bI'?m ([A-Z][a-z]+(?: [A-Z][a-z]+)?)\b")
# the video page's own lines after the transcript: the programme is over
CAPTION_END = re.compile(r"[ \t]*-?[ \t]*(?:New Episode\b|News and Public Affairs|Video has Closed Captions|"
                         r"Clip: S\d+ Ep\d+|Providing Support for PBS|Support for PBS provided by)")


def _caption_units(rec, a, b, paras):
    """(start, end) of each speaker's turn; the page's own lines after the transcript are one last unit (None)."""
    cuts, end = {a}, b
    for i, s, e in _pieces(rec, a, b):
        p, ps = paras[i], rec.pspans[i][0]
        if s == ps and CAPTION_END.match(p):
            end = s
            break
        if s == ps and re.match(r"[ \t]*-[ \t]+\S", p):
            cuts.add(s)
        for m in CAPTION_MIDLINE.finditer(p):
            if s < ps + m.end() < e:
                cuts.add(ps + m.end())
    cuts = sorted(c for c in cuts if a <= c < end)
    units = [(c, d) for c, d in zip(cuts, cuts[1:] + [end]) if d > c]
    return units + ([(end, b, None)] if end < b else [])


FP_WORDS = re.compile(r"\b(?:I|I'm|I've|I'd|I'll|my|me|mine|myself)\b")
# "you" said to the other person ("you brought them over with you"), not "you know" or "thank you"
YOU_WORDS = re.compile(r"(?i)(?<!thank )\byou(?:r|'re|'ve|'d|'ll|rself)?\b(?! know\b| guys\b)")


def _strip_directions(s):
    """A turn's words without stage directions at its end ("...do another one? (audience laughing)")."""
    return re.sub(r"(?:\s*\((?:[^()\n]{1,40})\))+\s*$", "", s.strip()).strip()


def _caption_evidence(body, host_first, guest_names):
    """(for the host or the audience, for the guest): the evidence in one caption turn."""
    s = _strip_directions(body)
    h = g = 0.0
    if s.endswith("?"):
        h += 1.5
    you = len(YOU_WORDS.findall(s))
    fp = len(FP_WORDS.findall(s))
    if you >= 2 and you > fp:
        h += 1.0
    elif you == 1 and fp == 0:
        h += 0.5
    if fp >= 2 and fp > you and len(s.split()) >= 8:
        g += 1.0
    if host_first and re.search(r"\b" + re.escape(host_first) + r"\b", s):
        g += 1.5
    if any(re.search(r"\b" + re.escape(x) + r"\b", s) for x in guest_names):
        h += 1.5
    return h, g, s.endswith("?"), fp, len(s.split())


def _decode(em, trans, prior):
    """Viterbi over two roles (0 the host or the audience, 1 the guest): em[k] the evidence for each role, trans[t][s]
    the score of moving from role t to role s at a dash, prior the score of each role for the first turn."""
    if not em:
        return []
    score = [prior[0] + em[0][0], prior[1] + em[0][1]]
    back = []
    for k in range(1, len(em)):
        new, bk = [], []
        for s in (0, 1):
            val, t = max(((score[t] + trans[t][s], t) for t in (0, 1)), key=lambda z: (z[0], -z[1]))
            new.append(val + em[k][s])
            bk.append(t)
        score = new
        back.append(bk)
    s = 0 if score[0] >= score[1] else 1
    roles = [s]
    for bk in reversed(back):
        s = bk[s]
        roles.append(s)
    return roles[::-1]


def _alternation(items, host_first, guest_names, stay=3.0):
    """An interview: the most likely host/guest sequence (0 host, 1 guest). Each dash marks a new speaker, so the two
    alternate; keeping the same speaker across a dash costs `stay`. The evidence for each turn: a name in brackets
    (decisive), a question and "you" said to the other (the host), the guest's name (the host), the host's first name
    (the guest), several words in the first person (the guest). Returns the roles and whether each turn kept the
    speaker of the turn before it."""
    em = []
    for body, lr in items:
        h, g = _caption_evidence(body, host_first, guest_names)[:2]
        em.append((h + (50.0 if lr == "host" else 0.0), g + (50.0 if lr == "guest" else 0.0)))
    roles = _decode(em, [[-stay, 0.0], [0.0, -stay]], (1.0, 0.0))     # the host opens the programme
    return roles, [False] + [roles[k] == roles[k - 1] for k in range(1, len(roles))]


def _audience_turns(items, host_first, guest_names):
    """A Q&A with an audience: which turns are the guest's (1) and which are not (0), kept to what the format shows.
    The guest answers: the first turn of four words or more after a question (short interjections such as "Great
    question." do not count), and the longer turns in the first person that go on from the guest's answer across the
    host's or the audience's short interjections ("Yep.", "Wow."). A turn addressed to the guest ("you" more than "I")
    or naming the guest is not the guest's; nor is a question. Everything else is left to the host or the audience."""
    out, pending, in_answer = [], False, False
    for body, lr in items:
        h, g, is_q, fp, n = _caption_evidence(body, host_first, guest_names)
        s = _strip_directions(body)
        you = len(YOU_WORDS.findall(s))
        short = n <= 3
        if lr == "guest":
            out.append(1)
            pending, in_answer = False, n >= 6
            continue
        if lr in ("host", "other"):
            out.append(0)
            if is_q:
                pending, in_answer = True, False
            elif not short and n > 5:
                in_answer = False
            continue
        named = any(re.search(r"\b" + re.escape(x) + r"\b", s) for x in guest_names)
        if is_q:
            out.append(0)
            pending, in_answer = True, False
        elif short:
            out.append(0)
        elif pending and not named and you <= max(fp, 1):
            out.append(1)
            pending, in_answer = False, True
        elif in_answer and n >= 15 and fp >= 1 and you <= fp and not named:
            out.append(1)
        else:
            out.append(0)
            pending, in_answer = False, False
    return out


def build_caption(rec, a, b, bld, paras, labels):
    """Version 2: caption transcripts, where a dash starts each new speaker and only a few turns carry a name
    ("- [Evan] Still love it?"). In an interview the speakers alternate at each dash; the most likely sequence is
    worked out over the whole programme from the names in brackets, the host's own introduction ("I'm Evan Smith"),
    the questions and the names the two use for each other, so that a missing dash does not swap the speakers for the
    rest of the programme. In a Q&A with an audience (the title says Q&A), a turn without a name is the guest's only
    when it answers a question; questions are the audience's; other turns without a name stay unknown. The video
    page's own lines after the programme are page furniture."""
    text = rec.text
    who_a = rec.speaker_for_who()
    guest = who_a if who_a[0] not in (UNKNOWN, AMBIGUOUS) else None
    audience = bool(re.search(r"(?i)\bq\s?&\s?a\b", rec.title))
    units = []
    for unit in _caption_units(rec, a, b, paras):
        u0, u1 = unit[:2]
        ut = text[u0:u1]
        m = CAPTION_LABEL.match(ut)
        lab = m.group("lab").strip() if m else None
        v = u0 + (m.end() if m else CAPTION_DASH.match(ut).end())
        body = text[v:u1].strip()
        speech = bool(ut.strip()) and len(unit) == 2 and bool(body) and not STAGE_DIRECTION.match(body)
        units.append((u0, u1, lab, v, body, speech))
    # the host introduces himself in his first turns ("I'm Evan Smith"); the guest's names
    host_name = None
    for u0, u1, lab, v, body, speech in [x for x in units if x[5] and not x[2]][:3]:
        hm = HOST_INTRO.search(body)
        if hm and not rec.names.mentions(hm.group(1)):
            host_name = hm.group(1)
            break
    host_first = host_name.split()[0] if host_name else None
    gnames = []
    if guest and guest[1]:
        for n in guest[1].split(" / "):
            gnames += [n] + ([n.split()[0]] if len(n.split()) >= 2 and len(n.split()[0]) >= 3 else [])
    host = (INTERVIEWER, host_name or rec._interviewer, "medium",
            "caption transcript: the host (the speakers alternate at each dash)")

    def label_role(lab):
        res = rec.resolve_label(lab)
        if fold(lab) in ("narrator", "announcer"):
            return None, res
        if res[0] == INTERVIEWER or (host_first and fold(lab.split()[0]) == fold(host_first)):
            return "host", (INTERVIEWER, host_name or res[1], "high", "caption transcript: the host's name in "
                                                                       "brackets")
        if guest and res[1] == guest[1]:
            return "guest", res
        return "other", res

    who_of = {}
    if not audience:
        chain = []
        for k, (u0, u1, lab, v, body, speech) in enumerate(units):
            if not speech:
                continue
            lr, res = label_role(lab) if lab else ("", None)
            if lr in (None, "other"):
                who_of[k] = res               # the announcer, "[Both]", "[Audience Member]": outside the alternation
                continue
            chain.append((k, body, lr or None, res))
        roles, stayed = _alternation([(body, lr) for k, body, lr, res in chain], host_first, gnames)
        for (k, body, lr, res), r, st in zip(chain, roles, stayed):
            if lr:
                who_of[k] = res
            elif r == 0:
                who_of[k] = host if not st else (host[0], host[1], "low", host[3] + "; two turns in a row")
            elif guest:
                conf = _min_level(guest[2], "medium") if not st else "low"
                who_of[k] = (guest[0], guest[1], conf, "caption transcript: the guest (the speakers alternate at each "
                                                       "dash)" + ("; two turns in a row" if st else ""))
            else:
                who_of[k] = (UNKNOWN, None, "unknown", "caption transcript: the guest, who is not named")
    else:
        chain = []
        for k, (u0, u1, lab, v, body, speech) in enumerate(units):
            if not speech:
                continue
            lr, res = label_role(lab) if lab else ("", None)
            if lr is None:
                who_of[k] = res               # the announcer: outside the exchange
                continue
            if lr == "other" and res and res[0] in BAND_VOICE | {OTHER_ERA}:
                lr = "guest"                  # "[Emily]" or "[Amy]" of a duo answering together
            chain.append((k, body, lr or None, res))
        roles = _audience_turns([(body, lr) for k, body, lr, res in chain], host_first, gnames)
        for (k, body, lr, res), r in zip(chain, roles):
            if lr:
                who_of[k] = res
            elif r == 1 and guest:
                who_of[k] = (guest[0], guest[1], "low", "caption transcript: the guest's answer to a question in a Q&A "
                                                        "with the audience")
            elif r == 1:
                who_of[k] = (UNKNOWN, None, "unknown", "caption transcript: the guest, who is not named")
            elif _caption_evidence(body, host_first, gnames)[2]:
                who_of[k] = (OTHER, None, "low", "caption transcript: a question from the audience or the host")
            else:
                who_of[k] = (UNKNOWN, None, "unknown", "caption transcript: the host or the audience (not named)")
    for k, (u0, u1, lab, v, body, speech) in enumerate(units):
        if k not in who_of:
            bld.add(u0, u1, "non-speech")
            continue
        who = who_of[k]
        is_q = body.rstrip().endswith("?")
        bld.add(u0, u1, "turn", who, label=lab, speech_start=v, new_turn=True,
                question=bool(is_q and who[0] == INTERVIEWER))


def build_unlabelled(rec, a, b, bld, paras, labels):
    units = S.qa_units(paras, min_answer=8, prompts=True)
    starts = {u[0]: u for u in units}
    answer_of = {}
    for u in units:
        for j in range(u[0] + 1, u[1] + 1):
            answer_of.setdefault(j, u)
    runs = dict(S.run_in_pairs(paras))
    who_a = rec.speaker_for_who()
    spans_all = rec.pspans
    # a member's label inside a paragraph, after the question ("What is your worst lyric? Corinne: Why would I...")
    inl = defaultdict(list)
    for i, lab, val, style, off in S.inline_labels(paras):
        if rec.resolve_label(lab)[0] in (MEMBER, OTHER_ERA, MEMBERS_UNSPEC):
            inl[i].append((lab, off))
    cur = None
    for i, s, e in _pieces(rec, a, b):
        p = paras[i]
        ps = spans_all[i][0]
        if not p.strip() or i in rec.field_paras:
            bld.add(s, e, "non-speech")
            continue
        if i in inl and s <= ps + min(o for _, o in inl[i]) < e:
            first = ps + min(o for _, o in inl[i])
            head = rec.text[s:first]
            if i in labels:
                lab, off, style = labels[i]
                lead = rec.resolve_label(lab)
                if lead[0] == OTHER and head.rstrip().endswith("?"):
                    lead = (INTERVIEWER, lead[1], "high", "a label on the interviewer's question")
            elif head.rstrip().endswith("?"):
                lead = interviewer(rec)
            else:
                lead = cur or who_a
            bld.add(s, first, "turn", lead, new_turn=True, question=lead[0] == INTERVIEWER)
            cuts = []
            for lab, off in inl[i]:
                o = ps + off
                m = re.match(r"[^:\n]{1,60}:\s*", rec.text[o:e])
                cuts.append((o, lab, o + (m.end() if m else 0)))
            cur = _split_turns(rec, bld, first, e, cuts, lead)
            continue
        if STAGE_DIRECTION.match(p) or furniture_line(p) and i not in starts:
            bld.add(s, e, "non-speech")
            continue
        if i in labels:
            lab, off, style = labels[i]
            res = rec.resolve_label(lab)
            if res[0] == OTHER and res[3].startswith("a label naming nobody") and i in starts:
                res = (INTERVIEWER, res[1], "high", "a label on the interviewer's question")
            if style == "alone":
                bld.add(s, e, "label", res, label=lab, new_turn=True)
                cur = res
                continue
            v = max(s, ps + off)
            cur = res
            bld.add(s, e, "turn", res, label=lab, speech_start=v, new_turn=True, question=res[0] == INTERVIEWER)
            continue
        if i in starts:
            if i in runs and s <= ps + runs[i] < e:
                cut = ps + runs[i]
                bld.add(s, cut, "turn", interviewer(rec), new_turn=True, question=True)
                cur = who_a
                bld.add(cut, e, "turn", who_a, new_turn=True)
            else:
                cur = interviewer(rec)
                bld.add(s, e, "turn", cur, new_turn=True, question=True)
            continue
        if i in answer_of:
            if cur is None or cur[0] == INTERVIEWER or (cur[0] != who_a[0] or cur[1] != who_a[1]):
                if not (cur and cur[3] and "label" in (cur[3] or "") and cur[0] != INTERVIEWER):
                    cur = who_a
                    bld.add(s, e, "turn", cur, new_turn=True)
                    continue
            bld.add(s, e, "turn", cur)
            continue
        if heading_like(p):
            bld.add(s, e, "heading")
            continue
        bld.add(s, e, "turn", cur or who_a)


def _section_heads(rec, paras):
    """First-person accounts with a heading per speaker: [(paragraph, name)]."""
    bare_title = re.sub(r"\"[^\"]*\"|(?:(?<=^)|(?<=[\s(]))'[^']*'(?=[\s:,.;!?)]|$)", " ", rec.title)
    cue = bool(S.TITLE_SECTIONS.search(bare_title))
    roles, bare, topics = S.headed_sections(paras, rec.title, single_names=cue)
    return sorted(roles + bare)


def resolve_name(rec, name):
    """A person named in a heading: a member of the act, or someone else."""
    m = rec.names.mentions(name)
    if m:
        n = min(m, key=m.get)
        return rec.member_type(n), n, "medium", "the heading above the section names this member"
    toks = name.split()
    if len(toks) == 1:
        low = fold(name)
        c = rec.names.first.get(low, set()) | rec.names.surname.get(low, set())
        c = [x for x in c if x in rec.full]
        if len(c) == 1:
            return rec.member_type(c[0]), c[0], "medium", "the heading names this member by one name"
    return OTHER, name, "medium", "the heading above the section names this person"


PROMPT_PHRASE = re.compile(r"(?i)\b(?:tell (?:me|us)|talk (?:to us |to me )?about|can you|could you|would you|"
                           r"what was it like|how did you|why did you|do you|did you|have you|were you)\b")
INTERVIEW_PAGE = re.compile(r"(?i)\binterview|\bq\s?&\s?a\b|/interviews?[-/]")


def build_first_person(rec, a, b, bld, paras, labels):
    detail = rec.detail
    # an interview printed as one account: its questions to "you" are the interviewer's
    interview_page = bool(INTERVIEW_PAGE.search(rec.title) or INTERVIEW_PAGE.search(rec.url))
    heads = dict(_section_heads(rec, paras)) if detail.startswith("sections headed") else {}
    narrator = rec.narrator()
    cur = narrator if not heads else None
    for i, s, e in _pieces(rec, a, b):
        p = paras[i]
        if not p.strip() or i in rec.field_paras:
            bld.add(s, e, "non-speech")
            continue
        if i in heads:
            cur = resolve_name(rec, heads[i])
            bld.add(s, e, "heading", cur, label=heads[i], new_turn=True)
            continue
        if STAGE_DIRECTION.match(p) or furniture_line(p):
            bld.add(s, e, "non-speech")
            continue
        if i in labels and labels[i][2] != "alone":
            lab, off, style = labels[i]
            res = rec.resolve_label(lab)
            if res[0] in (OTHER, UNKNOWN) and (S.is_question(p) or S.SECOND_PERSON.search(p)):
                res = (INTERVIEWER, rec._interviewer, "high", "a labelled question inside the account")
            elif res[0] not in (INTERVIEWER, MEMBER, OTHER_ERA, MEMBERS_UNSPEC):
                if heading_like(p) or len(p.strip()) <= 100:
                    bld.add(s, e, "heading")
                else:
                    bld.add(s, e, "turn", cur or narrator)
                continue
            ps = rec.pspans[i][0]
            bld.add(s, e, "turn", res, label=lab, speech_start=max(s, ps + off), new_turn=True,
                    question=res[0] == INTERVIEWER)
            continue
        if heading_like(p) and not S.fp_outside_quotes(p, None, True):
            bld.add(s, e, "heading")
            continue
        if interview_page and (S.is_question(p) or S.interviewer_statement(p) or S.interviewer_prompt(p)) and \
                S.SECOND_PERSON.search(p) and len(p.strip()) <= 400 and \
                (not S.fp_outside_quotes(p, None, True) or PROMPT_PHRASE.search(p)):
            bld.add(s, e, "turn", interviewer(rec), new_turn=True, question=True)
            continue
        if cur is None:
            bld.add(s, e, "narration", (OTHER, rec.author or rec.outlet or None, "medium",
                                        "the journalist's words before the first section"))
            continue
        if heads and not S.fp_outside_quotes(p) and cur[1] and \
                (re.search(r"\b(?:he|she|his|her)\b", p, re.I) or S._mentions(cur[1], p, rec.title)):
            # a short third-person introduction under the heading
            bld.add(s, e, "narration", (OTHER, rec.author or rec.outlet or None, "medium",
                                        "a third-person line about the section's speaker"))
            continue
        bld.add(s, e, "turn", cur)


# outlets' usual initials in transcripts ("RS:", "RC:", "ESQ:")
OUTLET_INITIALS = {"rs": "rollingstone", "rc": "recordcollector", "esq": "esquire", "sde": "superdeluxeedition",
                   "bb": "billboard", "ew": "ew.com", "nme": "nme", "ucr": "ultimateclassicrock", "cr": "loudersound",
                   "pkm": "pleasekillme", "im": "international musician", "mt": "music technology",
                   "sos": "sound on sound", "emm": "electronics & music maker", "ott": "one two testing"}
JOURNALIST_RULES = ("a label naming nobody on the line-up",)


def journalist_label(rec, lab):
    """A label that names the author or the outlet ("Gregg Shapiro:", "RATH:", "DK:" for Dale Kawashima, "ESQ:")."""
    low = fold(lab).strip(" .:")
    words = {w for w in re.findall(r"[a-z0-9'’-]+", low) if len(w) > 1}
    aw = _author_words(rec.author_raw)
    if words and aw and words <= aw:
        return "the label names the author"
    auth = rec.author or ""
    toks = auth.split()
    if len(toks) >= 2 and low.replace(".", "").replace(" ", "") == "".join(t[0] for t in toks).lower():
        return "the label is the author's initials"
    outlet = re.sub(r"^www\.|\.(?:com|co\.uk|org|net|pub|co|uk|fm)(?:/.*)?$", "", rec.outlet.lower())
    compact = re.sub(r"[^a-z0-9]", "", low)
    if len(compact) >= 3 and compact == re.sub(r"[^a-z0-9]", "", outlet):
        return "the label names the outlet"
    where = (rec.outlet + " " + rec.title).lower()
    if compact in OUTLET_INITIALS and (OUTLET_INITIALS[compact] in where or
                                       OUTLET_INITIALS[compact].replace(" ", "") in where.replace(" ", "")):
        return "the label is the outlet's initials"
    return None


WHO_RULES = ("interviewee named", "answer label", "several members named", "the act is named", "the act in the title",
             "the member the introduction is about", "named in the title", "narrator named")


def _third_person(rec, spans):
    """An answer given to the interviewee from the title or introduction that speaks of them in the third person and
    never in the first is the journalist's narration ("Neneh's first single was..."), not their words."""
    for x in spans:
        if x["kind"] != "turn" or x["speaker_type"] not in (MEMBER, OTHER_ERA, MEMBERS_UNSPEC) or \
                not any(r in (x["rule"] or "") for r in WHO_RULES) or \
                (x["label"] and fold(x["label"]) not in ("a", "answer", "answers", "interviewee", "guest")):
            continue
        words = rec.text[x["speech_start"]:x["end"]]
        # any first person, even inside quotation marks (answers are often printed whole in quotation marks)
        if len(words.strip()) < 80 or S.FIRST_PERSON.search(words) or words.lstrip()[:1] in "\"“'‘":
            continue
        plain = words.strip()
        names = [x["member"]] if x["member"] else [n for n in rec.names.members if rec.names.era(n) == "Nomination-era"]
        hit = False
        for n in names:
            toks = [t for t in _tokens(n) if t.casefold() not in SUFFIX]
            forms = set(rec.names.members[n]["forms"]) | ({toks[-1]} if len(toks) >= 2 and len(toks[-1]) >= 4 else set()) | \
                set(_short_names(rec.names, n))
            # a name at the very start is an unmarked label ("Slim Jim Phantom No, it seems too hard...")
            body = plain
            for f in sorted(forms, key=len, reverse=True):
                m0 = re.match(loose(f) + r"\b", body)
                if m0:
                    body = body[m0.end():]
                    break
            if any(re.search(r"(?<![\w'’])" + loose(f) + r"(?![\w])", body) for f in forms):
                hit = True
                break
        if hit:
            x["speaker_type"], x["speaker"], x["member"], x["era"] = UNKNOWN, None, None, None
            x["band_voice"], x["confidence"] = False, "unknown"
            x["rule"] = "the answer speaks of the interviewee in the third person"
    return spans


def _journalists(rec, spans):
    """In labelled Q&As, a label naming no member whose lines are mostly questions, or that names the author or the
    outlet, is the interviewer's."""
    by = defaultdict(list)
    for x in spans:
        if x["kind"] == "turn" and x["speaker_type"] == OTHER and x["rule"] in JOURNALIST_RULES:
            by[x["speaker"]].append(x)
    # "A:" lines that are mostly questions are the interviewer's (Arthur magazine's "A:" against "CH:")
    gen = [x for x in spans if x["kind"] == "turn" and x["label"] and fold(x["label"]) in ("a", "answer")]
    if len(gen) >= 3 and sum(1 for x in gen if S.is_question(rec.text[x["speech_start"]:x["end"]].split("\n")[0])) \
            >= 0.5 * len(gen):
        for x in gen:
            x["speaker_type"], x["speaker"], x["member"], x["era"] = INTERVIEWER, rec._interviewer, None, None
            x["band_voice"], x["confidence"], x["rule"] = False, "medium", "the label's lines are mostly questions"
            x["question"] = S.is_question(rec.text[x["speech_start"]:x["end"]])
    for spk, xs in by.items():
        lab = next((x["label"] for x in xs if x["label"]), None)
        if not lab:
            continue
        labelled = [x for x in xs if x["label"]]
        q = sum(1 for x in labelled if S.is_question(rec.text[x["speech_start"]:x["end"]].split("\n")[0]))
        why = journalist_label(rec, lab)
        if why and lab and any(fold(lab) in fold(n) for n in rec.who["who"]):
            why = None                         # the label names the interviewee ("Jim Salestrom" for an oral history)
        if not why and len(labelled) >= 3 and q >= 0.5 * len(labelled):
            why = "the label's lines are mostly questions"
        if why:
            for x in xs:
                x["speaker_type"], x["rule"], x["confidence"] = INTERVIEWER, why, "high" if "author" in why else "medium"
                x["question"] = S.is_question(rec.text[x["speech_start"]:x["end"]])
    return spans


def speakers(rec):
    """All spans for the record's blocks, in order."""
    paras = rec.text.split("\n")
    labels = _label_index(rec, paras)
    inline = defaultdict(list)
    if "labels inside paragraphs" in rec.detail:
        for i, lab, val, style, off in S.inline_labels(paras):
            inline[i].append((lab, off))
    out = []
    for a, b in rec.blocks:
        bld = Builder(rec)
        if rec.route == "labelled Q&A" and rec.detail.startswith("caption"):
            build_caption(rec, a, b, bld, paras, labels)
        elif rec.route == "labelled Q&A":
            build_labelled(rec, a, b, bld, paras, labels, inline)
        elif rec.route == "unlabelled Q&A":
            build_unlabelled(rec, a, b, bld, paras, labels)
        else:
            build_first_person(rec, a, b, bld, paras, labels)
        spans = bld.finish()
        # the spans must cover the block exactly
        pos = a
        for x in spans:
            assert x["start"] == pos and x["end"] > x["start"], (rec.rid, a, b, x["start"], pos)
            pos = x["end"]
        assert pos == b, (rec.rid, pos, b)
        out += spans
    if rec.route == "labelled Q&A":
        out = _journalists(rec, out)
    return _third_person(rec, out)
'''
SOURCES["stage4"] = r'''"""Stage 4 of the turn build: the quotations in the prose (technical appendix, Step 2, stage 4).

The prose segments of every record (prose records, and the prose introductions and tails of Q&A and first-person
records) are cut into spans that cover them exactly once:
    speech        a quotation of someone's spoken words
    non-speech    a title or a scare quote in quotation marks ("Walk of Life", the band's "comeback")
    undecided     a quotation that is neither clearly speech nor clearly a title
    attribution   the clause that says who spoke ("she says", "Knopfler told Rolling Stone:")
    unresolved    a paragraph whose quotation marks do not balance (kept out of both layers)
    narration     everything else: the journalist's own words
Speech marks follow each record's convention: double quotes (American style) or single quotes (British style).
Stage 5 gives each speech span its speaker; this stage only finds and classifies the quotations.
"""
import re

# ── speech verbs ─────────────────────────────────────────────────────────────
SPEECH_VERBS = ("say says said saying tell tells told add adds added explain explains explained recall recalls recalled "
                "admit admits admitted laugh laughs laughed continue continues continued note notes noted insist "
                "insists insisted remember remembers remembered reflect reflects reflected joke jokes joked agree agrees "
                "agreed ask asks asked reply replies replied confess confesses confessed muse muses mused observe "
                "observes observed reveal reveals revealed offer offers offered concede concedes conceded quip quips "
                "quipped declare declares declared claim claims claimed state states stated argue argues argued wrote "
                "writes write sigh sighs sighed enthuse enthuses enthused gush gushes gushed beam beams beamed smile "
                "smiles smiled grin grins grinned chuckle chuckles chuckled shrug shrugs shrugged warn warns warned "
                "predict predicts predicted promise promises promised acknowledge acknowledges acknowledged recount "
                "recounts recounted remark remarks remarked comment comments commented stress stresses stressed "
                "emphasise emphasises emphasised emphasize emphasizes emphasized suggest suggests suggested maintain "
                "maintains maintained boast boasts boasted complain complains complained reminisce reminisces "
                "reminisced marvel marvels marvelled marveled exclaim exclaims exclaimed shout shouts shouted yell "
                "yells yelled whisper whispers whispered cry cries cried tease teases teased deadpan deadpans "
                "deadpanned counter counters countered respond responds responded answer answers answered post posts "
                "posted tweet tweets tweeted describe describes described announce announces announced confide "
                "confides confided concludes concluded conclude opines opined opine wonders "
                "wondered wonder jokes ponders pondered ponder grumbles grumbled grumble adds explains reasons reasoned "
                "reason elaborates elaborated elaborate clarifies clarified clarify acknowledges furthers furthered "
                "reports reported asserts asserted snorts snorted chirps chirped interjects interjected retorts retorted "
                "ventures ventured volunteers volunteered posits posited surmises surmised reckons reckoned confirms "
                "confirmed reminds reminded cautions cautioned urges urged hints hinted vows vowed crows crowed raves "
                "raved purrs purred drawls drawled murmurs murmured mutters muttered giggles giggled sniffs sniffed "
                "scoffs scoffed snaps snapped barks barked groans groaned moans moaned rants ranted protests protested "
                "concurs concurred figures figured guesses guessed recollects recollected wisecracks wisecracked "
                "summarises summarised summarizes summarized proclaims proclaimed brags bragged laments lamented "
                "screams screamed emails emailed").split()
SV = r"(?:" + "|".join(sorted(set(SPEECH_VERBS), key=len, reverse=True)) + r")"
SV_PHRASE = (r"(?:" + SV + r"(?:\s+(?:of|about|on|in|to|with|at)\b)?|put(?:s)? it|puts? it|went on|goes on|points? out|pointed "
             r"out|sums? (?:it )?up(?![ \t]+(?:the|a|an|his|her|their|its|this|that|what|how|why)\b)|summed (?:it )?up"
             r"(?![ \t]+(?:the|a|an|his|her|their|its|this|that|what|how|why)\b)|chimes? in|chimed in|cuts? in|"
             r"jumps? in|jumped in|"
             r"according to)")

# who can speak: a pronoun, a name (one to four capitalised words), or a description ("the singer", "his bandmate")
PRON = r"(?:he|she|they|we|I|you)"
NAME = r"(?:[A-Z][\w'’.-]*(?:\s+(?:[A-Z][\w'’.-]*|de|van|von|la|le|da|di|del|Mc[A-Z]\w*)){0,3})"
DESC = (r"(?:(?:the|his|her|their|our|my|a|one|another)\s+(?:[\w'’-]+\s+){0,3}?"
        r"(?:singer|guitarist|bassist|drummer|keyboardist|frontman|frontwoman|vocalist|producer|songwriter|star|band|"
        r"group|duo|trio|artist|musician|member|bandmate|manager|friend|brother|sister|father|mother|wife|husband|son|"
        r"daughter|spokesman|spokeswoman|spokesperson|statement|label|rapper|dj|pianist|engineer|director|actor|"
        r"actress|writer|author|journalist|critic|fan|executive|founder|co-founder|leader|legend|icon|former|"
        r"latter|pair|couple|men|women|man|woman|boys|girls|lads|four|three|five|quartet|quintet|outfit|act)s?\b|"
        r"[A-Z][\w'’.-]*(?:[ \t]+[A-Z][\w'’.-]*){0,2}['’]s[ \t]+(?:family|estate|camp|management|label|team|"
        r"spokesman|spokeswoman|spokesperson|representative|rep|publicist|manager|widow|daughter|daughters|son|sons)\b)")
ROLE = (r"(?:lead singer|singer|vocalist|lead guitarist|rhythm guitarist|guitarist|guitar player|bassist|bass player|"
        r"drummer|keyboardist|keyboard player|frontman|frontwoman|bandleader|bandmate|producer|songwriter|manager|engineer|"
        r"rapper|DJ|pianist|saxophonist|percussionist|co-founder|founder|member|director|executive|journalist|critic|"
        r"writer|author|actor|actress|comedian|host|presenter|promoter|photographer|editor|publicist|spokesman|"
        r"spokeswoman|spokesperson|friend|brother|sister|father|mother|wife|husband|son|daughter|partner|girlfriend|"
        r"boyfriend|cousin|widow|star|legend|icon|artist|musician|composer|arranger|collaborator|colleague|biographer|"
        r"historian|chairman|boss|president|designer|choreographer|dancer|poet|novelist|filmmaker|co-star)")
ROLE_PREFIX = (r"(?:(?:former|ex-|late|then-|then|longtime|long-time|veteran|original|current|new|fellow|founding|lead|bass|"
               r"rhythm|backing|touring|session|co-)[ \t-]*){0,2}" + ROLE + r"[ \t]+")
POSS_PRE = r"(?:(?:[Tt]he[ \t]+)?" + NAME + r"(?:['’]s|s?['’])[ \t]+)"
NAME_PLUS = (r"(?:" + POSS_PRE + r"(?:" + ROLE_PREFIX + r")?" + NAME + r"|(?:(?:" + NAME + r"|the|his|her|their|our|my)[ \t]+)?" +
             ROLE_PREFIX + NAME + r"|" + NAME + r")")
SUBJ = r"(?:" + PRON + r"|" + NAME_PLUS + r"|" + DESC + r")"

QUOTE_TAIL = re.compile(
    # after the closing mark: , she says / says Knopfler / Knopfler told Rolling Stone / the band said in a statement
    r"^[ \t]*[,;]?[ \t]*(?:[-—–][ \t]*)?(?:(?P<subj1>" + SUBJ + r")(?:[ \t]*,[ \t]*[^,\n]{1,100},)?[ \t]+(?:(?:later|once|now|also|"
    r"simply|then|still|recently|previously|famously|memorably|proudly|candidly|frankly|happily|wryly|ruefully|"
    r"quickly|quietly|firmly|again|has|had|have|would|will|was|is|had once|has since|went on to|would later|"
    r"later|once)[ \t]+){0,2}" + SV_PHRASE + r"\b|(?:so[ \t]+)?(?P<verb2>" + SV_PHRASE + r")[ \t]+(?P<subj2>(?:(?!(?:of|in|at|for|from|with|about|on|to|by|after|before|during|since)\b)[a-z][\w'’-]*,?[ \t]+){0,8}?" +
    SUBJ + r")\b)")
QUOTE_HEAD = re.compile(
    # before the opening mark: Knopfler says: / He said, / As Hynde put it: / She told the Guardian: / Hay recalls -
    r"(?P<subj>" + SUBJ + r")(?:[ \t]*,[ \t]*[^,\n]{1,100},)?[ \t]+(?:(?:later|once|now|also|simply|then|still|"
    r"recently|previously|famously|quickly|immediately|finally|candidly|frankly|happily|wryly|ruefully|quietly|firmly|"
    r"proudly|has|had|would|will|was|is)[ \t]+){0,2}" + SV_PHRASE +
    r"(?:[ \t]+(?:[\w'’.-]+[ \t]+){0,8}?[\w'’.-]+)?[ \t]*[:,—–-]?[ \t]*$")
# a sentence ending in a colon may run longer after the verb ("He noted another drawback, the negative attitudes that
# permeate the Norwegian musical scene:")
QUOTE_HEAD_COLON = re.compile(
    r"(?:^|(?<=[.!?][ \t])|(?<=[.!?][\"”'’][ \t]))(?P<subj>" + SUBJ + r")(?:[ \t]*,[ \t]*[^,\n]{1,100},)?[ \t]+(?:(?:later|once|now|"
    r"also|simply|then|still|recently|previously|famously|quickly|immediately|finally|has|had|would|will|was|is)[ \t]+)"
    r"{0,2}(?:" + SV_PHRASE + r"|believes|believed|thinks|thought|feels|felt)" +
    r"(?:[ \t]+[^.!?\n\"“”:]{1,140}?)?[ \t]*:[ \t]*$")
# ", saying:", ", adding that", "before telling the crowd:" (the speaker is the sentence's subject)
PARTICIPLE_HEAD = re.compile(r"(?i)(?:,|\band|\bbefore|\bwhile|\bwhen)[ \t]+(?P<verb>(?:(?<=and )(?:said|says|added|adds|"
                             r"explained|recalled|laughed|joked|replied|told)|saying|adding|explaining|noting|recalling|insisting|"
                             r"admitting|joking|laughing|writing|telling|declaring|confessing|stating|reflecting|remembering|"
                             r"recounting|reminiscing|revealing|conceding|claiming|arguing|remarking|commenting|tweeting|"
                             r"announcing|answering|replying|musing|quipping|sighing|shrugging|smiling|grinning|explaining|"
                             r"proclaiming|warning|insisting|observing|acknowledging|continuing|concluding)"
                             r"(?:[ \t]+[\w'’.-]+){0,4})[ \t]*[:,]?[ \t]*$")
# lyrics quoted with slashes for line breaks: "Music scene is crazy/ Bands start up each and every day/"
LYRIC_SLASHES = re.compile(r"[a-z,.!?'’)\]][ \t]*/[ \t]*[A-Z]")
DASH_CREDIT = re.compile(r"^[ \t]*[—–-][ \t]*(?P<subj>" + NAME + r")\b")
_INV = ("says said adds added explains explained recalls recalled notes noted laughs laughed admits admitted insists "
        "insisted continues continued concludes concluded remembers remembered reflects reflected offers offered quips "
        "quipped jokes joked muses mused observes observed confides confided agrees agreed counters countered asks asked "
        "replies replied emphasises emphasised emphasizes emphasized enthuses enthused concedes conceded sighs sighed warns "
        "warned declares declared comments commented remarks remarked recounts recounted reminisces reminisced states "
        "stated writes wrote tells told furthers reports asserts explains").split()
INVERTED_HEAD = re.compile(r"(?:^|(?<=[.!?:;])[ \t]+|(?<=[.!?:;][\"”'’])[ \t]+)(?:As[ \t]+)?(?P<verb>" +
                           "|".join(sorted({v.capitalize() for v in _INV}, key=len, reverse=True)) +
                           r")[ \t]+(?P<subj>" + SUBJ + r")(?:[ \t]*,[ \t]*[^\n\"“”]{1,80}?|[ \t]+(?:with|of|in|on|about|"
                           r"after|before|when|while|as)[ \t][^\n\"“”]{1,80}?)?[ \t]*[:,—–-]?[ \t]*$")
# a speaker's name as a label at the start of the paragraph ("LEE ROCKER: "...", "Thompson: '...", "Lee Rocker (bass): ")
LABEL_HEAD = re.compile(r"^[ \t]*(?P<subj>" + NAME + r")(?:[ \t]*\([^)\n]{1,40}\))?[ \t]*:[ \t]*$")
SUNG = r"(?:sings|sang|croons|crooned|raps|rapped|wails|wailed|chants|chanted|belts|belted|intones|intoned|warbles|warbled)"
SUNG_TAIL = re.compile(r"^[ \t]*[,;]?[ \t]*(?:(?:" + SUBJ + r")[ \t]+" + SUNG + r"\b|" + SUNG + r"[ \t]+(?:" + SUBJ + r")\b)")
SUNG_HEAD = re.compile(r"(?:" + SUBJ + r")[ \t]+" + SUNG + r"(?:[ \t]+[\w'’.-]+){0,6}?[ \t]*[:,—–-]?[ \t]*$")

# ── words that make a quotation a title rather than speech ────────────────────
TITLE_WORDS = (r"song|songs|single|singles|album|albums|lp|ep|track|tracks|record|records|hit|hits|tune|tunes|ballad|"
               r"anthem|video|videos|film|films|movie|movies|book|books|memoir|novel|show|shows|series|tour|tours|play|"
               r"musical|documentary|cover|version|remix|mix|b-side|ep's|title|titled|entitled|called|named|dubbed|"
               r"nicknamed|known as|so-called|term|phrase|word|words|slogan|motto|headline|chapter|poem|catchphrase|"
               r"label|magazine|newspaper|column|programme|program|special|sketch|episode|opera|soundtrack|"
               r"compilation|collection|box set|boxset|reissue|edition|chorus|refrain|lyric|lyrics|line|lines|verse")
TITLE_BEFORE = re.compile(r"(?i)\b(?:" + TITLE_WORDS + r")(?:\s+(?:of|from|for|like|as|is|was|called|titled|entitled|"
                          r"named))?[\s,:]*$")
TITLE_AFTER = re.compile(r"(?i)^[\s,]*(?:\(\d{4}\)\s*)?(?:" + TITLE_WORDS + r")\b")
FIRST_PERSON = re.compile(r"\b(?:I|I'm|I've|I'd|I'll|me|my|mine|myself|we|we're|we've|we'd|we'll|us|our|ours|"
                          r"ourselves|you|your|you're|you've|you'd|you'll|yourself)\b")
VERBISH = re.compile(r"(?i)\b(?:is|are|was|were|be|been|being|am|do|does|did|have|has|had|can|could|will|would|shall|"
                     r"should|may|might|must|get|got|go|went|gone|make|made|know|knew|think|thought|want|wanted|feel|"
                     r"felt|said|say|says|like|liked|love|loved|need|needed|see|saw|seen|come|came|take|took|give|gave|"
                     r"tell|told|let|keep|kept|look|looked|play|played|write|wrote|written|sing|sang|sung|try|tried|"
                     r"mean|meant|work|worked|start|started|put|call|called|seem|seemed|find|found|believe|believed|"
                     r"remember|remembered|happen|happened|use|used|\w+n't|\w+'(?:s|re|ve|ll|d|m))\b")

ELISIONS = re.compile(r"(?i)^(?:em|til|till|cause|cos|coz|n|n'|bout|round|tis|twas|nuff|ere|appen|ow|im|er|e|ad|"
                      r"\d0s|\d\ds|\d\d)\b")


def paragraphs(text, a, b):
    """(start, end) of each paragraph of text[a:b], without its line break."""
    out, pos = [], a
    while pos < b:
        nl = text.find("\n", pos, b)
        end = b if nl < 0 else nl
        out.append((pos, end))
        pos = end + 1
    return out


def _mark_type(text, i, lo, hi, mark):
    """Whether the mark at i opens ('O') or closes ('C') a quotation, or could do either ('?')."""
    prev = text[i - 1] if i > lo else " "
    nxt = text[i + 1] if i + 1 < hi else " "
    space_before = prev.isspace() or prev in "([{—–-/—–" or i == lo
    space_after = nxt.isspace() or nxt in ".,;:!?)]}—–-…/" or i + 1 == hi
    if mark == "'" and prev.isalnum() and nxt.isalnum():
        return None                                    # an apostrophe: don't, Adamson's
    if space_before and not space_after:
        return "O"
    if space_after and not space_before:
        return "C"
    if prev in "\"'" and not space_after:
        return "O"
    return "?"


def _single_marks(text, lo, hi):
    """Single quotation marks in a British-style paragraph, without apostrophes, elisions ('em, '80s, rock 'n' roll)
    and plural possessives (the boys' lives)."""
    out = []
    for m in re.finditer("'", text[lo:hi]):
        i = lo + m.start()
        t = _mark_type(text, i, lo, hi, "'")
        if t is None:
            continue
        after = text[i + 1:i + 8]
        before = text[max(lo, i - 3):i]
        if t == "O" and ELISIONS.match(after):
            continue
        if t in ("C", "?") and re.search(r"(?i)s$", before) and re.match(r"\s+[a-z]", text[i + 1:i + 4]) and \
                not re.match(r"\s+(?:he|she|they|we|i|says|said|adds|added|recalls|explains)\b", text[i + 1:i + 14]):
            continue                                   # the boys' lives
        if re.search(r"(?i)\bn$", before) and t == "C":
            continue                                   # rock 'n' roll
        out.append((i, t))
    return out


INCH = re.compile(r'(?<![\w.,])\d{1,2}(?:\.\d)?$')


def _double_marks(text, lo, hi):
    """Double quotation marks in a paragraph. When they do not balance, an inch mark (18", 24") is left out."""
    ms = [(lo + m.start(), _mark_type(text, lo + m.start(), lo, hi, '"')) for m in re.finditer('"', text[lo:hi])]
    if len(ms) % 2:
        for i, t in ms:
            if INCH.search(text[max(lo, i - 6):i]) and not (i + 1 < hi and text[i + 1].isalpha()):
                return [x for x in ms if x[0] != i]
    return ms


def _first_char(text, s, e):
    i = s
    while i < e and text[i] in " \t":
        i += 1
    return i


def find_quotes(text, a, b, mark='"', lookahead=1500):
    """Quotation spans in the prose segment text[a:b]: [(start, end, paragraphs, how)], end just after the closing mark,
    and the unresolved paragraphs [(start, end)].
    Marks are paired in order within a paragraph. A quotation left open at the end of a paragraph runs on when
    (a) the next paragraph opens with a quotation mark (British magazines reopen every paragraph of a long quotation),
    or (b) the paragraphs after it carry no marks (up to 1,500 characters of them) and the next mark found closes it,
    as on hard-wrapped pages that break every printed line. Otherwise the paragraph is
    unresolved, as is the first paragraph of a running quotation that never closes."""
    paras = [p for p in paragraphs(text, a, b)]
    marks_of = [(_double_marks if mark == '"' else _single_marks)(text, s, e) for s, e in paras]
    quotes, unresolved = [], []
    k, n = 0, len(paras)
    open_at = open_k = None
    how = ""
    while k < n:
        s, e = paras[k]
        ms = marks_of[k]
        j = 0
        if open_at is not None:
            if ms and ms[0][0] == _first_char(text, s, e):
                j = 1                                  # the paragraph reopens the running quotation
            if j >= len(ms):
                k += 1                                 # no closing mark here: the quotation runs on
                continue
            quotes.append((open_at, ms[j][0] + 1, k - open_k + 1, how))
            open_at = None
            j += 1
        rest = ms[j:]
        for q in range(0, len(rest) - 1, 2):
            quotes.append((rest[q][0], rest[q + 1][0] + 1, 1, "paired"))
        if len(rest) % 2:
            i, t = rest[-1]
            carry = None
            if t != "C":
                budget = 0                             # characters of unmarked paragraphs the quotation may run over
                for m in range(k + 1, n):
                    if not text[paras[m][0]:paras[m][1]].strip():
                        continue
                    if marks_of[m] and marks_of[m][0][0] == _first_char(text, *paras[m]) and budget == 0:
                        carry = "reopened"             # the next paragraph opens with a mark
                        break
                    if marks_of[m]:
                        if marks_of[m][0][1] in ("C", "?"):
                            carry = "closed later"     # hard-wrapped pages: a quotation over several lines
                        break
                    budget += paras[m][1] - paras[m][0]
                    if budget > lookahead:
                        break
            if carry:
                open_at, open_k, how = i, k, carry
            else:
                quotes = [x for x in quotes if not (s <= x[0] < e)]
                unresolved.append((s, e))
        k += 1
    if open_at is not None:
        # a running quotation that never closed: its first paragraph is unresolved; read the rest again
        s0, e0 = paras[open_k]
        quotes = [x for x in quotes if x[0] < s0]
        unresolved = [u for u in unresolved if u[0] < s0]
        unresolved.append((s0, e0))
        if e0 + 1 < b:
            q2, u2 = find_quotes(text, e0 + 1, b, mark, lookahead)
            quotes += q2
            unresolved += u2
    return quotes, unresolved


# ── step 1: the record's quotation convention ─────────────────────────────────
def _attr_after(text, i, hi):
    return bool(QUOTE_TAIL.match(text[i:min(hi, i + 120)]))


def convention(text, segs):
    """'british' when single-quote attributions outnumber double-quote ones and there are at least two; otherwise
    'american'. Single marks inside a double-quoted passage are not counted. Returns (convention, evidence)."""
    d = s = 0
    for a, b in segs:
        for ps, pe in paragraphs(text, a, b):
            dq = _double_marks(text, ps, pe)
            inside = [(dq[k][0], dq[k + 1][0]) for k in range(0, len(dq) - 1, 2)]
            for i, t in dq:
                if t in ("C", "?") and i > ps and text[i - 1] in ",.?!…-—" and _attr_after(text, i + 1, pe):
                    d += 1
            for i, t in _single_marks(text, ps, pe):
                if any(x < i < y for x, y in inside):
                    continue
                if t in ("C", "?") and i > ps and text[i - 1] in ",.?!…-—" and _attr_after(text, i + 1, pe):
                    s += 1
    conv = "british" if s > d and s >= 2 else "american"
    return conv, {"double-quote attributions": d, "single-quote attributions": s}


# ── step 4: speech, non-speech or undecided ──────────────────────────────────
CLAUSE_END = re.compile(r"[.;:!?]|,(?=\s)|\n|[\"'‘“]")
SMALL_WORDS = {"a", "an", "the", "of", "in", "on", "and", "or", "to", "for", "at", "by", "with", "from", "is", "it",
               "as", "but", "nor", "so", "up", "out", "my", "me", "be", "are", "was", "do", "de", "la", "le", "vs", "vs."}


def title_case(s):
    words = re.findall(r"[A-Za-z][\w'’-]*", s)
    if not words or len(words) > 10:
        return False
    big = [w for w in words if w.lower() not in SMALL_WORDS or w is words[0]]
    return all(w[0].isupper() for w in big) and (len(words) >= 2 or (words[0][0].isupper() and len(words[0]) > 1))


def _clause_end(text, i, hi):
    m = CLAUSE_END.search(text, i, hi)
    if not m:
        return hi
    j = m.start()
    if text[j] in ".;:!?,":
        j += 1
    while j < hi and text[j] in " \t":
        j += 1
    return j


def classify(text, a, b, quotes, mark):
    """Each quotation in the segment as (start, end, kind, attribution span or None, evidence, name label or
    None). An undecided quotation after a name label keeps the label for stage 5, which decides whether it is speech."""
    out = []
    prev_speech_end = None
    for n, (s, e, paras, how) in enumerate(quotes):
        inner = text[s + 1:e - 1]
        # the paragraph around the quotation (the head stops at the previous quotation)
        ps = text.rfind("\n", a, s) + 1 if text.rfind("\n", a, s) >= 0 else a
        ps = max(ps, a, out[-1][1] if out else a)
        pe = text.find("\n", e, b)
        pe = b if pe < 0 else pe
        nxt_q = quotes[n + 1][0] if n + 1 < len(quotes) else b
        tail_end = min(pe, nxt_q, e + 160)
        tail = text[e:tail_end]
        head = text[max(ps, s - 160):s]
        attr, kind, ev, label = None, None, "", None
        m = SUNG_TAIL.match(tail)
        if m:
            kind, ev = "non-speech", "sung (a verb of singing beside it)"
        m = QUOTE_TAIL.match(tail) if not kind else None
        if m:
            end = _clause_end(text, e + m.end(), tail_end)
            attr = (e, end)
            kind, ev = "speech", "attribution after: " + text[e:end].strip()[:80]
        if not kind:
            m = QUOTE_HEAD.search(head)
            if m and not re.search(r"(?i)\b(?:" + TITLE_WORDS + r")[\s,:]*$", head):
                st = max(ps, s - 160) + m.start("subj")
                attr = (st, s)
                kind, ev = "speech", "attribution before: " + text[st:s].strip()[:80]
        if not kind:
            m = QUOTE_HEAD_COLON.search(head)
            if m and not re.search(r"(?i)\b(?:" + TITLE_WORDS + r")\s*:\s*$", head):
                st = max(ps, s - 160) + m.start("subj")
                attr = (st, s)
                kind, ev = "speech", "attribution before: " + text[st:s].strip()[:80]
        if not kind:
            m = INVERTED_HEAD.search(head)
            if m:
                st = max(ps, s - 160) + m.start("verb")
                attr = (st, s)
                kind, ev = "speech", "attribution before: " + text[st:s].strip()[:80]
        if not kind and SUNG_HEAD.search(head):
            kind, ev = "non-speech", "sung (a verb of singing beside it)"
        if not kind:
            m = PARTICIPLE_HEAD.search(head)
            if m:
                st = max(ps, s - 160) + m.start()
                while st < s and text[st] in " \t,":
                    st += 1
                attr = (st, s)
                kind, ev = "speech", "attribution before: " + text[st:s].strip()[:80]
        if not kind:
            m = DASH_CREDIT.match(tail)
            if m and len(inner) > 20 and not title_case(inner.strip(" ,.!?")):
                attr = (e, _clause_end(text, e + m.end(), tail_end))
                kind, ev = "speech", "credited after a dash: " + m.group("subj")
        if kind == "speech" and attr and title_case(inner.strip(" ,.!?")) and len(inner.split()) <= 8 and \
                not FIRST_PERSON.search(inner) and not re.search(r"[.!?]\s+[A-Z]", inner) and \
                re.search(r"(?i)\b(?:wrote|writes|penned|pens|composed|recorded|records|covered|sang|sings|released|"
                          r"releases|produced|co-wrote|cowrote)\b", text[attr[0]:attr[1]]):
            kind, attr, ev = "non-speech", None, "a title after a verb of writing or recording"
        plain = re.sub(r"[\"“”][^\"“”]*[\"“”]|(?<!\w)'[^'\n]*'(?!\w)", " ", inner) if mark == '"' else inner
        if not kind and (TITLE_BEFORE.search(head) or TITLE_AFTER.match(tail)):
            kind, ev = "non-speech", "a title word beside it"
        if not kind and LYRIC_SLASHES.search(inner):
            kind, ev = "non-speech", "lyrics (lines run together with slashes)"
        if not kind and prev_speech_end is not None and paras >= 1:
            gap = text[prev_speech_end:s]
            if not gap.strip() or re.fullmatch(r"\s*[,.;:—–-]?\s*", gap):
                kind, ev = "speech", "continues the quotation before it"
        if not kind and title_case(inner.strip(" ,.!?")) and not re.search(r"[.!?]\s+[A-Z]", inner):
            kind, ev = "non-speech", "title case"
        if not kind and FIRST_PERSON.search(plain) and len(plain.split()) >= 4 and VERBISH.search(plain):
            kind, ev = "speech", "a sentence in the first or second person"
        if not kind:
            kind, ev = "undecided", ""
        # a speaker's name as a label at the start of the paragraph ("LEE ROCKER: "...")
        para_start = text.rfind("\n", a, s) + 1 if text.rfind("\n", a, s) >= 0 else a
        if attr is None and kind in ("speech", "undecided") and ps == para_start and s - para_start <= 80:
            m = LABEL_HEAD.match(text[para_start:s])
            if m:
                label = (para_start + m.start("subj"), s)
                if kind == "speech":
                    attr = label
                    ev += "; a name label before it"
        out.append((s, e, kind, attr, ev, label))
        if kind == "speech":
            prev_speech_end = attr[1] if attr and attr[0] >= e else e
        elif kind != "speech":
            prev_speech_end = None
    return out


def spans(text, segs, conv=None):
    """Stage 4 for one record: its prose segments cut into spans that cover them exactly, and the record's convention.
    Each span is a dict: start, end, kind, evidence; a speech span also has "attribution": the (start, end) of its
    attribution clause, or None."""
    if conv is None:
        conv, _ = convention(text, segs)
    mark = "'" if conv == "british" else '"'
    out = []
    for a, b in segs:
        quotes, unresolved = find_quotes(text, a, b, mark)
        cls = classify(text, a, b, quotes, mark)
        pieces = [dict(start=u0, end=u1, kind="unresolved", evidence="quotation marks do not balance", attribution=None)
                  for u0, u1 in unresolved]
        for s_, e_, kind, attr, ev, label in cls:
            if e_ - s_ > 1500:
                ev += " (over 1,500 characters: a closing mark may be missing)"
            pieces.append(dict(start=s_, end=e_, kind=kind, evidence=ev, attribution=attr, label=label))
            if attr:
                pieces.append(dict(start=attr[0], end=attr[1], kind="attribution", evidence="", attribution=None))
        pieces.sort(key=lambda x: (x["start"], x["end"]))
        cur = a
        for x in pieces:
            st = max(x["start"], cur)
            if x["end"] <= st:
                continue
            if st > cur:
                out.append(dict(start=cur, end=st, kind="narration", evidence="", attribution=None))
            y = dict(x)
            y["start"] = st
            out.append(y)
            cur = x["end"]
        if cur < b:
            out.append(dict(start=cur, end=b, kind="narration", evidence="", attribution=None))
    return out, conv
'''
SOURCES["gender_names"] = r'''"""First names that are clearly a man's or a woman's in English, for the pronoun check on people who are not members
(from the first-name list of stage 2, with genders from the gender-guesser data; names used for both are left out)."""
import base64
import zlib

_GENDER_Z = (
    "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"
)
GENDER = dict(x.split(":") for x in zlib.decompress(base64.b64decode(_GENDER_Z)).decode().split("|"))
'''
SOURCES["stage5"] = r'''"""Stages 5 to 9 of the turn build (technical appendix, Step 2): who says each quotation in the prose (the attribution cascade,
stage 5, with speaker tracking, stage 6), the turns (stage 7, with second-hand self narrative), the interview date of
each turn (stage 8) and its confidence (stage 9). The Q&A and first-person blocks keep the speakers stage 3 gave them.

Speaker types: band member, band member (unspecified), band collective, other-era member, interviewer, journalist,
other person, unknown, ambiguous. Band voice = band member, band member (unspecified) and band collective (collective
turns carry their own type so they can be left out).
"""
import bisect
import functools
import re
from collections import defaultdict

import stage3 as T
import stage4 as Q
from gender_names import GENDER

S = T.S
MEMBER, MEMBERS_UNSPEC, OTHER_ERA, OTHER_ERA_UNSPEC = T.MEMBER, T.MEMBERS_UNSPEC, T.OTHER_ERA, T.OTHER_ERA_UNSPEC
INTERVIEWER, OTHER, UNKNOWN, AMBIGUOUS = T.INTERVIEWER, T.OTHER, T.UNKNOWN, T.AMBIGUOUS
COLLECTIVE = "band collective"
JOURNALIST = "journalist"
BAND_VOICE = {MEMBER, MEMBERS_UNSPEC, COLLECTIVE}
RULES_VERSION = "turns_v7: stages 4-9 version 1.3, stage 3 version 3"
LEVELS = T.LEVELS

COLLECTIVE_WORDS = {"band", "group", "duo", "trio", "quartet", "quintet", "outfit", "act", "pair", "boys", "girls", "lads",
                    "four", "three", "five", "foursome", "threesome", "members", "bandmates", "couple"}
# version 1.2: a word that gives the group's size ("the duo") must fit the act's nomination-era line-up
SIZE_WORDS = {"duo": 2, "pair": 2, "couple": 2, "twosome": 2, "two-piece": 2, "trio": 3, "threesome": 3, "three": 3,
              "three-piece": 3, "quartet": 4, "foursome": 4, "four": 4, "four-piece": 4, "quintet": 5, "five": 5,
              "five-piece": 5, "sextet": 6, "six-piece": 6}
# a capitalised name that is the subject of a verb ("Hoodoo Gurus have", "The Living End have also denounced")
NAME_RUN = re.compile(r"(?<![\w'’])(?:[Tt]he\s+)?[A-Z][\w'’.&-]*(?:\s+(?:[A-Z][\w'’.&-]*|of|&|de|van|von)){0,4}")
STRICT_VERB = re.compile(
    r"(?:\s*\([^)\n]{1,60}\))?(?:\s*[,–—][^,–—\n\"“”]{1,80}[,–—])?\s+(?:(?:also|still|now|later|then|once|never|"
    r"always|just|really|even|already|recently|finally|quickly|publicly|have|has|had)\s+)?"
    r"(?P<verb>is|was|were|are|has|had|have|will|would|can|could|did|do|does|says|said|announced|announces|released|"
    r"releases|wrote|writes|posted|posts|shared|shares|added|adds|explained|explains|told|tells|insisted|stated|"
    r"claimed|revealed|confirmed|denied|condemned|denounced|published|issued|thanked|joined|formed|split|played|plays|"
    r"performed|toured|recorded|covered|paid|took|made|became|returned|opened|headlined|unveiled|dropped|debuted|"
    r"launched|teamed|reunited|cancelled|canceled|postponed|rose|began|started|continued|signed|won|celebrated|spoke|"
    r"responded|replied|reacted|tweeted|referenced|asked|called|urged|apologised|apologized)\b")
SENTENCE_STARTERS = {"The", "A", "An", "But", "And", "So", "Then", "Now", "In", "On", "At", "For", "With", "After",
                     "Before", "When", "While", "As", "If", "It", "Its", "This", "That", "These", "Those", "There",
                     "Here", "He", "She", "They", "We", "I", "You", "His", "Her", "Their", "Our", "My", "Your", "Back",
                     "Last", "Next", "Earlier", "Later", "Meanwhile", "Still", "Yet", "Also", "One", "Some", "Many",
                     "Most", "All", "Both", "Each", "Every", "No", "Not", "Even", "Only", "Just", "What", "Who", "How",
                     "Why", "Where", "Mr", "Ms", "Mrs", "Dr", "Sir", "Related", "Read", "Advertisement", "Photo",
                     "Image", "Credit", "Source", "Getty", "Trending"}
ROLE_WORDS = {"singer", "guitarist", "bassist", "drummer", "keyboardist", "frontman", "frontwoman", "vocalist",
              "songwriter", "star", "artist", "musician", "rapper", "pianist", "leader", "legend", "icon", "former",
              "latter", "founder", "co-founder", "member", "bandmate", "man", "woman"}
# a one-word name after these is a relative or an associate, not a member who shares the first name ("his wife Jane")
RELATION_WORDS = {"wife", "husband", "son", "sons", "daughter", "daughters", "brother", "sister", "father", "mother",
                  "dad", "mum", "mom", "friend", "partner", "girlfriend", "boyfriend", "cousin", "widow", "manager",
                  "producer", "engineer", "fan", "family", "uncle", "aunt", "grandson", "granddaughter", "nephew",
                  "niece", "agent", "lawyer", "publicist", "spokesman", "spokeswoman", "spokesperson", "assistant"}
NOT_SPEAKERS = {"Answer", "Question", "Q", "A", "ID", "States", "Kingdom", "Times", "News", "Press", "Records", "Music",
                "Magazine", "Fans", "Critics", "Reports", "Sources", "Officials", "Police", "Speaking", "Talking",
                "Asked", "Interviewed", "Writing", "Recalling", "Explaining", "Describing", "Discussing"}
PRONOUNS = {"he": "he", "she": "she", "they": "they", "we": "we", "i": "i", "you": "you"}
DETERMINERS = ("the", "his", "her", "their", "our", "my", "a", "one", "another")
LEAD_WORDS = {"as", "and", "but", "so", "then", "now", "or", "yet", "when", "while", "here", "there", "still", "later",
              "finally", "also", "indeed", "instead", "meanwhile", "says", "said", "adds", "added", "explains",
              "explained", "recalls", "recalled", "notes", "noted"}
POSSESSIVE_TOKEN = re.compile(r"(?:['’]s|s['’])$")


def fold(s):
    return T.fold(s)


def lower_level(c):
    return T.lower_level(c)


def min_level(a, b):
    return a if LEVELS.index(a) <= LEVELS.index(b) else b


def person_pronoun(name):
    """'he' or 'she' for a person whose first name is clearly a man's or a woman's; otherwise None."""
    toks = re.sub(r"^(?:Mr|Ms|Mrs|Dr|Miss|Sir)\.?\s+", "", name or "").split()
    if not toks:
        return None
    if re.match(r"(?:Mr|Sir)\b", name or ""):
        return "he"
    if re.match(r"(?:Ms|Mrs|Miss)\b", name or ""):
        return "she"
    g = GENDER.get(toks[0].title()) if len(toks) >= 2 or toks[0].title() in GENDER else None
    return {"m": "he", "f": "she"}.get(g)


def named_part(subj):
    """The name a subject ends with, and the words before it: "the Bangles' Susanna Hoffs" gives ("Susanna Hoffs",
    "the Bangles'"); "vocalist Ian Kenny" gives ("Ian Kenny", "vocalist"); "As Hynde" gives ("Hynde", "")."""
    toks = subj.split()
    while len(toks) > 1 and toks[0].lower() in LEAD_WORDS:
        toks = toks[1:]
    run = []
    for tok in reversed(toks):
        if run and POSSESSIVE_TOKEN.search(tok):
            break
        if tok[:1].isupper() or tok in T.PARTICLES:
            run.insert(0, tok)
        else:
            break
    while run and run[0] in T.PARTICLES:
        run.pop(0)
    return " ".join(run), " ".join(toks[:len(toks) - len(run)])


# ── who is named where ───────────────────────────────────────────────────────
@functools.lru_cache(maxsize=None)
def _word_rx(word, accents):
    return re.compile(r"(?<![\w'’-])" + (T.loose(word) if accents else re.escape(word)) + r"(?![\w-])")


class Mentions:
    """Every person named in the record, with positions: members by any full form or by a first name or surname only
    they have in the act; other people named in full, and later by their surname alone."""

    def __init__(self, rec_text, title, names, stop, act, solo=None):
        self.names = names
        items = []
        for a, b, n, m in T.person_spans(rec_text, names, stop, act, lines=True):
            cut = rec_text.find("--", a, b)
            if cut > a and m is None:
                n = re.split(r"--", n)[0].strip()
                b = cut
                if len(n.split()) < 2:
                    continue
            items.append((a, b, n, m))
        full = sorted((a, b) for a, b, n, m in items)
        full_starts = [x[0] for x in full]

        def in_full(pos):
            k = bisect.bisect_right(full_starts, pos) - 1
            return k >= 0 and full[k][0] <= pos < full[k][1]

        # a member's first name or surname alone ("Hornsby considers...", "Pål surprised fans"): a name only one member
        # has is that member; a name several members share ("Watson" for Bruce and Jamie Watson) goes to the one named in
        # full most recently before it, as in rule 3
        full_members = sorted((a, m) for a, b, n, m in items if m)
        owners = defaultdict(set)
        for mem in names.members:
            for tk in _short_tokens(names, mem):
                owners[tk].add(mem)
        if solo:
            owners[act].add(solo)
        else:
            # version 1.2: the act's name that is also its singer's name ("Sade") names the singer, as in stage 3
            bare = re.sub(r"^the\s+", "", act, flags=re.I)
            singer = names.form.get(fold(bare))
            if singer and len(singer) == 1 and " " not in bare:
                owners[bare].add(next(iter(singer)))
        for tk, mems in sorted(owners.items()):
            for mm in _word_rx(tk, True).finditer(rec_text):
                if in_full(mm.start()):
                    continue
                if len(mems) == 1:
                    mem = next(iter(mems))
                else:
                    prev = [m2 for a2, m2 in full_members if a2 < mm.start() and m2 in mems]
                    if not prev:
                        continue
                    mem = prev[-1]
                items.append((mm.start(), mm.end(), mem, mem))
        # a Les Paul is a guitar ("his Les Paul", "a Gibson Les Paul Custom")
        items = [x for x in items if not (x[2] == "Les Paul" and x[3] is None and
                                          (re.search(r"(?i)(?:gibson|\b(?:a|his|her|my|the|their|old|vintage|black|gold))\s+$",
                                                     rec_text[max(0, x[0] - 12):x[0]]) or
                                           re.match(r"(?i)\s*(?:custom|junior|standard|special|guitar|gold|model|copy|s\b)",
                                                    rec_text[x[1]:x[1] + 12])))]
        # other people by their surname after they have been named in full ("Raitt" after "Bonnie Raitt")
        firsts = {}
        for a, b, n, m in sorted(items, key=lambda x: (x[0], x[1], x[2], x[3] or "")):
            if m is None:
                last = n.split()[-1]
                if len(last) >= 3 and last not in S.FUNC and last.title() not in stop and \
                        T.fold(last) not in names.surname and T.fold(last) not in names.first:
                    firsts.setdefault(last, (a, n))
        owners = defaultdict(set)
        for last, (a, n) in firsts.items():
            owners[last].add(n)
        for last, (a0, n) in sorted(firsts.items()):
            if len(owners[last]) != 1:
                continue
            for mm in _word_rx(last, False).finditer(rec_text):
                if mm.start() <= a0 or in_full(mm.start()):
                    continue
                items.append((mm.start(), mm.end(), n, None))
        self.items = sorted(set(items), key=lambda x: (x[0], x[1], x[2], x[3] or ""))
        self.title_members = set(names.mentions(title or ""))
        self.text = rec_text

    def before(self, pos, limit=None):
        if not hasattr(self, "_starts") or len(self._starts) != len(self.items):
            self._starts = [x[0] for x in self.items]
        hi = bisect.bisect_left(self._starts, pos)
        lo = 0 if limit is None else bisect.bisect_left(self._starts, limit)
        return self.items[lo:hi]

    def full_named(self, member, before=None):
        return any(x[3] == member and (before is None or x[0] < before) for x in self.items) or \
            member in self.title_members

    def incidental(self, pos, end=None):
        """A name in passing: someone's relation ("son of the late John Lennon", "John Lennon's son") or someone dead
        ("the late Prince")."""
        if end is not None and INCIDENTAL_AFTER.match(self.text, end):
            return True
        return bool(INCIDENTAL.search(self.text, max(0, pos - 40), pos))

    def verb_subject(self, end):
        """Whether the name ending at end is followed by a verb ("Hay was", "Hooper also takes", "Watson – on paper,
        a young punk – was coming")."""
        return bool(VERB_AFTER.match(self.text, end))

    def subject_position(self, pos):
        """Whether a name at pos opens its sentence (at most three words before it)."""
        st = max(self.text.rfind("\n", 0, pos), max((self.text.rfind(c, 0, pos) for c in ".!?"), default=-1))
        return len(self.text[st + 1:pos].split()) <= 3


def _tokens(name):
    return [t for t in T._tokens(name) if t.casefold() not in T.SUFFIX]


def _short_tokens(names, mem):
    """A member's first name (four letters or more) and surname (three or more), and those of the member's aliases, when
    they are not ordinary words, whether or not another member shares them."""
    out = []
    for form in [mem] + sorted(names.members[mem]["forms"] - {mem}):
        toks = _tokens(form)
        if len(toks) < 2:
            continue
        for t, minlen, surname in ((toks[0], 4, False), (toks[-1], 3, True)):
            if surname and fold(t) in names.no_bare:
                continue
            if len(t) >= minlen and t[0].isupper() and t not in S.FUNC and t not in S.WORDLIKE_NAMES and t not in out:
                out.append(t)
    return out


INCIDENTAL = re.compile(r"(?i)(?:\bthe late|\b(?:son|sons|daughter|daughters|wife|husband|widow|father|mother|brother|"
                        r"sister|friend|fan|fans|cousin|nephew|niece|grandson|granddaughter)\s+of(?:\s+the\s+late)?)\s+$")
INCIDENTAL_AFTER = re.compile(r"(?i)['’]s\s+(?:late\s+|former\s+|ex-|longtime\s+|long-time\s+)?(?:son|sons|daughter|"
                              r"daughters|wife|husband|widow|father|dad|mother|mum|mom|brother|sister|friend|manager|"
                              r"producer|label|band|bandmate|bandmates|girlfriend|boyfriend|partner|family|estate|"
                              r"biographer|lawyer|spokesman|spokeswoman|publicist|engineer|guitarist|drummer|bassist|"
                              r"keyboardist|assistant|children|kids)\b")
VERB_AFTER = re.compile(
    r"(?:\s*\([^)\n]{1,60}\))?(?:\s*[,–—][^,–—\n\"“”]{1,80}[,–—])?\s+(?:(?:also|still|now|later|then|once|never|always|"
    r"just|really|even|already|recently|finally|quickly|often|soon|himself|herself|too)\s+)?(?:is|was|were|are|has|had|"
    r"have|does|did|do|will|would|can|could|should|may|might|must|says|said|tells|told|recalls|recalled|remembers|"
    r"remembered|explains|explained|admits|admitted|thinks|thought|feels|felt|knows|knew|wants|wanted|believes|believed|"
    r"hopes|hoped|insists|insisted|agrees|agreed|laughs|laughed|smiles|smiled|notes|noted|adds|added|seems|seemed|"
    r"looks|looked|sounds|sounded|takes|took|makes|made|gets|got|goes|went|comes|came|sees|saw|became|becomes|began|"
    r"begins|started|starts|joined|left|formed|wrote|writes|released|recorded|played|plays|sang|sings|met|meets|moved|"
    r"moves|lives|lived|spent|spends|worked|works|grew|grows|returned|returns|turned|turns|decided|decides|considers|"
    r"considered|describes|described|reflects|reflected|leans|leaned|pauses|paused|sits|sat|stands|stood|shrugs|"
    r"shrugged|nods|nodded|sighs|sighed|grins|grinned|chuckles|chuckled|speaks|spoke|talks|talked|answers|answered|"
    r"asks|asked|replies|replied|continues|continued|points|pointed|reveals|revealed|confesses|confessed|found|kept|"
    r"held|brought|bought|built|caught|taught|fought|sought|lost|won|ran|gave|knew|heard|meant|sent|shot|understood|"
    r"wore|became|isn['’]t|"
    r"wasn['’]t|doesn['’]t|didn['’]t|hasn['’]t|hadn['’]t|won['’]t|wouldn['’]t|can['’]t|couldn['’]t|"
    r"(?!(?:this|his|its|plus|thus|perhaps|always|news|towards|besides|across|unless|less|yes|us|bus|lyrics|songs|"
    r"albums|records|guitars|fans|days|years|times|ones|others|has|was|is|as|shows|concerts|gigs|hits|tours|videos|"
    r"singles|tracks|friends|parents|brothers|sisters|kids|children|members|bandmates|vocals|drums|sessions|nights|"
    r"weeks|months|fame|awards|tapes|demos|lines|words|lyrics|movies|films|books|photos|pictures|clothes|stars|"
    r"and|or|but|for|in|on|at|to|of|with|from|by)\b)(?![a-z]+(?:ous|ss|ics|ies)\b)"
    r"[a-z]{2,}(?:ed|s))\b")


# ── stage 5: one attribution clause to a speaker ─────────────────────────────
_SENT_BREAK = re.compile(r"(?<=[a-z]{2})[.!?]\s+(?=[A-Z])")


def clause_subject(clause):
    """The speaker words in an attribution clause: a pronoun, a name or a description, or None. A name that runs over a
    sentence break is cut there ("Gibbon. But" after the quotation is Gibbon; "The Electric Company. She said:" is
    she)."""
    c = clause.strip()
    m = Q.QUOTE_TAIL.match(" " + c)
    if m:
        return _SENT_BREAK.split((m.group("subj1") or m.group("subj2") or "").strip())[0]
    m = Q.QUOTE_HEAD.search(c) or Q.QUOTE_HEAD_COLON.search(c)
    if m:
        return _SENT_BREAK.split(m.group("subj").strip())[-1]
    m = Q.INVERTED_HEAD.search(c)
    if m:
        return _SENT_BREAK.split(m.group("subj").strip())[0]
    m = Q.DASH_CREDIT.match(" " + c)
    if m:
        return m.group("subj").strip()
    m = Q.LABEL_HEAD.match(c)
    if m:
        return m.group("subj").strip()
    m = Q.PARTICIPLE_HEAD.search(c)
    if m:
        return ""                        # ", saying:": the subject is the sentence's, before the clause
    return None


class Tracker:
    """The current speaker as the article is read (stage 6)."""

    def __init__(self):
        self.cur = None                  # the speaker dict of the last named speaker, or None

    def set(self, who):
        self.cur = who

    def clear(self):
        self.cur = None


class Resolver:
    def __init__(self, names, stop, mentions, author, outlet, act, solo):
        self.names, self.stop, self.m = names, stop, mentions
        self.author, self.outlet, self.act, self.solo = author, outlet, act, solo
        self.journalist = author or outlet or None
        # relatives and namesakes: non-members named in full in this record who share a member's surname
        self.namesake_surnames = defaultdict(set)
        for a, b, n, mm in mentions.items:
            if mm is None:
                toks = n.split()
                self.namesake_surnames[fold(toks[-1])].add(n)
        self.relatives = {mem for mem in names.members
                          if _tokens(mem) and len(_tokens(mem)) >= 2 and
                          self.namesake_surnames.get(fold(_tokens(mem)[-1]))}
        self.subject = None
        self.subject_args = None
        self.title = ""                  # version 1.2: for act_in_focus
        self.in_quote = None

    def member(self, n, conf, rule, evidence, by_part=False):
        """A member as speaker. Stage 9: one level lower when the member was named by a surname or a first name and a
        non-member with the member's surname (a relative, a namesake) is named in the record."""
        typ = MEMBER if self.names.era(n) == "Nomination-era" else OTHER_ERA
        if by_part and n in self.relatives:
            conf, rule = lower_level(conf), rule + "; a relative or namesake with the same surname is named in the record"
        return dict(type=typ, name=n, member=n, conf=conf, rule=rule, evidence=evidence,
                    pronoun=self.names.members[n].get("pronoun"))

    def other(self, name, conf, rule, evidence):
        return dict(type=OTHER, name=name, member=None, conf=conf, rule=rule, evidence=evidence,
                    pronoun=person_pronoun(name))

    def unknown(self, rule, evidence=""):
        return dict(type=UNKNOWN, name=None, member=None, conf="unknown", rule=rule, evidence=evidence, pronoun=None)

    def collective(self, evidence):
        return dict(type=COLLECTIVE, name=self.act, member=None, conf="medium", rule="4: the band named as speaker",
                    evidence=evidence, pronoun="they")

    # ── version 1.2: "the band said" must mean the act ─────────────────────
    def _act_or_member(self, name):
        """Whether a name is the act, or one of its members (in full, or by a first name or surname of theirs)."""
        if T.act_in(name, self.act):
            return True
        if self.names.mentions(name):
            return True
        for tok in re.findall(r"[A-Z][\w'’.-]+", name):
            low = fold(tok.strip(".'’"))
            if self.names.surname.get(low) or self.names.first.get(low):
                return True
        return False

    def _group_like(self, name, verb):
        """Whether a name that is the subject of a verb can be another group: "The Living End have", "Hoodoo Gurus
        have", "Gardenhead are"; not a person ("Kaya Stewart will open"), nor one capitalised word with a singular verb
        ("Refunds is", "Schneider died")."""
        toks = name.split()
        the = toks[0].lower() == "the"
        core = toks[1:] if the else toks
        if not core:
            return False
        if the:
            return True
        if core[0] in S.FIRST_NAMES_STRICT or core[0].title() in S.FIRST_NAMES_STRICT or core[0] in GENDER:
            return False                              # a person: "Kaya Stewart", "Dusty Hill"
        # without "the", only a plural verb shows a group ("Hoodoo Gurus have"); "Best Coast will" could be anyone
        return verb in ("are", "were", "have")

    def act_in_focus(self, pos):
        """Whether "the band" just before pos can only be the act. The act (or one of its members) is named in this
        paragraph or the two before it, or in the title; and the last name in that stretch that is the subject of a
        verb, outside quotations, is the act or a member rather than another group ("Hoodoo Gurus have publicly
        condemned ...", "The Linda Lindas have paid tribute ...")."""
        text = self.m.text
        lo = text.rfind("\n", 0, pos) + 1             # this paragraph
        back = 0
        while lo > 0 and back < 2:                    # and the two non-empty paragraphs before it
            prev = text.rfind("\n", 0, lo - 1) + 1
            if text[prev:lo].strip():
                back += 1
            lo = prev
        window = text[lo:pos]
        named = T.act_in(window, self.act) or T.act_in(self.title or "", self.act) or \
            any(mm and lo <= a < pos for a, b, n, mm in self.m.items)
        if not named:
            return False
        events = []
        # a heading in capitals starts a section about whoever it names ("AISLES - THINK I'M A GHOST" in a list of new
        # tracks): another act's section unless it names this one
        for m in re.finditer(r"(?m)^[ \t]*([^\n]{3,90})$", window):
            line = m.group(1)
            letters = [ch for ch in line if ch.isalpha()]
            if len(letters) >= 4 and sum(ch.isupper() for ch in letters) / len(letters) >= 0.8 and \
                    re.search(r"\s[-–—]\s", line):
                events.append((m.start(), "act" if T.act_in(line, self.act) else "other"))
        for m in NAME_RUN.finditer(window):
            if self.in_quote and self.in_quote(lo + m.start()):
                continue
            toks = m.group(0).split()
            while toks and toks[0] in SENTENCE_STARTERS and toks[0] != "The":
                toks = toks[1:]
            name = " ".join(toks)
            if not name or name in S.FUNC or name in ("The", "the"):
                continue
            v = STRICT_VERB.match(window, m.end())
            if not v:
                continue
            if self._act_or_member(name):
                events.append((m.start(), "act"))
            elif self._group_like(name, v.group("verb").lower()):
                events.append((m.start(), "other"))
        events.sort()
        return not events or events[-1][1] == "act"

    def appositive(self, a0, a1, evidence):
        """A member named right beside the band in the attribution: "the duo's Dave Stewart", "Carl Palmer of Asia",
        "Knack front man Doug Fieger". The member is the speaker (rule 1)."""
        text = self.m.text
        act = re.escape(re.sub(r"^the\s+", "", self.act, flags=re.I))
        nouns = r"(?i:band|group|duo|trio|quartet|quintet|pair|outfit|" + act + r")"
        # only role words between the band and the name ("Knack front man Doug Fieger"), never a verb ("Musical Youth
        # said Waite Jr was ...": the band speaking about a member)
        roles = (r"(?:front|man|frontman|frontwoman|singer|lead|vocalist|guitarist|bassist|bass|drummer|keyboardist|"
                 r"keyboard|keyboards|player|member|founder|co-founder|founding|leader|songwriter|chief|and|former|"
                 r"original|new|late)")
        cands = []
        # the clause, and what follows it up to the next quotation mark or line break ("according to the duo's Dave
        # Stewart."), never the quotation itself ("Musical Youth said: "... Musical Youth's drummer Frederick Waite")
        tail = re.match(r"[^\"“”\n]{0,80}", text[a1:])
        region = text[a0:a1] + (tail.group(0) if tail else "")
        for m in re.finditer(r"\b" + nouns + r"(?:['’]s)?\s+(?:" + roles + r"\s+){0,4}?(?P<name>" + Q.NAME + r")",
                             region):
            cands.append(m.group("name"))
        for m in re.finditer(r"(?P<name>" + Q.NAME + r")\s+(?:of|from)\s+(?i:the\s+)?" + act + r"\b",
                             text[max(0, a0 - 80):a1]):
            cands.append(m.group("name"))
        for c in cands:
            hit = self.names.mentions(c)
            if len(hit) == 1:
                return self.member(next(iter(hit)), "high", "1: a member's name beside a speech verb", evidence)
        # "the group's Dennis Baptiste", someone who is not a member: another group's member speaking
        m = re.search(r"\b(?i:band|group|duo|trio|quartet|quintet|pair|outfit)['’]s\s+(?:[a-z][\w-]*\s+){0,3}?(?P<name>" +
                      Q.NAME + r")", region)
        if m and m.group("name").split()[0] not in S.FUNC:
            return self.other(m.group("name"), "low", "2: a named non-member beside a speech verb (one of another "
                                                      "group)", evidence)
        return None

    def group(self, s, words, pos, evidence):
        """Rule 4 for a group named by a common noun ("the band said", "the duo explained"). It is the act only when
        it can be: never in a solo act's record, only with "the", a size word that fits the act's line-up, and the
        act in focus; otherwise an unnamed speaker (another group), low."""
        n_noms = sum(1 for mm in self.names.members if self.names.era(mm) == "Nomination-era")
        bare = [re.sub(r"['’]s?$", "", w.strip(",.;:\"“”")) for w in words]
        sizes = [SIZE_WORDS[w] for w in bare if w in SIZE_WORDS]
        why = None
        if self.solo:
            why = "a solo act's record"
        elif words[0] != "the":
            why = "not 'the'"
        elif sizes and any(z != n_noms for z in sizes):
            why = "the size does not fit the act"
        elif not self.act_in_focus(pos):
            why = "the act is not in focus"
        if why:
            return self.other(s, "low", f"2: an unnamed speaker beside a speech verb ({s}: another group, {why})",
                              evidence)
        return self.collective(evidence)

    def resolve_name(self, name, pos, evidence):
        """Rules 1 to 4 for a name beside a speech verb."""
        names = self.names
        clean = re.sub(r"^(?:Mr|Ms|Mrs|Dr|Miss|Sir)\.?\s+", "", name.strip(" ,.:;'’\"")).strip()
        clean = re.sub(r"['’]s$", "", clean)
        low = fold(clean)
        if not clean:
            return None
        hit = names.mentions(clean)
        if hit and min(hit.values()) == 0:
            n = min(hit, key=lambda k: (hit[k], k))
            return self.member(n, "high", "1: a member's name beside a speech verb", evidence)
        # the act's own name: the band speaking as one, or a solo artist
        bare_act = re.sub(r"^the\s+", "", self.act, flags=re.I)
        if T.act_in(clean, self.act) and len(clean) >= len(bare_act) - 1:
            # version 1.2: only the name itself ("Culture Club said"), not "Asia's manager", nor a longer name that
            # holds it ("Sade Sanchez", the singer of L.A. Witch)
            if re.search(r"['’]s?$", name.strip(" ,.:;\"“”")):
                return None
            if len(re.sub(r"^the\s+", "", clean, flags=re.I)) <= len(bare_act) + 2:
                if self.solo:
                    return self.member(self.solo, "high", "1: a member's name beside a speech verb", evidence)
                # the act's name is also its singer's name: "Sade says" is Sade Adu, as in stage 3
                singer = sorted(names.form.get(fold(bare_act), ()))
                if len(singer) == 1:
                    return self.member(singer[0], "medium", "1: the act's name, which is also its singer's name",
                                       evidence)
                return self.collective(evidence)
        if low in names.form and names.form[low]:
            cands = sorted(names.form[low])
            if len(cands) == 1:
                return self.member(cands[0], "high", "1: a member's name beside a speech verb", evidence)
        toks = clean.split()
        if len(toks) >= 2:
            first, last = fold(toks[0]), fold(toks[-1])
            both = names.first.get(first, set()) & names.surname.get(last, set())
            if len(both) == 1:
                return self.member(next(iter(both)), "high", "1: a member's name beside a speech verb", evidence)
            if toks[0] in S.FIRST_NAMES or toks[0].title() in S.FIRST_NAMES:
                return self.other(clean, "high", "2: a named non-member beside a speech verb", evidence)
            return self.other(clean, "low", "2: a named non-member beside a speech verb (name not checked)", evidence)
        # one word: a surname or first name
        if clean.title() in self.stop or clean in S.FUNC:
            return None
        cands_m = sorted((names.surname.get(low, set()) if low not in names.no_bare else set()) |
                         names.first.get(low, set()))
        cands_o = sorted({n for a, b, n, mm in self.m.items if mm is None and
                          (fold(n.split()[-1]) == low or fold(n.split()[0]) == low)})
        if cands_m and not cands_o:
            if len(cands_m) == 1:
                c = cands_m[0]
                if self.m.full_named(c, pos) or self.m.full_named(c):
                    return self.member(c, "high", "1: a member's name beside a speech verb (a name the record gives "
                                                  "in full)", evidence, by_part=True)
                return dict(type=AMBIGUOUS, name=clean, member=None, conf="low",
                            rule="3: a member's one name, but the record never names the member in full",
                            evidence=evidence, pronoun=None)
        if cands_m or len(cands_o) > 1:
            # rule 3: a shared name goes to the person the title or introduction is about, if one of them ("Cherry says"
            # in "Neneh Cherry resets her career"), else to the nearest earlier full name
            w = self.interviewee()
            if w and w["kind"] == "member" and len(w["who"]) == 1 and w["who"][0] in cands_m:
                return self.member(w["who"][0], "medium", "3: a shared name, the member the title or introduction is "
                                                          "about", evidence, by_part=True)
            if w and w["kind"] == "other" and len(w["who"]) == 1 and w["who"][0] in cands_o:
                return self.other(w["who"][0], "medium", "3: a shared name, the person the title or introduction is "
                                                         "about", evidence)
            prev = [x for x in self.m.before(pos) if (x[3] in cands_m) or (x[3] is None and x[2] in cands_o)]
            if prev:
                a, b, n, mm = prev[-1]
                if mm:
                    return self.member(mm, "medium", "3: a shared name, settled by the nearest earlier full name",
                                       evidence, by_part=True)
                return self.other(n, "medium", "3: a shared name, settled by the nearest earlier full name", evidence)
            if cands_m and not cands_o:
                eras = {names.era(c) for c in cands_m}
                if eras == {"Nomination-era"}:
                    return dict(type=MEMBERS_UNSPEC, name=" / ".join(cands_m), member=None, conf="medium",
                                rule="3: a name shared by members, none named in full before", evidence=evidence,
                                pronoun=None)
            return dict(type=AMBIGUOUS, name=" / ".join(cands_m + cands_o), member=None, conf="low",
                        rule="3: a shared name, not settled", evidence=evidence, pronoun=None)
        if len(cands_o) == 1:
            return self.other(cands_o[0], "high", "2: a named non-member beside a speech verb (named in full before)",
                              evidence)
        # a single capitalised word that names nobody the record names in full
        if re.fullmatch(r"[A-Z][\w'’.-]+", clean) and clean not in S.FUNC and clean not in NOT_SPEAKERS:
            return self.other(clean, "low", "2: a named non-member beside a speech verb (one name)", evidence)
        return None

    def resolve(self, subj, pos, tracker, evidence, sentence_before=""):
        """Rules 1 to 6 for the speaker words of a clause."""
        if subj is None:
            return None
        s = subj.strip()
        low = s.lower()
        if low in PRONOUNS:
            return self.pronoun(low, pos, tracker, evidence)
        if s == "":
            # ", saying:": the subject of the sentence
            m = re.search(r"(" + Q.NAME + r"|\b[Hh]e\b|\b[Ss]he\b|\b[Tt]hey\b)[^.!?]*$", sentence_before)
            if m:
                return self.resolve(m.group(1), pos, tracker, evidence, "")
            return None
        name, before = named_part(s)
        if name and before:
            bw = [w.strip("'’").lower() for w in before.split()]
            if len(name.split()) == 1 and any(w in RELATION_WORDS or w.rstrip("s") in RELATION_WORDS for w in bw):
                return self.other(s, "low", "2: an unnamed speaker beside a speech verb (" + s + ")", evidence)
            r = self.resolve_name(name, pos, evidence)
            if r and r["type"] == COLLECTIVE and bw and bw[-1] in ("a", "an", "another", "one"):
                return None                   # version 1.2: "a Musical Youth potted history" is not the band
            if r:
                return r
        words = low.split()
        if words and words[0] in DETERMINERS:
            head = words[-1].rstrip("’'")
            # version 1.2: "her band member" is one person, not the band
            if head in COLLECTIVE_WORDS or (len(words) >= 2 and words[1] in COLLECTIVE_WORDS and head not in ROLE_WORDS):
                return self.group(s, words, pos, evidence)
            if head in ROLE_WORDS:
                return self.pronoun("member" if T.act_in(s, self.act) else "role", pos, tracker, evidence)
            return self.other(s, "low", "2: an unnamed speaker beside a speech verb (" + s + ")", evidence)
        return self.resolve_name(name or s, pos, evidence)

    def pronoun(self, p, pos, tracker, evidence):
        """Rule 5: a pronoun (or 'the singer') beside a speech verb goes to the current speaker. He or she must match a
        member's pronoun; for anyone else the gender is not checked (one level lower). Never surer than the current
        speaker."""
        if p == "i":
            return dict(type=JOURNALIST, name=self.journalist, member=None, conf="medium",
                        rule="5: 'I' in the narration is the writer", evidence=evidence, pronoun=None)
        cur = tracker.cur
        if p in ("they", "we"):
            if cur and cur["type"] in (COLLECTIVE, MEMBERS_UNSPEC):
                return dict(cur, conf=min_level("medium", cur["conf"]), rule="5: a pronoun for the current speaker",
                            evidence=evidence)
            return None
        if not cur or cur["type"] == COLLECTIVE or (p == "member" and cur["type"] not in (MEMBER, OTHER_ERA)):
            return None
        if p in ("he", "she") and cur.get("pronoun") in ("he", "she") and cur["pronoun"] != p:
            return None
        checked = cur.get("pronoun") in ("he", "she") or p in ("role", "member")
        r = dict(cur, rule="5: a pronoun for the current speaker" + ("" if checked else " (gender not checked)"),
                 evidence=evidence)
        r["conf"] = min_level("medium" if checked else "low", cur["conf"])
        return r

    def fits(self, x, p):
        """Whether a person named in the narration can be the 'he', 'she', 'member' or 'role' of an attribution: a
        member's pronoun must fit; another person's must fit when their first name is clearly a man's or a woman's;
        never the writer or a byline."""
        mm = x[3]
        if mm:
            pr = self.names.members[mm].get("pronoun")
            return not (p in ("he", "she") and pr and pr != p)
        if self.m.incidental(x[0], x[1]):
            return False
        if p == "member":
            return False
        g = person_pronoun(x[2])
        if p in ("he", "she") and g and g != p:
            return False
        if self.author and fold(x[2]) == fold(self.author):
            return False
        if re.search(r"(?i)(?:\bby|words|photos?|interview(?:ed)? by|written by)[:\s]+$",
                     self.m.text[max(0, x[0] - 20):x[0]]):
            return False
        return True

    def last_named(self, pos, limit, p):
        """5b: the person the narration last spoke about: the latest name at the start of a sentence, else the latest
        name; a member's pronoun must fit, other people's genders are unknown (low confidence)."""
        prev = self.m.before(pos, limit)
        if not prev:
            return None
        text = self.m.text
        cands = [x for x in prev if self.fits(x, p)]
        if not cands:
            return None
        # a name in the possessive ("John Lennon's son", "Louis Theroux's podcast") only when there is no other
        plain = [x for x in cands if not re.match(r"['’]s?\b", text[x[1]:x[1] + 2])]
        cands = plain or cands
        pick = ([x for x in cands if self.m.verb_subject(x[1])] or
                [x for x in cands if self.m.subject_position(x[0])] or cands)[-1]
        a, b, n, mm = pick
        if mm:
            return self.member(mm, "medium", "5b: a pronoun for the person the narration is about", "",
                               by_part=len(text[a:b].split()) == 1)
        return self.other(n, "low", "5b: a pronoun for the person the narration is about (not a member)", "")

    def interviewee(self):
        """Who the title and introduction present the piece as being about (stage 3's interviewee rule)."""
        if self.subject is None and self.subject_args:
            self.subject = T.interviewee(*self.subject_args)
        return self.subject

    def article_subject(self, p):
        """5c: the member the title is about, when the pronoun fits."""
        w = self.interviewee()
        if w and w["kind"] == "member":
            mm = w["who"][0]
            pr = self.names.members[mm].get("pronoun")
            if not (p in ("he", "she") and pr and pr != p):
                return self.member(mm, "low", "5c: a pronoun for the member the title is about", "")
        return None


PARTICIPLE_WORDS = {"speaking", "talking", "asked", "interviewed", "writing", "looking", "reflecting", "recalling",
                    "explaining", "describing", "discussing", "reminiscing", "chatting", "sitting", "standing", "after",
                    "before", "when", "while", "since", "during", "in", "on", "at", "for", "with"}


def sentence_subject(text, a0, before):
    """The subject of the sentence an attribution such as ", saying:" or "and said," belongs to: the last person in the
    sentence who is followed by a verb ("Speaking to Rolling Stone, Knopfler addressed the issue and said:"), else the
    last he, she or they."""
    st = max(before.rfind("\n"), max((before.rfind(c + " ") for c in ".!?"), default=-1),
             before.rfind("\""), before.rfind("”"))
    sent = before[st + 1:]
    base = a0 - len(before) + st + 1
    best, pron = None, None
    for m in re.finditer(r"\b(?:" + Q.NAME_PLUS + r")|\b(?:[Hh]e|[Ss]he|[Tt]hey)\b", sent):
        w = m.group(0)
        if w.split()[0].lower() in PARTICIPLE_WORDS or w in S.FUNC:
            continue
        if w.lower() in ("he", "she", "they"):
            pron = w
            if VERB_AFTER.match(text, base + m.end()):
                best = w
            continue
        if VERB_AFTER.match(text, base + m.end()):
            best = w
    return best or pron


def pronoun_kind(subj, act):
    """'he', 'she', 'member' (a description that names the act: "the ex-Eurythmics singer"), 'role' ("the singer") or
    None."""
    low = subj.lower().strip()
    if low in ("he", "she"):
        return low
    words = low.split()
    if words and words[0] in DETERMINERS and words[-1].rstrip("’'") in ROLE_WORDS:
        return "member" if T.act_in(subj, act) else "role"
    return None


# ── stage 7: second-hand self narrative ──────────────────────────────────────
SECOND_HAND = re.compile(r"(?:^|(?<=[.!?]\s)(?<![A-Z]\.[A-Z]\.\s)(?<!\b[A-Z]\.\s)|(?<=[.!?][\"”’']\s)|(?<=\n))(?P<sent>(?:[^.!?\n\"“”]{0,80}?\b)?(?P<subj>" + Q.SUBJ +
                         r")(?:\s*,[^,\n]{1,60},)?\s+(?:(?:also|later|once|now|still|has|had|have|would|will|previously|recently|"
                         r"repeatedly|often|frequently|openly|publicly|famously)\s+){0,3}"
                         r"(?:says|said|told|tells|recalls|recalled|explains|explained|admits|admitted|insists|"
                         r"insisted|claims|claimed|reckons|reckoned|reveals|revealed|(?:maintains|maintained)(?=\s+that\b)|"
                         r"notes|noted|adds|added|remembers|remembered|describes|described|"
                         r"talks about|talked about|opens up|opened up|discusses|discussed|mentions|mentioned|announces|"
                         r"announced|confirms|confirmed|denies|denied|acknowledges|acknowledged|argues|argued|stresses|"
                         r"stressed|suggests|suggested|warns|warned|promises|promised|vows|vowed|hints|hinted|jokes|"
                         r"joked|(?:writes|wrote)(?=\s+(?:that\b|on\s|in\s+(?:a|an|his|her|their|the)\s+(?:post|statement|letter|"
                         r"blog|memoir|book|autobiography|note|message|tweet|essay|column|caption)))|tweets|tweeted|"
                         r"reflects|reflected|reminisces|reminisced|concedes|"
                         r"conceded|complains|complained|recounts|recounted|credits|credited|praises|praised)"
                         r"\b(?!\s+(?:nothing|no one|nobody)\b)(?![^.!?\n]*[\"“”])[^.!?\n]*[.!?]?)")
SECOND_HAND_MID = re.compile(r"(?:,|—|–)[ \t]*(?P<subj>" + Q.SUBJ + r")[ \t]+(?:says|said|recalls|recalled|explains|"
                             r"explained|adds|added|admits|admitted|notes|noted|insists|insisted|remembers|"
                             r"remembered|reckons|reckoned|reflects|reflected|laughs|laughed|jokes|joked|"
                             r"maintains|maintained|claims|claimed)[ \t]*(?=[,.;:!?—–]|$)")
_SENTENCE = re.compile(r"[^.!?\n]+(?:[.!?]+|$)")
ACCORDING = re.compile(r"(?:^|(?<=[.!?]\s)|(?<=\n))(?P<sent>[^.!?\n\"“”]*\b[Aa]ccording to (?P<subj>" + Q.NAME +
                       r"|the band|the group)\b[^.!?\n\"“”]*[.!?]?)")


# ── stage 8: interview date ──────────────────────────────────────────────────
YEAR = r"(?:19[5-9]\d|20[0-2]\d)"
_SPEAKING = (r"(?:told|tells|said|says|recalled|recalls|explained|explains|admitted|admits|remembered|reflected|added|"
             r"insisted|wrote|confessed|revealed|joked|declared|claimed|noted|commented|remarked|observed|recounted|"
             r"reminisced|mused|quipped|stated|summed up)")
_TO_WHOM = (r"(?:\s+(?:the|a|an|to|me|us|him|her|them|magazine|newspaper|paper|radio|television|TV|reporter|"
            r"interviewer|journalist|writer|fans|audience|crowd|[A-Z][\w'’&.-]*)){0,5}")
# the year a quotation was given, bound to a speaking verb in the sentence that names the speaker: "told NME in 1983",
# "said in a 1985 interview", "Interviewed by Classic Rock in 2019, Knopfler claimed"
STATED_YEAR = re.compile(
    r"\b" + _SPEAKING + r"\b" + _TO_WHOM + r"\s+(?:back\s+)?in\s+(?:early\s+|late\s+|mid-?\s*)?(?P<y1>" + YEAR +
    r")\b(?![-–/]\d)"
    r"|\b(?:[Ii]n|[Dd]uring)\s+(?:a|an|his|her|their|one|this|that|the)\s+(?:(?:early|late|mid|rare|candid|radio|TV|"
    r"television|phone|telephone|[A-Z][\w'’&.-]*)\s+){0,3}(?P<y2>" + YEAR + r")\s+(?:[\w'’&.-]+\s+){0,3}?interview\b"
    r"|\b[Ii]nterview(?:ed)?\b(?:(?!\b(?:who|which|whose|died|passed|born|released|formed|split)\b)[^.!?\n,;()]){0,40}?"
    r"\bin\s+(?:early\s+|late\s+|mid-?\s*)?(?P<y3>" + YEAR + r")\b(?![-–/]\d)")
_MONTHS = r"(?:january|february|march|april|may|june|july|august|september|october|november|december)"
# a Q&A or first-person record whose title or introduction says when its own interview took place: "This interview was
# originally published August 2, 2013", "At the time of this interview in 1992", "This 1998 interview"
QA_STATED = re.compile(
    r"(?i)\bthis\s+(?:[\w'’&.-]+\s+){0,3}?(?P<a>" + YEAR + r")\s+(?:[\w'’&.-]+\s+){0,3}?"
    r"(?:interview|conversation|q&a|chat)\b"
    r"|\b(?:this|our)\s+(?:[\w'’&.-]+\s+){0,4}?(?:interview|conversation|q&a|chat)\b[^.!?\n]{0,80}?\b(?:in|from|on|"
    r"published|ran|aired|broadcast|conducted|recorded|done|held)\s+(?:(?:early|late|mid|some\s*time\s+in|back\s+in|"
    r"the\s+(?:spring|summer|autumn|fall|winter)\s+of)\s+)?(?:(?:a|an|the)\s+)?(?:(?:" + _MONTHS +
    r"|spring|summer|autumn|fall|winter)\s+)?(?:\d{1,2}(?:st|nd|rd|th)?,?\s+)?(?P<b>" + YEAR + r")\b(?![-–/]\d)")
QA_TITLE = re.compile(r"(?i)\b(?:the|this|a|an)\s+(?:[\w'’&.-]+\s+){0,3}?(?P<a>" + YEAR + r")\s+(?:[\w'’&.-]+\s+){0,3}?"
                      r"interview\b")
_SENT_END = re.compile(r"[.!?](?=\s)|\n|[\"“”]")
_SENT_END_NARR = re.compile(r"(?<![A-Z])[.!?](?=\s)|\n")


def attribution_sentence(text, a0, a1, q0, q1):
    """The sentence of an attribution clause: from the sentence start to the clause's end when the clause comes before
    the quotation; from the clause to the sentence's end when it comes after."""
    if a1 <= q0 + 1:
        lo = max(text.rfind("\n", 0, a0), max(text.rfind(c + " ", 0, a0) for c in ".!?"))
        lo = max(lo, text.rfind("\"", 0, a0), text.rfind("”", 0, a0))
        return text[lo + 1:a1]
    m = _SENT_END.search(text, a1)
    return text[a0:m.end() if m else len(text)]


def stated_year(sentences, pub_year):
    """The first year a sentence gives for when the words were said, if it is not after publication."""
    for s in sentences:
        for m in STATED_YEAR.finditer(s or ""):
            y = m.group("y1") or m.group("y2") or m.group("y3")
            if not pub_year or int(y) <= pub_year:
                return y, s[max(0, m.start() - 40):m.end() + 20].strip()
    return None


def qa_stated_year(title, intro, pub_year):
    for rx, s in ((QA_TITLE, title or ""), (QA_STATED, title or ""), (QA_STATED, intro or "")):
        for m in rx.finditer(s):
            y = m.group("a") if "a" in rx.groupindex and m.group("a") else m.group("b") if "b" in rx.groupindex else None
            if y and (not pub_year or int(y) <= pub_year):
                return y, s[max(0, m.start() - 30):m.end() + 20].strip()
    return None


def _pub_year(pub_date):
    m = re.match(r"(\d{4})", pub_date or "")
    return int(m.group(1)) if m else None


# ── one record: spans to turns ───────────────────────────────────────────────
STAGE3_TYPES = {T.MEMBER: MEMBER, T.MEMBERS_UNSPEC: MEMBERS_UNSPEC, T.OTHER_ERA: OTHER_ERA,
                T.OTHER_ERA_UNSPEC: OTHER_ERA_UNSPEC, T.INTERVIEWER: INTERVIEWER, T.OTHER: OTHER, T.UNKNOWN: UNKNOWN,
                T.AMBIGUOUS: AMBIGUOUS}


def _same(a, b):
    return a and b and a["type"] == b["type"] and a.get("name") == b.get("name")


def _embedded(text, s, e, conv):
    """Whether a quotation holds someone else's words in the inner marks (single inside double in American style,
    double inside single in British), at least three words long."""
    if e - s < 8:
        return False
    if conv == "british":
        ms = Q._double_marks(text, s + 1, e - 1)
        pairs = [(ms[k][0], ms[k + 1][0]) for k in range(0, len(ms) - 1, 2)]
    else:
        ms = [x for x in Q._single_marks(text, s + 1, e - 1)]
        pairs, op = [], None
        for i, t in ms:
            if t in ("O", "?") and op is None:
                op = i
            elif t in ("C", "?") and op is not None:
                pairs.append((op, i))
                op = None
    return any(len(text[a + 1:b].split()) >= 3 for a, b in pairs)


def _split_second_hand(text, x, R, tracker_cur, para_of, para_starts):
    """Cut a narration span at sentences that report a member's words without quotation marks."""
    out, cur = [], x["start"]
    seg = text[x["start"]:x["end"]]
    hits = []
    found = []
    for rx in (SECOND_HAND, ACCORDING):
        for m in rx.finditer(seg):
            found.append((m.group("subj"), x["start"] + m.start("sent"), x["start"] + m.end("sent")))
    for m in _SENTENCE.finditer(seg):
        sent = m.group(0)
        if re.search(r"[\"“”]", sent):
            continue
        mm = SECOND_HAND_MID.search(sent)
        if mm:
            lead = len(sent) - len(sent.lstrip())
            found.append((mm.group("subj"), x["start"] + m.start() + lead, x["start"] + m.end()))
    for subj, st, en in found:
        low = subj.lower().strip()
        res = None
        if low in ("he", "she"):
            if tracker_cur:
                if tracker_cur["type"] in (MEMBER, OTHER_ERA) and tracker_cur.get("pronoun") == low:
                    res = dict(tracker_cur, conf=min_level("medium", tracker_cur["conf"]),
                               rule="second-hand: a pronoun for the current speaker")
            else:
                para = para_of(st)
                r = R.last_named(st, para_starts[max(0, para - 2)], low)
                if r and r["type"] in (MEMBER, OTHER_ERA):
                    res = dict(r, rule="second-hand: " + r["rule"].split(": ", 1)[-1])
                elif r is None:
                    r = R.article_subject(low)
                    if r and r["type"] in (MEMBER, OTHER_ERA):
                        res = dict(r, rule="second-hand: " + r["rule"].split(": ", 1)[-1])
        elif low.startswith(("the band", "the group")):
            # version 1.2: only when "the band" can be the act (see Resolver.group)
            g = R.group(subj.strip(), low.split(), st, subj)
            if g["type"] == COLLECTIVE:
                res = dict(g, rule="second-hand: the band named")
        else:
            name, before = named_part(subj)
            r = R.resolve_name(name or subj, st, subj) if (name or subj) else None
            if r and r["type"] in (MEMBER, OTHER_ERA, COLLECTIVE):
                res = dict(r, rule="second-hand: " + r["rule"].split(": ", 1)[-1])
        if res and res["type"] in (MEMBER, OTHER_ERA, COLLECTIVE):
            hits.append((st, en, res))
    hits.sort(key=lambda h: (h[0], h[1]))
    for st, en, res in hits:
        if st < cur or en <= st:
            continue
        if st > cur:
            out.append(dict(x, start=cur, end=st))
        out.append(dict(start=st, end=en, kind="second-hand", evidence=text[st:en].strip()[:120], attribution=None,
                        reported=res))
        cur = en
    if cur < x["end"]:
        out.append(dict(x, start=cur, end=x["end"]))
    return out


def title_or_fragment(q):
    """Version 1.2: a quotation given to the band that is a title, a lyric or a fragment rather than something said: one
    or two words ("Fire", "Club Zero"), title case ("Closer to Fine"), capitals, or lines run together with slashes."""
    s = q.strip().strip("\"'“”‘’ ,.;:!?…")
    words = re.findall(r"[A-Za-z0-9][\w'’-]*", s)
    if len(words) <= 2 or Q.title_case(s):
        return True
    letters = [c for c in s if c.isalpha()]
    if letters and sum(c.isupper() for c in letters) / len(letters) > 0.8:
        return True
    return bool(Q.LYRIC_SLASHES.search(s))


def _is_gap(text, x):
    """A narration span with no words (a line break or a comma between two quotations): never a turn of its own."""
    return x["kind"] == "narration" and not re.search(r"\w", text[x["start"]:x["end"]])


def record_turns(rid, text, title, act, names, stop, author_raw, outlet, pub_date, prose_spans, block_spans=(),
                 solo=None, conv="american"):
    """All spans and turns of one record, in reading order."""
    author = T.person_from_author(author_raw)
    journalist = author or outlet or None
    mentions = Mentions(text, title, names, stop, act, solo)
    R = Resolver(names, stop, mentions, author, outlet, act, solo)
    R.subject_args = (names, stop, title, text[:1500], act, author_raw, (), "")
    tracker = Tracker()
    para_starts = [0] + [m.end() for m in re.finditer("\n", text)]
    pub_year = _pub_year(pub_date)

    def para_of(pos):
        return bisect.bisect_right(para_starts, pos) - 1

    # stage 5 and 6 over the prose spans, in order. Names inside quotations are someone's words, not the narration's:
    # they neither clear the current speaker nor stand for a pronoun.
    prose = sorted((dict(x) for x in prose_spans), key=lambda x: x["start"])
    quoted = [(x["start"], x["end"]) for x in prose if x["kind"] in ("speech", "non-speech", "undecided", "unresolved")]
    quoted += [(y["start"], y["end"]) for y in block_spans]
    quoted.sort()
    starts_q = [q[0] for q in quoted]

    def in_quote(pos):
        k = bisect.bisect_right(starts_q, pos) - 1
        return k >= 0 and quoted[k][0] <= pos < quoted[k][1]

    mentions.items = [m for m in mentions.items if not in_quote(m[0])]
    R.title, R.in_quote = title or "", in_quote

    # an undecided quotation after a name label is speech when the label names a member or the band
    fixed = []
    for x in prose:
        if x["kind"] == "undecided" and x.get("label"):
            l0, l1 = x["label"]
            subj = clause_subject(text[l0:l1])
            r = R.resolve(subj, x["start"], Tracker(), text[l0:l1].strip()) if subj else None
            if r and r["type"] in (MEMBER, MEMBERS_UNSPEC, COLLECTIVE, OTHER_ERA):
                x["kind"], x["attribution"] = "speech", (l0, l1)
                x["evidence"] = "a name label before it that names a member or the band"
                fixed.append((l0, l1))
    for l0, l1 in fixed:
        for k, y in enumerate(prose):
            if y["kind"] == "narration" and y["start"] <= l0 < y["end"] and l1 <= y["end"]:
                parts = [dict(y, end=l0)] if l0 > y["start"] else []
                parts.append(dict(start=l0, end=l1, kind="attribution", evidence="", attribution=None))
                if l1 < y["end"]:
                    parts.append(dict(y, start=l1))
                prose[k:k + 1] = parts
                break

    attr_spans = defaultdict(list)
    for y in prose:
        if y["kind"] == "attribution":
            attr_spans[(y["start"], y["end"])].append(y)
    last_quote_para, last_speech_end, last_tail_open = None, None, False
    spans = []
    for x in prose:
        if x["kind"] == "speech":
            para = para_of(x["start"])
            if last_quote_para is not None and para - last_quote_para > 2:
                tracker.clear()
            if tracker.cur and last_speech_end is not None:
                named = [m for m in mentions.before(x["start"], last_speech_end)
                         if not (x["attribution"] and x["attribution"][0] <= m[0] < x["attribution"][1])]
                cur_id = tracker.cur.get("member") or tracker.cur.get("name")
                # the rest of the previous attribution's sentence ("..., Stewart tells Billboard, referring to his and
                # Annie Lennox's first appearance") is not a narration naming someone else
                rest_end = last_speech_end
                if last_tail_open:
                    se = _SENT_END_NARR.search(text, last_speech_end)
                    rest_end = se.end() if se else last_speech_end
                others = [m for m in named if (m[3] or m[2]) != cur_id and m[0] >= rest_end and
                          not mentions.incidental(m[0], m[1])]
                if others:
                    pk = pronoun_kind(clause_subject(text[x["attribution"][0]:x["attribution"][1]]) or "", act) \
                        if x["attribution"] else None
                    if pk in ("he", "she", "member"):
                        # a pronoun goes back to the current speaker unless the narration has named someone else it
                        # can stand for
                        if any(R.fits(m, pk) for m in others):
                            tracker.clear()
                    elif x["attribution"]:
                        tracker.clear()
                    elif any(mentions.subject_position(m[0]) or mentions.verb_subject(m[1]) for m in others):
                        # an unattributed quotation continues the current speaker unless someone else has become the
                        # subject of the narration
                        tracker.clear()
            res = None
            if x["attribution"]:
                a0, a1 = x["attribution"]
                clause = text[a0:a1]
                subj = clause_subject(clause)
                before = text[max(para_starts[para_of(a0)], a0 - 200):a0]
                if subj == "":
                    subj = sentence_subject(text, a0, before)
                res = R.resolve(subj, x["start"], tracker, clause.strip(), before)
                if res is None and subj is not None:
                    # no current speaker, or the pronoun does not fit them: the person the narration is about
                    p = pronoun_kind(subj, act)
                    if tracker.cur and p not in ("he", "she"):
                        p = None
                    if p:
                        # names since the last quotation, at most two paragraphs back
                        limit = last_speech_end if last_speech_end is not None else 0
                        window = max(limit, para_starts[max(0, para - 2)])
                        res = R.last_named(x["start"], window, p) or R.article_subject(p)
                        if res:
                            res["evidence"] = clause.strip()
                x["sentence"] = attribution_sentence(text, a0, a1, x["start"], x["end"])
            elif tracker.cur and tracker.cur["type"] not in (JOURNALIST,):
                res = dict(tracker.cur, rule="5: continues the current speaker", evidence="",
                           conf=min_level("medium", tracker.cur["conf"]))
            else:
                r = R.article_subject("member")
                if r:
                    res = dict(r, rule="5c: an unattributed quotation in a piece about one member")
            if res is None:
                res = R.unknown("6: no rule applied", text[x["attribution"][0]:x["attribution"][1]].strip()
                                if x["attribution"] else "")
            # version 1.2: a member named beside the band ("the duo's Dave Stewart", "Carl Palmer of Asia") is the
            # speaker; and what the band (or "Sade") is given must be something said, not a title, a lyric or a fragment
            band_named = res["type"] == COLLECTIVE or res["rule"].startswith("1: the act's name")
            # a quotation split by its attribution ('"Sometimes," Sade tells me, "I think ..."') is not a fragment
            split = bool(x["attribution"] and x["attribution"][0] >= x["end"] - 1 and
                         re.match(r"[ \t]*[,:]?[ \t]*[\"“'‘]", text[x["attribution"][1]:x["attribution"][1] + 6]))
            fragment = band_named and not split and title_or_fragment(text[x["start"]:x["end"]])
            if band_named and not fragment and x["attribution"] and res["type"] == COLLECTIVE:
                ap = R.appositive(x["attribution"][0], x["attribution"][1], res["evidence"])
                if ap:
                    res = ap
            if fragment:
                # not speech, but the band (or the singer) is still who the narration is talking about
                x["kind"] = "non-speech"
                x["evidence"] = "a title, lyric or fragment given to the band (version 1.2: not speech)"
                tracker.set(res)
                last_quote_para = para
                last_speech_end = max(x["end"], x["attribution"][1] if x["attribution"] else 0)
                last_tail_open = False
                if x["attribution"]:
                    for y in attr_spans.get(tuple(x["attribution"]), []):
                        y["kind"] = "narration"
                    x["attribution"] = None
                spans.append(x)
                continue
            if res["type"] in (MEMBER, OTHER_ERA, OTHER, COLLECTIVE, MEMBERS_UNSPEC) and \
                    not res["rule"].startswith("5: continues"):
                tracker.set(res)
            elif res["type"] in (UNKNOWN, AMBIGUOUS):
                tracker.clear()
            x["speaker"] = res
            x["embedded"] = _embedded(text, x["start"], x["end"], conv)
            last_quote_para = para
            last_speech_end = max(x["end"], x["attribution"][1] if x["attribution"] else 0)
            last_tail_open = bool(x["attribution"] and x["attribution"][0] >= x["end"] and
                                  not re.search(r"[.!?]\s*$", text[x["attribution"][0]:x["attribution"][1]]))
            spans.append(x)
        elif x["kind"] == "narration":
            pieces = _split_second_hand(text, x, R, tracker.cur, para_of, para_starts)
            spans.extend(pieces)
            for y in pieces:
                if y["kind"] == "second-hand" and y["reported"]["type"] in (MEMBER, OTHER_ERA, COLLECTIVE):
                    tracker.set(dict(y["reported"], rule="5: the member a second-hand sentence reports"))
                    last_quote_para = para_of(y["start"])
                    last_speech_end = y["end"]
                    last_tail_open = False
        else:
            spans.append(x)
    # the Q&A and first-person blocks keep their stage 3 speakers
    for y in block_spans:
        z = dict(start=y["start"], end=y["end"], kind="block " + y["kind"], evidence=y.get("rule") or "",
                 attribution=None)
        if y["kind"] == "turn":
            # version 1.2: a passage inside a Q&A that stage 2 left as prose carries its own rule (see qa_passages)
            z["speaker"] = dict(type=STAGE3_TYPES.get(y["speaker_type"], y["speaker_type"]), name=y["speaker"],
                                member=y.get("member"), conf=y["confidence"],
                                rule=y["rule"] if y.get("passage") else "stage 3: " + (y["rule"] or ""),
                                evidence=y.get("evidence") or y.get("label") or "", pronoun=None)
            z["question"] = y.get("question")
            z["embedded"] = bool(y.get("embedded_quotes", 0))
            z["speech_start"] = y.get("speech_start", y["start"])
        spans.append(z)
    spans.sort(key=lambda x: x["start"])

    # stage 7: turns
    turns = []

    def new_turn(kind, who, start, end, evidence="", reported=None):
        turns.append(dict(kind=kind, who=who, start=start, end=end, parts=[(start, end)], evidence=[evidence] if
                          evidence else [], reported=reported, undecided=False, embedded=False, sentences=[],
                          question=False))
        return turns[-1]

    def only_gaps(i, j):
        return all(spans[q].get("gap") for q in range(i, j))

    pending_attr = []
    for k, x in enumerate(spans):
        kind = x["kind"]
        if _is_gap(text, x):
            x["gap"] = True
            x["turn"] = None
            continue
        if kind in ("unresolved", "block sung", "block non-speech"):
            x["turn"] = None
            pending_attr = []
            continue
        if kind == "speech":
            who = x["speaker"]
            last = turns[-1] if turns else None
            joinable = last is not None and last["kind"] == "speech" and _same(last["who"], who) and \
                not last.get("block") and \
                all(s["kind"] == "attribution" or s.get("gap") for s in spans[last["_last"] + 1:k])
            if joinable:
                last["parts"].append((x["start"], x["end"]))
                last["end"] = x["end"]
                t = last
            else:
                t = new_turn("speech", who, x["start"], x["end"])
            if x["attribution"]:
                t["evidence"].append(text[x["attribution"][0]:x["attribution"][1]].strip())
                t["sentences"].append(x.get("sentence") or "")
            t["_last"] = k
            t["embedded"] = t["embedded"] or x.get("embedded", False)
            x["turn"] = len(turns) - 1
            for p in pending_attr:
                p["turn"] = len(turns) - 1
            pending_attr = []
            continue
        if kind == "attribution":
            # belongs to the speech next to it
            prev = next((spans[j] for j in range(k - 1, -1, -1) if not spans[j].get("gap")), None)
            if prev is not None and prev["kind"] == "speech" and prev["attribution"] and \
                    prev["attribution"][0] <= x["start"] < prev["attribution"][1]:
                x["turn"] = prev.get("turn")
            else:
                pending_attr.append(x)
                x["turn"] = None
            continue
        if kind == "second-hand":
            rep = x["reported"]
            t = new_turn("second-hand", dict(type=JOURNALIST, name=journalist, member=None, conf=rep["conf"],
                                             rule=rep["rule"], evidence=x["evidence"], pronoun=None),
                         x["start"], x["end"], x["evidence"], reported=rep.get("member") or rep.get("name"))
            t["sentences"].append(text[x["start"]:x["end"]])
            t["_last"] = k
            x["turn"] = len(turns) - 1
            continue
        if kind == "block turn":
            who = x["speaker"]
            t = new_turn("speech", who, x["start"], x["end"], who.get("evidence") or "")
            t["block"] = True
            t["question"] = bool(x.get("question"))
            t["embedded"] = x.get("embedded", False)
            t["speech_start"] = x.get("speech_start", x["start"])
            t["_last"] = k
            x["turn"] = len(turns) - 1
            for p in pending_attr:
                p["turn"] = len(turns) - 1
            pending_attr = []
            continue
        if kind == "block label":
            pending_attr.append(x)
            x["turn"] = None
            continue
        # narration, non-speech and undecided quotations, block headings and narration: the journalist's turn
        last = turns[-1] if turns else None
        if last is not None and last["kind"] == "narration" and only_gaps(last["_last"] + 1, k) and not pending_attr:
            last["end"] = x["end"]
            last["parts"][-1] = (last["parts"][-1][0], x["end"])
            t = last
        else:
            t = new_turn("narration", dict(type=JOURNALIST, name=journalist, member=None, conf="high",
                                           rule="7: the writer's own words", evidence="", pronoun=None),
                         x["start"], x["end"])
        if kind == "undecided":
            t["undecided"] = True
        t["_last"] = k
        x["turn"] = len(turns) - 1
    # labels and clauses waiting for the turn they open; gaps go with the turn before them
    for k, x in enumerate(spans):
        if x["kind"] in ("block label", "attribution") and x.get("turn") is None:
            nxt = next((j for j in range(k + 1, len(spans)) if spans[j].get("turn") is not None), None)
            if nxt is not None:
                x["turn"] = spans[nxt]["turn"]
    last_turn = None
    for x in spans:
        if x.get("gap"):
            x["turn"] = last_turn
        elif x.get("turn") is not None:
            last_turn = x["turn"]
    first = next((x["turn"] for x in spans if x.get("turn") is not None), None)
    for x in spans:
        if x.get("gap") and x["turn"] is None:
            x["turn"] = first

    # stage 8: interview date. Stage 9: the confidence each speaker came with (rules 1 to 6, with the drops above)
    qa = None
    if block_spans:
        intro = text[:min(y["start"] for y in block_spans)][-1500:]
        qa = qa_stated_year(title, intro, pub_year)
    for t in turns:
        st = None
        if t["kind"] in ("speech", "second-hand"):
            st = qa if t.get("block") else stated_year(t["sentences"], pub_year)
        t["stated"] = st
    return spans, turns


# ── version 1.2: passages inside a Q&A that stage 2 left as prose ────────────
QA_KINDS = ("labelled Q&A", "unlabelled Q&A")
PASSAGE_ANSWER = "Q&A passage: an answer, by whoever answered before it"
PASSAGE_QUESTION = "Q&A passage: a question, by the interviewer"
PASSAGE_ANSWER_NO_FP = "Q&A passage: a member's answer going on, without the first person"
PASSAGE_EVIDENCE = "a passage inside the Q&A that the format check (stage 2) had left as prose"


def _paragraph_pieces(text, a, b):
    """The non-empty paragraphs of text[a:b] as [start, end, stripped text]. Blank lines go with the paragraph before,
    leading white space with the first, so the pieces cover a to b exactly."""
    out, pos = [], a
    while pos < b:
        nl = text.find("\n", pos, b)
        end = b if nl < 0 else nl + 1
        if text[pos:end].strip():
            out.append([pos, end, text[pos:end].strip()])
        elif out:
            out[-1][1] = end
        pos = end
    if out:
        out[0][0] = a
        out[-1][1] = b
    return out


def _names_speaker(body, who, names):
    """Whether a paragraph names the person who answers (outside quotation marks): "McIntosh laughs" is the journalist
    writing about them, not them answering."""
    if not who:
        return False
    forms = set(names.members[who]["forms"]) if who in names.members else {who}
    toks = set()
    for f in forms:
        ft = [t for t in _tokens(f) if len(t) >= 3 and t[0].isupper()]
        if len(ft) >= 2:
            toks.add(ft[-1])
        toks.add(f)
    return any(re.search(r"(?<![\w'’-])" + re.escape(t) + r"(?![\w'’-])", body) for t in toks if t)


# a quotation followed by a third-person attribution ('"We were both like, Barbie?" the singer said'): the journalist
# quoting someone, not an answer
THIRD_PARTY_ATTR = re.compile(r"[\"”][ \t]*,?[ \t]*(?:the[ \t]+[\w-]+|he|she|they|[A-Z][\w'’-]+(?:[ \t]+[A-Z][\w'’-]+)?)"
                              r"[ \t]+(?:said|says|told|tells|recalled|recalls|explained|explains|added|adds|laughed|"
                              r"laughs|admitted|admits|noted|notes|remembered|remembers|insisted|insists|joked|jokes)\b")
NAME_LABEL_START = re.compile(r"^[A-Z][\w'’.-]*(?:[ \t]+[A-Z][\w'’.-]*){0,2}[ \t]*:[ \t]")
# the first person, with "We" and "Our" at the start of a sentence too
FIRST_PERSON_ANY = re.compile(r"\b(?:I|I'm|I've|I'd|I'll|I’m|I’ve|I’d|I’ll)\b|\b(?:[Mm]e|[Mm]y|[Mm]ine|[Mm]yself|[Ww]e|"
                              r"[Ww]e['’](?:re|ve|d|ll)|us|[Oo]ur|[Oo]urs|[Oo]urselves)\b")
DASH_QUESTION = re.compile(r"^[–—-][ \t]+\S[^\n]*\?")


YOU = re.compile(r"(?i)\b(?:you|your|you're|you've|you'd|you'll|yourself|yourselves)\b")
I_SINGULAR = re.compile(r"\b(?:I|I'm|I've|I'd|I'll|me|my|mine|myself)\b")


def _passage_class(p, prev, who, names, nxt="", member=False):
    """'question', 'answer' or 'narration' for one paragraph of a passage. prev is the class of the paragraph (or Q&A
    turn) before it, nxt the text of the paragraph after it, member whether the person answering is a band member."""
    if T.STAGE_DIRECTION.match(p) or T.furniture_line(p) or S.FURNITURE.match(p) or NAME_LABEL_START.match(p) or \
            S.READER_Q.match(p) or (p.rstrip().endswith(":") and len(p.split()) <= 6):
        return "narration"                            # page furniture, a label ("Slim: ...", "Air4495 asks:")
    body = S._unquote_para(p)
    fp = bool(FIRST_PERSON_ANY.search(body))
    prev_ans = prev in ("answer", "answer, no first person")
    # a heading, unless it is a hard-wrapped line of an answer ("No, I've never been chucked into that particular")
    if T.heading_like(p) and not p.rstrip().endswith("?") and not fp and prev != "question" and \
            not (prev_ans and p[:1].islower()):
        return "narration"
    # the writer writing about the person who answers ("McIntosh laughs"), or quoting someone ('...," she said')
    if _names_speaker(body, who, names) or (not fp and (S.THIRD_ATTR_QUOTE.search(p) or THIRD_PARTY_ATTR.search(p))):
        return "narration"
    n = len(p.split())
    first = p.split()[0] if p.split() else ""
    # "Yes ...", "Absolutely ..." open an answer; "CORRECT." in capitals is a quiz master's verdict
    opener = bool(T.ANSWER_OPENER.match(p)) and (n <= 40 or bool(I_SINGULAR.search(body))) and \
        not (len(first) >= 3 and first.strip(".,!").isupper())
    # a question, also when the interviewer says "I" ("But I read that you love playing ... isn't it?")
    question = (S.is_question(p) and len(YOU.findall(body)) >= len(I_SINGULAR.findall(body))) or \
        bool(DASH_QUESTION.match(p))
    if prev == "question" and not question:
        return "answer"                               # the answer to the question before it
    if prev == "question" and (fp or opener):
        return "answer"                               # "Haven't a clue. Have you?"
    if opener:
        return "answer"                               # "Absolutely. No other way to survive that ..."
    if question or ((S.interviewer_statement(p) or S.interviewer_prompt(p)) and
                    FIRST_PERSON_ANY.search(S._unquote_para(nxt or ""))):
        return "question"                             # including "On this tour you're revisiting ..." before an answer
    if fp or (prev_ans and n <= 15):
        return "answer"
    if prev_ans and member:
        return "answer, no first person"              # a member's answer goes on until the next question
    return "narration"


def qa_passages(text, segments, blocks, names, journalist):
    """Version 1.2. In a Q&A record, a passage between two Q&A blocks that stage 2 left as prose (usually an answer
    paragraph that reports a conversation: 'They said, "Do you want to join the band?" I said, "All right then."') is
    read as part of the interview. A question line goes to the interviewer; an answer (a paragraph after a question, in
    the first person, or opening with "Yes", "Absolutely" ...) goes to whoever answered before it (or, at the start,
    after it), at that answer's confidence and never above medium. A member's answer going on without the first person
    is kept apart, at low confidence. Page furniture, labels and paragraphs that write about the person answering in
    the third person stay narration. Returns the prose segments still to be read as prose, the extra Q&A turns (stage
    3's span layout) and the number of passages read."""
    segs = [(a, b, k) for a, b, k in segments if b > a]
    turns = sorted((y for y in blocks if y["kind"] == "turn"), key=lambda y: y["start"])
    asker = next((y for y in turns if y["speaker_type"] == T.INTERVIEWER), None)
    ask = (T.INTERVIEWER, asker["speaker"]) if asker else (T.INTERVIEWER, journalist)
    keep, extra, n_read = [], [], 0
    for i, (a, b, k) in enumerate(segs):
        if k != "prose":
            continue
        interior = 0 < i < len(segs) - 1 and segs[i - 1][2] in QA_KINDS and segs[i + 1][2] in QA_KINDS
        before = [y for y in turns if y["end"] <= a]
        answers = [y for y in before if y["speaker_type"] != T.INTERVIEWER and not y.get("question")]
        answer = answers[-1] if answers else next((y for y in turns if y["start"] >= b and
                                                   y["speaker_type"] != T.INTERVIEWER and not y.get("question")), None)
        if not interior or answer is None:
            keep.append((a, b))
            continue
        who = answer.get("member") or answer.get("speaker")
        prev = "question" if before and (before[-1]["speaker_type"] == T.INTERVIEWER or before[-1].get("question")) \
            else "answer"
        runs = []                                     # [start, end, class]
        paras = _paragraph_pieces(text, a, b)
        for j, (s, e, p) in enumerate(paras):
            pieces = [(s, e, p)]
            ri = S.run_in_pairs([p])                  # a question run into its answer on one line
            if ri:
                cut = text.find(p, s) + ri[0][1]
                if s < cut < e:
                    pieces = [(s, cut, "?"), (cut, e, text[cut:e].strip())]
            nxt = paras[j + 1][2] if j + 1 < len(paras) else text[b:b + 400]
            for s2, e2, p2 in pieces:
                c = "question" if p2 == "?" else _passage_class(p2, prev, who, names, nxt,
                                                                 bool(answer.get("band_voice")))
                if runs and runs[-1][2] == c and c != "question":
                    runs[-1][1] = e2
                else:
                    runs.append([s2, e2, c])
                prev = c
        if all(c == "narration" for _, _, c in runs):
            keep.append((a, b))
            continue
        n_read += 1
        for s, e, c in runs:
            if c == "narration":
                keep.append((s, e))
                continue
            body_start = s + len(text[s:e]) - len(text[s:e].lstrip())
            if c == "question":
                typ, spk, conf, rule, member = ask[0], ask[1], "medium", PASSAGE_QUESTION, None
            else:
                typ, spk, member = answer["speaker_type"], answer["speaker"], answer.get("member")
                conf, rule = min_level("medium", answer["confidence"]), PASSAGE_ANSWER
                if c == "answer, no first person":
                    # 14 of 20 right in the held-out check (7 Oct): below the 80% for medium, so low (outside the main
                    # set); the misses were mostly the interviewer's statements
                    rule, conf = PASSAGE_ANSWER_NO_FP, "low"
            extra.append(dict(start=s, end=e, kind="turn", speaker_type=typ, speaker=spk, member=member,
                              era=answer.get("era") if c != "question" else None,
                              band_voice=c != "question" and bool(answer.get("band_voice")), confidence=conf, rule=rule,
                              label=None, speech_start=body_start, question=c == "question", quoted=False,
                              embedded_quotes=T.embedded_quotes(text[body_start:e]) if c != "question" else 0,
                              passage=True, evidence=PASSAGE_EVIDENCE))
    return keep, extra, n_read


# ── the three tables ─────────────────────────────────────────────────────────
SPAN_TYPES = {"speech": "speech", "narration": "narration", "second-hand": "second-hand", "attribution": "attribution",
              "non-speech": "non-speech quote", "undecided": "undecided quote", "unresolved": "unresolved",
              "block label": "label", "block heading": "narration", "block narration": "narration",
              "block sung": "sung", "block non-speech": "stage direction or page furniture"}
SPAN_DETAIL = {"block heading": "a heading inside a Q&A or first-person block",
               "block narration": "the journalist's words inside a Q&A or first-person block",
               "block sung": "lyrics sung in a transcript",
               "block non-speech": "a stage direction or page furniture inside a block"}
TURN_FIELDS = ["Record ID", "Turn ID", "Turn order", "Turn type", "Speaker type", "Speaker name", "Speaker candidates",
               "Reported member", "Band voice", "Main set", "Confidence", "Attribution rule", "Attribution evidence",
               "Embedded quote", "Question", "Holds an undecided quote", "Turn text", "Words", "Repeat of",
               "Span start", "Span end",
               "Interview date", "Interview date source", "Interview date precision", "Interview date evidence",
               "Route", "Quotation convention", "Focal Act", "Grammy Year", "Outlet / Source Name", "Publication Date",
               "Date Confidence", "Date Precision", "Content Era", "Temporal Class", "Rules version", "Human verified"]
SPAN_FIELDS = ["Record ID", "Span order", "Span type", "Span start", "Span end", "Turn ID", "Detail"]
CARRIED = ["Focal Act", "Grammy Year", "Outlet / Source Name", "Publication Date", "Date Confidence", "Date Precision",
           "Content Era", "Temporal Class"]


def _yes(b):
    return "yes" if b else "no"


def turn_text(text, t):
    """The turn's own words: its speech parts, with any attribution clause between them left out (it is the evidence).
    A Q&A turn starts after its label."""
    parts = list(t["parts"])
    if t.get("speech_start") is not None:
        parts[0] = (t["speech_start"], parts[0][1])
    out = [text[parts[0][0]:parts[0][1]]]
    for (a0, a1), (b0, b1) in zip(parts, parts[1:]):
        gap = text[a1:b0]
        out.append(gap if not gap.strip() else " ")
        out.append(text[b0:b1])
    return "".join(out)


def speaker_fields(who):
    typ = who["type"]
    if typ == MEMBERS_UNSPEC:
        return "band member, unspecified", who.get("name") or ""
    if typ == OTHER_ERA_UNSPEC:
        return "other-era member, unspecified", who.get("name") or ""
    if typ == AMBIGUOUS:
        return "", who.get("name") or ""
    if typ == UNKNOWN:
        return "", ""
    return who.get("name") or "", ""


def tables(rid, text, spans, turns, rec, route, conv):
    """The record's rows of the turns and spans tables, and its quality-control fields for the records table.
    rec: the record's corpus fields (CARRIED, plus 'Date Precision')."""
    pub = rec.get("Publication Date") or ""
    tid = {}
    turn_rows = []
    for i, t in enumerate(turns, 1):
        tid[i - 1] = f"{rid}-{i}"
        who = t["who"]
        name, cands = speaker_fields(who)
        bv = t["kind"] == "speech" and who["type"] in BAND_VOICE
        conf = who["conf"]
        if t["kind"] == "second-hand":
            reported = t["reported"] or ""
        else:
            reported = ""
        st = t.get("stated")
        if st:
            idate, isrc, iprec, iev = st[0], "stated in the text", "year", st[1]
        else:
            idate, isrc, iprec, iev = pub, "publication-date proxy", rec.get("Date Precision") or "", ""
        body = turn_text(text, t) if t["kind"] == "speech" else text[t["start"]:t["end"]]
        row = {"Record ID": rid, "Turn ID": tid[i - 1], "Turn order": i, "Turn type": t["kind"],
               "Speaker type": who["type"], "Speaker name": name, "Speaker candidates": cands,
               # version 1.2: the main set is a set of band-voice turns, so the column is blank for other turns
               "Reported member": reported, "Band voice": _yes(bv),
               "Main set": _yes(conf in ("high", "medium")) if bv else "",
               "Confidence": conf, "Attribution rule": who["rule"],
               "Attribution evidence": " | ".join(e for e in t["evidence"] if e)[:500],
               "Embedded quote": _yes(t.get("embedded")), "Question": _yes(t.get("question")),
               "Holds an undecided quote": _yes(t.get("undecided")), "Turn text": body, "Words": len(body.split()),
               "Repeat of": "",
               "Span start": t["parts"][0][0] if t.get("speech_start") is None else t["speech_start"],
               "Span end": t["end"], "Interview date": idate, "Interview date source": isrc,
               "Interview date precision": iprec, "Interview date evidence": iev, "Route": route,
               "Quotation convention": conv}
        for k in CARRIED:
            row[k] = rec.get(k)
        row["Rules version"] = RULES_VERSION
        row["Human verified"] = "no"
        turn_rows.append(row)
    # spans: gaps (a line break or a comma between two quotations) join the span before them
    merged = []
    for x in spans:
        if x.get("gap") and merged:
            merged[-1]["end"] = x["end"]
            continue
        y = dict(start=x["start"], end=x["end"], kind=x["kind"], turn=x.get("turn"), evidence=x.get("evidence") or "")
        if x.get("gap"):
            y["kind"] = "gap"
        if merged and merged[-1]["kind"] == "gap":
            y["start"] = merged[-1]["start"]
            merged.pop()
        merged.append(y)
    span_rows = []
    order = 0
    for y in merged:
        pieces = [y]
        if y["kind"] == "block turn":
            t = turns[y["turn"]] if y["turn"] is not None else None
            ss = t.get("speech_start") if t else None
            if ss is not None and y["start"] < ss < y["end"]:
                pieces = [dict(y, end=ss, kind="block label", evidence="the turn's label"), dict(y, start=ss)]
        for p in pieces:
            order += 1
            k = p["kind"]
            typ = "speech" if k == "block turn" else SPAN_TYPES.get(k, k)
            det = SPAN_DETAIL.get(k) or p["evidence"]
            span_rows.append({"Record ID": rid, "Span order": order, "Span type": typ, "Span start": p["start"],
                              "Span end": p["end"], "Turn ID": tid.get(p["turn"], "") if p["turn"] is not None else "",
                              "Detail": (det or "")[:160]})
    # round trip: the spans cover the text exactly once, in order
    pos, gap_at = 0, None
    for r in span_rows:
        if r["Span start"] != pos:
            gap_at = pos
            break
        pos = r["Span end"]
    if gap_at is None and pos != len(text):
        gap_at = pos
    unres = [r for r in span_rows if r["Span type"] == "unresolved"]
    unres_chars = sum(r["Span end"] - r["Span start"] for r in unres)
    prose_chars = sum(r["Span end"] - r["Span start"] for r in span_rows
                      if r["Span type"] in ("speech", "narration", "second-hand", "attribution", "non-speech quote",
                                            "undecided quote", "unresolved"))
    flags = []
    if gap_at is not None:
        flags.append(f"round trip failed at character {gap_at}")
    if unres and unres_chars > 0.5 * max(1, prose_chars):
        flags.append(f"most of the prose is unresolved ({100 * unres_chars // max(1, prose_chars)}%)")
    if any("over 1,500 characters" in (x.get("evidence") or "") for x in spans if x["kind"] in ("speech", "undecided",
                                                                                                  "non-speech")):
        flags.append("a quotation over 1,500 characters (a closing mark may be missing)")
    speech = [r for r in turn_rows if r["Turn type"] == "speech"]
    qc = {"Spans": len(span_rows), "Turns": len(turn_rows), "Speech turns": len(speech),
          "Band-voice turns": sum(r["Band voice"] == "yes" for r in turn_rows),
          "Band-voice turns (main set)": sum(r["Band voice"] == "yes" and r["Main set"] == "yes" for r in turn_rows),
          "Second-hand turns": sum(r["Turn type"] == "second-hand" for r in turn_rows),
          "Unknown speech turns": sum(r["Speaker type"] in (UNKNOWN, AMBIGUOUS) for r in speech),
          "Unresolved spans": len(unres), "Unresolved characters": unres_chars,
          "Round trip": "pass" if gap_at is None else f"fail at character {gap_at}",
          "Flags": "; ".join(flags)}
    return turn_rows, span_rows, qc
'''


def _load(name):
    mod = types.ModuleType(name)
    mod.__file__ = name + ".py"
    sys.modules[name] = mod
    exec(compile(SOURCES[name], name + ".py", "exec"), mod.__dict__)
    return mod


S = _load("stage2")
_load("gender_names")
T = _load("stage3")
Q = _load("stage4")
F = _load("stage5")

# ═════════════════════════════════════════════════════════════════════════════
# THE NAME LISTS (final, 5 Oct 2026: line-ups, aliases and the name-list review's decisions)
# ═════════════════════════════════════════════════════════════════════════════
_NAME_LISTS_Z = "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"
NAME_LISTS = json.loads(zlib.decompress(base64.b64decode(_NAME_LISTS_Z)).decode("utf-8"))


def drop_colliding_aliases(nl):
    """Version 1.2: an alias whose first and last names make the name of someone on the act's non-member list is not
    used: "John Charles Julian Lennon" made "John Lennon" (on Julian Lennon's non-member list) read as Julian Lennon,
    and "John" alone as him too. Returns the aliases left out."""
    non = defaultdict(set)
    for x in nl["nonmembers"]:
        non[x["act"]].add(T.fold(x["name"]))
    dropped = []
    for m in nl["members"]:
        keep = []
        for a in m.get("aliases") or []:
            toks = re.sub(r"[()]", " ", a).replace("/", " ").split()
            short = T.fold(toks[0] + " " + toks[-1]) if len(toks) >= 2 else T.fold(a)
            if short in non[m["act"]] or T.fold(a) in non[m["act"]]:
                dropped.append(f"{m['member']} ({m['act']}): {a}")
            else:
                keep.append(a)
        m["aliases"] = keep
    return dropped


ALIASES_LEFT_OUT = drop_colliding_aliases(NAME_LISTS)


# ═════════════════════════════════════════════════════════════════════════════
# RUN
# ═════════════════════════════════════════════════════════════════════════════
def text_of(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return str(v)


def stop(msg):
    print("\nSTOPPED, nothing was written:\n  " + msg.replace("\n", "\n  "))
    sys.exit(1)


def stage2_fingerprint(st2):
    return hashlib.md5("\n".join(f"{i}|{st2[i]['Route']}|{json.dumps(st2[i]['Segments'])}"
                                 for i in sorted(st2)).encode()).hexdigest()


def stage3_fingerprint(st3):
    lines = []
    for i in sorted(st3):
        for s in st3[i]["Spans"]:
            lines.append(f"{i}|{s['start']}|{s['end']}|{s['kind']}|{s.get('speaker_type')}|{s.get('speaker')}|"
                         f"{s.get('confidence')}|{s.get('speech_start')}")
    return hashlib.md5("\n".join(lines).encode()).hexdigest()


def load_input():
    for path, made_by in ((IN_JSON, "the CLEAN7 step"), (IN_STAGE1, "turns_v7.py"), (IN_STAGE2, "turns_v7.py"),
                          (IN_STAGE3, "turns_v7_stage3.py")):
        if not os.path.exists(path):
            stop(f"{path} not found. Put {os.path.basename(path)} (made by {made_by}) in {DIR} (or set HARRISON_DIR).")
    with open(IN_JSON, encoding="utf-8") as f:
        recs = json.load(f)
    problems = []
    if len(recs) != EXPECTED_RECORDS:
        problems.append(f"{len(recs)} records, expected {EXPECTED_RECORDS}")
    try:
        ids = [int(r.get("Record ID")) for r in recs]
    except (TypeError, ValueError):
        ids = []
    if ids != list(range(1, len(recs) + 1)):
        problems.append("Record IDs are not 1, 2, 3 ... in order")
    ts = [r for r in recs if text_of(r.get("In Turn Set")) == "Yes"]
    if len(ts) != EXPECTED_TURN_SET:
        problems.append(f"turn set is {len(ts)}, expected {EXPECTED_TURN_SET}")
    fp = hashlib.md5("\n".join(f"{int(r['Record ID'])}|{text_of(r.get('URL'))}" for r in ts).encode()).hexdigest()
    if not problems and fp != TURN_SET_FINGERPRINT:
        problems.append("the turn set's IDs and URLs are not the ones this script was built on")
    if problems:
        stop("This is not the CLEAN7 file the script expects:\n" + "\n".join(problems))
    del recs
    with open(IN_STAGE1, encoding="utf-8") as f:
        st1 = {x["Record ID"]: x for x in json.load(f)}
    with open(IN_STAGE2, encoding="utf-8") as f:
        st2 = {x["Record ID"]: x for x in json.load(f)}
    with open(IN_STAGE3, encoding="utf-8") as f:
        st3 = {x["Record ID"]: x for x in json.load(f)}
    ids = {int(r["Record ID"]) for r in ts}
    if set(st1) != ids or set(st2) != ids:
        stop("The stage 1 and stage 2 files are not for this CLEAN7 file's turn set. Run turns_v7.py again first.")
    raw = {int(r["Record ID"]): text_of(r.get("Extracted Text")) for r in ts}
    if any(len(st1[i]["Normalised Text"]) != len(raw[i]) for i in ids):
        stop("The stage 1 texts do not match CLEAN7. Run turns_v7.py again first.")
    counts = Counter(x["Route"] for x in st2.values())
    if dict(counts) != STAGE2_COUNTS or stage2_fingerprint(st2) != STAGE2_FINGERPRINT:
        stop("The stage 2 file is not the one made by version 2 of turns_v7.py (its routes or segments differ). Run "
             "version 2 first.")
    qa = {i for i in ids if st2[i]["Route"] in ("labelled Q&A", "unlabelled Q&A", "first-person")}
    if set(st3) != qa or stage3_fingerprint(st3) != STAGE3_FINGERPRINT:
        stop("The stage 3 file is not the one made by version 3 of turns_v7_stage3.py on these files. Run version 3 "
             "of turns_v7_stage3.py first (it overwrites harrison_v7_TURNS_stage3.json).")
    return ts, st1, st2, st3


RECORD_FIELDS = ["Record ID", "Focal Act", "Grammy Year", "Title / Headline", "Outlet / Source Name", "URL",
                 "Publication Date", "Date Precision", "Temporal Class", "Content Era"]


def excerpt(s, n):
    s = (s or "").strip()
    return s if len(s) <= n else s[:n].rstrip() + " …"


LEAD_IN_START = re.compile(r"(?:\n|[.!?][\"'”’)]?\s)\s*(?=\S)")


def lead_in(t, start, n=500):
    """The text before a turn, for the Turn review: at most n characters, starting at a sentence or a paragraph (or
    else at a word) instead of in the middle of a word, with "… " in front when the record's text goes back further."""
    a = max(0, start - n)
    s = t[a:start]
    if a == 0:
        return s.strip()
    m = LEAD_IN_START.search(s)
    if m and len(s) - m.end() >= n // 3:
        s = s[m.end():]
    elif not t[a - 1].isspace() and re.search(r"\s", s):
        s = s[re.search(r"\s", s).end():]
    return "… " + s.strip()


def process_record(rid, m, x, st1, st3, names, stoplist):
    """One record: its rows of the turns and spans tables, its row of the records table, and its text (None when the
    record is not processed)."""
    route = x["Route"]
    base = {f: (m.get(f) if f in ("Record ID", "Grammy Year") else text_of(m.get(f))) for f in RECORD_FIELDS}
    base["Record ID"] = rid
    base.update({"Route": route, "Detail": x.get("Detail") or ""})
    if route == FORUM:
        return [], [], dict(base, **{"Processed": "no: forum thread (posts by several users)"}), None
    if rid in SKIP:
        return [], [], dict(base, **{"Processed": "no: " + SKIP[rid]}), None
    text = st1[rid]["Normalised Text"]
    act = text_of(m.get("Focal Act"))
    nm = names[act]
    segs_all = [(a, b) for a, b, kind in x["Segments"] if kind == "prose" and b > a]
    blocks = list(st3[rid]["Spans"]) if rid in st3 else []
    segs, n_passages = segs_all, 0
    if route in F.QA_KINDS and blocks:
        # version 1.2: passages inside the Q&A that stage 2 left as prose are read as part of the interview
        journalist = T.person_from_author(text_of(m.get("Author / Interviewer"))) or \
            text_of(m.get("Outlet / Source Name")) or None
        segs, extra, n_passages = F.qa_passages(text, x["Segments"], blocks, nm, journalist)
        segs = sorted(segs)
        blocks += extra
    conv, cev = Q.convention(text, segs_all) if segs_all else ("american", {})
    sp, conv = Q.spans(text, segs, conv) if segs else ([], conv)
    noms = [n for n in nm.members if nm.era(n) == "Nomination-era"]
    solo = noms[0] if len(noms) == 1 else None
    spans, turns = F.record_turns(rid, text, text_of(m.get("Title / Headline")), act, nm, stoplist,
                                  text_of(m.get("Author / Interviewer")), text_of(m.get("Outlet / Source Name")),
                                  text_of(m.get("Publication Date")), sp, blocks, solo, conv)
    rec = {f: (m.get(f) if f == "Grammy Year" else text_of(m.get(f))) for f in F.CARRIED}
    # version 1.3: no convention on the turns of a record with no prose to read, as in the records table
    trows, srows, qc = F.tables(rid, text, spans, turns, rec, route, conv if segs_all else "")
    row = dict(base, **{"Processed": "yes", "Quotation convention": conv if segs_all else "",
                        "Double-quote attributions": cev.get("double-quote attributions", ""),
                        "Single-quote attributions": cev.get("single-quote attributions", ""),
                        "Q&A passages read as interview": n_passages}, **qc)
    return trows, srows, row, text


def _repeat_key(s):
    s = s.casefold().replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"[–—]", "-", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip(" \"'.,;:!?…-")


def mark_repeats(turn_rows):
    """Version 1.2: 'Repeat of' is the Turn ID of the earliest turn of the same type with the same text (at least
    REPEAT_MIN_WORDS words; earliest by interview date, then publication date, then Record ID and turn order)."""
    groups = defaultdict(list)
    for r in turn_rows:
        r["Repeat of"] = ""
        if r["Words"] >= REPEAT_MIN_WORDS:
            groups[(r["Turn type"], _repeat_key(r["Turn text"]))].append(r)
    for rs in groups.values():
        if len(rs) < 2:
            continue
        rs.sort(key=lambda r: (str(r["Interview date"] or "9999"), str(r["Publication Date"] or "9999"),
                               r["Record ID"], r["Turn order"]))
        for r in rs[1:]:
            r["Repeat of"] = rs[0]["Turn ID"]


def main():
    t0 = time.time()
    print(VERSION)
    ts, st1, st2, st3 = load_input()
    print(f"Read and checked the CLEAN7, stage 1, stage 2 and stage 3 files: turn set {len(ts):,} [{time.time() - t0:.0f}s]")
    names, stoplist = T.load_names(NAME_LISTS)
    meta = {int(r["Record ID"]): r for r in ts}

    turn_rows, span_rows, rec_rows, texts = [], [], [], {}
    todo = sorted(meta)
    for k, rid in enumerate(todo, 1):
        trows, srows, row, text = process_record(rid, meta[rid], st2[rid], st1, st3, names, stoplist)
        turn_rows.extend(trows)
        span_rows.extend(srows)
        rec_rows.append(row)
        if text is not None:
            texts[rid] = text
        if k % 1000 == 0:
            print(f"  {k:,} of {len(todo):,} records [{time.time() - t0:.0f}s]")
    mark_repeats(turn_rows)
    print(f"Stages 4 to 9: {sum(r['Processed'] == 'yes' for r in rec_rows):,} records, {len(turn_rows):,} turns, "
          f"{len(span_rows):,} spans [{time.time() - t0:.0f}s]")

    counts = summary_counts(turn_rows, span_rows, rec_rows)
    fp = turns_fingerprint(turn_rows)
    if counts == CHECK_COUNTS and fp == CHECK_FINGERPRINT:
        dev_note = ("These counts match the ones the script gave on the same files when it was checked on 12 Oct, and "
                    "every turn is the same.")
    elif counts == CHECK_COUNTS:
        dev_note = ("These counts match the ones the script gave when it was checked on 12 Oct, but some turns differ. "
                    "Please report this with the run's output.")
    else:
        dev_note = ("These counts differ from the ones the script gave on the same files when it was checked on 12 Oct. "
                    "Please report this with the run's output.")
    if DRY_RUN:
        report(counts, dev_note, wrote=False)
        return
    write_json_tables(turn_rows, span_rows, rec_rows)
    print(f"Wrote the JSON tables [{time.time() - t0:.0f}s]")
    write_turns_workbook(turn_rows)
    print(f"Wrote the turns workbook [{time.time() - t0:.0f}s]")
    write_dashboard(turn_rows, rec_rows, texts, counts, dev_note)
    with open(OUT_TURNS, encoding="utf-8") as f:
        back = json.load(f)
    if len(back["data"]) != len(turn_rows):
        stop("The written turns file did not read back correctly; delete it and run again.")
    report(counts, dev_note, wrote=True)
    print(f"[{time.time() - t0:.0f}s]")


# ── counts ───────────────────────────────────────────────────────────────────
def summary_counts(turn_rows, span_rows, rec_rows):
    route_of = {r["Record ID"]: r["Route"] for r in rec_rows}
    speech = [r for r in turn_rows if r["Turn type"] == "speech"]
    bv = [r for r in speech if r["Band voice"] == "yes"]
    return {
        "records processed": sum(r.get("Processed") == "yes" for r in rec_rows),
        "turns": dict(sorted(Counter(r["Turn type"] for r in turn_rows).items())),
        "spans": len(span_rows),
        "band-voice turns by route": {x: sum(route_of[r["Record ID"]] == x for r in bv) for x in ROUTES},
        "band-voice turns by confidence": {c: sum(r["Confidence"] == c for r in bv) for c in ("high", "medium", "low")},
        "records with band-voice turns": len({r["Record ID"] for r in bv}),
        "unknown speech turns": sum(r["Speaker type"] in ("unknown", "ambiguous") for r in speech),
        "stated interview dates": sum(r["Interview date source"] == "stated in the text" for r in turn_rows),
        "round-trip failures": sum(r.get("Round trip", "pass") != "pass" for r in rec_rows if r.get("Processed") == "yes"),
        "unresolved spans": sum(r.get("Unresolved spans", 0) for r in rec_rows if r.get("Processed") == "yes"),
        "Q&A passages read as interview": sum(r.get("Q&A passages read as interview", 0) or 0 for r in rec_rows),
        "band-voice turns from Q&A passages": sum(1 for r in bv if r["Attribution rule"].startswith("Q&A passage")),
        "turns that repeat an earlier turn": sum(1 for r in turn_rows if r["Repeat of"]),
    }


def turns_fingerprint(turn_rows):
    return hashlib.md5("\n".join(
        f"{r['Turn ID']}|{r['Turn type']}|{r['Speaker type']}|{r['Speaker name']}|{r['Confidence']}|{r['Span start']}|"
        f"{r['Span end']}|{r['Interview date']}|{r['Reported member']}|{r['Main set']}|{r['Repeat of']}"
        for r in turn_rows).encode()).hexdigest()


# ── writing ──────────────────────────────────────────────────────────────────
def write_json_tables(turn_rows, span_rows, rec_rows):
    with open(OUT_TURNS, "w", encoding="utf-8") as f:
        json.dump({"columns": F.TURN_FIELDS, "data": [[r[c] for c in F.TURN_FIELDS] for r in turn_rows]}, f,
                  ensure_ascii=False)
    with open(OUT_SPANS, "w", encoding="utf-8") as f:
        json.dump({"columns": F.SPAN_FIELDS, "data": [[r[c] for c in F.SPAN_FIELDS] for r in span_rows]}, f,
                  ensure_ascii=False)
    with open(OUT_RECORDS, "w", encoding="utf-8") as f:
        json.dump(rec_rows, f, ensure_ascii=False)


def _safe(v, cap=EXCEL_CELL_CAP):
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
        return v if len(v) <= cap else v[:cap - 40] + " … [cut here: the full text is in the JSON]"
    return v


# ── the turns workbook's Read me: what the table is, every column, and how speakers are found ─────────────────────
TURNS_README_INTRO = [
    ("The turns table", "One row per turn. A turn is a stretch of words by one speaker: a quotation in an article, an "
     "answer or a question in a Q&A, a first-person account, or a stretch of the journalist's own words. Turns are in "
     "each record's reading order. The same table, with no text cut, is harrison_v7_turns_final.json; the summaries "
     "and checks are in harrison_v7_dashboard_final.xlsx."),
    ("Status", "The rules were checked blind by AI agents (Claude) and against human coding of held-out records; "
     "both checks are reported in the technical appendix (Validation and checks). The dashboard's Read me gives the "
     "blind checks and the known limits."),
    ("Band voice", "The members' own words: speech turns whose speaker is a band member, a band member who cannot be "
     "told apart from the others, or the band speaking as a group. Filter Band voice = yes."),
    ("The main set", "The band-voice turns with high or medium confidence, which the analysis uses: filter Band voice = "
     "yes and Main set = yes (Main set is blank for turns that are not band voice). Low-confidence band-voice turns "
     "stay in the table for the robustness checks, as does a high-only run."),
    ("Public narrative", "Everything else: the journalist's narration, the interviewers' questions, and the words of "
     "other people (producers, managers, relatives, other musicians). Second-hand turns, in which the journalist "
     "reports a member's words without quotation marks, are a category of their own."),
    ("Republished articles", "Some articles were published more than once. Repeat of points from a later copy of a turn "
     "to the earliest one, so that a quotation can be counted once and dated by its first appearance."),
    ("Long texts", "Texts longer than 32,000 characters (an Excel cell holds at most 32,767) are cut and marked here; "
     "the JSON holds them in full. Every other value is as in the JSON."),
]
TURN_COLUMNS = {
    "Record ID": "The record's ID in CLEAN7. It joins this table to the records and spans tables and to CLEAN7.",
    "Turn ID": "The record ID and the turn's number in the record: 1837-55 is record 1837's 55th turn.",
    "Turn order": "The turn's position in the record's reading order (1 is the first).",
    "Turn type": "speech: someone's own words (a quotation in an article, an answer or a question in a Q&A, a "
                 "first-person account). narration: the journalist's own words, including the titles, lyrics and scare "
                 "quotes they quote. second-hand: a sentence of narration that reports a member's words without "
                 "quotation marks ('Adam says they have been practising for 10 hours every day').",
    "Speaker type": "band member: a member of the focal act at the time of its nomination. band member, unspecified: "
                    "one of several members, but not which one. band collective: the act speaking as a group ('the "
                    "band said in a statement'). other-era member: a member who joined after the nomination or had "
                    "left before it (not band voice); other-era member, unspecified: one of several of them, not which "
                    "one. interviewer: the person asking the questions in a Q&A. "
                    "journalist: the writer (their narration and second-hand turns, and their own 'I said' in an "
                    "article). other person: anyone else (producers, managers, relatives, fans, other musicians and "
                    "groups). unknown: no rule found the speaker. ambiguous: a name that several people share and that "
                    "could not be settled.",
    "Speaker name": "Who speaks: members as named on the line-up list; other people as the text names them; the "
                    "journalist is the byline, or the outlet when there is none; 'band member, unspecified' when the "
                    "member cannot be told; blank when the speaker is unknown or ambiguous, and for an audience member "
                    "that a caption transcript does not name.",
    "Speaker candidates": "For ambiguous speakers, the people it could be ('Mark Knopfler / David Knopfler'), or the "
                          "one name the record uses when it never names the member in full. For unspecified members, "
                          "the members it could be when the text narrows them down; blank when any member could be "
                          "speaking (a Q&A whose title names the band). Blank for every other speaker.",
    "Reported member": "Second-hand turns only: the member, or the act, whose words the journalist reports.",
    "Band voice": "yes for speech turns by a band member, a band member (unspecified) or the band collective; no for "
                  "every other turn.",
    "Main set": "For band-voice turns: yes when the confidence is high or medium (the main analysis set), no when it is "
                "low. Blank for turns that are not band voice.",
    "Confidence": "How sure the speaker is: high, medium or low. The targets, to be tested in the validation, are about "
                  "95% right for high and 80% for medium. 'unknown' when no speaker was found. The journalist's "
                  "narration is high.",
    "Attribution rule": "The rule that gave the speaker: '1:' to '6:' for quotations in prose (articles, and the prose "
                        "parts of Q&As and first-person accounts; see How speakers are found, below); 'stage 3:' for "
                        "Q&A and first-person turns; 'Q&A passage:' for passages inside a Q&A that the format check had "
                        "left as prose; '7: the writer's own words' for narration; 'second-hand:' for second-hand "
                        "turns. A note after a semicolon adds detail, often why the confidence was lowered.",
    "Attribution evidence": "The words that decided the speaker: the attribution clause (', Hynde says'), the Q&A label, "
                            "or a note.",
    "Embedded quote": "yes when the turn quotes someone else's words inside it (a member recounting what someone said "
                      "to them).",
    "Question": "yes when the turn was read as a question line: every interviewer turn in unlabelled Q&As and in Q&A "
                "passages, but in labelled Q&As many of the interviewer's labelled lines are not marked. For all of "
                "the interviewer's words, filter Speaker type = interviewer.",
    "Holds an undecided quote": "yes for a narration turn that contains a quotation that is neither clearly speech nor "
                                "a title. Such quotations are never band voice.",
    "Turn text": "The turn's own words. Attribution clauses between its parts and Q&A labels are left out; they are in "
                 "Attribution evidence.",
    "Words": "The number of words in Turn text.",
    "Repeat of": "When this turn repeats an earlier one (the same words, ignoring case, spacing, the style of "
                 "quotation marks and dashes, and punctuation at either end), usually because the article was "
                 "republished: the Turn ID of the earliest copy, by interview date, then publication date, then Record "
                 "ID. Only turns of at least 8 words, compared with turns of the same type. Blank for the first copy "
                 "and for turns that are not repeated.",
    "Span start": "Where the turn's own words start in the record's text, as a character position. The text is CLEAN7's "
                  "text with its line breaks and spaces tidied (saved as harrison_v7_TURNS_stage1.json); the spans "
                  "table covers every character of it.",
    "Span end": "Where the turn ends in the record's text (a character position, as for Span start).",
    "Interview date": "When the words were said: a year stated in the text, when the attribution ties one to a speaking "
                      "verb ('told NME in 1983', 'in a 1985 interview') or a Q&A's title or introduction dates the "
                      "interview (written as the year, '1983'); otherwise the publication date.",
    "Interview date source": "'stated in the text' or 'publication-date proxy'.",
    "Interview date precision": "Whether the interview date is known to the day, the month or the year.",
    "Interview date evidence": "The words that state the year (stated dates only).",
    "Route": "The record's format, from the format check (stage 2): prose (an article), labelled Q&A, unlabelled Q&A, or "
             "first-person (an account in the speaker's own words).",
    "Quotation convention": "american (quotations in double marks) or british (single marks), as read for the record's "
                            "prose; blank when the record has no prose (most first-person accounts).",
    "Focal Act": "The Best New Artist nominee the record was collected for (CLEAN7).",
    "Grammy Year": "The year of the act's Best New Artist ceremony (CLEAN7).",
    "Outlet / Source Name": "The publication, as CLEAN7 names it (usually its web domain).",
    "Publication Date": "The record's publication date (CLEAN7). A date known only to the month or the year is "
                        "written as its first day ('1979-05-01', with Date Precision 'month').",
    "Date Confidence": "How sure CLEAN7 is of the publication date. Every record in the turn set is high.",
    "Date Precision": "Whether CLEAN7's publication date is known to the day, the month or the year.",
    "Content Era": "The year the record's content dates from, when CLEAN7 gives one (for example, an older interview "
                   "printed again later); blank otherwise. Temporal Class uses it.",
    "Temporal Class": "When the record dates from, compared with the nomination (CLEAN7): pre-nomination (before the "
                      "Grammy year), contemporaneous (that year to three years after), near-term (four to ten years "
                      "after) or retrospective (more than ten years after). It uses Content Era when there is one, "
                      "else the publication year.",
    "Rules version": "The version of the rules that made the turn.",
    "Human verified": "'no' until a person has checked the turn.",
}
TURNS_README_RULES = [
    ("Quotations in prose", "In articles, and in the prose parts of Q&As and first-person accounts. A quotation is "
     "speech when a speech verb or a speaker's name label is beside it, when it "
     "continues the speech before it, or when it is a sentence in the first or second person. A title, lyrics or a scare "
     "quote is narration; anything else is undecided (never band voice). Its speaker comes from the first rule that "
     "applies. 1: a member's name beside a speech verb (high; medium for Sade, whose act's name is also her own). 2: "
     "someone else named (high; low for one name, a first "
     "name not on the list, or an unnamed speaker). 3: a name several people share goes to the person the title or "
     "introduction is about, else to the nearest earlier full name (medium; low when it cannot be settled). 4: the "
     "band named as speaker (medium; "
     "'the band said' only when it can be the act). 5: a pronoun for the current speaker, or a quotation that continues "
     "them (medium; low without a gender check); 'I' in the narration is the writer, the journalist's own words. 5b: "
     "with no current speaker, a pronoun goes to the person the "
     "narration is about (members medium, others low). 5c: an unattributed quotation in a piece about one member goes "
     "to that member (low). 6: unknown."),
    ("Q&As and first-person accounts", "Stage 3 (version 3) gives each answer, question and account its speaker, from "
     "the labels ('LAUPER:'), the interviewee the title or introduction names, and the questions: where the questions "
     "carry no label, a paragraph that ends with a question mark is read as the interviewer's (see the known limits in "
     "the dashboard's Read me). A passage inside a Q&A that the "
     "format check had left as prose is read as part of the interview: a question line is the interviewer's, the rest "
     "continues the answer before it (rule names 'Q&A passage')."),
    ("Lowered confidence", "Confidence drops one level when a member was named by a surname or first name that a "
     "relative or namesake named in the record shares, and a turn is never surer than the speaker it continues."),
]


TURN_WIDTHS = {
    "Record ID": 9, "Turn ID": 10, "Turn order": 7, "Turn type": 11, "Speaker type": 16, "Speaker name": 20,
    "Speaker candidates": 18, "Reported member": 16, "Band voice": 7, "Main set": 7, "Confidence": 10,
    "Attribution rule": 28, "Attribution evidence": 28, "Embedded quote": 9, "Question": 9,
    "Holds an undecided quote": 10, "Turn text": 80, "Words": 7, "Repeat of": 10, "Span start": 9, "Span end": 9,
    "Interview date": 11, "Interview date source": 14, "Interview date precision": 10, "Interview date evidence": 24,
    "Route": 13, "Quotation convention": 11, "Focal Act": 20, "Grammy Year": 8, "Outlet / Source Name": 22,
    "Publication Date": 11, "Date Confidence": 10, "Date Precision": 9, "Content Era": 9, "Temporal Class": 14,
    "Rules version": 18, "Human verified": 9}


def write_turns_workbook(turn_rows):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Read me")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 120
    wrap = Alignment(wrap_text=True, vertical="top")

    def cell(v, bold=False):
        c = WriteOnlyCell(ws, value=v)
        c.alignment = wrap
        if bold:
            c.font = Font(bold=True)
        return c

    ws.append([cell("harrison_v7_turns_final", True), cell(VERSION)])
    ws.append([])
    ws.append([cell("WHAT THIS IS", True)])
    for k, v in TURNS_README_INTRO:
        ws.append([cell(k, True), cell(v)])
    ws.append([])
    ws.append([cell("THE COLUMNS", True), cell("In the order of the Turns sheet.")])
    for col in F.TURN_FIELDS:
        ws.append([cell(col, True), cell(TURN_COLUMNS[col])])
    ws.append([])
    ws.append([cell("HOW SPEAKERS ARE FOUND", True)])
    for k, v in TURNS_README_RULES:
        ws.append([cell(k, True), cell(v)])
    ws = wb.create_sheet("Turns")
    ws.freeze_panes = "C2"
    for i, col in enumerate(F.TURN_FIELDS, 1):
        ws.column_dimensions[get_column_letter(i)].width = TURN_WIDTHS.get(col, 12)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(F.TURN_FIELDS))}{len(turn_rows) + 1}"
    ws.append([cell(col, True) for col in F.TURN_FIELDS])
    for r in turn_rows:
        ws.append([_safe(r[c]) for c in F.TURN_FIELDS])
    wb.save(OUT_XLSX)


def _sheet(df):
    return df.apply(lambda col: col.map(_safe))


def outlet_key(o):
    """An outlet's name for grouping: without "www." and trailing full stops."""
    o = (o or "").strip().lower()
    o = re.sub(r"^www\d?\.", "", o).rstrip(".")
    return o or "(none)"


def write_dashboard(turn_rows, rec_rows, texts, counts, dev_note):
    route_of = {r["Record ID"]: r["Route"] for r in rec_rows}
    rec_by_id = {r["Record ID"]: r for r in rec_rows}
    speech = [r for r in turn_rows if r["Turn type"] == "speech"]
    # 1. the dashboard tables
    tables = []
    kinds = ["speech", "narration", "second-hand"]
    tables.append(("Turns by type and route", pd.DataFrame(
        [dict({"Turn type": k}, **{x: sum(1 for r in turn_rows if r["Turn type"] == k and route_of[r["Record ID"]] == x)
                                   for x in ROUTES}, Total=sum(1 for r in turn_rows if r["Turn type"] == k))
         for k in kinds])))
    types_ = ["band member", "band member, unspecified", "band collective", "other-era member",
              "other-era member, unspecified", "interviewer", "journalist", "other person", "unknown", "ambiguous"]
    tables.append(("Speech turns by speaker type and route (band voice: the first three)", pd.DataFrame(
        [dict({"Speaker type": t}, **{x: sum(1 for r in speech if r["Speaker type"] == t and route_of[r["Record ID"]] == x)
                                      for x in ROUTES}, Total=sum(1 for r in speech if r["Speaker type"] == t))
         for t in types_])))
    tables.append(("Confidence of speech turns by route", pd.DataFrame(
        [dict({"Route": x}, **{c: sum(1 for r in speech if r["Confidence"] == c and route_of[r["Record ID"]] == x)
                               for c in ("high", "medium", "low", "unknown")},
              **{"Band voice: high": sum(1 for r in speech if r["Band voice"] == "yes" and r["Confidence"] == "high" and
                                         route_of[r["Record ID"]] == x),
                 "Band voice: medium": sum(1 for r in speech if r["Band voice"] == "yes" and r["Confidence"] == "medium"
                                           and route_of[r["Record ID"]] == x),
                 "Band voice: low": sum(1 for r in speech if r["Band voice"] == "yes" and r["Confidence"] == "low" and
                                        route_of[r["Record ID"]] == x)}) for x in ROUTES])))
    sh = [r for r in turn_rows if r["Turn type"] == "second-hand"]
    member_names = {m["member"] for m in NAME_LISTS["members"] if m["era"] != "Nomination-era"}
    tables.append(("Second-hand self narrative (narration reporting a member's words without quotation marks)",
                   pd.DataFrame([{"Second-hand turns": len(sh),
                                  "Of them, reporting the band as a whole": sum(1 for r in sh if "band named" in
                                                                                r["Attribution rule"]),
                                  "Of them, reporting an other-era member": sum(1 for r in sh if r["Reported member"] in
                                                                                member_names),
                                  "Records with second-hand turns": len({r["Record ID"] for r in sh}),
                                  "Band-voice speech turns, for comparison": sum(1 for r in speech if r["Band voice"] ==
                                                                                 "yes")}])))
    tables.append(("Interview dates", pd.DataFrame(
        [dict({"Turn type": k}, **{"Stated in the text": sum(1 for r in turn_rows if r["Turn type"] == k and
                                                             r["Interview date source"] == "stated in the text"),
                                   "Publication-date proxy": sum(1 for r in turn_rows if r["Turn type"] == k and
                                                                 r["Interview date source"] != "stated in the text")})
         for k in kinds])))
    pas = [r for r in speech if r["Attribution rule"].startswith("Q&A passage")]
    tables.append(("Passages inside Q&As that the format check had left as prose, read as part of the interview",
                   pd.DataFrame([{"Passages": sum(r.get("Q&A passages read as interview", 0) or 0 for r in rec_rows),
                                  "Records": sum(1 for r in rec_rows if r.get("Q&A passages read as interview")),
                                  "Answer turns": sum(1 for r in pas if r["Question"] != "yes"),
                                  "Question turns": sum(1 for r in pas if r["Question"] == "yes"),
                                  "Band-voice turns": sum(1 for r in pas if r["Band voice"] == "yes"),
                                  "Band-voice words": sum(r["Words"] for r in pas if r["Band voice"] == "yes")}])))
    tables.append(("Repeated turns (the same text as an earlier turn: Repeat of)", pd.DataFrame(
        [dict({"Turn type": k},
              **{"Turns": sum(1 for r in turn_rows if r["Turn type"] == k),
                 "Repeats of an earlier turn": sum(1 for r in turn_rows if r["Turn type"] == k and r["Repeat of"]),
                 "Words in the repeats": sum(r["Words"] for r in turn_rows if r["Turn type"] == k and r["Repeat of"]),
                 "Band-voice repeats": sum(1 for r in turn_rows if r["Turn type"] == k and r["Repeat of"] and
                                           r["Band voice"] == "yes"),
                 "Band-voice words in the repeats": sum(r["Words"] for r in turn_rows if r["Turn type"] == k and
                                                        r["Repeat of"] and r["Band voice"] == "yes")})
         for k in kinds])))
    by_act = defaultdict(Counter)
    heard = defaultdict(Counter)
    year = {}
    for r in rec_rows:
        a = r["Focal Act"]
        year.setdefault(a, r.get("Grammy Year"))
        by_act[a]["Records"] += 1
        by_act[a]["Records processed"] += r.get("Processed") == "yes"
    for r in turn_rows:
        a = r["Focal Act"]
        c = by_act[a]
        c["Turns"] += 1
        if r["Turn type"] == "speech":
            c["Speech turns"] += 1
            if r["Band voice"] == "yes":
                c["Band-voice turns"] += 1
                c["Band-voice words"] += r["Words"]
                c["Band-voice turns (main set)"] += r["Main set"] == "yes"
                heard[a][r["Speaker name"] or "(unspecified)"] += 1
            if r["Speaker type"] in ("other-era member", "other-era member, unspecified"):
                c["Other-era member turns"] += 1
            if r["Speaker type"] in ("unknown", "ambiguous"):
                c["Unknown speech turns"] += 1
        elif r["Turn type"] == "second-hand":
            c["Second-hand turns"] += 1
    acts = sorted(by_act, key=lambda a: (year.get(a) or 0, a))
    none = [a for a in acts if not by_act[a]["Band-voice turns"]]
    tables.append(("Acts with no band-voice turns at all", pd.DataFrame([{"Focal Act": a, "Grammy Year": year.get(a),
                                                                          "Records": by_act[a]["Records"]} for a in none]
                                                                        or [{"Focal Act": "(none)"}])))
    act_df = pd.DataFrame([dict({"Focal Act": a, "Grammy Year": year.get(a)},
                                **{k: by_act[a][k] for k in ("Records", "Records processed", "Turns", "Speech turns",
                                                             "Band-voice turns", "Band-voice turns (main set)",
                                                             "Band-voice words", "Second-hand turns",
                                                             "Other-era member turns", "Unknown speech turns")},
                                **{"Band-voice share of speech turns": round(by_act[a]["Band-voice turns"] /
                                                                             by_act[a]["Speech turns"], 3)
                                   if by_act[a]["Speech turns"] else 0,
                                   "Members heard (turns)": "; ".join(f"{k} ({v})" for k, v in heard[a].most_common())})
                           for a in acts])
    # by publication (version 1.2: "www.nme.com" is nme.com, "julianlennon-photography.com." without its full stop)
    pub = defaultdict(Counter)
    for r in rec_rows:
        o = outlet_key(r["Outlet / Source Name"])
        pub[o]["Records"] += 1
        pub[o]["Unresolved spans"] += r.get("Unresolved spans", 0) or 0
        pub[o]["Records with unresolved spans"] += bool(r.get("Unresolved spans"))
        pub[o]["Records flagged"] += bool(r.get("Flags"))
    for r in speech:
        o = outlet_key(r["Outlet / Source Name"])
        pub[o]["Speech turns"] += 1
        pub[o]["Unknown speech turns"] += r["Speaker type"] in ("unknown", "ambiguous")
        pub[o]["Band-voice turns"] += r["Band voice"] == "yes"
    pub_df = pd.DataFrame([dict({"Outlet / Source Name": o}, **{k: c[k] for k in (
        "Records", "Speech turns", "Band-voice turns", "Unknown speech turns", "Unresolved spans",
        "Records with unresolved spans", "Records flagged")},
        **{"Unknown share of speech turns": round(c["Unknown speech turns"] / c["Speech turns"], 3) if c["Speech turns"]
           else 0}) for o, c in pub.items()])
    pub_df = pub_df.sort_values(["Records", "Outlet / Source Name"], ascending=[False, True])
    # by publication year
    yr = defaultdict(Counter)
    for r in turn_rows:
        y = (r["Publication Date"] or "")[:4] or "(none)"
        yr[y]["Turns"] += 1
        if r["Turn type"] == "speech":
            yr[y]["Speech turns"] += 1
            yr[y]["Band-voice turns"] += r["Band voice"] == "yes"
            yr[y]["Unknown speech turns"] += r["Speaker type"] in ("unknown", "ambiguous")
    for r in rec_rows:
        yr[(r.get("Publication Date") or "")[:4] or "(none)"]["Records"] += 1
    yr_df = pd.DataFrame([dict({"Publication year": y}, **{k: yr[y][k] for k in (
        "Records", "Turns", "Speech turns", "Band-voice turns", "Unknown speech turns")}) for y in sorted(yr)])
    # the watch list: sources by outlet (not by any URL that contains the name), and the acts with the most relatives
    # or namesakes on the name list (ties in alphabetical order)
    rel_acts = sorted({x["act"] for x in NAME_LISTS["nonmembers"]},
                      key=lambda a: (-sum(1 for x in NAME_LISTS["nonmembers"] if x["act"] == a), a))[:8]
    watch = []
    for w in WATCH_SOURCES:
        rs = [r for r in rec_rows if outlet_key(r["Outlet / Source Name"]) == w]
        ids = {r["Record ID"] for r in rs}
        sp_ = [r for r in speech if r["Record ID"] in ids]
        watch.append({"Watched": "source: " + w, "Records": len(rs), "Speech turns": len(sp_),
                      "Band-voice turns": sum(r["Band voice"] == "yes" for r in sp_),
                      "Unknown speech turns": sum(r["Speaker type"] in ("unknown", "ambiguous") for r in sp_),
                      "Unresolved spans": sum(r.get("Unresolved spans", 0) or 0 for r in rs),
                      "British-style records": sum(r.get("Quotation convention") == "british" for r in rs)})
    for a in rel_acts:
        rs = [r for r in rec_rows if r["Focal Act"] == a]
        ids = {r["Record ID"] for r in rs}
        sp_ = [r for r in speech if r["Record ID"] in ids]
        watch.append({"Watched": "act with relatives or namesakes on the list: " + a, "Records": len(rs),
                      "Speech turns": len(sp_), "Band-voice turns": sum(r["Band voice"] == "yes" for r in sp_),
                      "Unknown speech turns": sum(r["Speaker type"] in ("unknown", "ambiguous") for r in sp_),
                      "Unresolved spans": sum(r.get("Unresolved spans", 0) or 0 for r in rs),
                      "Band-voice turns lowered for a relative": sum(1 for r in sp_ if "relative" in
                                                                     r["Attribution rule"])})
    # records
    rec_cols = ["Record ID", "Focal Act", "Grammy Year", "Title / Headline", "Outlet / Source Name", "URL",
                "Publication Date", "Route", "Detail", "Processed", "Quotation convention", "Spans", "Turns",
                "Speech turns", "Band-voice turns", "Band-voice turns (main set)", "Second-hand turns",
                "Unknown speech turns", "Unresolved spans", "Unresolved characters", "Round trip", "Flags",
                "Q&A passages read as interview"]
    rec_df = pd.DataFrame([{c: r.get(c, "") for c in rec_cols} for r in rec_rows])
    flagged = rec_df[(rec_df["Flags"].astype(str) != "") | (rec_df["Round trip"].astype(str).str.startswith("fail"))]
    sh_df = pd.DataFrame([{c: r[c] for c in ("Record ID", "Turn ID", "Focal Act", "Reported member", "Confidence",
                                             "Attribution rule", "Turn text", "Interview date", "Outlet / Source Name",
                                             "Publication Date")} for r in sh])
    stated = [r for r in turn_rows if r["Interview date source"] == "stated in the text"]
    st_df = pd.DataFrame([{c: r[c] for c in ("Record ID", "Turn ID", "Focal Act", "Turn type", "Speaker type",
                                             "Speaker name", "Interview date", "Publication Date",
                                             "Interview date evidence", "Route")} for r in stated])
    sample = draw_turn_sample(turn_rows, texts, route_of)
    # for the known limits: records whose byline is a member (by name or alias, in any case)
    member_keys = {m["member"].lower() for m in NAME_LISTS["members"]} | {
        a.lower() for m in NAME_LISTS["members"] for a in (m.get("aliases") or [])}
    named = [r for r in turn_rows if r["Speaker type"] in ("interviewer", "journalist") and
             (r["Speaker name"] or "").lower() in member_keys]
    byline_recs = {r["Record ID"] for r in named}
    limits = {"byline records": len(byline_recs),
              "byline narration words": sum(r["Words"] for r in turn_rows if r["Record ID"] in byline_recs and
                                            r["Turn type"] == "narration"),
              "byline turns": len(named),
              "records with unresolved": sum(1 for r in rec_rows if r.get("Unresolved spans")),
              "unresolved on musicmusingsandsuch.com": pub["musicmusingsandsuch.com"]["Unresolved spans"]}
    readme = readme_rows(counts, dev_note, limits)
    with pd.ExcelWriter(OUT_DASH, engine="openpyxl") as xw:
        _sheet(readme).to_excel(xw, sheet_name="Read me", index=False, header=False)
        r0 = 1
        for title, tab in tables:
            _sheet(tab).to_excel(xw, sheet_name="Dashboard", index=False, startrow=r0)
            xw.book["Dashboard"].cell(row=r0, column=1, value=title).font = Font(bold=True)
            r0 += len(tab) + 4
        _sheet(act_df).to_excel(xw, sheet_name="By act", index=False)
        _sheet(pub_df).to_excel(xw, sheet_name="By publication", index=False)
        _sheet(yr_df).to_excel(xw, sheet_name="By year", index=False)
        _sheet(pd.DataFrame(watch)).to_excel(xw, sheet_name="Watch list", index=False)
        _sheet(rec_df).to_excel(xw, sheet_name="Records", index=False)
        _sheet(flagged).to_excel(xw, sheet_name="Flagged records", index=False)
        _sheet(sh_df).to_excel(xw, sheet_name="Second-hand", index=False)
        _sheet(st_df).to_excel(xw, sheet_name="Stated dates", index=False)
        _sheet(sample).to_excel(xw, sheet_name="Turn review", index=False, startrow=3)
        wb = xw.book
        ws = wb["Read me"]
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 120
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            row[0].font = Font(bold=True)
        ws = wb["Dashboard"]
        ws.column_dimensions["A"].width = 44
        for col in "BCDEFGHI":
            ws.column_dimensions[col].width = 17
        for name, widths in (("By act", [26, 8, 8, 9, 8, 9, 10, 11, 11, 10, 11, 10, 11, 80]),
                             ("By publication", [34, 8, 9, 10, 10, 10, 11, 9, 11]),
                             ("By year", [12, 9, 9, 10, 11, 11]),
                             ("Watch list", [60, 9, 10, 10, 11, 10, 12]),
                             ("Records", [9, 22, 8, 50, 26, 30, 12, 14, 18, 12, 11, 7, 7, 8, 9, 10, 10, 10, 10, 11, 10, 50,
                                          12]),
                             ("Flagged records", [9, 22, 8, 50, 26, 30, 12, 14, 18, 12, 11, 7, 7, 8, 9, 10, 10, 10, 10, 11, 10,
                                                  50, 12]),
                             ("Second-hand", [9, 11, 22, 22, 10, 40, 90, 12, 24, 12]),
                             ("Stated dates", [9, 11, 22, 11, 16, 22, 10, 12, 70, 14])):
            ws = wb[name]
            ws.freeze_panes = "B2"
            for i, w in enumerate(widths[:26]):
                ws.column_dimensions[chr(65 + i)].width = w
            ws.auto_filter.ref = ws.dimensions
        ws = wb["Turn review"]
        ws["A1"] = ("Optional. 66 turns drawn at random (seeded) across the rules and confidence levels, at most one per "
                    "record: 60 from articles and 6 from passages inside Q&As. For each, read the turn and the text "
                    "before it, then say in 'Right speaker?' whether the speaker given is right.")
        ws["A2"] = ("Band voice = band member, band member (unspecified), band collective. A second-hand turn's speaker "
                    "is the journalist; judge its 'Reported member'.")
        ws["A1"].font = Font(bold=True)
        last = 4 + len(sample)
        dv = DataValidation(type="list", formula1='"yes,no: another member,no: not a member,no: a member (given as '
                                                  'someone else),not speech,not sure"', allow_blank=True)
        ws.add_data_validation(dv)
        col_r = chr(65 + list(sample.columns).index("Right speaker?"))
        dv.add(f"{col_r}5:{col_r}{last}")
        fill = PatternFill("solid", fgColor="FFF2CC")
        for row in ws.iter_rows(min_row=5, max_row=last):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            for c in row[-2:]:
                c.fill = fill
        for i, w in enumerate([4, 9, 11, 18, 30, 26, 10, 34, 60, 90, 18, 30]):
            ws.column_dimensions[chr(65 + i)].width = w
        for c in ws[4]:
            c.font = Font(bold=True)
        ws.freeze_panes = "C5"


def draw_turn_sample(turn_rows, texts, route_of):
    """66 turns: from the prose, band voice by rule (30), other people (16), unknown (8) and second-hand (6); and band
    voice from passages inside Q&As (6). One per record."""
    rng = random.Random(SAMPLE_SEED)
    pools = defaultdict(list)
    for r in turn_rows:
        rule = r["Attribution rule"]
        if len(r["Turn text"].strip()) < 30:
            continue
        if rule.startswith("Q&A passage"):
            if r["Band voice"] == "yes":
                pools["band voice, Q&A passage"].append(r)
            continue
        if route_of[r["Record ID"]] != "prose":
            continue
        if r["Turn type"] == "second-hand":
            g = "second-hand"
        elif r["Turn type"] != "speech":
            continue
        elif r["Band voice"] == "yes":
            g = ("band voice, rule 1" if rule.startswith("1") else "band voice, rule 5" if rule.startswith("5:") else
                 "band voice, rule 5b" if rule.startswith("5b") else "band voice, rule 5c" if rule.startswith("5c") else
                 "band voice, other rules")
        elif r["Speaker type"] in ("unknown", "ambiguous"):
            g = "unknown"
        elif rule.startswith("2"):
            g = "other person, rule 2"
        else:
            g = "other person, other rules"
        pools[g].append(r)
    quota = {"band voice, rule 1": 6, "band voice, rule 5": 8, "band voice, rule 5b": 6, "band voice, rule 5c": 4,
             "band voice, other rules": 6, "other person, rule 2": 8, "other person, other rules": 8, "unknown": 8,
             "second-hand": 6, "band voice, Q&A passage": 6}
    picked, used = [], set()
    for g, n in quota.items():
        pool = sorted(pools.get(g, []), key=lambda z: (z["Record ID"], z["Turn order"]))
        rng.shuffle(pool)
        k = 0
        for r in pool:
            if k >= n:
                break
            if r["Record ID"] in used:
                continue
            used.add(r["Record ID"])
            picked.append(r)
            k += 1
    rng.shuffle(picked)
    out = []
    for i, r in enumerate(picked, 1):
        before = lead_in(texts[r["Record ID"]], r["Span start"])
        who = r["Speaker type"] + (f": {r['Speaker name']}" if r["Speaker name"] else "")
        if r["Turn type"] == "second-hand":
            who = "journalist, reporting " + r["Reported member"]
        out.append({"#": i, "Record ID": r["Record ID"], "Turn ID": r["Turn ID"], "Focal Act": r["Focal Act"],
                    "Outlet / Source Name": r["Outlet / Source Name"], "Speaker given": who,
                    "Confidence": r["Confidence"], "Rule": r["Attribution rule"],
                    "Text before the turn": before, "The turn": excerpt(r["Turn text"], 1500),
                    "Right speaker?": "", "Your note": ""})
    return pd.DataFrame(out)


DASHBOARD_SHEETS = [
    ("Dashboard", "The main counts: turns by type and route, speech turns by speaker type, confidence by route, "
     "second-hand turns, interview dates, the passages inside Q&As read as part of the interview, repeated turns, and "
     "any act with no band voice."),
    ("By act", "One row per act: its records, turns, band-voice turns (all, and the main set) and band-voice words, "
     "second-hand, other-era and unknown turns, and the members heard, to compare the acts."),
    ("By publication", "One row per outlet: its records, speech and band-voice turns, unknown speakers and unresolved "
     "passages, to spot outlets the rules read badly."),
    ("By year", "Records and turns by publication year."),
    ("Watch list", "The sources and acts watched closely: NPR, the Guardian, Record Collector, "
     "musicmusingsandsuch.com and electricityclub.co.uk, and the eight acts with the most relatives or namesakes on "
     "the name list (ties in alphabetical order)."),
    ("Records", "Every record of the turn set, the 50 not processed included: its route, quotation convention, numbers "
     "of spans and turns, and the "
     "checks (round trip, unresolved passages, flags)."),
    ("Flagged records", "The records with a flag (a quotation over 1,500 characters, or most of the prose unresolved), "
     "to check by hand."),
    ("Second-hand", "Every second-hand turn (the journalist reporting a member's words without quotation marks), with "
     "the member it reports, to judge whether it should stay a category of its own."),
    ("Stated dates", "Every turn whose interview year is stated in the text, to check before turns are ordered by "
     "date."),
    ("Turn review", "An optional check by a person: 66 turns drawn at random across the rules, each with the text before "
     "it and a column to mark whether the speaker is right."),
]


def readme_rows(counts, dev_note, lim):
    """The dashboard's Read me, as (label, text) rows."""
    bv = counts["band-voice turns by route"]
    rows = [
        ("harrison_v7_dashboard_final", VERSION),
        ("", ""),
        ("WHAT THIS IS", ""),
        ("This workbook", "The summary and checks of the turns table (harrison_v7_turns_final.xlsx, and "
         "harrison_v7_turns_final.json in full), with its two linked tables: the spans (harrison_v7_spans_final.json, "
         "every character of every text in one span) and the records (harrison_v7_records_final.json). The Read me of "
         "harrison_v7_turns_final.xlsx defines every column of the turns table."),
        ("Records and turns", f"Records processed: {counts['records processed']:,} (the 49 forum threads and record "
         f"11698 are not). Turns: " + "; ".join(f"{k} {v:,}" for k, v in counts["turns"].items()) +
         f". Spans: {counts['spans']:,}."),
        ("Band voice", "Band-voice turns by route: " + "; ".join(f"{k} {v:,}" for k, v in bv.items()) +
         ". By confidence: " + "; ".join(f"{k} {v:,}" for k, v in counts["band-voice turns by confidence"].items()) +
         " (the main set is high and medium)."),
        ("Reproducibility", dev_note),
        ("", ""),
        ("THE SHEETS", "What each sheet is for."),
    ] + list(DASHBOARD_SHEETS) + [
        ("", ""),
        ("THE CHECKS", ""),
        ("Articles: blind check by AI agents (6 Oct 2026)", "Four samples of prose records (144 records, 1,958 quotations) were coded quotation by quotation by AI agents (Claude) that did not see the script's answers. The first three (112 records) were used to develop the rules; the fourth (32 records, 412 quotations) was held out and measured once. Held-out sample: every one of the 64 quotations the script gave to the band was a member's words (precision 100%), and it found 64 of the 68 member quotations that carry quotation marks (recall 94%; about 80% counting 13 member quotations on pages whose marks do not pair up). In the main set (high and medium) precision was 100% and recall 84%. On the development samples precision was 89% to 97% (main set 98% to 99%) and recall 78% to 84%. Accuracy of the speaker by confidence, over all four samples: high 96% (219 of 227), medium 96% (279 of 292), low 73% (140 of 191), against targets of 95% and 80%. Speech or not: 92% to 97% of the quotations the script calls speech are speech; it finds 78% to 85% of the speech (most of the rest are short unattributed fragments it leaves undecided, never band voice). Second-hand self narrative is found narrowly: about a third of the sentences the agents listed, two thirds of what the script finds being ones they listed. These were AI agents, not people; the validation against human coding is reported in the technical appendix."),
        ("Q&A passages: blind check by AI agents (7 Oct 2026)",
         "AI agents (Claude) that did not see the script's answers coded who speaks in samples of the turns the rule "
         "makes and of the narration it leaves inside the same passages. Two samples were used to develop the rule; a "
         "third was held out and measured once: 70 turns, made of 30 answers, 8 questions, 12 narration turns and 20 "
         "answers going on without the first person. 27 of the 30 answers the rule gave to a member were that member's "
         "words (90%), all 8 questions were the interviewer's, and 1 of the 12 narration turns was mostly a member's "
         "words. A member's answer going on without the first person was right 14 times in 20 (70%, below the 80% for "
         "medium), so those turns are low confidence (rule 'Q&A passage: a member's answer going on, without the first "
         "person'). These were AI agents, not people; the validation against human coding is reported in the "
         "technical appendix."),
        ("Q&A speakers: blind check by AI agents (12 Oct 2026)",
         "AI agents (Claude) that did not see the script's answers coded who speaks in 1,268 Q&A turns: every "
         "interviewer turn of 20 words or more in the first person in records with band voice (767), the interviewer "
         "turns that carry a member's name (77), 200 of the other interviewer turns drawn at random, and the 224 "
         "band-voice turns that end with a question to 'you'. 106 interviewer turns were wholly or partly a member's "
         "words (about 8,600 words); counting from the random 200, about 230 turns and 14,000 words in all (1% of the "
         "band-voice words), mostly answers with a question mark in them. Of the 224 band-voice turns, 14 were the "
         "interviewer's question and 35 ran the next question into the answer. Version 3 of stage 3 fixes the cases the "
         "page's labels make certain: 16 of the 106, and 27 stretches of text in 14 records in all. None of the 865 "
         "interviewer turns in the check moved, and a second blind check confirmed all 18 changes that are more than a "
         "label. These were AI agents, not people; the validation against human coding is reported in the technical "
         "appendix."),
        ("", ""),
        ("KNOWN LIMITS", "What to keep in mind when using the tables."),
        ("Human verified", "Every turn's Human verified is 'no': no turn in the table was checked one by one. The "
         "confidence levels rest on the blind checks above, made by AI agents; the validation against human coding of "
         "held-out records is reported in the technical appendix."),
        ("Low confidence and unknown speakers", f"Low-confidence band-voice turns "
         f"({counts['band-voice turns by confidence']['low']:,}) are outside the main set and stay in the table for the "
         f"robustness checks. Speech turns whose speaker was not found or could not be settled "
         f"({counts['unknown speech turns']:,}), and quotations left undecided, are never band voice, so some member "
         f"speech is missed."),
        ("Unresolved paragraphs", f"{counts['unresolved spans']:,} paragraphs in {lim['records with unresolved']:,} "
         f"records (the Unresolved spans columns) have quotation marks that do not pair up, so they are in neither band "
         f"voice nor public narrative. "
         f"{lim['unresolved on musicmusingsandsuch.com']:,} of them are on musicmusingsandsuch.com, which quotes long "
         f"excerpts from other writers. By publication and Records give them by outlet and by record."),
        ("Q&A answers read as questions", "Where a Q&A's questions carry no label, a guest's paragraph that contains a "
         "question mark (a rhetorical question, or one the guest quotes) is read as the interviewer's question. After "
         "version 3's fix, about 210 such turns remain (about 12,000 words, 0.8% of the band-voice words): no rule tested "
         "could find them without also moving real questions, so they are left as they are."),
        ("Questions inside answers", "When an answer and the next question share a paragraph, the whole paragraph goes "
         "to one speaker. Of the 224 band-voice turns that end with a question to 'you', 35 hold the next question and "
         "14 are the interviewer's question."),
        ("Second-hand turns", "The script finds about a third of the second-hand sentences the AI agents listed, so the "
         "second-hand counts are low. Second-hand turns are a category of their own, never band voice."),
        ("Bylines that name a member", f"{lim['byline records']:,} records give a band member as the writer "
         f"({lim['byline narration words']:,} words of narration). Some are the member's own writing (Rickie Lee "
         f"Jones's tribute to Walter Becker, record 554); in others the byline is wrong (record 1674). Their narration "
         f"counts as the journalist's, not as band voice, until a person checks them. In these records "
         f"{lim['byline turns']:,} interviewer and journalist turns carry the member's name as Speaker name, so filter "
         f"by Speaker type as well as by name."),
        ("Republished articles", f"{counts['turns that repeat an earlier turn']:,} turns repeat an earlier turn (the "
         f"same words), for example because an article was published more than once. To count each quotation once, "
         f"keep the turns whose Repeat of is blank."),
    ]
    return pd.DataFrame(rows, columns=["A", "B"])


def report(counts, dev_note, wrote):
    print(f"\n{'=' * 70}")
    print(f"  Records processed {counts['records processed']:,}; turns: "
          + "; ".join(f"{k} {v:,}" for k, v in counts["turns"].items()) + f"; spans {counts['spans']:,}")
    print("  Band-voice turns by route: " + "; ".join(f"{k} {v:,}" for k, v in counts["band-voice turns by route"].items()))
    print("  Band-voice turns by confidence: "
          + "; ".join(f"{k} {v:,}" for k, v in counts["band-voice turns by confidence"].items())
          + f"; records with band voice {counts['records with band-voice turns']:,}")
    print(f"  Unknown speech turns {counts['unknown speech turns']:,}; stated interview dates "
          f"{counts['stated interview dates']:,}; unresolved spans {counts['unresolved spans']:,}; round-trip failures "
          f"{counts['round-trip failures']}")
    print(f"  Q&A passages read as interview {counts['Q&A passages read as interview']:,} (band-voice turns from them "
          f"{counts['band-voice turns from Q&A passages']:,}); turns that repeat an earlier turn "
          f"{counts['turns that repeat an earlier turn']:,}")
    print("\n  " + dev_note)
    if wrote:
        print(f"\nWrote:\n  {OUT_TURNS}\n  {OUT_SPANS}\n  {OUT_RECORDS}\n  {OUT_XLSX}\n  {OUT_DASH}")
    else:
        print("\nDRY RUN: every check passed; nothing was written.")


if __name__ == "__main__":
    main()
