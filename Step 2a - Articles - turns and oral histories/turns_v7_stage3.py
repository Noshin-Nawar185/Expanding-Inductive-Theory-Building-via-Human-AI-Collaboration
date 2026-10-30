"""
===============================================================================
turns_v7_stage3.py: STAGE 3 OF THE TURN BUILD
WHO SPEAKS IN THE Q&As AND FIRST-PERSON ACCOUNTS
Step 2 (articles), stage 3. Version 3, 12 Oct 2026.
===============================================================================
Runs after turns_v7.py. The stages are those of Step 2 (articles) in the
technical appendix.
Reads   harrison_v7_CLEAN7.json           (never changed)
        harrison_v7_TURNS_stage1.json     the cleaned texts (from turns_v7.py)
        harrison_v7_TURNS_stage2.json     the routes and segments (from turns_v7.py)
Writes  harrison_v7_TURNS_stage3.json     for every labelled Q&A, unlabelled Q&A
                                          and first-person record: its blocks
                                          cut into turns, each with a speaker
        harrison_v7_TURNS_stage3.xlsx     summary, by act, every record, the
                                          labels, what is unresolved, and an
                                          optional sample of turns to check
The final name lists are built into this script. Prose records are left for
stage 4 (record 11698, machine-generated text, is prose; stage 4 skips it).

STAGE 3. WHO SPEAKS
   Each block is cut into spans that cover it exactly once:
       turn         one speaker's words (a labelled line, a question, an
                    answer, an account)
       sung         lyrics in a transcript ("BROWN: (Singing) ...", a song
                    clip after "(SOUNDBITE OF SONG ...)")
       heading      a section heading inside the block
       label        a speaker's label standing on its own line
       narration    the journalist's own words inside a block
       non-speech   stage directions and page furniture
   Each turn gets a speaker:
       band member                 a nomination-era member      (band voice)
       band member, unspecified    one of the nomination-era
                                   members, not said which      (band voice)
       other-era member            a member outside the nomination era
       interviewer                 the journalist or host
       other person                anyone else
       unknown / ambiguous         not settled by the rules
   with a confidence (high, medium, low) and the rule that decided it.

THE RULES
     1. Labels are matched to the line-up with the final name lists: the
        name-list review's decisions first; then full names, also written
        slightly differently ("Pål Waaktaar Savoy" for Pål Waaktaar-Savoy,
        "Paul S Denman" for Paul S. Denman); then first names, surnames,
        initials and the first words of a name ("Slim Jim" for Slim Jim
        Phantom), which count only when the member is named in full in the
        title, the introduction or before the label is first used (the
        full-name rule). A namesake who is not a member is someone else; one
        member of the act sharing a surname with a relative named in the
        record is given at a lower confidence, lowered once even when the
        namesake and the relative are the same person. Labels stage 2 does
        not read are read here: initials it keeps out as metadata ("CD:" for
        Corinne Drewery, "PR:" for Peter Relic) when the record names
        someone with them; a label with a semicolon ("Knopfler;"); initials
        run into the first word ("JLYeah."); a label the record uses,
        written with another separator or after stray characters ("CL– It's
        a story", "PWS. Yes."); and a guest's label run into the
        interviewer's line ("JR: Is there a contagious factor? NJ: Not for
        my type.").
     2. The interviewee (or narrator) comes from the title and introduction:
        the person the title is about ("Interview: Slim Jim Phantom",
        "Annie Lennox and Dave Stewart on ...", "Gary Younge interviews
        Tracy Chapman"); otherwise the person the introduction presents as
        speaking ("Knopfler met RC", "says Ray"); weaker fallbacks give a low
        confidence. Authors, hosts and named journalists are not interviewees.
     3. First-person accounts belong to a member only when the title or
        introduction presents the piece as theirs (Soundtrack of my life,
        "X on Y", "X: '...'", an interview), or it is on the member's own
        website. A person merely named at the head of the title (an obituary,
        a fan's tribute) does not make the account theirs.
     4. Questions are the interviewer's; so is a label whose lines are mostly
        questions, or that names the author or the outlet ("RS:",
        "Songfacts:", "Gregg Shapiro:"). An unlabelled question between
        labelled answers is the interviewer's, but not one just before the
        interviewer's own label: that is the guest asking back ("Do you know
        what The Onion is?"). In a Q&A whose answers carry the guest's label
        and whose questions carry none, a paragraph of the guest's answer
        with a question mark in it ("...the start of the year, wasn't it?",
        or a question the guest reports in quotation marks) stays the
        guest's when more of the answer follows without a label. It must be
        in the first person, say "you" to no one and not name the guest (the
        writer's own comments between answers do).
     5. An answer given to the interviewee that speaks of them only in the
        third person is not theirs (unknown).
     6. An answer printed without its label, after the interviewer's labelled
        question, request ("Tell me a little about that song.") or short
        statement put to the guest ("By my count the only bands not
        represented on your new record are..."), is the answerer's: the one
        person the other labels name, or the interviewee (medium). Not in
        radio transcripts, and not for photo captions, closing notes or lines
        that name the answerer.
     7. Caption transcripts (a dash starts each speaker): in an interview the
        speakers alternate at each dash, worked out over the whole programme;
        in a Q&A with an audience a turn is the guest's only when it answers a
        question (low). The video page's lines after the programme, and a
        tour listing's "Address:" and "Details:" lines, are page furniture.
   The workbook's Labels sheet says whether a label's lines are spoken or
   only sung.

THE CHECKS THE SCRIPT MAKES (it stops without writing if one fails)
   - CLEAN7 is the expected file (record count, turn set, fingerprint);
   - the stage 1 and stage 2 files belong to it, and stage 2 was made by
     version 2 of turns_v7.py: the route and segments of every Q&A and
     first-person record are worked out again and must match the file;
   - every block is covered by its spans exactly once.

OPTIONAL HAND CHECK
   The workbook's "Turn review" sheet has 60 turns, drawn with a fixed seed,
   to check by hand.

COLAB (one cell)
  from google.colab import drive
  drive.mount('/content/drive')
  %run /content/drive/MyDrive/harrison_v7/turns_v7_stage3.py

Set DRY_RUN = True to run every check and print the summary without writing.
===============================================================================
"""

import base64
import functools
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
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "turns_v7 stage 3, version 3 (12 Oct 2026)"
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN7.json")
IN_STAGE1 = os.path.join(DIR, "harrison_v7_TURNS_stage1.json")
IN_STAGE2 = os.path.join(DIR, "harrison_v7_TURNS_stage2.json")
OUT_STAGE3 = os.path.join(DIR, "harrison_v7_TURNS_stage3.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_TURNS_stage3.xlsx")
DRY_RUN = False

EXPECTED_RECORDS = 13993
EXPECTED_TURN_SET = 9316
# MD5 of "ID|URL" for every turn-set record in ID order, from CLEAN7 as written on 4 Oct 2026
TURN_SET_FINGERPRINT = "e07271c9aacf5776a0d593bbd64cf7c0"
# the routes version 2 of turns_v7.py gave on that file (5 Oct 2026)
STAGE2_COUNTS = {"labelled Q&A": 631, "unlabelled Q&A": 652, "first-person": 170, "prose": 7814, "forum thread": 49}
SKIP = {11698: "machine-generated text"}
SPEAKER_ROUTES = ["labelled Q&A", "unlabelled Q&A", "first-person"]
SAMPLE_SEED = 20261005
EXCEL_CELL_CAP = 32000
# counts this script produced on the same files when it was checked on 12 Oct 2026; a run on those files should
# give exactly these numbers
CHECK_COUNTS = {"band-voice turns": {"labelled Q&A": 6061, "unlabelled Q&A": 3901, "first-person": 161}, "records with band-voice turns": {"labelled Q&A": 281, "unlabelled Q&A": 317, "first-person": 59}, "turns": {"labelled Q&A": 35276, "unlabelled Q&A": 17160, "first-person": 579}}
# a fingerprint of version 2's stage 3 file on the same files: when the file in the folder is version 2's, the workbook
# gets a sheet of the changes before the file is overwritten
PREV_FINGERPRINT = "522a4c3de8b8df411f08d8c8ddc1b432"


# ═════════════════════════════════════════════════════════════════════════════
# STAGE 2 (version 2, as in turns_v7.py; stage 3 uses its label and question tests)
# ═════════════════════════════════════════════════════════════════════════════
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


S = types.SimpleNamespace(**{k: v for k, v in globals().items() if not k.startswith("__")})


# ═════════════════════════════════════════════════════════════════════════════
# STAGE 3
# ═════════════════════════════════════════════════════════════════════════════
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


# ═════════════════════════════════════════════════════════════════════════════
# THE NAME LISTS (final, 5 Oct 2026: line-ups, aliases and the name-list review's decisions)
# ═════════════════════════════════════════════════════════════════════════════
_NAME_LISTS_Z = "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"
NAME_LISTS = json.loads(zlib.decompress(base64.b64decode(_NAME_LISTS_Z)).decode("utf-8"))


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


def load_input():
    for path, made_by in ((IN_JSON, "the CLEAN7 step"), (IN_STAGE1, "turns_v7.py"), (IN_STAGE2, "turns_v7.py")):
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
    with open(IN_STAGE1, encoding="utf-8") as f:
        st1 = {x["Record ID"]: x for x in json.load(f)}
    with open(IN_STAGE2, encoding="utf-8") as f:
        st2 = {x["Record ID"]: x for x in json.load(f)}
    ids = {int(r["Record ID"]) for r in ts}
    if set(st1) != ids or set(st2) != ids:
        stop("The stage 1 and stage 2 files are not for this CLEAN7 file's turn set. Run turns_v7.py again first.")
    raw = {int(r["Record ID"]): text_of(r.get("Extracted Text")) for r in ts}
    if any(len(st1[i]["Normalised Text"]) != len(raw[i]) for i in ids):
        stop("The stage 1 texts do not match CLEAN7. Run turns_v7.py again first.")
    counts = Counter(x["Route"] for x in st2.values())
    if dict(counts) != STAGE2_COUNTS:
        stop("The stage 2 file was not made by version 2 of turns_v7.py (its route counts are " +
             "; ".join(f"{k} {v:,}" for k, v in counts.most_common()) + "). Run version 2 first.")
    return ts, st1, st2


SPEAKER_ORDER = [MEMBER, MEMBERS_UNSPEC, OTHER_ERA, OTHER_ERA_UNSPEC, INTERVIEWER, OTHER, UNKNOWN, AMBIGUOUS]
SPAN_KINDS = ["turn", "sung", "heading", "label", "narration", "non-speech"]


def words(text, x):
    return len(text[x["speech_start"]:x["end"]].split())


def excerpt(s, n):
    s = s.strip()
    return s if len(s) <= n else s[:n].rstrip() + " …"


def main():
    t0 = time.time()
    print(VERSION)
    ts, st1, st2 = load_input()
    print(f"Read the CLEAN7, stage 1 and stage 2 files: turn set {len(ts):,}")
    names, stoplist = load_names(NAME_LISTS)
    meta = {int(r["Record ID"]): r for r in ts}

    out, rows, problems, skipped = [], [], [], []
    todo = [i for i in sorted(st2) if st2[i]["Route"] in SPEAKER_ROUTES]
    for k, rid in enumerate(todo, 1):
        m, x = meta[rid], st2[rid]
        if rid in SKIP:
            skipped.append(rid)
            out.append({"Record ID": rid, "Route": x["Route"], "Detail": x["Detail"], "Skipped": SKIP[rid],
                        "Interviewee": None, "Spans": []})
            continue
        text = st1[rid]["Normalised Text"]
        title = text_of(m.get("Title / Headline"))
        # stage 2 again: the label and question tests stage 3 uses depend on the record (handles, Q&A titles)
        res = route(text, title, text_of(m.get("Page Type")).strip().lower(), text_of(m.get("Language")).strip() or "en")
        if res["route"] != x["Route"] or [list(s) for s in res["segments"]] != x["Segments"]:
            problems.append(f"record {rid}: stage 2 gives {res['route']} with other segments than the stage 2 file")
            continue
        act = text_of(m.get("Focal Act"))
        rec = Record(rid, text, title, act, x["Route"], x["Detail"], x["Segments"], names[act], stoplist,
                     text_of(m.get("Author / Interviewer")), text_of(m.get("Outlet / Source Name")), text_of(m.get("URL")))
        try:
            spans = speakers(rec)
        except AssertionError as ex:
            problems.append(f"record {rid}: the spans do not cover a block exactly ({ex})")
            continue
        out.append({"Record ID": rid, "Route": x["Route"], "Detail": x["Detail"], "Interviewee": rec.who,
                    "Spans": spans})
        rows.append((rec, m, spans))
        if k % 250 == 0:
            print(f"  stage 3: {k:,} of {len(todo):,} [{time.time() - t0:.0f}s]")
    if problems:
        stop(f"Stage 3 could not be checked in {len(problems)} records, e.g.\n" + "\n".join(problems[:5]) +
             "\nIf the message is about stage 2, run version 2 of turns_v7.py first; otherwise please report it "
             "with this output.")
    print(f"Stage 3: {len(rows):,} records cut into turns and checked [{time.time() - t0:.0f}s]")

    # summaries
    turns = Counter()
    kinds = Counter()
    bv_turns, bv_words, bv_recs, conf_bv = Counter(), Counter(), Counter(), Counter()
    who_kinds = Counter()
    for rec, m, spans in rows:
        who_kinds[(rec.route, rec.who["kind"])] += 1
        has = False
        for x in spans:
            kinds[(rec.route, x["kind"])] += 1
            if x["kind"] == "turn":
                turns[(rec.route, x["speaker_type"])] += 1
                if x["band_voice"]:
                    has = True
                    bv_turns[rec.route] += 1
                    bv_words[rec.route] += words(rec.text, x)
                    conf_bv[x["confidence"]] += 1
        bv_recs[rec.route] += has
    counts = {"band-voice turns": {r: bv_turns[r] for r in SPEAKER_ROUTES},
              "records with band-voice turns": {r: bv_recs[r] for r in SPEAKER_ROUTES},
              "turns": {r: sum(v for (rr, t), v in turns.items() if rr == r) for r in SPEAKER_ROUTES}}
    if counts != CHECK_COUNTS:
        dev_note = ("These counts differ from the ones the script gave on the same files when it was checked on 12 Oct. "
                    "Please report this with the run's output.")
    else:
        dev_note = "These counts match the ones the script gave on the same files when it was checked on 12 Oct."
    changes = changes_from_previous(out, rows)

    if DRY_RUN:
        report(counts, conf_bv, skipped, dev_note, wrote=False, changes=changes)
        return
    with open(OUT_STAGE3, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    write_workbook(rows, turns, kinds, bv_turns, bv_words, bv_recs, conf_bv, who_kinds, skipped, dev_note, changes)
    with open(OUT_STAGE3, encoding="utf-8") as f:
        back = json.load(f)
    if len(back) != len(todo) or sum(len(x["Spans"]) for x in back) != sum(len(s) for _, _, s in rows):
        stop("The written file did not read back correctly; delete it and run again.")
    report(counts, conf_bv, skipped, dev_note, wrote=True, changes=changes)
    print(f"[{time.time() - t0:.0f}s]")


def stage3_fingerprint(st3):
    lines = []
    for i in sorted(st3):
        for s in st3[i]["Spans"]:
            lines.append(f"{i}|{s['start']}|{s['end']}|{s['kind']}|{s.get('speaker_type')}|{s.get('speaker')}|"
                         f"{s.get('confidence')}|{s.get('speech_start')}")
    return hashlib.md5("\n".join(lines).encode()).hexdigest()


def changes_from_previous(out, rows):
    """When the stage 3 file in the folder is version 2's, the records whose turns changed: band-voice turns and words
    before and after, and what changed. None when there is no version 2 file to compare with."""
    if not os.path.exists(OUT_STAGE3):
        return None
    try:
        with open(OUT_STAGE3, encoding="utf-8") as f:
            old = {x["Record ID"]: x for x in json.load(f)}
    except (OSError, ValueError):
        return None
    if stage3_fingerprint(old) != PREV_FINGERPRINT:
        return None
    new = {x["Record ID"]: x for x in out}
    text = {rec.rid: rec.text for rec, m, s in rows}
    meta = {rec.rid: (rec, m) for rec, m, s in rows}

    def bv(spans, t):
        xs = [x for x in spans if x["kind"] == "turn" and x["band_voice"]]
        return len(xs), sum(len(t[x["speech_start"]:x["end"]].split()) for x in xs)

    def key(x):
        return (x["start"], x["end"], x["kind"], x.get("speaker_type"), x.get("speaker"), x.get("confidence"))

    out_rows = []
    for rid in sorted(new):
        if rid not in old or rid not in text:
            continue
        a, b = old[rid]["Spans"], new[rid]["Spans"]
        if [key(x) for x in a] == [key(x) for x in b]:
            continue
        t = text[rid]
        (n1, w1), (n2, w2) = bv(a, t), bv(b, t)
        before = {key(y) for y in a}
        changed = [x for x in b if x["kind"] == "turn" and key(x) not in before]
        rules = Counter(x["rule"] for x in ([x for x in changed if x["band_voice"]] or changed))
        rec, m = meta[rid]
        out_rows.append({"Record ID": rid, "Focal Act": rec.act, "Title / Headline": rec.title, "Route": rec.route,
                         "Detail": rec.detail, "Band-voice turns, version 2": n1, "Band-voice turns, version 3": n2,
                         "Band-voice words, version 2": w1, "Band-voice words, version 3": w2,
                         "Main rule of the changed turns (band voice first)": rules.most_common(1)[0][0] if rules else
                         "(boundaries or rule texts only)"})
    return pd.DataFrame(out_rows)


def _safe(v, cap=EXCEL_CELL_CAP):
    if isinstance(v, str):
        return ILLEGAL_CHARACTERS_RE.sub("", v)[:cap]
    return v


def _sheet(df):
    return df.apply(lambda col: col.map(_safe))


def readme_rows(n_records, skipped, dev_note, conf_bv):
    rows = [
        VERSION,
        "",
        "WHAT THIS FILE IS",
        f"Who speaks in the {n_records:,} labelled Q&A, unlabelled Q&A and first-person records of the turn set, worked "
        "out by rules (technical appendix, Step 2, stage 3). Each block that stage 2 found is cut into spans that cover it exactly "
        "once; each turn has a speaker, a confidence and the rule that decided it. The full results, with the exact "
        "character positions, are in harrison_v7_TURNS_stage3.json; this workbook is for reading and checking them. "
        "Prose records, and the prose parts of these records, are left for stage 4."
        + (f" Skipped: record {', '.join(str(x) for x in skipped)} ({'; '.join(SKIP[x] for x in skipped)})." if skipped else ""),
        "",
        "SPANS",
        "turn: one speaker's words. sung: lyrics in a transcript. heading: a section heading inside the block. label: a "
        "speaker's label on its own line. narration: the journalist's words inside a block. non-speech: stage "
        "directions and page furniture.",
        "",
        "SPEAKERS",
        "band member: a nomination-era member (band voice). band member, unspecified: one of the nomination-era "
        "members, not said which (band voice). other-era member: a member outside the nomination era. interviewer: the "
        "journalist or host. other person: anyone else. unknown: the rules could not settle it. ambiguous: a label "
        "that could be members of different eras.",
        "Band voice = turns by band members (named or unspecified). Band-voice turns by confidence: "
        + "; ".join(f"{k} {conf_bv.get(k, 0):,}" for k in ("high", "medium", "low")) + ".",
        "",
        "CONFIDENCE",
        "high: a label naming the member in full, or a first name, surname or initials of a member named in full "
        "before; the interviewer's labels and questions. medium: the interviewee or narrator named as the subject of "
        "the title or presented as speaking in the introduction; a label resolved with help (a namesake who is the "
        "interviewee, a relative; lowered once even when the namesake is the relative); an answer printed without its "
        "label after the interviewer's question; a caption transcript's turns, from the alternation of speakers. low: "
        "weaker evidence (a member only mentioned in the title, the member the introduction is about, a band named as "
        "the subject without saying which member answers, several members answering without saying which, an answer "
        "in a caption transcript of a Q&A with an audience).",
        "",
        "THE RULES, IN SHORT",
        "1. Labels: the name-list review's decisions; full names, also written slightly differently ('Pål Waaktaar "
        "Savoy', 'Paul S Denman'); then first names, surnames, initials and the first words of a name ('Slim Jim') only "
        "when the member is named in full in the title, the introduction or before the label is first used. A namesake "
        "who is not a member is someone else. 'CHRIS (Caller):' is a caller. Labels stage 2 does not read are read "
        "here: initials kept out as metadata ('CD:', 'PR:') when the record names someone with them, a label with a "
        "semicolon ('Knopfler;'), initials run into the first word ('JLYeah.'); from version 3, a label the record "
        "uses written with another separator or after stray characters ('CL– ', 'PWS. '), and a guest's label run into "
        "the interviewer's line ('JR: Is there a contagious factor? NJ: Not for my type.').",
        "2. Interviewee: the person the title is about; otherwise the person the introduction presents as speaking. "
        "Authors, hosts and journalists named as interviewing are not interviewees.",
        "3. First-person accounts belong to a member only when the title or introduction presents the piece as theirs, "
        "or it is on the member's own website (an obituary or a fan's tribute does not count).",
        "4. Questions are the interviewer's; so is a label whose lines are mostly questions or that names the author or "
        "the outlet. An unlabelled question between labelled answers is the interviewer's, but not one just before "
        "the interviewer's own label (the guest asking back). Version 3: in a Q&A whose answers carry the guest's "
        "label and whose questions carry none, a paragraph of the guest's answer with a question mark in it (a "
        "rhetorical question, or one the guest reports) stays the guest's when more of the answer follows without a "
        "label; it must be in the first person, say 'you' to no one and not name the guest.",
        "5. An answer given to the interviewee that speaks of them only in the third person is not theirs (unknown).",
        "6. An answer printed without its label, after the interviewer's labelled question (or, from version 3, a "
        "request such as 'Tell me a little about that song.' or a short statement put to the guest), is the "
        "answerer's: the one person the other labels name, or the interviewee (medium). Not in radio transcripts, and "
        "not for photo captions, closing notes or lines that name the answerer.",
        "7. Caption transcripts (a dash starts each speaker): in an interview the speakers alternate at each dash, "
        "worked out over the whole programme from the names in brackets, the host's own introduction, the questions and "
        "the names the two use for each other; in a Q&A with an audience a turn is the guest's only when it answers a "
        "question (low). The video page's own lines after the programme are page furniture.",
        "",
        dev_note,
        "",
        "SHEETS",
        "Summary: turns by route and speaker; spans by kind; band voice by route; interviewees found.",
        "By act: band voice by act, and which members are heard.",
        "Records: every record with its interviewee and its turns by speaker.",
        "Labels: every speaker label in every record, who it was matched to, and whether its lines are spoken or only "
        "sung. 'First use' is a short excerpt from where the label first appears, to judge the match; the full words "
        "of every turn are in the JSON file (and, from stages 4 to 9, in the turns table).",
        "Unresolved: records with turns whose speaker is unknown or ambiguous.",
        "Turn review: an optional check of 60 turns (seeded): read the turn, say whether the speaker is right.",
        "Changes from version 2: when the folder held version 2's stage 3 file, every record whose turns changed, with "
        "its band-voice turns and words before and after.",
    ]
    return pd.DataFrame({"A": rows})


def write_workbook(rows, turns, kinds, bv_turns, bv_words, bv_recs, conf_bv, who_kinds, skipped, dev_note, changes=None):
    # Summary tables
    t1 = pd.DataFrame([dict(Speaker=t, **{r: turns.get((r, t), 0) for r in SPEAKER_ROUTES},
                            Total=sum(turns.get((r, t), 0) for r in SPEAKER_ROUTES)) for t in SPEAKER_ORDER])
    t2 = pd.DataFrame([dict(Span=k, **{r: kinds.get((r, k), 0) for r in SPEAKER_ROUTES}) for k in SPAN_KINDS])
    t3 = pd.DataFrame([{"Route": r, "Records": sum(1 for rec, _, _ in rows if rec.route == r),
                        "Records with band-voice turns": bv_recs[r], "Band-voice turns": bv_turns[r],
                        "Band-voice words": bv_words[r],
                        "When checked (12 Oct): band-voice turns": CHECK_COUNTS["band-voice turns"].get(r, "")}
                       for r in SPEAKER_ROUTES])
    wk = ["member", "members", "act", "other", "mixed", "others", "unknown"]
    t4 = pd.DataFrame([dict(Interviewee=w, **{r: who_kinds.get((r, w), 0) for r in SPEAKER_ROUTES}) for w in wk])
    # By act, records, labels, unresolved
    acts = defaultdict(lambda: Counter())
    heard = defaultdict(Counter)
    year = {}
    rec_rows, lab_rows, unres = [], [], []
    labs = defaultdict(lambda: [0, None, Counter()])
    for rec, m, spans in rows:
        a = rec.act
        year.setdefault(a, m.get("Grammy Year"))
        acts[a]["Records"] += 1
        c = Counter(x["speaker_type"] for x in spans if x["kind"] == "turn")
        bv = [x for x in spans if x["kind"] == "turn" and x["band_voice"]]
        acts[a]["Records with band-voice turns"] += bool(bv)
        acts[a]["Band-voice turns"] += len(bv)
        acts[a]["Band-voice words"] += sum(words(rec.text, x) for x in bv)
        acts[a]["Other-era member turns"] += c.get(OTHER_ERA, 0) + c.get(OTHER_ERA_UNSPEC, 0)
        acts[a]["Unknown turns"] += c.get(UNKNOWN, 0) + c.get(AMBIGUOUS, 0)
        for x in spans:
            if x["kind"] == "turn" and x["speaker_type"] in (MEMBER, OTHER_ERA):
                heard[a][x["speaker"]] += 1
            if x["kind"] == "turn" and x["speaker_type"] == MEMBERS_UNSPEC:
                heard[a]["(unspecified)"] += 1
            if x["label"]:
                key = (rec.rid, x["label"])
                labs[key][0] += 1
                labs[key][2][x["kind"]] += 1
                if labs[key][1] is None or (labs[key][1][2]["kind"] == "sung" and x["kind"] != "sung"):
                    labs[key][1] = (rec, m, x)
        heard_rec = Counter(x["speaker"] or "(unspecified)" for x in bv)
        w = rec.who
        rec_rows.append({"Record ID": rec.rid, "Focal Act": rec.act, "Grammy Year": m.get("Grammy Year"),
                         "Title / Headline": rec.title, "URL": text_of(m.get("URL")), "Route": rec.route,
                         "Detail": rec.detail, "Interviewee": w["kind"], "Named": "; ".join(w["who"]),
                         "Found from": w["how"], "Confidence": w["confidence"], "Turns": sum(c.values()),
                         "Band-voice turns": len(bv), "Members heard": "; ".join(f"{k} ({v})" for k, v in heard_rec.most_common()),
                         "Interviewer turns": c.get(INTERVIEWER, 0), "Other-person turns": c.get(OTHER, 0),
                         "Other-era turns": c.get(OTHER_ERA, 0) + c.get(OTHER_ERA_UNSPEC, 0),
                         "Unknown turns": c.get(UNKNOWN, 0) + c.get(AMBIGUOUS, 0)})
        unk = [x for x in spans if x["kind"] == "turn" and x["speaker_type"] in (UNKNOWN, AMBIGUOUS)]
        if unk:
            unres.append({"Record ID": rec.rid, "Focal Act": rec.act, "Title / Headline": rec.title,
                          "URL": text_of(m.get("URL")), "Route": rec.route, "Unknown turns": len(unk),
                          "Of turns": sum(c.values()), "Why": Counter(x["rule"] for x in unk).most_common(1)[0][0],
                          "First unknown turn": excerpt(rec.text[unk[0]["speech_start"]:unk[0]["end"]], 400)})
    for (rid, lab), (n, (rec, m, x), kinds_l) in sorted(labs.items()):
        sung, spoken = kinds_l.get("sung", 0), n - kinds_l.get("sung", 0)
        lab_rows.append({"Record ID": rid, "Focal Act": rec.act, "Route": rec.route, "Label": lab, "Spans": n,
                         "Lines": "sung only" if not spoken else ("spoken and sung" if sung else "spoken"),
                         "Speaker": x["speaker_type"], "Name": x["speaker"] or "", "Confidence": x["confidence"],
                         "Rule": x["rule"], "First use": excerpt(rec.text[x["start"]:x["end"]], 200)})
    order = sorted(acts, key=lambda a: (year.get(a) or 0, a))
    by_act = pd.DataFrame([dict({"Focal Act": a, "Grammy Year": year.get(a)}, **acts[a],
                                **{"Members heard (turns)": "; ".join(f"{k} ({v})" for k, v in heard[a].most_common())})
                           for a in order])
    sample = draw_turn_sample(rows)
    readme = readme_rows(len(rows), skipped, dev_note, conf_bv)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        _sheet(readme).to_excel(xw, sheet_name="Read me", index=False, header=False)
        r0 = 1
        for title, tab in (("Turns by route and speaker", t1), ("Spans by kind", t2), ("Band voice by route", t3),
                           ("Interviewee or narrator found (records)", t4)):
            tab.to_excel(xw, sheet_name="Summary", index=False, startrow=r0)
            xw.book["Summary"].cell(row=r0, column=1, value=title).font = Font(bold=True)
            r0 += len(tab) + 4
        _sheet(by_act).to_excel(xw, sheet_name="By act", index=False)
        _sheet(pd.DataFrame(rec_rows)).to_excel(xw, sheet_name="Records", index=False)
        _sheet(pd.DataFrame(lab_rows)).to_excel(xw, sheet_name="Labels", index=False)
        _sheet(pd.DataFrame(unres)).to_excel(xw, sheet_name="Unresolved", index=False)
        _sheet(sample).to_excel(xw, sheet_name="Turn review", index=False, startrow=3)
        if changes is not None:
            _sheet(changes).to_excel(xw, sheet_name="Changes from version 2", index=False)
        wb = xw.book
        ws = wb["Read me"]
        ws.column_dimensions["A"].width = 120
        for c in ws["A"]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.value and str(c.value).isupper() and len(str(c.value)) < 60:
                c.font = Font(bold=True)
        ws = wb["Summary"]
        ws.column_dimensions["A"].width = 34
        for col in "BCDEF":
            ws.column_dimensions[col].width = 18
        for name, widths in (("By act", [26, 8, 9, 12, 12, 12, 12, 10, 90]),
                             ("Records", [9, 22, 8, 50, 30, 15, 30, 11, 30, 34, 10, 7, 10, 40, 10, 10, 10, 10]),
                             ("Labels", [9, 22, 15, 24, 7, 14, 22, 26, 10, 60, 70]),
                             ("Unresolved", [9, 22, 50, 30, 15, 9, 8, 50, 90])) + \
                (((("Changes from version 2", [9, 22, 50, 15, 30, 12, 12, 12, 12, 60]),) if changes is not None else ())):
            ws = wb[name]
            ws.freeze_panes = "B2"
            for i, w in enumerate(widths):
                ws.column_dimensions[chr(65 + i)].width = w
            ws.auto_filter.ref = ws.dimensions
        ws = wb["Turn review"]
        ws["A1"] = ("Optional. 60 turns drawn at random (seeded) across the routes and speakers. For each, read the "
                    "turn (and the text before it), then say in 'Right speaker?' whether the speaker given is right, "
                    "and add a note if anything is odd.")
        ws["A2"] = ("'band member' and 'band member, unspecified' are band voice; 'other-era member' is a member outside "
                    "the nomination era; 'interviewer' is the journalist or host; 'other person' is anyone else.")
        ws["A1"].font = Font(bold=True)
        last = 4 + len(sample)
        dv = DataValidation(type="list", formula1='"yes,no: another member,no: not a member,no: the interviewer,'
                                                  'no: a member (given as someone else),not speech,not sure"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        col_r = chr(65 + list(sample.columns).index("Right speaker?"))
        dv.add(f"{col_r}5:{col_r}{last}")
        fill = PatternFill("solid", fgColor="FFF2CC")
        for row in ws.iter_rows(min_row=5, max_row=last):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            for c in row[-2:]:
                c.fill = fill
        for i, w in enumerate([4, 9, 18, 34, 26, 14, 24, 10, 34, 50, 90, 18, 30]):
            ws.column_dimensions[chr(65 + i)].width = w
        for c in ws[4]:
            c.font = Font(bold=True)
        ws.freeze_panes = "C5"


def draw_turn_sample(rows):
    """60 turns: band voice in each route (30, a third of them at low or medium confidence), the interviewer, other
    people and unknown turns (30); at most one turn per record."""
    rng = random.Random(SAMPLE_SEED)
    pools = defaultdict(list)
    for rec, m, spans in rows:
        for x in spans:
            if x["kind"] != "turn" or len(rec.text[x["speech_start"]:x["end"]].strip()) < 40:
                continue
            if x["band_voice"]:
                g = ("band voice, " + ("high" if x["confidence"] == "high" else "lower") + " confidence", rec.route)
            elif x["speaker_type"] == INTERVIEWER:
                g = ("interviewer", "")
            elif x["speaker_type"] in (UNKNOWN, AMBIGUOUS):
                g = ("unknown", "")
            else:
                g = ("other person or other-era member", "")
            pools[g].append((rec, m, x))
    quota = {("band voice, high confidence", "labelled Q&A"): 8, ("band voice, high confidence", "unlabelled Q&A"): 4,
             ("band voice, lower confidence", "unlabelled Q&A"): 8, ("band voice, lower confidence", "labelled Q&A"): 2,
             ("band voice, high confidence", "first-person"): 2, ("band voice, lower confidence", "first-person"): 6,
             ("interviewer", ""): 8, ("other person or other-era member", ""): 12, ("unknown", ""): 10}
    picked, used = [], set()
    for g, n in quota.items():
        pool = sorted(pools.get(g, []), key=lambda z: (z[0].rid, z[2]["start"]))
        rng.shuffle(pool)
        k = 0
        for rec, m, x in pool:
            if k >= n:
                break
            if rec.rid in used:
                continue
            used.add(rec.rid)
            picked.append((rec, m, x, g[0]))
            k += 1
    rng.shuffle(picked)
    out = []
    for i, (rec, m, x, why) in enumerate(picked, 1):
        before = rec.text[max(rec.blocks[0][0], x["start"] - 400):x["start"]] if rec.blocks else ""
        out.append({"#": i, "Record ID": rec.rid, "Focal Act": rec.act, "Title / Headline": rec.title,
                    "URL": text_of(m.get("URL")), "Route": rec.route,
                    "Speaker given": x["speaker_type"] + (f": {x['speaker']}" if x["speaker"] else ""),
                    "Confidence": x["confidence"], "Rule": x["rule"],
                    "Text before the turn": excerpt(before, 400) if before.strip() else "",
                    "The turn": excerpt(rec.text[x["start"]:x["end"]], 1500), "Right speaker?": "", "Your note": ""})
    return pd.DataFrame(out)


def report(counts, conf_bv, skipped, dev_note, wrote, changes=None):
    print(f"\n{'=' * 66}")
    print(f"  {'':16} {'turns':>8} {'band voice':>11} {'records with band voice':>24}")
    for r in SPEAKER_ROUTES:
        print(f"  {r:16} {counts['turns'][r]:>8,} {counts['band-voice turns'][r]:>11,} "
              f"{counts['records with band-voice turns'][r]:>24,}")
    print("\n  Band-voice turns by confidence: " + "; ".join(f"{k} {conf_bv.get(k, 0):,}" for k in ("high", "medium", "low")))
    if skipped:
        print("  Skipped: " + ", ".join(f"{x} ({SKIP[x]})" for x in skipped))
    print("\n  " + dev_note)
    if changes is not None:
        print(f"  Compared with version 2's file: {len(changes):,} records changed (sheet 'Changes from version 2').")
    if wrote:
        print(f"\nWrote:\n  {OUT_STAGE3}\n  {OUT_XLSX}")
    else:
        print("\nDRY RUN: every check passed; nothing was written.")


if __name__ == "__main__":
    main()
