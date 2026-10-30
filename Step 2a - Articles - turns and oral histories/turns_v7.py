"""
===============================================================================
turns_v7.py: STAGES 1 AND 2 OF THE TURN BUILD
CLEAN THE TEXT, THEN SORT EACH RECORD BY FORMAT
Step 2 (articles), stages 1 and 2. Version 2, 5 Oct 2026.
===============================================================================
Runs after clean7_v7.py. The stages are those of Step 2 (articles) in the
technical appendix.
Reads   harrison_v7_CLEAN7.json           (never changed)
Writes  harrison_v7_TURNS_stage1.json     each turn-set record's cleaned text
        harrison_v7_TURNS_stage2.json     each turn-set record's format (route),
                                          the parts of its text, and the evidence
        harrison_v7_TURNS_stage2.xlsx     summary, every record's route, and a
                                          sample to check by hand
Only the 9,316 turn-set records ("In Turn Set" = Yes) are processed. CLEAN7 is
not changed, and the raw text is kept as it is: stage 1 writes a cleaned copy.

STAGE 1. CLEAN UP THE TEXT
   Every change swaps one character for one character, so position 1,000 in
   the cleaned text is position 1,000 in the raw text. Nothing else changes.
       curly double quotes, „ ‟ and ＂       -> "
       curly single quotes, ‛ ‚ ʼ ′ ´ `      -> '
       non-breaking and other odd spaces   -> ordinary space
       line and paragraph separators        -> line break
       « » in a pair within a paragraph     -> " "   (a lone » is often an
                                               arrow, "Read more »")
       7" 10" 12" (records), 12''           -> ″      (inch, not a quotation)
       ``x'' and ''x'' typed as quotes      -> "x"
   The script checks every record: same length, and every changed character
   is one of the swaps above. It stops if any record fails.

STAGE 2. SORT EACH RECORD BY FORMAT
   Four routes, plus forum threads, which are not processed further:
       labelled Q&A     each turn starts with a speaker label: "LAUPER:",
                        "Q:", "NVE:", "[Host]", "PHIL OAKEY –", NPR's
                        "KEN TUCKER, BYLINE:", a label on its own line, labels
                        inside paragraphs, lecture transcripts, oral histories
       unlabelled Q&A   questions followed by first-person answers, no labels;
                        a question may run straight into its answer
       first-person     one person speaking throughout (essay, diary,
                        "Soundtrack of my life"), or one person per section
                        under a heading naming them ("How we made" features)
       prose            everything else: the journalist's narration with
                        quotations; stage 4 works on these
       forum thread     posts by several users (Page Type)
   Each record is cut into segments that cover the text once, in order: a
   prose introduction, the Q&A or first-person block(s), and a prose tail.
   Quotations in the introduction and tail go through the prose route.
   The checks for "I", "we" and "you" follow each record's Language (English,
   Spanish, French, German, Italian, Danish).
   The "Read me" sheet of the workbook lists every sub-type and flag.

OPTIONAL HAND CHECK
   The workbook's "Routing review" sheet has 54 records, drawn with a fixed
   seed, to check the routes by hand. The check of version 1 found the
   problems this version fixes (technical appendix, Validation and checks).

COLAB (one cell)
  from google.colab import drive
  drive.mount('/content/drive')
  %run /content/drive/MyDrive/harrison_v7/turns_v7.py

Set DRY_RUN = True to run every check and print the summary without writing.
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
import zlib
from collections import Counter, defaultdict

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "turns_v7 stages 1-2, version 2 (5 Oct 2026, after the routing review)"
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
IN_JSON = os.path.join(DIR, "harrison_v7_CLEAN7.json")
OUT_STAGE1 = os.path.join(DIR, "harrison_v7_TURNS_stage1.json")
OUT_STAGE2 = os.path.join(DIR, "harrison_v7_TURNS_stage2.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_TURNS_stage2.xlsx")
DRY_RUN = False

EXPECTED_RECORDS = 13993
EXPECTED_TURN_SET = 9316
# MD5 of "ID|URL" for every turn-set record in ID order, from CLEAN7 as written on 4 Oct 2026
TURN_SET_FINGERPRINT = "e07271c9aacf5776a0d593bbd64cf7c0"
SAMPLE_SEED = 20261005
EXCEL_CELL_CAP = 32000
ROUTES = ["labelled Q&A", "unlabelled Q&A", "first-person", "prose", "forum thread"]
# counts this script produced on the same CLEAN7 file when it was checked on 5 Oct 2026; a run on that file
# should give exactly these numbers
CHECK_COUNTS = {"labelled Q&A": 631, "unlabelled Q&A": 652, "first-person": 170, "prose": 7814, "forum thread": 49}


# ═════════════════════════════════════════════════════════════════════════════
# STAGE 1
# ═════════════════════════════════════════════════════════════════════════════
# one character for one character
DOUBLE = {"“": '"', "”": '"', "‟": '"', "„": '"', "＂": '"'}
SINGLE = {"‘": "'", "’": "'", "‛": "'", "‚": "'", "ʼ": "'", "′": "'", "´": "'", "`": "'"}
SPACES = {c: " " for c in "               　"}
LINES = {" ": "\n", " ": "\n"}
INCH = "″"
SPEECH_VERB = r"(?:says?|said|asks?|asked|adds?|added|recalls?|recalled|explains?|explained|told|tells|laughs?|laughed|notes?|noted|admits?|admitted|continues?|continued|replies|replied|insists?|insisted)"


def normalise(raw):
    """Return (text, counts). The text has exactly the length of raw; only the swaps below happen."""
    t = list(raw)
    n = Counter()
    # 1-3: one-for-one character swaps
    for i, ch in enumerate(t):
        if ch in DOUBLE:
            t[i] = DOUBLE[ch]; n["curly double quote"] += 1
        elif ch in SINGLE and not (ch == "`" and i + 1 < len(t) and t[i + 1] == "`") and not (
                ch == "`" and i > 0 and raw[i - 1] == "`"):
            t[i] = SINGLE[ch]; n["curly or odd single quote"] += 1
        elif ch in SPACES:
            t[i] = " "; n["unusual space"] += 1
        elif ch in LINES:
            t[i] = "\n"; n["line separator"] += 1
    s = "".join(t)
    # 4: guillemets, converted only in pairs within a paragraph (a lone » is often an arrow: "Read more »")
    s, k = _pair_guillemets(s)
    n["guillemet pair"] += k
    # 5: inch marks after 7, 10, 12 (7" single, 12" mix); not after another digit, a decimal point, a slash
    #    (ratings such as 9/10") or a comma
    out = list(s)
    for m in re.finditer(r'(?<![\d./,])(?:7|10|12)(")', s):
        out[m.start(1)] = INCH; n["inch mark"] += 1
    for m in re.finditer(r"(?<![\d./,])(?:7|10|12)('')(?!')", s):      # 12'' with two apostrophes
        out[m.start(1)] = INCH; out[m.start(1) + 1] = " "; n["inch mark"] += 1
    s = "".join(out)
    # 6: double quotes typed as two characters: ``like this'' or ''like this''
    s, k = _ascii_double_quotes(s)
    n["two-character double quote"] += k
    assert len(s) == len(raw)
    return s, {k: v for k, v in n.items() if v}


def _pair_guillemets(s):
    out = list(s)
    k = 0
    pos = 0
    for para in s.split("\n"):
        stack = None
        for j, ch in enumerate(para):
            if ch == "«":
                stack = j
            elif ch == "»" and stack is not None:
                out[pos + stack] = '"'; out[pos + j] = '"'; k += 1
                stack = None
        pos += len(para) + 1
    return "".join(out), k


OPEN2 = re.compile(r"(?:^|(?<=[\s(\[—\-\"/]))(``|'')(?=[A-Za-z0-9(\[\"'])")
CLOSE2 = re.compile(r"(?<=[A-Za-z0-9.,!?;:)\]\-])('')(?=$|[\s,.;:!?)\]—\-\"])")


def _ascii_double_quotes(s):
    """``x'' and ''x'' are double quotes typed as two characters. A pair becomes ' "x" '
    (quote moved next to the words, the spare character turned into a space). Unpaired ones are
    left alone unless the record clearly uses this style and the mark ends a quote before a
    speech verb (''... he said'' spanning paragraphs)."""
    out = list(s)
    k = 0
    pos = 0
    paired_total = 0
    loose = []
    for para in s.split("\n"):
        toks = []
        for m in OPEN2.finditer(para):
            toks.append((m.start(1), "open", m.group(1)))
        for m in CLOSE2.finditer(para):
            toks.append((m.start(1), "close", m.group(1)))
        toks.sort()
        opener = None
        used = set()
        for p, kind, g in toks:
            if kind == "open":
                if opener is None:
                    opener = p
                continue
            if opener is not None and p > opener + 1:
                o = pos + opener
                c = pos + p
                out[o], out[o + 1] = " ", '"'
                out[c], out[c + 1] = '"', " "
                used |= {opener, p}
                k += 1
                paired_total += 1
                opener = None
        for p, kind, g in toks:
            if p not in used and g == "''":
                loose.append((pos + p, kind))
        pos += len(para) + 1
    # a leftover `` is always an opening double quote
    for m in re.finditer(r"``", "".join(out)):
        out[m.start()], out[m.start() + 1] = " ", '"'; k += 1
    if paired_total >= 2:          # the record types double quotes this way: settle obvious loose ends
        for p, kind in loose:
            after = s[p + 2:p + 40]
            before = s[max(0, p - 2):p]
            if kind == "close" and (re.match(r"\s*(?:\w+\s+){0,2}" + SPEECH_VERB + r"\b", after)
                                    or re.search(r"[,.!?]$", before)):
                out[p], out[p + 1] = '"', " "; k += 1
            elif kind == "open" and (p == 0 or s[p - 1] == "\n"):
                out[p], out[p + 1] = " ", '"'; k += 1
    return "".join(out), k


ALLOWED = {('"', c) for c in DOUBLE} | {("'", c) for c in SINGLE} | {(" ", c) for c in SPACES} | \
          {("\n", c) for c in LINES} | {('"', "«"), ('"', "»"), (INCH, '"'), (INCH, "'"), (" ", "'"),
                                        ('"', "'"), (" ", "`"), ('"', "`")}


def check(raw, norm):
    """Every changed position must be one of the allowed swaps."""
    if len(raw) != len(norm):
        return "length differs"
    bad = [(i, raw[i], norm[i]) for i in range(len(raw)) if raw[i] != norm[i] and (norm[i], raw[i]) not in ALLOWED]
    return f"{len(bad)} unexpected changes, first at {bad[0][0]}: {bad[0][1]!r} -> {bad[0][2]!r}" if bad else ""


# ═════════════════════════════════════════════════════════════════════════════
# STAGE 2
# ═════════════════════════════════════════════════════════════════════════════
# first names: the 1990 US Census first-name lists (public domain) plus names from the corpus's member tables
# and British, Irish and European names common in music writing; used to tell a person's name from a heading
_FIRST_NAMES_Z = (
    "eNpNfVlyG0mU5In6EK0iJJIAQTaBokr6CwJBZAiJCFQugJKmC8w1xmxspi4wX/2ni/Vz9xfJMkO6v9j3PTLxn6Er+dd/vr7GCZgi"
    "EOJ+bA2hjBRyMEyHkKDqQhNOZMN9wNMA2tfYDYVSyhFMcwtgv4fX+9gGYdrPgvSFSbopV5Z2K5SmrNKG7P0owNK+FdIZUenok2F2"
    "qsroPGu4TsxOVUnhCq/fBmTR22jiIQwNKMee6KoRqi6dGZ/DGPekfki5smk0p7AHRsNkP1hi6tPRIp7aiCgYIeB0isAe3qdhZyn6"
    "YdKRNo+pRW4fUwet0iNeR4DlkuUnggJFUCZQxzWk6IeQOggT1a8EKzunqmyqIF8lRZcKWAZ7AuLf7qEdQw9bcUdAUSh1bWSqjZiy"
    "Nv4Ied99SPJNMv2RmkFGd9I38q936mcvBmHjFJ1l7ydBormIXZW6Wa9zm5WsUK1O5b6I38kdio8UnFMVYJ7wQCPthJEoLTpRGpWN"
    "qafcS26kUIRT74S6Z+Q2vdwS05qQqPQOsRXArmd1Gw9MWuuq7AqE1SahKxRGOwmlOsH6Sak8KWtLFjJDPD/OjPS5Md1YBejTwRAs"
    "Ckq7SsULhe2pvQQCC/oSGdeL171LykIoJubhJH+mXkgt5djUuwl9nZhT1tSEBdgSEnUyDV5R/sB2AncFHrLNndQJnWLHEjwlPJFA"
    "zNREbE8nAg2L1ZXTmWk5mY8WEdjLgZ1nDiroHE6BdSyj/e1mifFHlXRiQlQxDaPQFaEXu66rriR50/3+x4mGiE0cBtg/sPcFuYp1"
    "gkIvTq5Ou8pV41iFPAvuSfp7rNLgZjWYIkK9yodEW0lmjlFIBfrQbP1aTq8osKwglXFqG1mtICucNDCDfghqGn/USCNMxrUWAvlD"
    "iJTa6KxwUELRWdkpiXoCazq9S/IjflgYBqfgTO1E2EkWMmkZHUz21OTqy2AtCGU2pMIQyfJpKMp0ctXKQld42MO/3OTkOunDUKUI"
    "oYgZ4hXdR57MqKQ3c3wurVyfO4z6524yRKZbo47o6cS9C+ZHt1Md6nYNEtPtbbgEua5G8s6H8s7H8o4DR7enR/tEF+rFO6+SxllD"
    "kI/kPpDXcbwO42qyzEeDVsggGJ5VtEyKIkcZ0b2PHuChMvWTkJ1Hp7pvXQcJXYpTcFZUIRVwkoG79jLqvIQ6VYAul3bvBEd55HSG"
    "k7NOOOBpxg7E+Awj+pvukpidqMBsJgaWW2CWSd+wgYGiKB0aCTJNQhlKCy22b1i52Dn1stuPGXHthy5ZXE21S7BivToSPL52sDVy"
    "pjey+xrVcYFowjYFkg3oYXrkFJyZfS5Fl9A7jN4PjvTaRdUQcgEXGRb0mz4DG4fxZITx5uI9mA00FvgFVe7Cyn1NnMT8NHEKmGVN"
    "rGiTKtikycc0Yjh4V6/xnt4Nxh3T/v5uSfoUXoOFbsSi/WRzZUv2p3B6TYbFnq7a6F4JoXOSXnSmX92rVHBrlQUedR26GLBs2Hzp"
    "gEKCSHMaSqe3VBEtez5FPoOwQ18koXd+B48EzC0/xVfY2B2TUAoLIUJCw/kU5a+ahPGRflnz/sTG/YntzvAC+UTv4CrbdGk3QEiM"
    "jVEB/QjWPCCwEwUz0JwRKHu5T7GLWTG3Ye+tdHtIOQj3ynGKqDSUqpVuH6qQ8r/E2VqhEN13NVAIVSNVqmqZI88xZ/7EWcsnzZdB"
    "su+xHrBIMp6QN31PoIWhIYTsKqtqzhM5wsHQR6p6ItM4yP9hUGKGQUaTux8mmI4sBC/NC56Ijq8KkwvG1ndar/kpoZImFKCBRARj"
    "NAmh6PbU6w7JUv2jdBbgj9//H6QVB6gDHinTY1DjBN2SOQyY0PfFcqadBuhbGZRXPIhH4UoUNAmpUPrL0Z58ALAciupKYV0pGf6U"
    "o03vPhXW7wJdTCENrT50QaAKZALzAiwioip27M9J0amymyYn1y6uTR+y5Y91UCggUHaqSjihd1GBgLK4iGVV3qj+W9snABNtJ/mf"
    "3OPkHrPWWg99oAPyh0CjDyHtnYLY9RWrlHvljuVzFmdX50QahMGpmpqUZiFWyc1iFWbb1TINyu4I3GP+TSlfJyS2lKMwiqxou5Ep"
    "HZn6MbMHh2BVoJuYQ5NnzaQCYTXo6N3EMpmQsSPCGfd4UHgjOvRP49sbZFUTNia1pZHVbOxawAna6ABGNXYbb1DxJozXf4Q0WO8o"
    "nsCtFpsmxFcgFnNgtDPQeAJd6EarEHIV5KkNQIBIbINoZmq7qa1D/kAd3kVxdsEWAR29RA0mugEqwR+q0KSe1ElzoqVzJ8vUtIpJ"
    "RDhdIuytdf3BiSwwC2kX82uSM5NuuQiQGKNTFmfXzq4/uBOF1yahDJNbqt4WYS+qmtWHoSgIecSyQVESpYitU3BOs5BnaajSbKuI"
    "5Ue16XE7KdDSCoNTdt3o5Jxy5VAFN3H3U64cqyAd4dSSVSxeHh529aNj8ns5QL01otbUCmXCgPpaEXrmfk8P+8R1sAkfxn0YZ1FW"
    "+yTcu4ouLdRqC9pD8IQOtb5YfFxooBUpRVUpCWGWqmE1m1ygz00SypLXrqGZnVXbtOy+eiUfPJQJhRW1sjG2eRXRFVFUzara/Is2"
    "z905HY1VpbHYHipH5zSbMF4eVOva7I7/0JLbCC0osgbF3iaQfzRodAZXBtPYAtYgEwJxJ529VMx4Y9Rx0syydiBakThP4ujUVv4Q"
    "aCRfMeSKonNf9avO4BoDPWbrb2q7b2rDb2rLp9BXVifQBF+YSXJ/kqtTdWbtfRZcq6+WJgV2wlBfBRmV2WuGqlir+oIV637WGFxn"
    "JFhWUzUhllZSsfIkZvBoagYSo5NyMaLBAl2heEbPTnBDbicnGbur6qybnFqyVFXh4Uwsv8gByxg5FSduev+BJoM9SINoKyGjTKAO"
    "q0fKIxV/jypNx2tU9WgLfCm0qgqB4mfkIExiuiG3LqgKUJq1Zp0cZsk9GlLlMAv5Q/qX5uxgFmZ7H9Y+bOXZw3Jm9khsZpkpmDzm"
    "KPmRLW/MBEt2YnMAmlbi3sgfKVLPygXEQUMNO2HPL1aNJIwkOMvcwf6DXS2n18QIyrsJJBNp2aC8cwGrEQmDhISNMApuReG33sO2"
    "ar2tD9Ztbb2tN1rjI1BF2QabMhFdoWYDKQXnqs5VKGDGeeIg3Kqna6MNmrtZGirHKrCYqihdznLaWAjnMDATYnEabDJIc0Y8YvvA"
    "yFb/2Kz5o01vb0JM7yAoNolBp+yqC3yQj6Xm1ezzBCworoJiKu0rMRJQbY04nSjqyIrPYGr/ojG+tNVyO1MUuy23xigVmyO+IinF"
    "CvnMzReJwywqDFtqOSWpMQ03wurHyGoua4lJI+c4EgqEIbLQmWyDVohuCpzE7wzO9/Uk0OLelbKmlHVKkVfDwvliQY9VOllj4RbV"
    "QVBVSo2jD3CuYUEYqxZDdX84OyAF56otYYjvpKwYAMaqsseWfdZyO42sbAnodrGm+MN7Le+0ap9V+yfvimoHU/uXuXuZexd2I6/V"
    "qXco3p90o0WOfYpTK0aFGm0WgWUGAPGa2EtM7B4m7x4m7GIZyuXUIYcmGt3wpOEm7Pc2M7gJb9N+b3Q42dTByFYzN6FpaQOxveFA"
    "aIgxCwQ8Ygp9Y7WAQHO5aDENMrJOAYRKesMNcyhP6I5BUSRDIp9AgGdZ0LZi1PwbLCEBQdiSdlLF1qkqo3PVnw3c5Ch1VVpNBzFR"
    "mXFRZOag0VxE7iQnxoa5lWVnlPNJkZk88MldnBvl47khdbDbvcI59shFUiUhbe2oJa86ZlvHDEDVBhbiiNztmEbNm2582nTjs6Yb"
    "nwbdBLWRG64LbrB/2Do5SxcTR7KUbneinas0qbjAy0vaCwMpixClayYEIqxOzIqJOpMtBG5iwOzzBvvwN9wdAVYtHHU5S0NKniPd"
    "cBPlRnsoN9pCubFltD2vSUiNV/qFbSPDKAXNi2SjBixVVRTyrmW09vBIcYLkAdeQ4166bkF66NluYpJeUjpSJ/4BvRaPWlZslW4d"
    "txuz6LU4MHpL4iTcObeVGjLGInBPou+ycXLsSFGEzGgLiojE4JQfIGnjKJHs1jpZw/EuKVVWGmyFCeoVdybfAt2hZtooPdBzcJqF"
    "XawSLTGVbP1WYVENjRhkls3MeOTkej3jkbmKuNHm8U1UiWS3Wu2ijcZ8YdrzRD1UaANa7+IRyCofvepjdx+VN/Y2y4riSmTrtfbg"
    "s0qnV7qVTGzYGBZHail3L17brmFirl3R4uMV22Y38Sdm3jfIH3aTnLDecK4KlI5rUQ/xTDawwpFqHwgm8VCAnavk74FuWyzib9Lp"
    "jIbIuBs0wEJg/Ngx87QRSEtFYXYIU008/USSy57ILSvQBFTdKV53yglFWDlAsEVD5SKWhfRhgcmjgNNkSlUYWPmLFYGbZTyB0O5B"
    "qvx++HnjR583WH/caI/YiI7kSl1iybP92QG7fW4kG5ZLoCMZTbDZCbBbYLxD07HE80rQjaY7N5rugKou7zzc+PznBjMfgESG0DFU"
    "TTBvNA+68bxkJ1+64n6V6lfxKLgtjtWlc1vVErdAyKABz8CgcGJwU8YDgUO3GMKFNthxW8rQnXRReDUYd1y4mMD5t3EvjZEVd+Ru"
    "+s1os+ubsUU7HtW92cIYScXa6AaHcsJMMgdXLMMNUT7eWK7p0FiUsY17M6FjXmC3EBCFroVMX/iwt9AuwMJHvwVGv36o7FrIpEXg"
    "unfxatVDaEZ7+5nRXkAdm2AtbGrLba4F9v0X+4N1rcR9MaYtXp9Z8OB6sWdPYTTOBHsZ7ukJSmyxH939FWdUIipTFsLmRP94OLJ4"
    "e0NIbzyZMEJMcB1t0WClvNDVsYWuji2O5fprYWViD6bjC10JW2j4WXD0WfA2GLB0oImqVxhwVFpg1rvghS7DIuyoxWMEE6K7JYcq"
    "yKQnMRzGyDAR5aI6qPaL65dqUNxEl6zA6V8CTOAd05aUnGSVeKHRcsFrVkApflgXt8AiatEqH5KCSx4IxlDiKw7kFryMRSykRqay"
    "9E6QCPuzQIfvsxBdOsIrFkXLiyUL3tBacK23aD2LmBjWW8MyOImRqra/og1Dmghwe4JLjO8Lr53tCfGlh56NRRlQ6ElRMotSV0Zn"
    "eNgRaUzNnlHpoTmw6C8w4gUtoySMxCyUlpsjtCs3G41LserN21qLFvl24pNH1NoT11YLba4vfGtdHMkTsEMfD0Y/DmZjO2FXa3HC"
    "dUjDlhCFInqZmKsgWYMTxuAUPJA5LieckS3YqQGlgP1yYBU+Yb24QASZMhxjU8H7jAvMUdVwsk2OF7nsGiBKPtsMo5AweDkjKQW5"
    "d05vXOIsOj49ChGbxgveZAE2xKOQWpQlZgJQDc1vqCy8/8OhxkKHFtQ9JYWVCewavWcMsxRdQkTYLfLYYNGxsLkYWGDLfGGz+UWv"
    "4jHH/dmkjBLm1GiBZaxBfEV1N35jM8W+nlPrVJV0g+rcYwcfpB2xBbtrnIkDUVRAJhbCRO0J2rqxuxj6EZUCY+2Cw91ihNvxLZ6Q"
    "9vHARjeqVEFVN1W1ucZ534LnfUBWphGe0CknqAvcQVmMva0uKsOd5RttX/gwephFAPzu3CyaHxc+chDZRV2qpdmG/Ig85xFTzXNg"
    "ZxqI6FWihAzApBQAt9eAx9SYyC3e4xGraOO/RwmWns/hFTMSUBG20Iz2JKv+htYNfQ7tDqacWn7WUvkzV8afA6aThqIhnYITlH6v"
    "x4UiAXTxYK7wDr5M5iF+PAI2G6g0BrxH/FnX+0hRVHV5+ChpCLPQUzpL4ywnvVTo28E/gejqPlsdMeAtkQy/OlxKNMLdFBFcdFRk"
    "9nEuxCoUCqZMNm8QBhGiaT2RLk6bhPkxtDCV/dwGFN1nblcCd2gEn23daw8FlA4Q23YQ4r+EFGapuFTNlOVVdNNhkF6qlMV7V+dZ"
    "kDeYn37GbQ3xRLDsKbSI2f8AcmYz/oyN788eXVKuQl8ZxehSkdR8aCa3zvhLqDp95dnu7EGphhznIB2F7t0RjUzCJKGfvfn9f+U6"
    "7gmBWInuI+aDINTLWQizdPyQpJlE1fZsZbbhwXCf2Jgl0DnKLA3vhgWz5M8jx7vPY0Yt/f0PKu6X8MqLjs6hCrEKH1rUS7BKPSox"
    "D/mis9cvPDAA2VTThtwv2GDKZFseOZG7WEk8AeUUYo+W9AUbQ4bQmBDSxECnVjJqAxi7718w+2zbymYQLWieGXzBysOAmAVdAUWC"
    "FR75EkQpVr5USUKiBRnLn6qYgOUiL4qtW4jYozahO2AhTiE6fWgE+QQxV1O0rCq5XgqV84eUP6R/aVYXlWczmnAiJPrQKOIpk7u9"
    "E7X3cG6rTQOdblKgTemiuHRn7YvurIG6cV951mA2dwN1mtRrJfEFKfMU6K7Rl3Qwf1HFkmZDzrEKBQKilXCk+wWXzQDogIxRykyz"
    "X242PlhXYoRSQu/1JWFNTisUEFrPWg+KIjLKIdmofD7HKtBrLIl60iSkYtqD0BZa3sciocuFIN3sCuRSqy2nL5pWfVFPCoKfk/Qm"
    "2ihMbeFCCgTXxaZJ1rF/sbqPhoKd9S+4Pk1MUiRXMW28lPZFb1yJaJoOwDwQL9zaNmmSR+/Yc/3SxQNBqGhKKC7A48gYcMzJFKjG"
    "oeWeQo/+pR6FfOkKdta+jGEf2hG5O+47mxt9GREDjZBVtkDGzOncFxtG7cEA8gX3fC8wsrCvCPGqPAfxWgsk4CRgr/T7v+XRbbBf"
    "tCfhwWbDbWjxSELoRq7C7P82nHBmfmtV7DbQNB8A1DoCpIldMRB3QiH1Bmxyt+wInaSk912RofWtQ+UPIUrqnQo96OhuLCfEvrvQ"
    "G9ydITGynLHeWoWMTjCbFODE1dRteI+t0Ah7T0JkTtwNBbS/WuUwMlfwJ+IlM2AiQjcdqZPfgXKkVn3L27232o6+1Uzb6ECFjHKn"
    "FwmqFCnB1+6VgKZv7FMdk9RQbtH9CDMpfuhji+/Wq484Oidq9DvM/SE4M3z2Xbe6w3Cra7G3Nm4moAWRdNFNXMQTaE9NZAoQezkf"
    "kjQxC7jFzpYcnOgAA9ttOtsMH8lKuNZrWI5UlL5BuD3ujt8Ws1+UD0W1EVumwijqSTDBKv4Wm5W35QxDdgYkRBrHlpziSepjFaiF"
    "Cya3ZeSYe1u4m3RbJgu2T+Ov21FRsHWnPRbUeGgA5utokRmZD+PJiwl7cwbMy5HRGa+/bm3VcDtZTDLqmY3ktxNLbyp4IFjbv8Ou"
    "1J1p35mHd3uul+54InynNz3ueHv+7nCYDDJ7OXH5dWcV8Q5WuC91hyjdtVEyNrXuWAPvtM9zR8/4Gtod9iru2hE1Bm5OeNhD3fmS"
    "/c4cw2LGNAzYgRgrvLxp8G5wgOYhEl5trKXQAfGKw52t5JX7d1ij3DFkbJTfwbOOL7HedWgqd5x03/Hs+Q7t/o53de5Q5Q1Yx+64"
    "iCZaOFxL39lSxJ6wA+J2hmjmUIUIgWOScbLOitSDRmpG2pUJD6JIBQRbp0DPi8LopOpiAPUI3iZAd1iqAiBifYQs1sryDp3VnVaS"
    "Rh0U9ryzI7h7Z+a8v0+/7oP9dpyQ3POVIDKbOXmoXMBHgpa6kmT/mITSrZp/jx820aEau1F5Bf49hmqOrRYRJvyzNGtO+UP6l2bV"
    "7Z0HRddWyE5uN3349S+JnjKhe2ihE4YCcjo5iujPEWB1vIclrknvsd0OsFIBCTtgJAipzTieInL8lAg0xL7QvS4Gg5RNJxlCzIFA"
    "ZYBjxj5HGUYpohStsFJwlh131/fSHoSNyFWDmw2yy+ASYUed5OgqxU35idkTUQr5xGjJloc4e+5+jeivjSd6+VHoc5mjXzcsRFd0"
    "inDnuhcG3Z9YcfqTPFBGYpfLaPQIXRKVCHGCpekkZFwmRlTVa5Lzdzzy9939fX9PQinMq92ve5aGwX/MtRGK1Sij/3hIHAwpP4VR"
    "UiBGIQspeilFL5FYiyTWMoleElFFET3zjfUu4r3WOPc6bRdVHz68qG7dseVhtGUUIGIZJYEZQEk6dKN1FhkEGwyO9VNxZ5WMqoWx"
    "JqyGzYXfPc4W7qOnJKc3lAm4CvKHNSkqORZ5N6wezz7KLyUpz559+Fa9y4wwSoZzxPt5WQbpBMOOMjYdGwnUTIRdILFKRm3d3tdF"
    "mglFPpaTVHLZKaqdvKh+qJhUdUFyyezsuew1ltW+J0ShG/VJ6Aq3mGSJfYnvCRuPvTCI2AS8OnDmdR+vrJVXRtQ8TmQt7+7TCQ+b"
    "RsLG+j2aAZs8l3f3Nh4U2CusyoVVuaheFrb8orpYVBeL17tiTTvlyjD3QadgBn2P63n3PLO+xyW9+wIhyj9EtqhfK6pYRTNc4/7c"
    "gJi6orpRUFVLo+g1HpnGY9NkQhByt9qlIsm9aryimZBE7mZWMpDcu5LNpiAN1Gi55Xiv+4X3uldoJLMkBROvsYyOGQBWN8Ce6PGb"
    "o0eHagrF2xjmf/c4HL0vipoimqsC0SwtvOv29IyL2XtulNzjXXrAmzCIVDwUZKyo2tK8cao2dMkEkjKrp4VmpCnbU1EG9e7FOVJv"
    "dod49SPF3/8YXhhFbonfF0R9ooAeAyhf0IXjqAAQiLSPTQ+DRJCIYW7cs2xG1KoROTbi5AmYRT0wEYIwO1VldK76s4GbOFYLcdYf"
    "nILzhyALhciheeRVyfuRTlGmI0vQR7AxH2EXO1hGPbV4LH/vm/X3/gru/chAfv/T/f6Hbfj3f9PRMtgP99gMEyECoYYS05Clvwa1"
    "9NegloGaegl5yYW7IWxzbbQMsnTCI6unRMBpzNLfZ1ry3aWlv7u01MtKS90kW/IVvCUmJw1VSdwJ4Dl2XEVSOdIC49MlQiSeqM8k"
    "cHq/1HXppa5JL/2W9BLXE5a8lbf0awlLv5aw5KtGS711sOSrRku9JrTk+0BLvf+zrK/5LP39niVv7S11TW/Jl3mWeplnyb2EZXCg"
    "F3TQ1zzomQ6lvE8EOOs/LEhjEKiEhqhgB02Gl3oLhzaa4EnS+zhhltxUweOuAnG2HJVx9ZWd5fzKztJf2VnWV3aWemVnyVd2lnpl"
    "Z1lf2ZHgztrq3mPb1DBmj6t/3eQa8ndSePQ+MTIKw2Pm3nWqU0ONsFuCF3xoakgisMynVjoe36nW8Mnj/T4SrMUteZK1jJGVP+Kg"
    "e8mbhMQMYsHFRMtsRRHL16U+m2NkXRARCjQe9tpEp4kkkyQ9aU1U9LLQJyH1sDhdRv8myhKJUJww5hBbJ3EnXR5qLCMn/0vOK5Z+"
    "Xc/4B/zP8qXayMmZgQyEwsCmIERfAyFKYEBsApg/LbErBFedFGwOnPIYsvNaeu7xOzrLOIw4jlxirmhAb5jSJux5N8QE3atcwjr7"
    "5yVf31jyhQ0gr3mbQH8T08G9x6XNauzBt2RELEpKLAiXZr00C7NhNRMzvBOtsexZA5PyGpsGy3S2hyrcJV7igp6Bx6/j/XOxq2FH"
    "sR7wyHyg9lQYAtceINPjmySGNmQs82guypvFpfwwe7y2v8TltaXfoF9iF2BZb9AvuSG9xM4HAHOhpW7NL3Vrfum35sVZzKbi7/ss"
    "/Tr9Utfpl/V9nqXfq1/6vfplvVe/rPfqKbTOBTP15cf9+qXu14vgelJsJo/N5IFO7vlUPZ/czwnuTGO0ssYl/KUu4S+vWJUvrxFt"
    "iD0ADreW6gUmlvB0wuOlrOyeMLFaTsjmCePxiqOnYcMWvMIwuuIwatixwq0wnq7CHufjRpyJrbB/swpNOUGGE3QSKy4gVuEHZzCr"
    "cIzuwVH9BgTWJQqTuBqwtZjgWPWrQ8wXVoGBYEtihXv8WC9QGED0D1Fk2Dj4Xeml+FVgrDjBXPnqdOVr0FVgJuREoFohZVdlpJ0X"
    "mlY2xef8F4IsdYJIVNpxX9ewKbLJIXrlQ/RK7zGtMKzKdHbVIZiuB4zQ8GFyhRFUdvvmX1pjrlzNrrnyrFGNJpJS3zclO7lx9aKo"
    "QPCavLhTzg/hg1k2gxfNIBXvP6w0hoHckN/tIcs7Hv6Aq7lrT9LmmCeuGtlZDkYqx86xZTaPHQt87KRAaya3Ile5y6gqMXKjlDwM"
    "oQrSSUK3l5yUQWOX3N0Efy+BuhevTZeo9PKWGik7VaVsDVIqwIvSeEnulkXDG9ykvXMUVTvVmPpX1farV5JrkPE11uReqzCx7Ux0"
    "9Y5T7JVJMeBpAJlAtepZ1DfwVlwArLQVs9JWzMovXa509XhVv+my4sRihZsDKxRQdLfVcXUdpUKTiHSQ6CK9AQ6NMDjRB3YvQLTU"
    "mOSSNSN6FxFbQUPMexD9xmrfENfsVqgSCpwfZ1npIv8qKnpq85wZrKInkJc0VxH5UBgCuzxc6AdIlFdF33lxgQ6YWCxpSUmGyS0l"
    "GSplvIK68huoK7+ACuYAY8KZn2VaRTdnrCni9uTKVnsrZafyomdYPVtJ5Dc0QElIvV6E05YVa2VUSXp7jt6QjQepBzrA0L2y+YvB"
    "NSHQn9R+p6+mC7vJFk8rzmNWXGWutKRcJb4jA5qAsRvIHVtSEtA9SgkfOdN706t0oG5LSASJqLOp9ZBamcuC22gVlbZa/bDLiGBK"
    "ukqEU2iskBJs0UbelZYKRivve/Si4OgsIqo1Y3pk0ADoMkc6ZTnj8oCBu74EhsXj1FVS/WCVrpeAIagRpr5qRIKCYtFoSMEhiTRR"
    "iOnChL/joVD9e688E928fwiIJq9VrbAPtioHIosF95JWvNG0stXiiu2g8PMTK7YGNAa2BfVZmGKssHezUmelXZsVd21WrL/Fxz1/"
    "sXPFyo4XGTJJiiiFLWPIUtE/9eCgd1e+F/Ig5TAzrdF9CsJMf5IMPPAkX1XYuna24vsPQPklG3O05zDkub4Gt9JNtZW/+bnC/HSl"
    "Vx5WZcSj0Ec1hzJy8w8sz0ZZZLgj64OuSRsx08ZuH2mmUEYfc8olEuTZxT270A73QlfcaVqxbC0ampwpLqM6Zo/R+Ap5J+iBiRCE"
    "uXIhR2F2ylK3stS6qipd7V4wo0ccvwJHBtVFD8jyVoLlHre8gBoqMTXyXPNMG3HlZ8U8U5Ypx8Yj43+ESNstEk4vztA7s1iZraP2"
    "SMAstZFXGVY1d0drSxOiO7Fv4iiKifaK59CriQ/Nsyxoyjl5lzF5l6ErZiu/YQZm9CfvPLDbAaBWlkY1cffqLaaeRfUQ7Iej1Ac/"
    "nDUm7vA0roUyesApJ8AaiZSTwV5v0VJIVbC0GO9pa+/XMSjFD6kKs44cHaiSSx76PeB7ViT2AQ8BRhF9ykN445Xth3DAthtpImUi"
    "bvcZywbjOQsfWvDtIJMDQz1IK/P280Noghj+S6B/idmC9esDv2BgiK4URF9+2CBMpOoY6ai1oeAkJike8r9NR6EUWQ5QEUliGbb0"
    "s+UtLWMPQ28vPPD49EEHpQ/MCOrKP2w1PuizSA/cW3zgUuUBx40tKVaC9WLPmQllOrs5T02qgjzudgS6BgVn5S8/pZnDLH3oFZeq"
    "u9lhW+201qYluV133cx2klAuk/Rkt8ip+4DU6u3+B26UGh4INVEUwizFWfow/ZfxbJ5mzfQvveKS3Erf436INcBYjVmSuGtJauLs"
    "qdKDb9SB5KkXuAnSrb7NMR0/3FOU7XGWGFgjG8w9ZR7OI0gnUXZyw+zFypHPedaILjCGSf0JeJCXO1msNSPt8I2Ah+AJSW69mtZ6"
    "kGoZp5rzqea72kmHrSUQhl+yjY5kpylXluAot151vT4m+VRkpRfKQl8jpiaOS3CKcO8WPHLDu/ii9NTWD4FZfiTUtBzd2VH1U0G0"
    "nt3ceiVXcm1FtlVKaltS22gVd090O9FbuctJSNt/f1QKvxlCSWnWhoQLkhR5T7uRWpDMVL6DzIYkzE6uqTgOk5CGFxp5Pl1qeV/k"
    "8qKY/ATQ0RSEfFGC0kmkUp1qRZxqRZzqUAZJUZx0IkrJ0zXhU8sSmH8TF3sUVJ0mz2Z8gYyk2jX5S2cUp1xZAj+//RB4ee4haABT"
    "9ths4EpKGrQowO5QNaoayWdTtspeSO50YAGOe4JE6bB/1rbFg+87iFmrR6/jo9dUfGv3ISibkcs/EyEST0Gk4jOhkOjBz6u6T5QK"
    "C4UZXRk2JsKJyNWscWHpTowjZw7zFGJ3TBlLyoeoQduoIdfuneMmCiMe8NDs4NTMBLYkoDbVyUbk+l4UnF0bweEF0QfesXvg2h7o"
    "iuzkyt5pZrc9kFp5Wq3N9nhmT05OUubKjMSFdlnrow/fUbeuHyJH52jjHWbqLrxToH7k2e0DjhEA++iqWZsVM3oXha17YRYzyM7t"
    "MNxOPrHHArEJUGjJDHWQqWLcXeR8ggVkh3mAzGFnb8BO07gBNkGIMhLPGlGCG7gyujpWcm2Vj0tVr9pxSxwUcJYDQO1KuIZHhAVO"
    "MbUr/JAOuuNpQkZTTQfOgEgK6ShQfI+R0BLhMf09EmmhtYlRqdxTgK+saH4X94FvDBjyg03GKiLe0KUavVweJEUSQmrp7cBInhAe"
    "BMaQ9YVf23rgPoJhZAeLfYQH7CM88ENED9pBeODVxQe8u4l0oFUmn8bZRDYpE/XGIngicU6QNCdIrDCJA0BS16+BkWckDzzyeNBB"
    "xwMPLB74rx9EBjnUMhtws4HEElP3ZqNoEkLBs4YHa0446hBDjTcH9mL8E4kE6jBB2GGwPNYd5QdeTn7AtsIDPqgAZB0t/JjOg3+M"
    "4UGbC0ZdIfX4vNlA0RG+0Hub/wLZERXNm/QBgwdeaCahpDgWFMWlbyDzrrsoOJv2ubFEjZbssQA0IRoteAsRP1VWHb+CeuDv/0cV"
    "9kge2LlOKqapawkwZFHxG66GA+WBg8OktjzxAAhE6zjHWQf77QkIbs0379f6FPWaq5+1n7WsecSy9pOVtU5W1haFteXLGhOZNQ5P"
    "omgiJaH0qMUZ3bpew1vruGQdsB0EhINzwYv0Jtj8vof2EHaNuE2Vo3iSum8qy0JTbTSMGwh9kSR8kskl6XmChnThm5hrP2xYo1da"
    "xx2uOKyttq0xPK35BZM1z77X3H42hLWWIiOo0WWNsl+z5Ncs97VOudcaR9Y82DbEpGGNCZ9BbIDMWO4XAy1LYuchqYNfYwGxRtvo"
    "jGiZjWmNo2IA7erNmnW8om6uDwEPhLIDWKtbH0abHP1aI8pWlSzo2q1CsDmvKPQuVH0JR0ISSqfaPbrd4wSUR5gViPrKrhGdWB8g"
    "UQdJ4K7uWn33Gq/GwekBZZYQOgNnnUyqi+yaDeUnA2SpJKkZYCa8AeUtHOIrSuuCjRBD9E7rAhEhFRWp7tutdd/OCKVSGhEN0POv"
    "WWToedboONb8SI0hX2Zwht+cuK21G2mUGCAnbUb0gm+aiPZgdDbrgtL1/br1aD+8wb3maxBmym2mNbaZHvEPWKZ4RMN73A18Q1dc"
    "fj1ahjzqpd5Hro4NUace9RXEx32SkQX4+CZrx1+P8BXz6Uf2SY9829eQHuMVm8c2nc5UpEzYU74Io7AjyRZvID+yITy2k9ye+Jj2"
    "KcJLvF3yaHn5qE8VPKKjeGTbeDxbw3w8N4pexwctC0QdxFSfsXnku0+W7O5MFXbQHzs1jcd+h9B4ovHYX/hS4mN/FQ/ynH4MyhP0"
    "nY+DleDjyAjx2y+PeBPskf38o7r5Rwxlj++O06+nYD8L9im82qj+FGw+8hQSEZp4RQdoqQVhQHuyTvVJl8ieAj/l/OSf1SZLnUQj"
    "SS/9GfcIr1CvMDTUnCcsQ+E9PqyT6U2n45qn0OPlA8QFQ/dTGPAknvw9BX4Ei7Sr7Nq72eAozlQzdOuKhLRLX8eWEIioQWSsy1yQ"
    "jnzRmwlkWnj/9YT3aoXRKbuQKlcNCy7i9PHJlhD2HGgB7+882ZrhCTex8a7SE7adJ1ISSg9aZwsUucW5t2FLdQcZ8eRXqp74BS0g"
    "42ys8ziTmMwGf0PxZCs+Qjo7naVZNfAuFVgv1D81Ra4wLBsW0qTvBFDoncRUw4iFgLeu96Cui6ABeYD3xgzhUeKr5E/Yg3tKeBvg"
    "qZgrzpWeSveWsCFiAl5XI1u9fyr8uDYIwXR43d3WqRAwpDyhGrAyWiagYlhV4za+C5SKDal6D+kJ7+XSHpavT7//T/vrv7jb9F9j"
    "5BvOxji0NHbthCkDiHrVjHeu/stGrvzrOWiJQRbFypWic1VLg7Ph5wCZJ83PQZuXz+FNHr5xCfAcDrKIUgZadhj1UOiWnPEb4AzR"
    "Gm8BlSwMIlTj55AFCAO9I7F1EkeZRVcmYpHXmOk/8z2X53D2RHsienxFG8SX8iEk8gjDC7MJbic4nbgKefarfcaK68RvPzlLn59/"
    "eubLIM+RD2YPzzqQf9bx/XN8JbzGXWUJsVLlo/ho4+EzY2lxBdCvtxHntGILMyqggwz1nZNntWSjJN3EL6hRcG3aStDDeGpIi7RH"
    "JQoBb+Ub2GraVmsmsIiiQyTKdg6DNAfpDtKcgIw+Q81RouRBThSQcolDy3OkjjJMlsZXFAtma8/xpz3w12b9QLqbMk86XLCw8V00"
    "g1eitcHnBmszIhayz02B9Ybp0Q2d52Yy/5Ap1nnjeoG4gBsC1p/PdcYHIZGOBESJ07tnTe+eOZt7xldNntOhaCbzzAnXM28UPmMn"
    "6ZmZzvbBbMDaj46R0nRBLSiveF6TMIoyaRJSwULRi6iiqixgOaARre8QraJWX7y1l5qygl0xQ8S/7PHsKenDFybgtUkjDAhG/DNO"
    "42raWXKND/yXJjAsc+vw2V9rEEdnWEKfUFq3hHwpekUGnIXU5JgOhiPVRpSmC6zhKkwcij/zQAwomZmHIeu52JrTpoXIjw5K5L8B"
    "dr3IqKDg5OpU1blqMKK9H0S55JYursGM7v3f5SgVd+W+5OxUlfKgWu+YfXrp03hXaCxwt7E6jtV19Lh7RvWxxouzILGrFT//SzuX"
    "mBnR/fSkRRWRPsj/7K9qPOP1DKJkOVGMk+x5AD3B0+zWEcxVwfwMlaIoO7mutOnsJ5zxwTjAo/bn8u5WNJV8VkcxormMbC0j6/6I"
    "tjLiv4UMf/8DnT0r3ejjBd7ueB6jXPMFFqPR4o73KZ7HM9vViMQYoNWQyANcDg0hCJEIcaQgRHTwkcNnbsw/TyO6kF+bwD9WJUVQ"
    "R0UnTazjgeiiN3jbdMPDX8MJeAC0eCKtt3LFW+EbnYSCqMBLZkYD4IIXjCUMuBKysQm0PSEPDO8UOpErmbebwLn2RmenG/7rjOHI"
    "8HnpaxMYuj6XtJkxRlISStHJCD5kDrEbXA810I1WF1CTNroAt8H/UwhlblOsQ5HgGlLZQLVR9LGYBCZiBlKU/U6e90zhMHZZ2iMe"
    "JGX0SOozJM7m4Y7/SFEZQhE4IoE7fndzEy1belAGYNtiYyOS7jqbpIggd/BZIEMVoiZWG22wbvid340MMreXNvE88O6+SV14k61O"
    "NnCSCrSGuInuvxKKDXADbiVu4ogI2ORlNE+QAU14fSXtARFAxw23sEid80BuHVnGJiBtJKlPNnMhJzFN5WGW926R1cLv24Jl5q5x"
    "XWvj/1uyqf9bAkFusLgiu+1kKzMJimPOlRR8ViQruaXB44N/P9no3082/u8nm/rvJ5v53082+sNFo79HT5BOBDf8WxQive+UPlyx"
    "IskXLLRBSagEoP8lKR06m9j4/6ds9L8pm/q/KZv6dymb+a9MNo2/RUShkofvQVX9SdGQ5V7+j1noir1zVVdLypiLPLrUAHQ91oSr"
    "Z8e1Fu21FsA1C13hplk5c50tKRwWwqQUT9kJNiPkyFrKabBRais1YKnaV6eJtGd8ORyC3A77x4avzYiiWLqD3Hgljl6JPdBcVSqd"
    "qPjHTshERS/uyMyPXs44fWLdiCrw6AUeeYa0qf8Bs/H/gNn4LQewF7jeMNnoT2E2/qcw4JP7q3h3c/Aefg0P/xvg7EJ1MDlJmxeg"
    "Nvw7mY3+TmbjfyOzaZKyOTkW5HyCL+ogEl+MAbdCZkiqdTfVpCRPSuqU90lJSu9jFKGbahi9wgz2+la8pyjZdZlD2H2m2nrvhpWH"
    "XWx6RdRtiWMPJqsb/SPLJuEOk2FHI2ywGl6CMItRXzBDN0AsEiebpCgiU4/DmDVEdJ62rMI0yfj3P1SXV/Z+qYw/MfnZJPXUyWKa"
    "ftLm7/9t5kcM5xYTy8422VjcToyBZULZFSsn4zdqMONLiwdflwRzI8O4MKb8yIyhCBoIz/MNWyQb3o82/EG0+fCGrw5sCkIpRwCc"
    "8UPwJITS4UB7cw7dEb3gGRsT5v/54w9tNmf8RemvzRBwpEGaSEkYRDKiSSZo9DehlX1uVxh3hJbOOip6+oEPjDqlylWfvkOYyLMO"
    "4zfEc1WfG3WclHLlWSPNQpil2XCSEHPlajT76DoXR1q8SO+KC0+bQRlVOsyghpF6Vh90N3TDOwubkQ8miptRVzM24+FgNXW0Ahpb"
    "Thj0lUAjljV67gl0ACAR/ANDw15xG/H9KWIQuWpWohsde48Fa+vYw4N3unuXg3d38O72dK/ZOBEomqNLRK/A1F/iwGs7myu+6LyZ"
    "XtENTTg82ExqlxPaHUn35CEh7yd+fnjjf3yymax7tSazxe+Vc9Mtb12Tk3hvT8KToeKRxdZfk9qGI26dgPAFp60fzG25zgNWlXw6"
    "4aGA2c6WL0NtNTXe+tTY+FXKiL58q3kPKYnlMrolBXBKhJ3kSJSzxOidPCKnGlZ1L4fuxOqOIQ2YNjbwLefUW82gtv5i9Dbo7+C3"
    "rNBbTZi2/ra08Q+iPMGG3ZYvJm0VOCbRW85btpi2bPEmFVBuMZHYBsk8SiTRq54q/TuX8YjsHCapRgVwlUozhq1PGLacKQAnIRT4"
    "fiQIJnGPB8nkPGDL7a+t3lHZcrq81dC91chtdEaWRf41wlZuPGf8/dWtNrRJLS0Vng+IC/gNHe8Wn9Al07CLOJPb+htAYFf2rkqu"
    "xDhOQeaJsJOcRTLpHOWdPqq69QGdPGu4/SSUJucPWw7nW0+frVi3HiwX31tM8QH7aEuTrW7vGuUGaCnFPMtgT2RmNpHaniMSIoWW"
    "asWl8WQ3Nd1NTa/mJSCqEh7FsSknQuidWDaU0r8k+KqwPagRDyc9xuYPglM9053LLV/q3WLvbYtr81u+3rvV6L/V8eg2cZgweuMH"
    "iySkytEFGSSy9yI8It3q9snWxuptUjXjx04M+YcNWyYgefWSe8QdXwrbuib+qWxbXhMhSI5As1x2O6bBJmL2AGDthIc2T6x5RfUc"
    "R94Al9GJFHUThXECUlFohJhxOrX1/88AwxqSz8nBlhexgIypduS3nDFscTdrywnDtpQBcWE8URN5n3PLVzW2ekNji/21bZEv6iBA"
    "jMYFCm/7ZRJY5nScP5AmUhJKj1qZcADikt7W22Gn/qCrOAhR+VRxvEJ23hQ67CcD0bd0DI1HR1s1yg5Xg4C840BhoJ2+EdIO1lBb"
    "f9F5izfEtviKqjCSqIWm0o2sslZ7RzL0RuyDbnmtYKvu0HPEO8Mru5sr5rRb80g3XbaTQp9a9NaTlxLuuGwxRBLRYfBiCxF62H/4"
    "c19+/Wl+/dnye/OgYzSaelzC+dMsWKB/dj3i9id8/dOqx5/X+OsFVyReQouHC9YXbri/YOMJefaidwVIOHeqQpglGVrVEAVn155A"
    "rVTFdUXwC73eC/PlBWOakzPm2y8c2l609H9h1gHdarXbNwOC71OLqxAvYTw05jv6uhfeRnnhMvGF08UXdn4vvIT4ghr1Et0bP7l4"
    "0YeoXrQQfNEIYmSN+oVd4ws6nRd+nBeIEgHTVicfeX/9hSemL77Ke9E3A0hSZgJtZPekxqBjqv3f61/07/Uv/qf14EK3RY51q0vC"
    "sQq42WVSL4cDqadrgGUSL9sF5wjGUccLjzpedNTxwqMOQ1yvEaVQBYSNXasXZjnPP174lc0XGbFogdZcxe+hCgUC06KreS/8fDww"
    "Dk5VyajxQzWgJJXNM0WpSGCInc6+XuTxACOUccIU9AXXLSpRM8XshNJAhpbMl65e2IW+6FzhK7Zmv9rE8yuPnwxxf+Ar/i1vF8kT"
    "cJAmuiRQh6vKXwMU8OsrKrOTOJE6WOoy3fIDMV/DBfr82+av3B75Gi1obXZ8jXxY38BJGEncJvrKy7OGqBhfo3y25v9V769+1Wn1"
    "1yYNVDb6K/iv/LCz00AeO6cMRgp4svXV726CcaPTuUDAnALFQLFKLUFikENuhYqjmFmS/M8yv+rF06/+4qlz70IWu7WIv2P5qssD"
    "oLKPbqNcSYzxib6dlEJ0lV91m/Qrl/eGnqTsScqpCix/UNVgtqbcu3MWusWxtG8HjFpfWXG+lrLvEAF8BMWQTduIjkHGf+m7g3/x"
    "RPevZNOzv1LhouCvgr/7bH/9hZHkW+AXT78F/IXvNxx1fwtNkNYPZ9iysA2k5kcPv/lHD41HOr3++mbDxzdtRH+zivTNgv/m35f7"
    "Vr8m941l8S3q5fpvyX7wOmXrVb5hivHNz/m++TnfNz/nMy7Sz3jMDSYDUUR3Y3WIVey30X4Kwxa99iRCg0sDEI4kgLsaMx445PbM"
    "N3xmAAA7uAv3zWPMN/2/aefyO45YeEVV0kSO4iNBmhbAd+6/f0fP8F3rre/YNfjOweM7h43vHDa+I/+++379d2aaIWZ5xsi87wk3"
    "bL6j/n9HI/9ezBvu8HznDs/3QsgE6nYCeojR+bv2AIwMf/8v/SfK/wD4sI35")
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
    if not os.path.exists(IN_JSON):
        stop(f"{IN_JSON} not found. Put harrison_v7_CLEAN7.json in {DIR} (or set HARRISON_DIR).")
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
    return recs, ts


def flags_for(rec, text):
    out = []
    lang = text_of(rec.get("Language")).strip().lower()
    if lang and lang != "en":
        out.append(f"not in English ({lang})")
    if text_of(rec.get("Media Page")).strip():
        out.append("media page: " + text_of(rec.get("Media Page")).strip())
    if text.strip() and "\n" not in text.strip():
        out.append("one paragraph: the page lost its line breaks")
    if len(text.strip()) < 300:
        out.append("very short text")
    return out


def check_segments(text, res):
    pos = 0
    for a, b, k in res["segments"]:
        if a != pos or b < a or k not in ("prose", res["route"]):
            return f"segment ({a}, {b}, {k}) does not follow on"
        pos = b
    return "" if pos == len(text) else f"segments end at {pos}, text length {len(text)}"


def block_excerpt(text, segs, before=300, after=700):
    blk = [s for s in segs if s[2] != "prose"]
    if not blk:
        return text[:after]
    a = blk[0][0]
    return (("…" if a > before else "") + text[max(0, a - before):a] + " ⟦BLOCK STARTS⟧ " + text[a:a + after] +
            ("…" if a + after < len(text) else ""))


def short_evidence(ev):
    keep = ["labels", "label style", "headings", "title cue", "question-answer units", "run-in questions",
            "answers opening in the first person", "questions put to someone", "speech share",
            "labels naming people", "fp_per_k", "third_per_k", "quote_per_k"]
    return "; ".join(f"{k}: {ev[k]}" for k in keep if k in ev and ev[k] not in ("", None))


def draw_sample(rows):
    """54 records: 12 labelled, 12 unlabelled, 12 first-person (spread over sub-types), 8 random prose and
    10 borderline prose (first person outside quotation marks, questions, or label-like lines)."""
    rng = random.Random(SAMPLE_SEED)
    by = defaultdict(list)
    for r in rows:
        by[r["Route"]].append(r)
    picked = []

    def spread(items, k, why):
        """One record from each of the largest sub-types (up to half the places), the rest at random."""
        groups = defaultdict(list)
        for x in sorted(items, key=lambda x: x["Record ID"]):
            groups[x["Detail"]].append(x)
        order = sorted(groups, key=lambda g: (-len(groups[g]), g))
        out = []
        for g in order[:max(1, k // 2)]:
            out.append(groups[g].pop(rng.randrange(len(groups[g]))))
        rest = [x for g in order for x in groups[g]]
        out += rng.sample(rest, min(k - len(out), len(rest)))
        for x in out:
            picked.append((x, why))

    spread(by["labelled Q&A"], 12, "labelled Q&A")
    spread(by["unlabelled Q&A"], 12, "unlabelled Q&A")
    spread(by["first-person"], 12, "first-person")
    prose = sorted(by["prose"], key=lambda x: x["Record ID"])
    for x in rng.sample(prose, min(8, len(prose))):
        picked.append((x, "prose, random"))
    chosen = {x["Record ID"] for x, _ in picked}
    risky = [x for x in prose if x["Record ID"] not in chosen and x["_risky"]]
    for x in rng.sample(risky, min(10, len(risky))):
        picked.append((x, "prose, borderline"))
    rng.shuffle(picked)          # mixed order, so the routes do not come in blocks
    return picked


def readme_rows(counts, detail_counts, n_records, swaps_total, flags_total, dev_note):
    rows = [
        VERSION,
        "",
        "WHAT THIS FILE IS",
        f"The format (route) of each of the {n_records:,} turn-set records in harrison_v7_CLEAN7.json, worked out by "
        "rules from how speakers are marked (technical appendix, Step 2, stages 1 and 2). CLEAN7 is not changed. The full "
        "results, with the cleaned texts and the exact segments, are in harrison_v7_TURNS_stage1.json and "
        "harrison_v7_TURNS_stage2.json; this workbook is for reading and checking them.",
        "",
        "STAGE 1: THE TEXT WAS CLEANED BY ONE-FOR-ONE CHARACTER SWAPS",
        "Curly quotes became straight quotes, odd spaces became spaces, paired « » became \" \", inch marks after "
        "7, 10 and 12 became ″, and quotes typed as ``x'' or ''x'' became \"x\". Every record was checked: same "
        "length, and only these swaps. Swaps made, all records together: "
        + "; ".join(f"{k} {v:,}" for k, v in swaps_total.most_common()) + ".",
        "",
        "STAGE 2: ROUTES",
        "labelled Q&A: each turn starts with a speaker label.",
        "unlabelled Q&A: questions followed by first-person answers, with no labels.",
        "first-person: one person speaking throughout, or one person per section under a heading with their name.",
        "prose: everything else; the journalist's narration with quotations. Stage 4 works on these.",
        "forum thread: posts by several users; not processed further.",
        "",
        "Counts: " + "; ".join(f"{r} {counts.get(r, 0):,}" for r in ROUTES) + ".",
        dev_note,
        "",
        "SUB-TYPES (column Detail)",
    ]
    for d, n in detail_counts.most_common():
        rows.append(f"{n:>5}  {d}")
    rows += [
        "",
        "Labelled Q&A sub-types: 'labelled turns' (two or more speakers, each line labelled), 'many speakers' (oral "
        "histories, radio shows), 'answers labelled' (questions are plain paragraphs), 'questions labelled' (the "
        "interviewer's lines carry the label), 'one speaker labelled' (one person's labelled lines, as in NPR's In "
        "Your Ear), 'quotes labelled with the speaker's name' ('John Illsley [bass]: \"...\"'), 'radio transcript' "
        "(NPR pages naming each speaker once, 'KEN TUCKER, BYLINE:'), 'labels inside paragraphs' (pages that lost "
        "their line breaks), 'lecture transcript' (Red Bull Music Academy: the name stands before each turn with no "
        "colon), 'caption transcript' (video captions, a dash for each new speaker).",
        "First-person sub-types: 'one speaker throughout'; '(title says so; quoted answers count)' for pieces such as "
        "Soundtrack of my life, where the speaker's words are printed in quotation marks; '(title names the person)' "
        "for 'X on Y' pieces; 'sections headed by the speaker's name' (How we made, tributes, oral histories by "
        "section); 'one speaker, sections headed by topic' (a heading per song); 'interview highlights' (NPR pages "
        "giving the interviewee's answers without the questions).",
        "Unlabelled Q&A sub-types: 'questions on their own lines'; 'questions run into the answers'. The record is cut "
        "into question-and-answer units; journalist narration between units stays prose.",
        "",
        "SEGMENTS",
        "Each record's text is cut into segments that cover it once, in order: a prose introduction, the block (or, in "
        "unlabelled Q&As, several blocks of question-and-answer units), and a prose tail. Positions are character "
        "positions in the text, the same in the raw and the cleaned text. 'Block share' is the share of the text "
        "inside the block(s).",
        "",
        "FLAGS",
        "not in English; media page (from CLEAN7); one paragraph (the page lost its line breaks, so formats are "
        "harder to see); very short text (under 300 characters). Flags: "
        + ("; ".join(f"{k} {v}" for k, v in flags_total.most_common()) or "none") + ".",
        "",
        "THE ROUTING REVIEW (sheet 'Routing review'): an optional check by hand. The check of version 1, on 5 Oct, "
        "found the problems version 2 fixes.",
        "54 records drawn at random with a fixed seed: 12 labelled Q&As, 12 unlabelled Q&As and 12 first-person "
        "records spread over their sub-types, 8 prose records at random, and 10 prose records that look borderline "
        "(first person outside quotation marks, several questions, or label-like lines).",
        "For each record: read the text shown (the block start is marked ⟦BLOCK STARTS⟧; open the URL if you need "
        "more), choose the right route in 'Your route', choose in 'Block starts right?' whether the Q&A or "
        "first-person part starts at the mark (an introduction in the third person should come before it), and add "
        "a note if anything is odd. Judge the format, not who is speaking: a blogger writing in the first person is "
        "first-person. Leave nothing blank; 'not sure' is fine.",
        "",
        "SHEETS",
        "Summary: counts by route and sub-type, and by act.",
        "Routes: every turn-set record with its route, sub-type, flags, block share, the evidence, and the text around "
        "the start of the block.",
        "Routing review: the optional check by hand.",
    ]
    return pd.DataFrame({"A": rows})


def main():
    t0 = time.time()
    print(VERSION)
    recs, ts = load_input()
    print(f"Read {os.path.basename(IN_JSON)}: {len(recs):,} records, turn set {len(ts):,}")

    # STAGE 1
    stage1, problems, swaps_total = [], [], Counter()
    for r in ts:
        raw = text_of(r.get("Extracted Text"))
        norm, counts = normalise(raw)
        msg = check(raw, norm)
        if msg:
            problems.append(f"record {r['Record ID']}: {msg}")
        swaps_total.update(counts)
        stage1.append({"Record ID": int(r["Record ID"]), "Text Length": len(raw), "Swaps": counts,
                       "Normalised Text": norm})
    if problems:
        stop(f"Stage 1 changed something it should not have, in {len(problems)} records, e.g.\n" +
             "\n".join(problems[:5]) + "\nThis is a bug; please report it with this output.")
    print(f"Stage 1: {len(stage1):,} texts cleaned and checked ({sum(swaps_total.values()):,} swaps) "
          f"[{time.time() - t0:.0f}s]")

    # STAGE 2
    stage2, rows, problems, flags_total = [], [], [], Counter()
    for k, (r, s1) in enumerate(zip(ts, stage1), 1):
        text = s1["Normalised Text"]
        title = text_of(r.get("Title / Headline"))
        res = route(text, title, text_of(r.get("Page Type")).strip().lower(), text_of(r.get("Language")).strip() or "en")
        msg = check_segments(text, res)
        if msg:
            problems.append(f"record {r['Record ID']}: {msg}")
        ev = res["evidence"]
        fl = flags_for(r, text)
        flags_total.update(f.split(":")[0].split(" (")[0] for f in fl)
        stage2.append({"Record ID": int(r["Record ID"]), "Route": res["route"], "Detail": ev.get("detail", ""),
                       "Quote Convention": ev.get("quote convention", ""),
                       "Segments": [[a, b, kind] for a, b, kind in res["segments"]],
                       "Block Share": ev.get("block share", 0.0), "Flags": fl,
                       "Evidence": {x: y for x, y in ev.items() if x not in ("detail", "quote convention", "block share")}})
        risky = res["route"] == "prose" and ((ev.get("fp_per_k") or 0) >= 20 or text.count("?") >= 6 or
                                            sum(1 for p in text.split("\n") if LABEL_LINE.match(p)) >= 3)
        blk = [s for s in res["segments"] if s[2] != "prose"]
        rows.append({"Record ID": int(r["Record ID"]), "Focal Act": text_of(r.get("Focal Act")),
                     "Grammy Year": r.get("Grammy Year"), "Title / Headline": title,
                     "Outlet / Source Name": text_of(r.get("Outlet / Source Name")), "URL": text_of(r.get("URL")),
                     "Route": res["route"], "Detail": ev.get("detail", ""),
                     "Block Starts At": blk[0][0] if blk else "", "Block Share": ev.get("block share", 0.0),
                     "Blocks": len(blk), "Quote Convention": ev.get("quote convention", ""), "Flags": "; ".join(fl),
                     "Evidence": short_evidence(ev), "Text Length": len(text),
                     "Text Around Block Start": block_excerpt(text, res["segments"]), "_risky": risky})
        if k % 1000 == 0:
            print(f"  stage 2: {k:,} of {len(ts):,} [{time.time() - t0:.0f}s]")
    if problems:
        stop(f"Stage 2 produced segments that do not cover the text, in {len(problems)} records, e.g.\n" +
             "\n".join(problems[:5]) + "\nThis is a bug; please report it with this output.")
    counts = Counter(x["Route"] for x in stage2)
    detail_counts = Counter(f"{x['Route']}: {x['Detail']}" if x["Detail"] else x["Route"] for x in stage2)
    print(f"Stage 2: {len(stage2):,} records routed and checked [{time.time() - t0:.0f}s]")

    # summaries
    df = pd.DataFrame(rows)
    by_act = df.pivot_table(index="Focal Act", columns="Route", values="Record ID", aggfunc="count", fill_value=0)
    by_act = by_act.reindex(columns=[c for c in ROUTES if c in by_act.columns])
    by_act["Total"] = by_act.sum(axis=1)
    first_year = df.groupby("Focal Act")["Grammy Year"].min()
    by_act = by_act.loc[first_year.sort_values(kind="stable").index].reset_index()
    route_table = pd.DataFrame([{"Route": r, "Records": counts.get(r, 0), "When checked (5 Oct)": CHECK_COUNTS.get(r, "")}
                                for r in ROUTES])
    detail_table = pd.DataFrame([{"Route: sub-type": d, "Records": n} for d, n in detail_counts.most_common()])
    diffs = {r: counts.get(r, 0) - CHECK_COUNTS.get(r, 0) for r in ROUTES}
    if any(diffs.values()):
        dev_note = ("These counts differ from the ones the script gave on the same file when it was checked on 5 Oct (" +
                    "; ".join(f"{r} {d:+d}" for r, d in diffs.items() if d) + "). Please report this with the run's output.")
    else:
        dev_note = "These counts match the ones the script gave on the same file when it was checked on 5 Oct."
    sample = draw_sample(rows)

    if DRY_RUN:
        report(counts, detail_counts, flags_total, dev_note, wrote=False)
        return

    # write
    with open(OUT_STAGE1, "w", encoding="utf-8") as f:
        json.dump(stage1, f, ensure_ascii=False)
    with open(OUT_STAGE2, "w", encoding="utf-8") as f:
        json.dump(stage2, f, ensure_ascii=False)
    write_workbook(df, route_table, detail_table, by_act, sample, readme_rows(
        counts, detail_counts, len(stage2), swaps_total, flags_total, dev_note))
    # read back and confirm
    with open(OUT_STAGE1, encoding="utf-8") as f:
        b1 = json.load(f)
    with open(OUT_STAGE2, encoding="utf-8") as f:
        b2 = json.load(f)
    raw_by_id = {int(r["Record ID"]): text_of(r.get("Extracted Text")) for r in ts}
    if len(b1) != EXPECTED_TURN_SET or len(b2) != EXPECTED_TURN_SET or \
            any(len(x["Normalised Text"]) != len(raw_by_id[x["Record ID"]]) for x in b1) or \
            Counter(x["Route"] for x in b2) != counts:
        stop("The written files did not read back correctly; delete them and run again.")
    report(counts, detail_counts, flags_total, dev_note, wrote=True)
    print(f"[{time.time() - t0:.0f}s]")


def _safe(v, cap=EXCEL_CELL_CAP):
    if isinstance(v, str):
        return ILLEGAL_CHARACTERS_RE.sub("", v)[:cap]
    return v


def write_workbook(df, route_table, detail_table, by_act, sample, readme):
    cols = ["Record ID", "Focal Act", "Grammy Year", "Title / Headline", "Outlet / Source Name", "URL", "Route",
            "Detail", "Flags", "Block Starts At", "Block Share", "Blocks", "Quote Convention", "Text Length",
            "Evidence", "Text Around Block Start"]
    routes_sheet = df[cols].apply(lambda col: col.map(_safe))
    review = []
    for i, (x, why) in enumerate(sample, 1):
        review.append({"#": i, "Record ID": x["Record ID"], "Focal Act": x["Focal Act"],
                       "Title / Headline": _safe(x["Title / Headline"]), "Outlet / Source Name": x["Outlet / Source Name"],
                       "URL": x["URL"], "Sampled as": why, "Route given": x["Route"], "Detail": x["Detail"],
                       "Text around the block start (or the opening, for prose)": _safe(x["Text Around Block Start"]),
                       "Your route": "", "Block starts right?": "", "Your note": ""})
    review = pd.DataFrame(review)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        readme.apply(lambda col: col.map(_safe)).to_excel(xw, sheet_name="Read me", index=False, header=False)
        route_table.to_excel(xw, sheet_name="Summary", index=False, startrow=1)
        detail_table.to_excel(xw, sheet_name="Summary", index=False, startrow=len(route_table) + 4)
        by_act.to_excel(xw, sheet_name="Summary", index=False, startrow=len(route_table) + len(detail_table) + 7)
        routes_sheet.to_excel(xw, sheet_name="Routes", index=False)
        review.to_excel(xw, sheet_name="Routing review", index=False, startrow=3)
        wb = xw.book
        ws = wb["Read me"]
        ws.column_dimensions["A"].width = 120
        for c in ws["A"]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.value and str(c.value).isupper() and len(str(c.value)) < 60:
                c.font = Font(bold=True)
        ws = wb["Summary"]
        ws["A1"] = "Records by route"
        ws.cell(row=len(route_table) + 4, column=1, value="Records by route and sub-type")
        ws.cell(row=len(route_table) + len(detail_table) + 7, column=1, value="Records by act and route")
        for cell in (ws["A1"], ws.cell(row=len(route_table) + 4, column=1),
                     ws.cell(row=len(route_table) + len(detail_table) + 7, column=1)):
            cell.font = Font(bold=True)
        ws.column_dimensions["A"].width = 70
        ws = wb["Routes"]
        ws.freeze_panes = "B2"
        for col, w in zip("ABCDEFGHIJKLMNOP", [9, 22, 8, 50, 22, 30, 15, 40, 30, 9, 9, 7, 9, 9, 60, 90]):
            ws.column_dimensions[col].width = w
        ws.auto_filter.ref = ws.dimensions
        ws = wb["Routing review"]
        ws["A1"] = ("Optional (the check was done on 5 Oct on version 1). Check 54 records. For each, read the text (the block start is marked ⟦BLOCK STARTS⟧; open the URL "
                    "if you need more), pick the right route in 'Your route', say whether the Q&A or first-person part "
                    "starts at the mark in 'Block starts right?' (choose 'not applicable' for prose), and add a note if "
                    "anything is odd. Judge the format, not who is speaking.")
        ws["A2"] = ("Routes: labelled Q&A (each turn starts with a speaker label) · unlabelled Q&A (questions, then "
                    "first-person answers, no labels) · first-person (one person speaking throughout, or one per "
                    "section) · prose (third-person journalism with quotations, lists, news) · forum thread.")
        ws["A1"].font = Font(bold=True)
        ws["A1"].alignment = ws["A2"].alignment = Alignment(wrap_text=False)
        last = 4 + len(review)
        dv1 = DataValidation(type="list", formula1='"labelled Q&A,unlabelled Q&A,first-person,prose,forum thread,not sure"',
                             allow_blank=True)
        dv2 = DataValidation(type="list", formula1='"yes,no: starts too early,no: starts too late,not applicable,not sure"',
                             allow_blank=True)
        ws.add_data_validation(dv1)
        ws.add_data_validation(dv2)
        dv1.add(f"K5:K{last}")
        dv2.add(f"L5:L{last}")
        fill = PatternFill("solid", fgColor="FFF2CC")
        for row in ws.iter_rows(min_row=5, max_row=last):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            for c in row[10:13]:
                c.fill = fill
        for col, w in zip("ABCDEFGHIJKLM", [4, 9, 18, 34, 18, 26, 16, 14, 26, 90, 16, 18, 30]):
            ws.column_dimensions[col].width = w
        for c in ws[4]:
            c.font = Font(bold=True)
        ws.freeze_panes = "C5"


def report(counts, detail_counts, flags_total, dev_note, wrote):
    print(f"\n{'=' * 66}")
    for r in ROUTES:
        print(f"  {r:16} {counts.get(r, 0):>6,}   (when checked on 5 Oct: {CHECK_COUNTS.get(r, 0):,})")
    print("\n  Sub-types:")
    for d, n in detail_counts.most_common():
        print(f"      {n:>5}  {d}")
    print("\n  Flags: " + ("; ".join(f"{k} {v}" for k, v in flags_total.most_common()) or "none"))
    print("\n  " + dev_note)
    if wrote:
        print(f"\nWrote:\n  {OUT_STAGE1}\n  {OUT_STAGE2}\n  {OUT_XLSX}")
    else:
        print("\nDRY RUN: every check passed; nothing was written.")


if __name__ == "__main__":
    main()
