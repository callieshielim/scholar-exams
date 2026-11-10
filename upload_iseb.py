#!/usr/bin/env python3
"""
Upload ISEB questions to admin.11plusonline.co.uk
Usage: python upload_iseb.py [--dry-run] [--no-submit] [--start N] [--end N]

  --dry-run    Parse + preflight only, no browser.
  --no-submit  Fill the form for each question but DO NOT click Finish
               (screenshots each to /tmp; safe for testing form fills).
  --start/--end  Upload only questions whose number is in the range.

Set only TOPIC and DOCX_PATH below per paper.
"""

import asyncio
import argparse
import base64
import random
import re
import zipfile
import os
import sys

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import docx
from docx.text.paragraph import Paragraph
from docx.table import Table
from playwright.async_api import async_playwright

# ── User config — only change these two lines per paper ──────────────────────
# Valid TOPIC values: "Non Verbal Reasoning", "English", "Mathematics", "Verbal Reasoning"
TOPIC     = "English"   # "Non Verbal Reasoning" | "English" | "Mathematics" | "Verbal Reasoning"
DOCX_PATH = r"C:\Users\angel\Documents\Exceptional Academics\ISEB\ISEB 7\ENGLISH.docx"
print(DOCX_PATH)
print(os.path.exists(DOCX_PATH))

# ── Delay between uploads (seconds) — set both to 0 to disable ───────────────
DELAY_MIN = 60
DELAY_MAX = 100
# ─────────────────────────────────────────────────────────────────────────────

LOGIN_URL = "https://admin.11plusonline.co.uk/login"
ADD_URL   = "https://admin.11plusonline.co.uk/question-bank/question/add"
EMAIL     = "admin1@11plusonline.co.uk"
PASSWORD  = "password1"
IMG_DIR   = os.path.join(os.environ.get("TEMP", "/tmp"), "iseb_upload_images")
SHOT_DIR  = os.path.join(os.environ.get("TEMP", "/tmp"), "iseb_shots")

TEST_TYPE  = "ISEB"
DIFFICULTY = "easy"

# ── Editable prompt text for the "spot the error" questions ──────────────────
SPELLING_PROMPT    = "Which part of the sentence contains a spelling mistake?"
PUNCTUATION_PROMPT = "Which part of the sentence contains a punctuation error?"

# ── Layout-card selectors (substring of the card's description text) ─────────
# NVR    cards: "image, 5 image-based answer options" | "image, 5 text-based answer options"
# ENG/VR cards: "Question text, 5 text-based answer options"
# Maths  cards: "Question text, 5 text-based answer options" |
#               "Question text, image on the top, 5 text-based answer options" |
#               "Question text, image on the right side, 5 text-based answer options"
LAYOUT_NVR_TEXT_ANSWERS  = "text-based answer options"     # picks the text-answer NVR card
LAYOUT_5_TEXT            = "5 text-based answer options"   # plain 5-option text card (Eng/Maths/VR)
LAYOUT_3_TEXT            = "3 text-based answer options"   # plain 3-option text card (Eng/Maths/VR)
LAYOUT_MATHS_IMAGE_TOP   = "image on the top, 5 text-based answer options"

# ── ISEB map: TOPIC → paper-level config ─────────────────────────────────────
# has_topic_dropdown : whether the site shows a Topic dropdown for this subject
# content            : "image" (question is an image) or "text" (question typed) — the
#                       default/most-common value for the paper; individual questions may
#                       still override this per-question (e.g. Maths chart questions).
# expected_count     : sanity-check total
ISEB_MAP = {
    "Non Verbal Reasoning": {
        "subject": "Non Verbal Reasoning",
        "has_topic_dropdown": False,
        "content": "image",
        "expected_count": 35,
        "has_images": True,
    },
    "English": {
        "subject": "English",
        "has_topic_dropdown": True,   # topic set per-question (Comprehension/Spelling/Punctuation/Grammar)
        "content": "text",
        "expected_count": 42,
        "has_images": False,
    },
    "Mathematics": {
        "subject": "Maths",
        "has_topic_dropdown": False,
        "content": "text",           # most questions are text-only; some override to "image"
        "expected_count": 36,
        "has_images": True,          # a handful of questions include a chart/diagram image
    },
    "Verbal Reasoning": {
        "subject": "Verbal Reasoning",
        "has_topic_dropdown": False,
        "content": "text",
        "expected_count": 36,
        "has_images": False,
    },
}

if TOPIC not in ISEB_MAP:
    print(f"ERROR: Unknown/unsupported TOPIC '{TOPIC}'. Valid: {list(ISEB_MAP)}")
    sys.exit(1)

CFG = ISEB_MAP[TOPIC]
SUBJECT = CFG["subject"]

LETTERS = ["A", "B", "C", "D", "E"]


# ── Parsing helpers ──────────────────────────────────────────────────────────

def iter_block_items(doc):
    """Yield paragraphs and tables in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def parse_docx(path):
    doc = docx.Document(path)
    if TOPIC == "Non Verbal Reasoning":
        return parse_nvr(doc)
    if TOPIC == "English":
        return parse_english(doc)
    if TOPIC == "Mathematics":
        return parse_maths(doc)
    if TOPIC == "Verbal Reasoning":
        return parse_vr(doc)
    raise ValueError(TOPIC)


# ── NVR: image questions ──────────────────────────────────────────────────────
#
# Two formats seen in the wild:
#   (a) inline   : "N." then "Answer: X <explanation>" right after each question.
#   (b) table-key: "N." markers only (no inline answers), followed later by an
#                  "ANSWERS AND EXPLANATIONS" heading and one table per section,
#                  each row = [num, letter, explanation].
# Both are handled here; the table-key answers are merged in after the main pass.

def parse_nvr(doc):
    questions = []
    current_title = "Non Verbal Reasoning"
    current_instruction = ""     # the prompt text shown under each section header
    pending = None
    mode = "questions"            # "questions" | "answers"
    answer_key = {}                # num -> (letter, explanation), from answer tables

    def flush():
        if pending is not None:
            questions.append(pending)

    for block in iter_block_items(doc):
        if isinstance(block, Table):
            if mode == "answers":
                for row in block.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) < 2:
                        continue
                    num_s, letter = cells[0].strip(), cells[1].strip().upper()
                    expl = cells[2].strip() if len(cells) > 2 else ""
                    if num_s.isdigit() and letter in LETTERS:
                        answer_key[int(num_s)] = (letter, expl)
            continue

        t = block.text.strip()
        if not t:
            continue

        if t.upper() == "ANSWERS AND EXPLANATIONS":
            flush()
            pending = None
            mode = "answers"
            continue

        if mode == "answers":
            # section sub-headers repeated above each table — nothing to parse
            continue

        num_m = re.match(r"^(\d+)\.?$", t)
        ans_m = re.match(r"^Answer:\s*([A-E])\b\s*(.*)$", t, re.IGNORECASE | re.DOTALL)

        if num_m:
            flush()
            pending = {
                "num": int(num_m.group(1)),
                "title": current_title,
                "site_topic": None,
                "content": "image",
                "layout": LAYOUT_NVR_TEXT_ANSWERS,
                "question_text": current_instruction,   # goes into the "Question" editor
                "options": list(LETTERS),      # A/B/C/D/E labels (choices are in the image)
                "answer": None,
                "correct": [],
                "explanation": "",
                "needs_manual": False,
            }
        elif ans_m and pending:
            letter = ans_m.group(1).upper()
            pending["answer"] = letter
            pending["correct"] = [LETTERS.index(letter)]
            pending["explanation"] = ans_m.group(2).strip()
        else:
            # header (short) vs instruction (long sentence)
            if len(t) <= 40 and not t.endswith(('.', '?', ':')) and not t.lower().startswith("answer"):
                current_title = t
            elif len(t) > 40:
                current_instruction = t

    flush()

    # merge in table-key answers for questions that had no inline "Answer:" line
    for q in questions:
        if not q["answer"] and q["num"] in answer_key:
            letter, expl = answer_key[q["num"]]
            q["answer"] = letter
            q["correct"] = [LETTERS.index(letter)]
            q["explanation"] = expl

    questions = [q for q in questions if q["answer"]]
    questions.sort(key=lambda q: q["num"])
    return questions


# ── English: Comprehension + Spelling + Punctuation + Grammar ────────────────
#
# Structure in the docx:
#   COMPREHENSION 1 / COMPREHENSION 2 / NON-FICTION COMPREHENSION / POEM COMPREHENSION
#       <passage paragraphs>
#       N. stem\na) ...\nb) ...\nc) ...\nd) ...\ne) ...   (x5 per section)
#       ANSWERS AND EXPLANATIONS
#       N. x) answer text\nExplanation: ...                (x5 per section)
#   SPELLING
#       Qn. stem with __________ blank
#       a) ...\nb) ...\nc) ...\nd) ...\ne) ...
#       Answer: x) word
#   PUNCTUATION
#       n                                  (bare number)
#       [table: 5 sentence fragments | A B C D E labels]
#       Answer: X fragment-text
#   GRAMMAR
#       Qn. stem
#       a) ...   (each option its own paragraph, or packed on one line)
#       ...
#       Answer: x)  [optional trailing text/explanation]

def parse_english(doc):
    questions = []
    answer_key = {}          # comprehension: num -> (letter, explanation)
    mode = None               # "comp" | "spelling" | "punctuation" | "grammar"
    current_title = ""
    passage_lines = []
    in_answers = False
    cur = None                # current spelling/punctuation/grammar question being built

    num_re = re.compile(r"^(\d+)\.?$")
    qnum_re = re.compile(r"^Q?(\d+)\.\s*(.*)$")
    opt_re = re.compile(r"^([a-e])\)\s*(.*)$", re.IGNORECASE)
    ans_re = re.compile(r"^Answer:?\s*([a-e])\)?\s*(.*)$", re.IGNORECASE | re.DOTALL)
    section_comp_nums = []
    unnumbered_idx = 0
    next_comp_num = 1

    def flush_cur():
        nonlocal cur
        if cur:
            questions.append(cur)
            cur = None

    def new_question(num, site_topic, layout):
        return {
            "num": num,
            "title": current_title,
            "site_topic": site_topic,
            "content": "text",
            "layout": layout,
            "question_text": None,
            "options": [],
            "answer": None,
            "correct": [],
            "explanation": "",
            "needs_manual": False,
            "_opts": {},          # letter -> text, collected before finalising
        }

    for block in iter_block_items(doc):
        if isinstance(block, Table):
            if mode == "punctuation" and cur is not None:
                # first row = the 5 sentence fragments; second row = A-E labels
                frags = [c.text.strip() for c in block.rows[0].cells]
                cur["fragments"] = frags
            continue

        t = block.text.strip()
        if not t:
            continue
        upper = t.upper()

        # ── Section headers ────────────────────────────────────────────────
        if upper == "SPELLING":
            flush_cur()
            mode, current_title = "spelling", "SPELLING"
            continue
        if upper == "PUNCTUATION":
            flush_cur()
            mode, current_title = "punctuation", "PUNCTUATION"
            continue
        if upper == "GRAMMAR" or upper.startswith("GRAMMAR AND"):
            flush_cur()
            mode, current_title = "grammar", "GRAMMAR"
            continue
        first_line, _, rest_after_hdr = t.partition("\n")
        hdr = first_line.strip()
        hdr_upper = hdr.upper()
        is_comp_header = bool(
            re.fullmatch(r"(NON-FICTION\s+)?COMPREHENSION(\s+\d+)?", hdr_upper)
            or hdr_upper == "POEM COMPREHENSION"
        )
        if is_comp_header:
            flush_cur()
            mode, current_title = "comp", hdr
            passage_lines = []
            in_answers = False
            section_comp_nums = []
            unnumbered_idx = 0
            if not rest_after_hdr.strip():
                continue
            t = rest_after_hdr.strip()
            upper = t.upper()

        # ── Comprehension mode ─────────────────────────────────────────────
        if mode == "comp":
            if "ANSWERS AND EXPLANATIONS" in upper:
                in_answers = True
                t = re.sub(r"ANSWERS AND EXPLANATIONS", "", t, flags=re.IGNORECASE).strip()
                if not t:
                    continue
            if in_answers:
                unnumbered_idx = _parse_comp_answer_block(
                    t, answer_key, section_comp_nums, unnumbered_idx
                )
                continue
            qm = re.match(r"^(\d+)\.\s", t)
            lines = t.split("\n")
            has_options = any(opt_re.match(line.strip()) for line in lines[1:])
            if has_options:
                num = int(qm.group(1)) if qm else next_comp_num
                lines = t.split("\n")
                stem = re.sub(r"^\d+\.\s*", "", lines[0]).strip()
                opt_map = {}
                for line in lines[1:]:
                    om = opt_re.match(line.strip())
                    if om:
                        opt_map[om.group(1).lower()] = om.group(2).strip()
                option_list = [opt_map.get(l, "") for l in "abcde"]
                passage = "\n".join(passage_lines).strip()
                questions.append({
                    "num": num,
                    "title": current_title,
                    "site_topic": "Comprehension",
                    "content": "text",
                    "layout": LAYOUT_5_TEXT,
                    # Comprehension is the only layout with a separate "Passage Text"
                    # editor on the site (see upload_question) — keep passage and
                    # the question stem apart instead of concatenating them.
                    "passage": passage,
                    "question_text": stem,
                    "options": option_list,
                    "answer": None,
                    "correct": [],
                    "explanation": "",
                    "needs_manual": False,
                })
                section_comp_nums.append(num)
                next_comp_num = max(next_comp_num, num + 1)
            else:
                passage_lines.append(t)
            continue

        # ── Spelling: "Qn. stem" then options (all on one packed paragraph) ──
        if mode == "spelling":
            qm = qnum_re.match(t)
            ans_m = ans_re.match(t)
            if qm:
                flush_cur()
                cur = new_question(int(qm.group(1)), "Spelling", LAYOUT_5_TEXT)
                cur["question_text"] = qm.group(2).strip()
                continue
            if ans_m and cur is not None:
                letter = ans_m.group(1).upper()
                cur["answer"] = letter
                cur["explanation"] = ans_m.group(2).strip()
                continue
            if cur is not None:
                for part in t.split("\n"):
                    om = opt_re.match(part.strip())
                    if om:
                        cur["_opts"][om.group(1).upper()] = om.group(2).strip()
            continue

        # ── Punctuation: bare number, table of fragments, inline Answer ──────
        if mode == "punctuation":
            num_m = num_re.match(t)
            ans_m = ans_re.match(t)
            if num_m:
                flush_cur()
                cur = new_question(int(num_m.group(1)), "Punctuation", LAYOUT_5_TEXT)
                cur["fragments"] = []
                continue
            if ans_m and cur is not None:
                letter = ans_m.group(1).upper()
                cur["answer"] = letter
                cur["explanation"] = ans_m.group(2).strip()
                continue
            continue

        # ── Grammar: "Qn. stem" then options (own paragraph or packed) ───────
        if mode == "grammar":
            qm = qnum_re.match(t)
            ans_m = ans_re.match(t)
            if qm:
                flush_cur()
                cur = new_question(int(qm.group(1)), "Grammar", LAYOUT_5_TEXT)
                cur["question_text"] = qm.group(2).strip()
                continue
            if ans_m and cur is not None:
                letter = ans_m.group(1).upper()
                cur["answer"] = letter
                cur["explanation"] = ans_m.group(2).strip()
                continue
            if cur is not None:
                for part in t.split("\n"):
                    om = opt_re.match(part.strip())
                    if om:
                        cur["_opts"][om.group(1).upper()] = om.group(2).strip()
                    elif part.strip() and not cur["_opts"]:
                        # extra stem line before the options, e.g. the statement to respond to
                        cur["question_text"] = (cur["question_text"] + "\n" + part.strip()).strip()
            continue

    flush_cur()

    # finalise comprehension correct answers
    for q in questions:
        if q["site_topic"] == "Comprehension" and q["num"] in answer_key:
            letter, expl = answer_key[q["num"]]
            if letter in LETTERS:
                q["correct"] = [LETTERS.index(letter)]
                q["answer"] = letter
            q["explanation"] = expl

    # finalise spelling/grammar questions (options collected in "_opts")
    for q in questions:
        if q["site_topic"] in ("Spelling", "Grammar"):
            q["options"] = [q["_opts"].get(l, "") for l in LETTERS]
            letter = q.get("answer")
            if letter and letter in LETTERS:
                q["correct"] = [LETTERS.index(letter)]
            else:
                q["needs_manual"] = True
            q.pop("_opts", None)

    # finalise punctuation questions (build options from table fragments)
    for q in questions:
        if q["site_topic"] == "Punctuation":
            frags = q.get("fragments", [])
            sentence = " ".join(frags).strip()
            q["options"] = frags
            q["question_text"] = f"{PUNCTUATION_PROMPT}\n\n{sentence}" if sentence else PUNCTUATION_PROMPT
            letter = q.get("answer")
            if letter and letter in LETTERS and LETTERS.index(letter) < len(frags):
                q["correct"] = [LETTERS.index(letter)]
            else:
                q["needs_manual"] = True
            q.pop("_opts", None)

    questions.sort(key=lambda q: q["num"])
    return questions


def _extract_comp_explanation(rest):
    expl_m = re.search(r"Explanation:\s*(.+)", rest, re.IGNORECASE | re.DOTALL)
    if expl_m:
        return expl_m.group(1).strip()
    parts = rest.split("\n", 1)
    if len(parts) > 1:
        return parts[1].strip()
    return ""


def _parse_comp_answer_block(text, answer_key, section_nums, next_idx):
    numbered = re.match(r"^(\d+)\.\s*([a-e])\)\s*(.*)$", text, re.IGNORECASE | re.DOTALL)
    unnumbered = re.match(r"^([a-e])\)\s*(.*)$", text, re.IGNORECASE | re.DOTALL)

    if numbered:
        num = int(numbered.group(1))
        letter = numbered.group(2).upper()
        rest = numbered.group(3).strip()
    elif unnumbered:
        while next_idx < len(section_nums) and section_nums[next_idx] in answer_key:
            next_idx += 1
        if next_idx >= len(section_nums):
            return next_idx
        num = section_nums[next_idx]
        letter = unnumbered.group(1).upper()
        rest = unnumbered.group(2).strip()
        next_idx += 1
    else:
        return next_idx

    answer_key[num] = (letter, _extract_comp_explanation(rest))
    return next_idx


# ── Mathematics: numbered questions, some with a chart/diagram image ─────────
#
#   Q1. stem\na) ...\nb) ...\n...\ne) ...        (stem + all options in one line, or)
#   Q4                                            (bare "Qn"/"n" marker)
#   <image>
#   [optional caption line(s)]
#   stem line
#   a) ...\nb) ...\n...
#   ...
#   ANSWERS AND EXPLANATIONS – SET n
#   Q1. c) 3,047 – explanation text...

def parse_maths(doc):
    questions = []
    answer_key = {}   # num -> (letter, explanation)
    mode = "questions"

    qmark_re = re.compile(r"^Q(\d+)\.?\s*(.*)$", re.IGNORECASE)
    barenum_re = re.compile(r"^(\d+)$")
    opt_re = re.compile(r"^([a-e])\)\s*(.*)$", re.IGNORECASE)
    ans_re = re.compile(r"^Q(\d+)\.\s*([a-e])\)\s*(.*)$", re.IGNORECASE | re.DOTALL)

    pending = None   # {"num":.., "stem_lines":[...], "opts":{}, "has_image":False}

    def flush():
        nonlocal pending
        if pending is None:
            return
        stem = "\n".join(pending["stem_lines"]).strip()
        options = [pending["opts"].get(l, "") for l in LETTERS]
        questions.append({
            "num": pending["num"],
            "title": f"Mathematics",
            "site_topic": None,
            "content": "image" if pending["has_image"] else "text",
            "layout": LAYOUT_MATHS_IMAGE_TOP if pending["has_image"] else LAYOUT_5_TEXT,
            "question_text": stem,
            "options": options,
            "answer": None,
            "correct": [],
            "explanation": "",
            "needs_manual": False,
        })
        pending = None

    for para in doc.paragraphs:
        if mode == "questions" and "graphic" in para._p.xml:
            if pending is not None:
                pending["has_image"] = True

        t = para.text.strip()
        if not t:
            continue

        if mode == "questions" and "ANSWERS AND EXPLANATIONS" in t.upper():
            flush()
            mode = "answers"
            continue

        if mode == "questions":
            qm = qmark_re.match(t)
            bn = barenum_re.match(t)
            if qm:
                flush()
                pending = {"num": int(qm.group(1)), "stem_lines": [], "opts": {}, "has_image": False}
                rest = qm.group(2).strip()
                if rest:
                    pending["stem_lines"].append(rest)
                continue
            if bn:
                flush()
                pending = {"num": int(bn.group(1)), "stem_lines": [], "opts": {}, "has_image": False}
                continue
            if pending is not None:
                matched_any_opt = False
                for part in t.split("\n"):
                    om = opt_re.match(part.strip())
                    if om:
                        pending["opts"][om.group(1).upper()] = om.group(2).strip()
                        matched_any_opt = True
                if not matched_any_opt:
                    pending["stem_lines"].append(t)
            continue

        if mode == "answers":
            am = ans_re.match(t)
            if am:
                num = int(am.group(1))
                letter = am.group(2).upper()
                expl = am.group(3).strip()
                answer_key[num] = (letter, expl)
            continue

    flush()

    for q in questions:
        if q["num"] in answer_key:
            letter, expl = answer_key[q["num"]]
            if letter in LETTERS:
                q["answer"] = letter
                q["correct"] = [LETTERS.index(letter)]
            q["explanation"] = expl
        else:
            q["needs_manual"] = True

    questions.sort(key=lambda q: q["num"])
    return questions


# ── Verbal Reasoning: five distinct question styles in one file ─────────────
#
#   ODD TWO OUT   : 5 words, 2 are the odd-ones-out (multi-select, 2 correct)
#   SYNONYMS      : two word-groups (A-C, X-Z); answer picks one from each —
#                   doesn't fit the site's flat 5-option layout, so these are
#                   parsed but flagged needs_manual (excluded from auto-upload).
#   ANALOGIES     : "X is to Y as Z is to ___", single correct answer
#   WORD CODES    : letter-substitution codes, single correct answer, no numbering
#   ANAGRAMS      : sentence with 2 jumbled words sharing a removed letter,
#                   single correct answer (the removed letter)

def parse_vr(doc):
    sections = {}
    current = None
    headers = ("ODD TWO OUT", "SYNONYMS", "ANALOGIES", "WORD CODES", "ANAGRAMS",
               "COMPOUND WORDS", "MOVE A LETTER", "MOVE TWO LETTERS", "FORM A WORD", "WORD MATCH",
               "CONNECTED WORDS", "INSERT A LETTER", "RELATED WORDS",
               "WORD COMPLETION", "WORD PAIRS", "WORD RELATIONSHIPS", "WORD OPPOSITES",
               "EMBEDDED WORDS", "MISSING LETTERS")
    header_set = set(headers)

    def _norm_header(text):
        # Papers vary: "MOVE 2 LETTERS" vs "MOVE TWO LETTERS", trailing ":" etc.
        t = text.upper().strip().rstrip(":.).!")
        return t.replace(" 2 ", " TWO ").replace(" 1 ", " ONE ")

    for block in iter_block_items(doc):
        if isinstance(block, Table):
            if current is not None:
                rows = [[c.text.strip() for c in row.cells] for row in block.rows]
                sections[current].append(("T", rows))
            continue
        t = block.text.strip()
        if not t:
            continue
        _t_norm = _norm_header(t)
        if _t_norm in header_set:
            current = _t_norm
            sections[current] = []
            continue
        if current is not None:
            sections[current].append(("P", t))

    questions = []
    questions += _parse_vr_odd_two_out(sections.get("ODD TWO OUT", []))

    # These topics use the same numbered-question format. The option labels
    # determine the option count, so the shared parser supports both 3- and
    # 5-option questions (including the sample paper's Synonyms format).
    numbered_sections = {
        "RELATED WORDS": "Related Words",
        "SYNONYMS": "Synonyms",
        "WORD COMPLETION": "Word Completion",
        "ANALOGIES": "Analogies",
        "WORD CODES": "Word Codes",
        "ANAGRAMS": "Anagrams",
        "COMPOUND WORDS": "Compound Words",
        "MOVE A LETTER": "Move a Letter",
        "MOVE TWO LETTERS": "Move Two Letters",
        "FORM A WORD": "Form a Word",
        "WORD MATCH": "Word Match",
        "CONNECTED WORDS": "Connected Words",
        "INSERT A LETTER": "Insert a Letter",
        "WORD PAIRS": "Word Pairs",
        "WORD RELATIONSHIPS": "Word Relationships",
        "WORD OPPOSITES": "Word Opposites",
        "EMBEDDED WORDS": "Embedded Words",
        "MISSING LETTERS": "Missing Letters",
    }
    # Process in the document's section order so assigned question numbers
    # follow the paper rather than the order of parser implementation details.
    for section, items in sections.items():
        title = numbered_sections.get(section)
        if title:
            questions += _parse_vr_numbered(items, title)

    for i, q in enumerate(questions, start=1):
        q["num"] = i
    return questions


_VR_ANS_RE = re.compile(r"^Answer\s*[:\t]?\s*([A-E]{1,2})\b[:,]?\s*(.*)$", re.IGNORECASE)


def _is_vr_instruction(t):
    """The one-time descriptive blurb at the top of each VR section always starts
    this way (unlike question stems), regardless of its length."""
    low = t.lower()
    return low.startswith("in this") or low.startswith("in these")


def _parse_vr_odd_two_out(items):
    questions = []
    instr = ""
    opts = {}
    for kind, val in items:
        if kind != "P":
            continue
        t = val
        m_ans = _VR_ANS_RE.match(t)
        if m_ans:
            letters = m_ans.group(1).upper()
            expl = m_ans.group(2).strip()
            options = [opts.get(l, "") for l in LETTERS]
            correct = [LETTERS.index(c) for c in letters if c in LETTERS]
            questions.append({
                "title": "Odd Two Out",
                "site_topic": None,
                "content": "text",
                "layout": LAYOUT_5_TEXT,
                "question_text": instr,
                "options": options,
                "answer": letters,
                "correct": correct,
                "explanation": expl,
                "needs_manual": False,
            })
            opts = {}
            continue
        if not opts and _is_vr_instruction(t):
            if not instr:
                instr = t
            continue
        for part in t.split("\n"):
            part = part.strip()
            if part and part[0] in LETTERS:
                opts[part[0]] = part[1:].lstrip(": ").strip()
    return questions


def _parse_vr_synonyms(items):
    questions = []
    instr = ""
    pending = None
    num_re = re.compile(r"^(\d+)\.?$")
    # answers here pick one letter from group1 (A-C) and one from group2 (X-Z),
    # e.g. "Answer: CX" — not the general A-E pattern used elsewhere.
    ans_re = re.compile(r"^Answer\s*[:\t]?\s*([A-CX-Z]{2})\b\s*(.*)$", re.IGNORECASE)
    for kind, val in items:
        if kind == "T":
            if pending is not None and val:
                row = val[0]
                pending["g1_labels"] = [x.strip() for x in row[0].split("\n") if x.strip()]
                pending["g1_words"] = [x.strip() for x in row[1].split("\n") if x.strip()]
                pending["g2_labels"] = [x.strip() for x in row[3].split("\n") if x.strip()]
                pending["g2_words"] = [x.strip() for x in row[4].split("\n") if x.strip()]
            continue
        t = val
        if not instr and _is_vr_instruction(t):
            instr = t
            continue
        if num_re.match(t):
            pending = {}
            continue
        m_ans = ans_re.match(t)
        if m_ans and pending is not None:
            letters = m_ans.group(1).upper()
            g1_labels = pending.get("g1_labels", [])
            g1_words = pending.get("g1_words", [])
            g2_labels = pending.get("g2_labels", [])
            g2_words = pending.get("g2_words", [])
            pairs = dict(zip(g1_labels, g1_words))
            pairs.update(dict(zip(g2_labels, g2_words)))
            l1, l2 = letters[0], letters[1] if len(letters) > 1 else ""
            group1_desc = ", ".join(f"{l}) {w}" for l, w in zip(g1_labels, g1_words))
            group2_desc = ", ".join(f"{l}) {w}" for l, w in zip(g2_labels, g2_words))
            qtext = (
                f"{instr}\n\nGroup 1: {group1_desc}\nGroup 2: {group2_desc}"
            )
            questions.append({
                "title": "Synonyms",
                "site_topic": None,
                "content": "text",
                "layout": LAYOUT_5_TEXT,
                "question_text": qtext,
                "options": [],
                "answer": letters,
                "correct": [],
                "explanation": f"{l1}) {pairs.get(l1, '?')} / {l2}) {pairs.get(l2, '?')}",
                "needs_manual": True,
            })
            pending = None
    return questions


def _parse_vr_numbered(items, title):
    """Shared parser for numbered verbal-reasoning sections with 3 or 5 options.

    Supports two formats:
    1. Traditional: numbered marker on its own line, stem on next line, options same/following lines, Answer+Explanation on one line.
    2. Your format: "1. Stem text" on same line as number, options on separate lines starting with A-E,
       "Answer: X" and "Explanation: ..." on separate lines.
    """
    questions = []
    instr = ""
    num_re = re.compile(r"^(\d+)\.?\s*(.*)$")  # captures stem on same line
    stem = None
    opts = {}
    pending_expl = None

    def _save():
        nonlocal stem, opts, pending_expl, instr
        if stem is None:
            return
        qtext = "\n\n".join(part for part in (instr, stem) if part)
        # Filter to only non-empty options (for 3-option layouts)
        option_letters = [letter for letter in LETTERS if opts.get(letter)]
        real_options = [opts[letter] for letter in option_letters]
        n_opts = len(real_options)
        layout = LAYOUT_3_TEXT if n_opts == 3 else LAYOUT_5_TEXT
        correct = (
            [option_letters.index(letter) for letter in pending_expl[0] if letter in option_letters]
            if pending_expl else []
        )
        questions.append({
            "title": title,
            "site_topic": None,
            "content": "text",
            "layout": layout,
            "question_text": qtext,
            "options": real_options,
            "answer": pending_expl[0] if pending_expl else "",
            "correct": correct,
            "explanation": pending_expl[1] if pending_expl else "",
            "needs_manual": False,
        })
        stem, opts, pending_expl = None, {}, None

    for kind, val in items:
        if kind != "P":
            continue
        t = val.strip()
        if not t:
            continue

        # Section instruction (e.g., "Choose the word that completes the sentence...")
        # Heuristic: long-ish sentence, no A-E option pattern, appears before any numbered Q
        if stem is None and not opts and not num_re.match(t) and not _VR_ANS_RE.match(t) and not t.startswith("Explanation"):
            if len(t) > 20 and " " in t:
                instr = "\n".join(part for part in (instr, t) if part)
            continue

        # Answer line
        m_ans = _VR_ANS_RE.match(t)
        if m_ans:
            letters = m_ans.group(1).upper()
            expl = m_ans.group(2).strip()
            pending_expl = (letters, expl)
            # If explanation is on NEXT line, we'll catch it below
            continue

        # Explanation line (separate from Answer)
        if t.startswith("Explanation"):
            if pending_expl:
                letters, _ = pending_expl
                expl = t.split(":", 1)[1].strip() if ":" in t else t
                pending_expl = (letters, expl)
            continue

        # Numbered question start
        m_num = num_re.match(t)
        if m_num:
            # Save previous if exists
            _save()
            stem = m_num.group(2).strip()
            opts = {}
            pending_expl = None
            continue

        # Option lines: "A something", "B something", etc.
        if stem is not None:
            m_opt = re.match(r"^([A-E])\s*[:\.]?\s*(.+)$", t)
            if m_opt:
                opts[m_opt.group(1)] = m_opt.group(2).strip()
                continue
            # Fallback: if line looks like option but no letter, ignore

    _save()
    return questions


# ── Image extraction (NVR / Maths) ───────────────────────────────────────────

def extract_images(docx_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        try:
            os.remove(os.path.join(out_dir, f))
        except OSError:
            pass
    paths = []
    with zipfile.ZipFile(docx_path) as z:
        media = sorted(
            [f for f in z.namelist() if f.startswith("word/media/")],
            key=lambda x: int(re.search(r"\d+", os.path.basename(x)).group()),
        )
        for i, m in enumerate(media, start=1):
            ext = os.path.splitext(m)[1]
            dest = os.path.join(out_dir, f"img{i}{ext}")
            with z.open(m) as src, open(dest, "wb") as dst:
                dst.write(src.read())
            paths.append(dest)
    return paths


# ── Preflight ────────────────────────────────────────────────────────────────

def preflight_check(questions, img_paths, confirm=True):
    print("\n" + "=" * 60)
    print("PRE-FLIGHT CHECK")
    print("=" * 60)
    print(f"  Test type        : {TEST_TYPE}")
    print(f"  TOPIC            : {TOPIC}")
    print(f"  Subject          : {SUBJECT}")
    print(f"  Questions parsed : {len(questions)}")
    print(f"  Expected count   : {CFG['expected_count']}")
    image_questions = [q for q in questions if q["content"] == "image"]
    if CFG["has_images"]:
        print(f"  Images extracted : {len(img_paths)}  (questions with an image: {len(image_questions)})")
    print()

    errors = []
    if len(questions) != CFG["expected_count"]:
        errors.append(f"  x Expected {CFG['expected_count']} questions, got {len(questions)}")
    if CFG["has_images"] and len(img_paths) != len(image_questions):
        errors.append(f"  x Expected {len(image_questions)} images (one per image question), got {len(img_paths)}")
    for q in questions:
        if q["content"] == "text" and not (q["question_text"] or "").strip():
            errors.append(f"  x Q{q['num']} has empty question text")
        if not q["needs_manual"] and len(q["options"]) not in (3, 4, 5):
            errors.append(f"  x Q{q['num']} has {len(q['options'])} options (expected 3 or 5)")

    if errors:
        print("\n".join(errors))
        print("\nAborting — fix the docx and try again.")
        return False

    print("  OK: counts look right\n")
    print("Questions to upload:")
    manual = []
    for q in questions:
        topic = q["site_topic"] or "-"
        bits = [f"Q{q['num']:>2}", f"[{topic}]", f"title='{q['title']}'"]
        if q["content"] == "image":
            bits.append(f"img=img{q['_img_no']}")
        if q["answer"]:
            bits.append(f"ans={q['answer']}")
        if q["needs_manual"]:
            bits.append("!! MANUAL — will NOT be auto-uploaded")
            manual.append(q["num"])
        print("   " + "  ".join(bits))
        if q.get("passage"):
            ppreview = q["passage"].replace("\n", " ")[:90]
            print(f"        Passage: {ppreview}")
        if (q["question_text"] or "").strip():
            preview = (q["question_text"] or "").replace("\n", " ")[:90]
            print(f"        Q: {preview}")
            if q["options"]:
                print(f"        options: {q['options']}")
    print()
    if manual:
        print(f"NOTE: Q{', Q'.join(map(str, manual))} need manual entry through the site's UI "
              f"and will be SKIPPED during upload (not created at all).\n")

    if not confirm:
        return True

    while True:
        ans = input("Proceed with upload? [y/n]: ").strip().lower()
        if ans == "y":
            return True
        if ans == "n":
            print("Aborted by user.")
            return False
        print("Please type y or n.")


# ── Browser helpers ──────────────────────────────────────────────────────────

async def login(page):
    await page.goto(LOGIN_URL)
    await page.wait_for_load_state("networkidle")
    await page.fill('input[type="email"], input[name="email"]', EMAIL)
    await page.fill('input[type="password"], input[name="password"]', PASSWORD)
    await page.click('button[type="submit"]')
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1500)
    if "login" in page.url:
        raise RuntimeError("Login failed — check credentials")
    print("Logged in\n")


async def select_dropdowns(page, site_topic):
    """Set Test Type, Subject, and (for English) Topic."""
    await page.wait_for_selector("select")
    selects = await page.query_selector_all("select")
    await selects[0].select_option(label=TEST_TYPE)
    await page.wait_for_timeout(400)

    selects = await page.query_selector_all("select")
    await selects[1].select_option(label=SUBJECT)
    await page.wait_for_timeout(600)

    if CFG["has_topic_dropdown"] and site_topic:
        await page.wait_for_function(
            "([t]) => Array.from(document.querySelectorAll('select'))"
            ".some(s => Array.from(s.options).some(o => o.text === t))",
            arg=[site_topic],
        )
        selects = await page.query_selector_all("select")
        await selects[2].select_option(label=site_topic)
        await page.wait_for_timeout(400)


async def select_layout(page, needle):
    ok = await page.evaluate(
        """(needle) => {
            const ps = [...document.querySelectorAll('p')].filter(p => (p.textContent||'').includes(needle));
            if (!ps.length) return false;
            let card = ps[0];
            for (let i = 0; i < 5 && card; i++) {
                card = card.parentElement;
                if (card && (card.className||'').includes('cursor-pointer')) break;
            }
            (card || ps[0]).click();
            return true;
        }""",
        needle,
    )
    if not ok:
        raise RuntimeError(f"Layout card matching '{needle}' not found")
    await page.wait_for_timeout(600)


async def set_editor(page, editor_index, text):
    editors = await page.query_selector_all(".rsw-ce")
    if len(editors) <= editor_index:
        raise RuntimeError(f"Editor #{editor_index} not found (found {len(editors)})")
    await page.evaluate(
        """([el, t]) => {
            el.focus();
            el.innerText = t;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        }""",
        [editors[editor_index], text],
    )


async def inject_image(page, file_input_index, img_path):
    with open(img_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode()
    name = os.path.basename(img_path)
    mime = "image/png" if name.lower().endswith(".png") else "image/jpeg"
    ok = await page.evaluate(
        """(args) => {
            const [idx, b64, fname, mime] = args;
            const inputs = document.querySelectorAll('input[type="file"]');
            if (!inputs[idx]) return false;
            const bin = atob(b64), ab = new ArrayBuffer(bin.length), ia = new Uint8Array(ab);
            for (let i = 0; i < bin.length; i++) ia[i] = bin.charCodeAt(i);
            const file = new File([ab], fname, {type: mime});
            const dt = new DataTransfer(); dt.items.add(file);
            inputs[idx].files = dt.files;
            inputs[idx].dispatchEvent(new Event('change', {bubbles: true}));
            inputs[idx].dispatchEvent(new Event('input', {bubbles: true}));
            return true;
        }""",
        [file_input_index, img_b64, name, mime],
    )
    if not ok:
        raise RuntimeError(f"File input [{file_input_index}] not found")


async def set_difficulty(page, value):
    await page.evaluate(
        """(val) => {
            const r = document.querySelector(`input[type=radio][value="${val}"]`);
            if (r && !r.checked) r.click();
        }""",
        value,
    )


async def set_correct(page, indices):
    # The correct-answer control is <button role="switch" data-testid="answer-N-correct">.
    # Radix ignores a bare JS .click(); use a real Playwright mouse click.
    for i in indices:
        sel = f'button[data-testid="answer-{i}-correct"]'
        btn = page.locator(sel)
        if await btn.count() == 0:
            raise RuntimeError(f"Correct toggle {sel} not found")
        await btn.scroll_into_view_if_needed()
        for _ in range(4):
            state = await btn.get_attribute("aria-checked")
            if state == "true":
                break
            try:
                await btn.click(timeout=3000)
            except Exception:
                await page.evaluate(
                    """(sel) => {
                        const el = document.querySelector(sel);
                        if (!el) return;
                        el.dispatchEvent(new PointerEvent('pointerdown', {bubbles: true}));
                        el.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                        el.dispatchEvent(new PointerEvent('pointerup', {bubbles: true}));
                        el.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
                        el.dispatchEvent(new MouseEvent('click', {bubbles: true}));
                    }""",
                    sel,
                )
            await page.wait_for_timeout(250)
        else:
            state = await btn.get_attribute("aria-checked")
            if state != "true":
                raise RuntimeError(f"Could not set correct answer on {sel} (aria-checked={state})")


async def ensure_option_count(page, needed):
    """Some layout cards (e.g. English Spelling/Punctuation/Grammar) render a
    variable '4-5 answers' widget that starts with only 4 answer inputs and a
    '+ Add answer option' button to reveal the 5th. Click it until we have
    enough inputs for this question."""
    for _ in range(needed + 2):
        count = await page.evaluate(
            """() => document.querySelectorAll('input[name^="answers."]').length"""
        )
        if count >= needed:
            return
        clicked = await page.evaluate(
            """() => {
                const b = [...document.querySelectorAll('button')]
                    .find(b => (b.innerText||'').includes('Add answer option'));
                if (b) { b.click(); return true; }
                return false;
            }"""
        )
        if not clicked:
            raise RuntimeError(
                f"Need {needed} answer options but found {count} and no "
                "'Add answer option' button to add more"
            )
        await page.wait_for_timeout(300)


async def fill_options(page, options):
    await ensure_option_count(page, len(options))
    for i, opt in enumerate(options):
        sel = f'input[name="answers.{i}.text"]'
        el = await page.query_selector(sel)
        if el is None:
            raise RuntimeError(f"Answer input {sel} not found")
        await page.fill(sel, opt)


# ── Upload one question ──────────────────────────────────────────────────────

async def upload_question(page, q, img_paths, dry_run, no_submit):
    topic = q["site_topic"] or "-"
    label = f"Q{q['num']} [{topic}]"
    print(f"-> {label} ...", end=" ", flush=True)

    if dry_run:
        print("[DRY RUN]")
        return

    await page.goto(ADD_URL)
    await page.wait_for_load_state("networkidle")

    await select_dropdowns(page, q["site_topic"])
    await select_layout(page, q["layout"])

    # difficulty
    await set_difficulty(page, DIFFICULTY)

    # title
    await page.fill('input[name="questionTitle"]', q["title"][:120])

    # question content
    # Editor order on the site is normally [Question, Explanation] (indices 0, 1).
    # English Comprehension is the one layout that also renders a "Passage Text"
    # editor *before* the Question editor, shifting everything: [Passage, Question,
    # Explanation] (indices 0, 1, 2). Only questions with a non-empty "passage"
    # trigger that 3-editor path.
    expl_editor = 1
    passage = (q.get("passage") or "").strip()
    if q["content"] == "image":
        img = img_paths[q["_img_idx"]]
        await inject_image(page, 0, img)
        try:
            async with page.expect_response(
                lambda r: "upload-image" in r.url and r.status == 200, timeout=20000
            ):
                pass
        except Exception:
            await page.wait_for_timeout(2000)
        await page.wait_for_timeout(800)
        # the "Question" editor is a separate REQUIRED field (the prompt text)
        if (q.get("question_text") or "").strip():
            await set_editor(page, 0, q["question_text"])
            await page.wait_for_timeout(150)
    elif passage:
        await set_editor(page, 0, passage)
        await page.wait_for_timeout(150)
        await set_editor(page, 1, q["question_text"])
        await page.wait_for_timeout(150)
        expl_editor = 2
    else:
        await set_editor(page, 0, q["question_text"])
        await page.wait_for_timeout(200)

    # options + correct
    await fill_options(page, q["options"])
    if q["correct"]:
        await set_correct(page, q["correct"])

    # explanation
    if q.get("explanation"):
        try:
            await set_editor(page, expl_editor, q["explanation"])
        except Exception:
            pass

    if no_submit:
        os.makedirs(SHOT_DIR, exist_ok=True)
        shot = os.path.join(SHOT_DIR, f"q{q['num']}_{topic}.png")
        await page.screenshot(path=shot, full_page=True)
        state = await page.evaluate(
            """() => {
                const val = s => { const e = document.querySelector(s); return e ? e.value : null; };
                const answers = [...document.querySelectorAll('input[name^="answers."][name$=".text"]')].map(e => e.value);
                const correct = [...document.querySelectorAll('button[role=switch][data-testid$="-correct"]')]
                    .map((e, i) => ({i, on: e.getAttribute('aria-checked')})).filter(x => x.on === 'true').map(x => x.i);
                const previewImgs = [...document.querySelectorAll('img')].map(i => i.src)
                    .filter(s => s && !s.includes('/images/iseb') && (s.startsWith('blob:') || s.includes('amazonaws') || s.includes('/upload')));
                const editors = [...document.querySelectorAll('.rsw-ce')].map(e => (e.innerText||'').slice(0,110));
                return {title: val('input[name=questionTitle]'), answers, correct, previewImgs, qeditor: editors[0]};
            }"""
        )
        print(f"[NO-SUBMIT] title={state['title']!r} answers={state['answers']} "
              f"correct={state['correct']} preview_imgs={state['previewImgs']}")
        print(f"           qeditor={state['qeditor']!r}")
        print(f"           screenshot -> {shot}")
        return

    # submit — real click (scrolled into view); JS fallback
    try:
        btn = page.locator('button', has_text="Finish").last
        await btn.scroll_into_view_if_needed()
        await btn.click(timeout=5000)
    except Exception:
        await page.evaluate(
            """() => { const b = [...document.querySelectorAll('button')]
                .find(b => (b.innerText||'').trim() === 'Finish'); if (b) b.click(); }"""
        )
    try:
        await page.wait_for_function(
            "document.body.innerText.includes('created successfully')"
            " || document.body.innerText.includes('successfully')"
            " || !window.location.href.includes('/add')",
            timeout=20000,
        )
        print("OK")
    except Exception:
        diag = await page.evaluate(
            """() => {
                const errs = [...document.querySelectorAll('[role=alert], .text-destructive, .text-red-500, p')]
                    .map(e => (e.innerText||'').trim())
                    .filter(t => t && /error|required|select|must|invalid|please/i.test(t));
                return {url: location.href, errs: [...new Set(errs)].slice(0, 10)};
            }"""
        )
        os.makedirs(SHOT_DIR, exist_ok=True)
        shot = os.path.join(SHOT_DIR, f"error_q{q['num']}.png")
        await page.screenshot(path=shot, full_page=True)
        raise RuntimeError(f"Submit failed url={diag['url']} errs={diag['errs']} shot={shot}")


# ── Main ─────────────────────────────────────────────────────────────────────

async def main(args):
    if not CFG.get("enabled", True) and not args.dry_run:
        print(f"\n'{TOPIC}' upload is ON HOLD (see ISEB_MAP['{TOPIC}']['enabled'] in the script).")
        print("(You can still use --dry-run to inspect parsing.)")
        sys.exit(0)

    print(f"Parsing {os.path.basename(DOCX_PATH)} ...")
    questions = parse_docx(DOCX_PATH)

    img_paths = []
    if CFG["has_images"]:
        print("Extracting images ...")
        img_paths = extract_images(DOCX_PATH, IMG_DIR)
        # map one image per question, in order, only to questions that actually have one
        img_idx = 0
        for q in questions:
            if q["content"] == "image":
                q["_img_idx"] = img_idx
                q["_img_no"] = img_idx + 1
                img_idx += 1

    if not preflight_check(questions, img_paths, confirm=not args.dry_run):
        sys.exit(1)

    upload_qs = [q for q in questions if not q["needs_manual"]]
    skipped_manual = [q["num"] for q in questions if q["needs_manual"]]
    if skipped_manual:
        print(f"Skipping {len(skipped_manual)} manual-only question(s): "
              f"Q{', Q'.join(map(str, skipped_manual))}\n")

    if args.start or args.end:
        lo = args.start or 1
        hi = args.end or max(q["num"] for q in questions)
        upload_qs = [q for q in upload_qs if lo <= q["num"] <= hi]
        print(f"\nFiltered to Q{lo}-Q{hi}: {len(upload_qs)} questions\n")

    if args.dry_run:
        print("\n[DRY RUN — no browser, no uploads]\n")
        for q in upload_qs:
            await upload_question(None, q, img_paths, True, False)
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        async def on_response(resp):
            if "upload-image" in resp.url:
                print(f"    [net] upload-image -> {resp.status}", flush=True)
        page.on("response", on_response)

        await login(page)

        failed = []
        for i, q in enumerate(upload_qs):
            try:
                await upload_question(page, q, img_paths, False, args.no_submit)
            except Exception as e:
                print(f"x  ({e})")
                failed.append(q["num"])

            if i < len(upload_qs) - 1 and DELAY_MIN > 0 and not args.no_submit:
                delay = random.randint(DELAY_MIN, DELAY_MAX)
                print(f"   waiting {delay}s before next question ...", flush=True)
                await asyncio.sleep(delay)

        await browser.close()

    print(f"\n{'=' * 60}")
    ok = len(upload_qs) - len(failed)
    print(f"Done: {ok}/{len(upload_qs)} processed")
    if failed:
        print(f"Failed: Q{', Q'.join(map(str, failed))}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-submit", action="store_true")
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    asyncio.run(main(parser.parse_args()))
