#!/usr/bin/env python3
"""
Upload CAT4 questions to admin.11plusonline.co.uk  (newer "SET 4"-style papers)

Difference vs cat4.py:
  * Papers now start with 2 PRACTICE questions ("EXAMPLE 1/2") before Q1.
  * Question titles are now built from the instructions sentence:
        regular  ->  "TOPIC - INSTRUCTIONS"
        practice ->  "PQ (TOPIC) - INSTRUCTIONS"
    where INSTRUCTIONS is the paragraph directly under the topic-name heading.
  * Images are associated with questions by document order (not naive index
    pairing), so the leading example images no longer shift the mapping.

Usage: python3 new_cat4.py [--dry-run] [--start N] [--end N] [--skip-practice]
"""

import asyncio
import argparse
import base64
import docx
import random
import re
import zipfile
import os
import sys
from docx.oxml.ns import qn
from playwright.async_api import async_playwright

# ── User config — only change these two lines per paper ──────────────────────
TOPIC     = "Figure Analysis"  # must match a key in CAT4_MAP below
DOCX_PATH = r"C:\Users\angel\Documents\Exceptional Academics\CAT 4\CAT 4 SET 4\Figure Analysis.docx"
print(DOCX_PATH)
print(os.path.exists(DOCX_PATH))
# ─────────────────────────────────────────────────────────────────────────────

# ── Delay between uploads (seconds) — set both to 0 to disable ───────────────
DELAY_MIN = 60   # 3 minutes
DELAY_MAX = 120   # 5 minutes
# ─────────────────────────────────────────────────────────────────────────────

LOGIN_URL = "https://admin.11plusonline.co.uk/login"
ADD_URL   = "https://admin.11plusonline.co.uk/question-bank/question/add"
EMAIL     = "admin1@11plusonline.co.uk"
PASSWORD  = "password1"
IMG_DIR   = os.path.join(os.environ.get("TEMP", "/tmp"), "fig_upload_images")

# ── CAT4 map: TOPIC → (subject, site_topic, expected_count, content_type, imgs_per_q)
#
# content_type "image" : question is an image file; answer inputs get A/B/C/D/E labels
# content_type "text"  : question is typed into a rich-text editor; answer inputs get
#                        the actual option values (numbers / words) from the docx
# imgs_per_q           : images per REGULAR question from the docx
#                        0 = text-based (no images), 1 = single question image,
#                        2 = question image + explanation image (Figure Recognition)
CAT4_MAP = {
    "Figure Classification": ("Non Verbal Reasoning",   "Figure Classification", 24, "image", 1),
    "Figure Matrices":       ("Non Verbal Reasoning",   "Figure Matrices",       24, "image", 1),
    "Verbal Classification": ("Verbal Reasoning",       "Verbal Classification", 24, "text",  0),
    "Verbal Analogies":      ("Verbal Reasoning",       "Verbal Analogies",      24, "text",  0),
    "Number Analogies":      ("Quantitative Reasoning", "Number Analogies",      18, "text",  0),
    "Number Series":         ("Quantitative Reasoning", "Number Series",         18, "text",  0),
    "Figure Analysis":       ("Spatial Reasoning",      "Figure Analysis",       18, "image", 1),
    "Figure Recognition":    ("Spatial Reasoning",      "Figure Recognition",    18, "image", 2),
}

if TOPIC not in CAT4_MAP:
    print(f"ERROR: Unknown topic '{TOPIC}'. Valid topics: {list(CAT4_MAP)}")
    sys.exit(1)

SUBJECT, SITE_TOPIC, EXPECTED_COUNT, CONTENT_TYPE, IMGS_PER_Q = CAT4_MAP[TOPIC]
TEST_TYPE  = "CAT4"
DIFFICULTY = "easy"

LETTERS = ["A", "B", "C", "D", "E"]
BLIP  = qn("a:blip")
EMBED = qn("r:embed")


# ── Small helpers ─────────────────────────────────────────────────────────────

def _clean_expl(s: str) -> str:
    """Strip the leading joiner junk from an explanation fragment."""
    s = (s or "").strip()
    s = s.lstrip(")=(-–—\u2212 \t").strip()
    s = s.strip("()").strip()
    return s


def extract_instructions(doc) -> str:
    """The paragraph directly under the topic-name heading (2nd non-empty para).

    Older SET-3 docs have no heading/instructions block — they open straight at
    'EXAMPLE 1' or a numbered question — so return "" for those.
    """
    nonempty = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if len(nonempty) < 2:
        return ""
    first = nonempty[0]
    # Old format: no topic heading, so there is no instructions paragraph.
    if re.match(r"^example\b", first, re.IGNORECASE) or re.match(r"^\d+\s*[.)]", first):
        return ""
    cand = nonempty[1]
    # Guard: never mistake an answer/example/numbered line for instructions.
    if re.match(r"^(answer\b|example\b|\d+\s*[.)])", cand, re.IGNORECASE):
        return ""
    return cand


def parse_answer_table(doc) -> dict:
    """Answer key stored as a table [num | letter | explanation] (old SET-3 docs).

    Returns {question_number: (letter, explanation)}.
    """
    key = {}
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if (len(cells) >= 2
                    and re.match(r"^\d+$", cells[0])
                    and re.match(r"^[A-E]$", cells[1], re.IGNORECASE)):
                num = int(cells[0])
                letter = cells[1].upper()
                expl = _clean_expl(cells[2]) if len(cells) >= 3 else ""
                key[num] = (letter, expl)
    return key


# ── Image extraction (document-order aware) ───────────────────────────────────

def extract_media(doc, docx_path: str, out_dir: str) -> dict:
    """Extract every embedded image to out_dir and return {relationship-id: path}."""
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        try:
            os.remove(os.path.join(out_dir, f))
        except OSError:
            pass

    base2path = {}
    with zipfile.ZipFile(docx_path) as z:
        for m in z.namelist():
            if m.startswith("word/media/"):
                dest = os.path.join(out_dir, os.path.basename(m))
                with z.open(m) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
                base2path[os.path.basename(m)] = dest

    rid2path = {}
    for rid, rel in doc.part.rels.items():
        try:
            if "image" in rel.reltype:
                b = os.path.basename(rel.target_ref)
                if b in base2path:
                    rid2path[rid] = base2path[b]
        except Exception:
            pass
    return rid2path


def para_images(par, rid2path: dict) -> list:
    """Extracted image paths anchored in this paragraph, in order."""
    out = []
    for blip in par._p.findall(".//" + BLIP):
        rid = blip.get(EMBED)
        if rid and rid in rid2path:
            out.append(rid2path[rid])
    return out


def body_image_paths(doc, rid2path: dict) -> list:
    """Every embedded image in whole-document (body) order.

    Some docs (e.g. SET-3 Figure Matrices) anchor a question's image to the
    *following* number paragraph, so per-paragraph anchoring misassigns them.
    Assigning by strict body order sidesteps that entirely.
    """
    out = []
    for blip in doc.element.body.findall(".//" + BLIP):
        rid = blip.get(EMBED)
        if rid and rid in rid2path:
            out.append(rid2path[rid])
    return out


# ── Parsing ───────────────────────────────────────────────────────────────────

def parse_docx(path: str):
    """Return (practice, regular, instructions)."""
    doc = docx.Document(path)
    instructions = extract_instructions(doc)
    if CONTENT_TYPE == "image":
        practice, regular = _parse_image_docx(doc, path)
    else:
        practice, regular = _parse_text_docx(doc)

    reg_title = f"{TOPIC} - {instructions}".strip(" -")
    pq_title  = f"PQ ({TOPIC}) - {instructions}".strip(" -")
    for q in practice:
        q["title"] = pq_title
    for q in regular:
        q["title"] = reg_title
    return practice, regular, instructions


# ── Image topics: FC, FM, FA, FR ─────────────────────────────────────────────

def _parse_ans_line(t: str, q: dict):
    """Answer info that lives on the same line as a number ('1 B ...', '3 Answer: A')."""
    m = re.match(r"^\d+\s*[.)]?\s*Answer:\s*([A-E])\b\s*(.*)$", t, re.IGNORECASE)
    if not m:
        m = re.match(r"^\d+\s*[.)]?\s+([A-E])\b\s*(.*)$", t)
    if m:
        if q["answer"] is None:
            q["answer"] = m.group(1).upper()
        rest = _clean_expl(m.group(2))
        if rest:
            q["explanation"] = (q["explanation"] + " " + rest).strip() if q["explanation"] else rest


def _parse_ans_text(t: str, q: dict):
    """Answer info on its own line ('Answer: B', 'Answer: C — explanation')."""
    m = re.match(r"^Answer:\s*([A-E])\b\s*(.*)$", t, re.IGNORECASE)
    if m:
        if q["answer"] is None:
            q["answer"] = m.group(1).upper()
        rest = _clean_expl(m.group(2))
        if rest:
            q["explanation"] = (q["explanation"] + " " + rest).strip() if q["explanation"] else rest


def _parse_image_docx(doc, path: str):
    rid2path = extract_media(doc, path, IMG_DIR)

    practice = []
    regular = {}

    def reg(num: int) -> dict:
        d = regular.get(num)
        if d is None:
            d = {"num": num, "is_practice": False, "label": f"Q{num}",
                 "answer": None, "answer_options": None, "question_text": None,
                 "explanation": "", "images": [], "needs_manual": False}
            regular[num] = d
        return d

    cur = None          # question currently receiving trailing images / answer text
    phase = "q"         # "q" = question section, "a" = answer section
    last_qnum = 0
    started = False

    for par in doc.paragraphs:
        t = par.text.strip()
        imgs = para_images(par, rid2path)

        # Skip any leading front-matter (topic-name heading + instructions
        # paragraph in SET-4 docs) until we reach the first EXAMPLE or the
        # first numbered question. Old SET-3 docs start straight at EXAMPLE 1,
        # so nothing is skipped for them. We fall through (no `continue`) once
        # started so the triggering EXAMPLE/question line is processed below.
        if not started:
            is_example  = re.match(r"^example\b", t, re.IGNORECASE)
            is_numbered = re.match(r"^\d+\s*[.)]", t) or re.match(r"^\d+\s+[A-Za-z]", t)
            if is_example or is_numbered:
                started = True
            else:
                continue

        # ── practice example heading ──────────────────────────────────────────
        if re.match(r"^example\b\s*:?\s*\d*", t, re.IGNORECASE) and "answer" not in t.lower():
            pnum = len(practice) + 1
            cur = {"num": 0, "is_practice": True, "practice_num": pnum,
                   "label": f"PQ{pnum}", "answer": None, "answer_options": None,
                   "question_text": None, "explanation": "", "images": list(imgs),
                   "needs_manual": False}
            practice.append(cur)
            continue

        # ── numbered item (question or answer, depending on phase) ────────────
        num_m = re.match(r"^(\d+)\s*[.)]", t) or re.match(r"^(\d+)\s+[A-Za-z]", t)
        if num_m:
            num = int(num_m.group(1))
            if phase == "q" and num <= last_qnum:
                phase = "a"       # numbering reset -> we've reached the answer key
            if phase == "q":
                last_qnum = num
                cur = reg(num)
                cur["images"] += imgs
            else:
                cur = reg(num)
                _parse_ans_line(t, cur)
                cur["images"] += imgs
            continue

        # ── continuation line (belongs to the current question) ───────────────
        if cur is not None:
            low = t.lower()
            if low.startswith("answer"):
                _parse_ans_text(t, cur)
            elif low.startswith("explanation"):
                e = _clean_expl(re.sub(r"^explanation:?", "", t, flags=re.IGNORECASE))
                if e:
                    cur["explanation"] = (cur["explanation"] + " " + e).strip() if cur["explanation"] else e
            cur["images"] += imgs

    regular_list = sorted(regular.values(), key=lambda q: q["num"])

    # Old SET-3 docs keep the regular answer key in a table
    # (num | letter | explanation) rather than inline "Answer:" lines. Fill in
    # any answers/explanations still missing after the paragraph scan.
    answer_table = parse_answer_table(doc)
    for q in regular_list:
        if q["num"] in answer_table:
            letter, expl = answer_table[q["num"]]
            if q["answer"] is None:
                q["answer"] = letter
            if expl and not q["explanation"]:
                q["explanation"] = expl

    # Prefer strict body-order image assignment when the image count is regular
    # (every example + question has exactly IMGS_PER_Q images). This fixes docs
    # whose drawings are anchored to the wrong paragraph. Falls back to the
    # per-paragraph anchoring above when counts are irregular.
    if IMGS_PER_Q:
        items = practice + regular_list          # document order (examples first)
        seq = body_image_paths(doc, rid2path)
        if len(seq) == len(items) * IMGS_PER_Q:
            it = iter(seq)
            for q in items:
                q["images"] = [next(it) for _ in range(IMGS_PER_Q)]

    # image questions answer with A/B/C/D/E labels
    for q in practice + regular_list:
        q["answer_options"] = list(LETTERS)
        if q["answer"] is None:            # e.g. FR practice shows the answer only as an image
            q["answer"] = "A"
            q["needs_manual"] = True

    return practice, regular_list


# ── Text topics: NA, NS, VA, VC ──────────────────────────────────────────────

def _split_q_opts(content: str):
    """Split a block of 'stem ... a) .. b) ..' into (question_text, [5 options])."""
    parts = re.split(r"\n(?=[a-e]\))", content, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) > 1:
        qt, opts_raw = parts[0].strip(), parts[1]
    else:
        pm = re.search(r"(?im)(?:^|\s)a\)\s", content)
        if pm:
            qt, opts_raw = content[:pm.start()].strip(), content[pm.start():]
        else:
            qt, opts_raw = content.strip(), ""
    opts = [o.strip() for o in re.split(r"[a-e]\)\s*", opts_raw, flags=re.IGNORECASE) if o.strip()]
    return qt, opts


def _collect_q_block(paragraphs, start: int):
    """Join a Q-line with following option/stem lines until the next question."""
    block = [paragraphs[start]]
    j = start + 1
    while j < len(paragraphs):
        nxt = paragraphs[j]
        if re.match(r"^Q\d+\.", nxt) or re.match(r"^example\b", nxt, re.IGNORECASE):
            break
        if re.match(r"^answers?\b", nxt, re.IGNORECASE) and len(nxt) < 60:
            break
        block.append(nxt)
        # one dedicated options line is enough (SET-4 VA/VC: stem then a)…e))
        if re.search(r"(?i)(?:^|\s)a\)\s", nxt) and re.search(r"(?i)(?:^|\s)e\)\s", nxt):
            j += 1
            break
        j += 1
    return "\n".join(block).strip(), j


def _parse_text_practice(doc):
    paras = [p.text.strip() for p in doc.paragraphs]
    n = len(paras)
    practice = []
    for idx, t in enumerate(paras):
        m = re.match(r"^example\b\s*:?\s*\d*\s*[.:]?\s*(.*)$", t, re.IGNORECASE | re.DOTALL)
        if not m:
            continue
        remainder = m.group(1).strip()
        block = [remainder] if remainder else []
        ans, expl = None, ""
        j = idx + 1
        while j < n:
            tj = paras[j].strip()
            am = re.match(r"^Answer:\s*([a-eA-E])\)?\s*(.*)$", tj)
            if am:
                ans = am.group(1).upper()
                expl = _clean_expl(am.group(2))
                break
            if re.match(r"^example\b", tj, re.IGNORECASE) or re.match(r"^Q\d+\.", tj):
                break
            if tj:
                block.append(tj)
            j += 1
        qt, opts = _split_q_opts("\n".join(block).strip())
        pnum = len(practice) + 1
        practice.append({
            "num": 0, "is_practice": True, "practice_num": pnum, "label": f"PQ{pnum}",
            "answer": ans or "A", "answer_options": (opts + [""] * 5)[:5],
            "question_text": qt, "explanation": expl, "images": [],
            "needs_manual": ans is None,
        })
        if len(practice) >= 2:
            break
    return practice


def _parse_text_docx(doc):
    practice = _parse_text_practice(doc)

    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # The real answer key sits under a standalone answers header, which varies
    # between papers: "Answers and Explanations", "ANSWERS & EXPLANATIONS" (SET-4)
    # or a singular "ANSWER" / "Answer Key" (SET-3). Match any line starting with
    # "Answer(s)" that is NOT a practice "Answer: <letter> …" line (those carry a
    # single A–E option right after the word and must not be mistaken for it).
    answer_key = {}
    header_idx = len(paragraphs)
    for i, para in enumerate(paragraphs):
        is_header = (
            re.match(r"^answers?\b", para, re.IGNORECASE)
            and len(para) < 60
            and not re.match(r"^answer\s*:?\s*[a-e]\b", para, re.IGNORECASE)
        )
        if is_header:
            header_idx = i
            ak_text = "\n".join(paragraphs[i + 1:])
            for line in ak_text.split("\n"):
                m = re.match(r"Q(\d+)\.?\s+(?:correct answer:\s*)?([a-e])\)", line, re.IGNORECASE)
                if m:
                    num = int(m.group(1))
                    letter = m.group(2).upper()
                    expl_m = re.search(r"[—–\u2212-]\s*(.+)", line)
                    expl = expl_m.group(1).strip().rstrip(".") if expl_m else ""
                    answer_key[num] = (letter, expl)
            break

    regular = []
    body = paragraphs[:header_idx]
    i = 0
    while i < len(body):
        m = re.match(r"Q(\d+)\.\s+(.+)", body[i], re.DOTALL)
        if not m:
            i += 1
            continue
        num = int(m.group(1))
        # SET-4 VA/VC put options on the next paragraph: "Q1. stem" then "a) … e)"
        joined, i = _collect_q_block(body, i)
        qt, opts = _split_q_opts(re.sub(r"^Q\d+\.\s+", "", joined, count=1))
        if len(opts) != 5:
            print(f"  WARNING: Q{num} parsed {len(opts)} options (expected 5): {opts}")
        if num in answer_key:
            letter, explanation = answer_key[num]
            needs_manual = False
        else:
            letter, explanation, needs_manual = "A", "", True   # answer missing from docx
        regular.append({
            "num": num, "is_practice": False, "label": f"Q{num}",
            "answer": letter, "answer_options": (opts + [""] * 5)[:5],
            "question_text": qt, "explanation": explanation, "images": [],
            "needs_manual": needs_manual,
        })

    regular.sort(key=lambda q: q["num"])
    return practice, regular


# ── Preflight ─────────────────────────────────────────────────────────────────

def preflight_check(practice, regular, instructions) -> bool:
    all_qs = practice + regular
    print("\n" + "=" * 60)
    print("PRE-FLIGHT CHECK")
    print("=" * 60)
    print(f"  Topic            : {TOPIC}")
    print(f"  Subject          : {SUBJECT}")
    print(f"  Content type     : {CONTENT_TYPE}")
    print(f"  Instructions     : {instructions[:80]}{'…' if len(instructions) > 80 else ''}")
    print(f"  Practice parsed  : {len(practice)}  (expected 2)")
    print(f"  Regular parsed   : {len(regular)}  (expected {EXPECTED_COUNT})")
    if IMGS_PER_Q:
        total_imgs = sum(len(q["images"]) for q in all_qs)
        print(f"  Images anchored  : {total_imgs}")
    print()
    print(f"  Regular title    : {regular[0]['title'] if regular else '(none)'}")
    print(f"  Practice title   : {practice[0]['title'] if practice else '(none)'}")
    print()

    errors, warnings = [], []
    if len(regular) != EXPECTED_COUNT:
        errors.append(f"  ✗ Expected {EXPECTED_COUNT} regular questions, got {len(regular)}")
    if len(practice) != 2:
        warnings.append(f"  ! Expected 2 practice questions, got {len(practice)}")

    for q in all_qs:
        if CONTENT_TYPE == "image":
            if not q["images"]:
                errors.append(f"  ✗ {q['label']} has no image")
            elif IMGS_PER_Q == 2 and len(q["images"]) < 2:
                warnings.append(f"  ! {q['label']} has {len(q['images'])} image(s), expected 2")
        else:
            if not (q["question_text"] or "").strip():
                errors.append(f"  ✗ {q['label']} has empty question text")
            if len([o for o in q["answer_options"] if o]) != 5:
                warnings.append(f"  ! {q['label']} has {len([o for o in q['answer_options'] if o])} options")
        if q.get("needs_manual"):
            warnings.append(f"  ! {q['label']} answer unknown from docx — defaulted to A (fix manually)")

    if warnings:
        print("\n".join(warnings) + "\n")
    if errors:
        print("\n".join(errors))
        print("\nAborting — fix the docx and try again.")
        return False

    print("  ✓ Counts look right\n")
    print("Questions to upload:")
    for q in all_qs:
        parts = [f"{q['label']:>4}", f"answer={q['answer']}"]
        if CONTENT_TYPE == "image":
            parts.append(f"imgs={len(q['images'])}")
        else:
            parts.append(f"opts={q['answer_options']}")
        if q.get("explanation"):
            parts.append("expl=✓")
        print("    " + "  ".join(parts))
    print()

    while True:
        ans = input("Proceed with upload? [y/n]: ").strip().lower()
        if ans == "y":
            return True
        if ans == "n":
            print("Aborted by user.")
            return False
        print("Please type y or n.")


# ── Browser helpers ───────────────────────────────────────────────────────────

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
    print("✓ Logged in\n")


async def _inject_file(page, file_input_index: int, img_path: str):
    """Inject an image into the Nth file input via JS DataTransfer (fires React onChange)."""
    with open(img_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode()
    img_name = os.path.basename(img_path)
    img_mime = "image/png" if img_name.lower().endswith(".png") else "image/jpeg"

    ok = await page.evaluate(f"""(idx) => {{
        const inputs = document.querySelectorAll('input[type="file"]');
        if (!inputs[idx]) return false;
        const b = atob("{img_b64}"), ab = new ArrayBuffer(b.length), ia = new Uint8Array(ab);
        for (let i = 0; i < b.length; i++) ia[i] = b.charCodeAt(i);
        const f = new File([ab], "{img_name}", {{type:"{img_mime}"}});
        const dt = new DataTransfer(); dt.items.add(f);
        inputs[idx].files = dt.files;
        inputs[idx].dispatchEvent(new Event('change', {{bubbles:true}}));
        inputs[idx].dispatchEvent(new Event('input',  {{bubbles:true}}));
        return true;
    }}""", file_input_index)
    if not ok:
        raise RuntimeError(f"File input [{file_input_index}] not found on page")
    return img_mime


async def _set_editor(page, editor_index: int, text: str):
    """Set text on the Nth .rsw-ce contenteditable editor via JS."""
    editors = await page.query_selector_all('.rsw-ce')
    if len(editors) <= editor_index:
        return
    await page.evaluate(
        "([el, t]) => { el.innerText = t; el.dispatchEvent(new Event('input', {bubbles:true})); }",
        [editors[editor_index], text]
    )


async def _select_option_js(page, label: str, retries: int = 8):
    """Set a <select> by option text entirely in JS (no Playwright handle).

    Cascading React dropdowns remount the nodes as soon as a value changes, so
    select_option() dies with 'Element is not attached to the DOM'.
    """
    last_err = None
    for _ in range(retries):
        try:
            await page.wait_for_function(
                """(lab) => Array.from(document.querySelectorAll('select')).some(
                    s => !s.disabled && Array.from(s.options).some(
                        o => (o.text || '').trim() === lab || o.value === lab
                    )
                )""",
                arg=label,
                timeout=15000,
            )
            await page.wait_for_timeout(200)
            ok = await page.evaluate(
                """(lab) => {
                    const s = Array.from(document.querySelectorAll('select')).find(sel =>
                        Array.from(sel.options).some(
                            o => (o.text || '').trim() === lab || o.value === lab
                        )
                    );
                    if (!s) return false;
                    const opt = Array.from(s.options).find(
                        o => (o.text || '').trim() === lab || o.value === lab
                    );
                    const last = s.value;
                    const desc = Object.getOwnPropertyDescriptor(
                        window.HTMLSelectElement.prototype, 'value'
                    );
                    desc.set.call(s, opt.value);
                    if (s._valueTracker) s._valueTracker.setValue(last);
                    opt.selected = true;
                    s.dispatchEvent(new Event('input', { bubbles: true }));
                    s.dispatchEvent(new Event('change', { bubbles: true }));
                    return s.value === opt.value;
                }""",
                label,
            )
            if ok:
                await page.wait_for_timeout(400)
                stuck = await page.evaluate(
                    """(lab) => Array.from(document.querySelectorAll('select')).some(s => {
                        const o = s.selectedOptions && s.selectedOptions[0];
                        return !!(o && ((o.text || '').trim() === lab || s.value === lab));
                    })""",
                    label,
                )
                if stuck:
                    return
        except Exception as e:
            last_err = e
        await page.wait_for_timeout(400)
    raise RuntimeError(f"Could not select '{label}': {last_err}")


# ── Upload ────────────────────────────────────────────────────────────────────

async def upload_question(page, q: dict, dry_run: bool):
    print(f"→ Uploading {q['label']} (answer={q['answer']}) …", end=" ", flush=True)

    if dry_run:
        print("[DRY RUN]")
        return

    await page.goto(ADD_URL)
    await page.wait_for_load_state("networkidle")

    # ── Dropdowns ─────────────────────────────────────────────────────────────
    await page.wait_for_selector("select")
    await _select_option_js(page, TEST_TYPE)
    await _select_option_js(page, SUBJECT)
    await _select_option_js(page, SITE_TOPIC)
    await page.wait_for_timeout(300)

    # ── Difficulty ────────────────────────────────────────────────────────────
    diff_radio = await page.query_selector(f'input[type="radio"][value="{DIFFICULTY}"]')
    if not await diff_radio.is_checked():
        await diff_radio.click()

    # ── Question Title (per-question: practice vs regular) ────────────────────
    await page.fill('input[name="questionTitle"]', q["title"])

    # ── Question content ──────────────────────────────────────────────────────
    if CONTENT_TYPE == "image":
        await _inject_file(page, 0, q["images"][0])
        try:
            await page.wait_for_selector('button:has-text("Change Image")', timeout=15000)
        except Exception:
            raise RuntimeError("Image upload timed out — 'Change Image' never appeared")
    else:
        await _set_editor(page, 0, q["question_text"])
        await page.wait_for_timeout(200)

    # ── Answer options ────────────────────────────────────────────────────────
    if CONTENT_TYPE == "image":
        for i, lbl in enumerate(LETTERS):
            await page.fill(f'input[name="answers.{i}.text"]', lbl)
        answer_index = LETTERS.index(q["answer"])
    else:
        for i, opt in enumerate(q["answer_options"]):
            await page.fill(f'input[name="answers.{i}.text"]', opt)
        answer_index = ord(q["answer"]) - ord("A")

    # ── Correct answer toggle ─────────────────────────────────────────────────
    toggle = await page.query_selector(f'button[data-testid="answer-{answer_index}-correct"]')
    if not toggle:
        raise RuntimeError(f"Correct-answer toggle not found for index {answer_index}")
    await toggle.click()
    await page.wait_for_timeout(200)

    # ── Explanation text (optional) ───────────────────────────────────────────
    if q.get("explanation"):
        expl_editor_idx = 1 if CONTENT_TYPE == "text" else 0
        await _set_editor(page, expl_editor_idx, q["explanation"])

    # ── Explanation image — topics with a worked-answer image (FR) ───────────
    if IMGS_PER_Q == 2 and len(q["images"]) >= 2:
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(500)
        await _inject_file(page, 1, q["images"][1])
        async with page.expect_response(
            lambda r: "upload-image" in r.url and r.status == 200, timeout=15000
        ) as response_info:
            pass
        await response_info.value
        await page.wait_for_timeout(2000)

    # ── Submit ────────────────────────────────────────────────────────────────
    await page.locator('button:has-text("Finish")').click()
    try:
        await page.wait_for_function(
            "document.body.innerText.includes('created successfully')"
            " || !window.location.href.includes('/add')",
            timeout=10000
        )
        print("✓")
    except Exception:
        ss = os.path.join(IMG_DIR, f"error_{q['label']}.png")
        await page.screenshot(path=ss)
        raise RuntimeError(f"Submit failed — screenshot at {ss}")


# ── Main ──────────────────────────────────────────────────────────────────────

async def main(args):
    print(f"Parsing {os.path.basename(DOCX_PATH)} …")
    practice, regular, instructions = parse_docx(DOCX_PATH)

    if not preflight_check(practice, regular, instructions):
        sys.exit(1)

    questions = ([] if args.skip_practice else list(practice)) + list(regular)
    if args.start or args.end:
        lo = args.start or 1
        hi = args.end or (regular[-1]["num"] if regular else 0)
        questions = [q for q in questions if q["is_practice"] or lo <= q["num"] <= hi]
        print(f"\nFiltered to Q{lo}–Q{hi}"
              f"{' (practice included)' if not args.skip_practice else ''}: "
              f"{len(questions)} questions\n")

    if args.dry_run:
        print("\n[DRY RUN — no uploads will happen]\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        async def on_response(resp):
            if "upload-image" in resp.url:
                print(f"    [net] upload-image → {resp.status}", flush=True)
        page.on("response", on_response)

        await login(page)

        failed = []
        for i, q in enumerate(questions):
            try:
                await upload_question(page, q, args.dry_run)
            except Exception as e:
                print(f"✗  ({e})")
                failed.append(q["label"])

            if i < len(questions) - 1 and DELAY_MIN > 0:
                delay = random.randint(DELAY_MIN, DELAY_MAX)
                print(f"   ⏳ waiting {delay}s before next question …", flush=True)
                await asyncio.sleep(delay)

        await browser.close()

    print(f"\n{'='*50}")
    ok = len(questions) - len(failed)
    print(f"Done: {ok}/{len(questions)} uploaded successfully")
    if failed:
        print(f"Failed: {', '.join(failed)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    parser.add_argument("--skip-practice", action="store_true",
                        help="do not upload the 2 practice questions")
    asyncio.run(main(parser.parse_args()))
