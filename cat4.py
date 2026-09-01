#!/usr/bin/env python3
"""
Upload CAT4 questions to admin.11plusonline.co.uk
Usage: python3 upload_questions.py [--dry-run] [--start N] [--end N]
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
from playwright.async_api import async_playwright

# ── User config — only change these two lines per paper ──────────────────────
TOPIC     = "Figure Analysis"  # must match a key in CAT4_MAP below
DOCX_PATH = r"C:\Users\angel\Downloads\CAT 4\CAT 4 SET 3\Figure Analysis.docx"
print(DOCX_PATH)
print(os.path.exists(DOCX_PATH))
# ─────────────────────────────────────────────────────────────────────────────

# ── Delay between uploads (seconds) — set both to 0 to disable ───────────────
DELAY_MIN = 60   # 3 minutes
DELAY_MAX = 100   # 5 minutes
# ─────────────────────────────────────────────────────────────────────────────

LOGIN_URL = "https://admin.11plusonline.co.uk/login"
ADD_URL   = "https://admin.11plusonline.co.uk/question-bank/question/add"
EMAIL     = "admin1@11plusonline.co.uk"
PASSWORD  = "password1"
IMG_DIR   = "/tmp/fig_upload_images"

# ── CAT4 map: TOPIC → (subject, site_topic, expected_count, content_type, imgs_per_q)
#
# content_type "image" : question is an image file; answer inputs get A/B/C/D/E labels
# content_type "text"  : question is typed into a rich-text editor; answer inputs get
#                        the actual option values (numbers / words) from the docx
# imgs_per_q           : images extracted per question from the docx
#                        0 = text-based (no images), 1 = single question image,
#                        2 = question image + explanation image (Figure Recognition)
# site_topic           : the exact label in the site's Topic dropdown — may differ from
#                        the TOPIC key (e.g. FR appears as "Figure Classification" under
#                        Spatial Reasoning on this site)
CAT4_MAP = {
    "Figure Classification": ("Non Verbal Reasoning",   "Figure Classification", 24, "image", 1),
    "Figure Matrices":       ("Non Verbal Reasoning",   "Figure Matrices",       24, "image", 1),
    "Verbal Classification": ("Verbal Reasoning",       "Verbal Classification", 24, "text",  0),
    "Verbal Analogies":      ("Verbal Reasoning",       "Verbal Analogies",      24, "text",  0),
    "Number Analogies":      ("Quantitative Reasoning", "Number Analogies",      18, "text",  0),
    "Number Series":         ("Quantitative Reasoning", "Number Series",         18, "text",  0),
    "Figure Analysis":       ("Spatial Reasoning",      "Figure Analysis",       18, "image", 1),
    # Site labels Figure Recognition as "Figure Classification" under Spatial Reasoning
    "Figure Recognition":    ("Spatial Reasoning",      "Figure Recognition", 18, "image", 2),
}

if TOPIC not in CAT4_MAP:
    print(f"ERROR: Unknown topic '{TOPIC}'. Valid topics: {list(CAT4_MAP)}")
    sys.exit(1)

SUBJECT, SITE_TOPIC, EXPECTED_COUNT, CONTENT_TYPE, IMGS_PER_Q = CAT4_MAP[TOPIC]
QUESTION_TITLE = TOPIC
TEST_TYPE      = "CAT4"
DIFFICULTY     = "easy"


# ── Parsing ───────────────────────────────────────────────────────────────────

def parse_docx(path: str) -> list[dict]:
    doc = docx.Document(path)
    return _parse_image_docx(doc) if CONTENT_TYPE == "image" else _parse_text_docx(doc)


def _parse_image_docx(doc) -> list[dict]:
    """FC, FM: table [num | letter | explanation].  FA, FR: inline 'Answer: X'."""
    # Try table format first
    for table in doc.tables:
        candidates = []
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if (len(cells) >= 2
                    and re.match(r'^\d+$', cells[0])
                    and re.match(r'^[A-E]$', cells[1], re.IGNORECASE)):
                expl = cells[2] if len(cells) >= 3 else ""
                candidates.append({
                    "num": int(cells[0]),
                    "answer": cells[1].upper(),
                    "answer_options": None,
                    "question_text": None,
                    "explanation": expl,
                })
        if candidates:
            return sorted(candidates, key=lambda x: x["num"])

    # Inline format: numbered paragraphs with "Answer: X" and optional "Explanation:" block
    questions = []
    pending = None  # {"num": int, "answer": str, "expl_lines": []}
    for para in doc.paragraphs:
        t = para.text.strip()
        num_m = re.match(r'^(\d+)\.?$', t)
        ans_m = re.match(r'^Answer:\s*([A-E])$', t, re.IGNORECASE)
        expl_hdr = re.match(r'^Explanation:?$', t, re.IGNORECASE)

        if num_m:
            if pending and pending["answer"]:
                questions.append(_make_image_q(pending))
            pending = {"num": int(num_m.group(1)), "answer": None,
                       "expl_lines": [], "in_expl": False}
        elif pending:
            if ans_m:
                pending["answer"] = ans_m.group(1).upper()
            elif expl_hdr:
                pending["in_expl"] = True
            elif pending["in_expl"] and t:
                pending["expl_lines"].append(t)

    if pending and pending["answer"]:
        questions.append(_make_image_q(pending))

    if not questions:
        # Fallback: docx has no "Answer:" lines — answer key is "N L" pairs
        # e.g. "1 B", "2 C", … after the numbered question paragraphs.
        nums, answer_key = [], {}
        for para in doc.paragraphs:
            t = para.text.strip()
            num_m = re.match(r'^(\d+)\.?$', t)
            ans_m = re.match(r'^(\d+)[.)]?\s+([A-E])$', t, re.IGNORECASE)
            if num_m:
                nums.append(int(num_m.group(1)))
            elif ans_m:
                answer_key[int(ans_m.group(1))] = ans_m.group(2).upper()
        for num in sorted(set(nums)):
            if num in answer_key:
                questions.append({
                    "num": num,
                    "answer": answer_key[num],
                    "answer_options": None,
                    "question_text": None,
                    "explanation": "",
                })

    return sorted(questions, key=lambda x: x["num"])


def _make_image_q(p: dict) -> dict:
    return {
        "num": p["num"],
        "answer": p["answer"],
        "answer_options": None,
        "question_text": None,
        "explanation": " ".join(p["expl_lines"]).strip(),
    }


def _parse_text_docx(doc) -> list[dict]:
    """NA, NS, VA, VC: each paragraph is one question with inline options;
    answer key block is a single paragraph after an ANSWERS header."""
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # Locate and parse the answer key paragraph; record its index so we
    # don't accidentally re-parse it as a question (it also starts with "Q1.")
    answer_key: dict[int, tuple[str, str]] = {}  # num → (letter, explanation)
    answers_header_idx = len(paragraphs)  # default: no header found, parse all
    for i, para in enumerate(paragraphs):
        if re.search(r'\bANSWERS?\b', para, re.IGNORECASE) and len(para) < 60:
            answers_header_idx = i
            # Answer lines may be one multi-line paragraph (joined by \n) or
            # one paragraph per line — handle both by scanning every following
            # paragraph.
            ak_text = "\n".join(paragraphs[i + 1:])
            # ak_text = paragraphs[i + 1] if i + 1 < len(paragraphs) else ""
            for line in ak_text.split('\n'):
                m = re.match(r'Q(\d+)\.\s+([a-e])\)', line, re.IGNORECASE)
                if m:
                    num = int(m.group(1))
                    letter = m.group(2).upper()
                    expl_m = re.search(r'[—–]\s*(.+)', line)
                    expl = expl_m.group(1).strip().rstrip('.') if expl_m else ""
                    answer_key[num] = (letter, expl)
            break

    # Parse only question paragraphs that appear before the ANSWERS header
    questions = []
    for para in paragraphs[:answers_header_idx]:
        m = re.match(r'Q(\d+)\.\s+(.+)', para, re.DOTALL)
        if not m:
            continue
        num = int(m.group(1))
        if num not in answer_key:
            continue
        content = m.group(2).strip()

        # Split at the first option line ("\na)" or the start of "a) ...")
        parts = re.split(r'\n(?=[a-e]\))', content, maxsplit=1, flags=re.IGNORECASE)
        question_text = parts[0].strip()
        options_raw = parts[1] if len(parts) > 1 else ""

        # Parse options — works for both inline "a) x b) y" and multiline "a) x\nb) y"
        opts = [o.strip() for o in re.split(r'[a-e]\)\s*', options_raw,
                                             flags=re.IGNORECASE) if o.strip()]

        if len(opts) != 5:
            print(f"  WARNING: Q{num} parsed {len(opts)} options (expected 5): {opts}")

        letter, explanation = answer_key[num]
        questions.append({
            "num": num,
            "answer": letter,
            "answer_options": (opts + [""] * 5)[:5],
            "question_text": question_text,
            "explanation": explanation,
        })

    return sorted(questions, key=lambda x: x["num"])


# ── Image extraction ──────────────────────────────────────────────────────────

def extract_images(docx_path: str, out_dir: str) -> list[str]:
    """Return sorted list of extracted image paths.
    For FR (IMGS_PER_Q=2) caller pairs them: img[2i], img[2i+1] per question."""
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    with zipfile.ZipFile(docx_path) as z:
        media = sorted(
            [f for f in z.namelist() if f.startswith("word/media/")],
            key=lambda x: int(re.search(r'\d+', os.path.basename(x)).group()),
        )
        for i, m in enumerate(media, start=1):
            ext = os.path.splitext(m)[1]
            dest = os.path.join(out_dir, f"img{i}{ext}")
            with z.open(m) as src, open(dest, "wb") as dst:
                dst.write(src.read())
            paths.append(dest)
    return paths


# ── Preflight ─────────────────────────────────────────────────────────────────

def preflight_check(questions: list[dict], img_paths: list[str]) -> bool:
    print("\n" + "=" * 50)
    print("PRE-FLIGHT CHECK")
    print("=" * 50)
    print(f"  Topic            : {TOPIC}")
    print(f"  Subject          : {SUBJECT}")
    print(f"  Content type     : {CONTENT_TYPE}")
    print(f"  Questions parsed : {len(questions)}")
    if IMGS_PER_Q:
        print(f"  Images extracted : {len(img_paths)}  ({IMGS_PER_Q} per question)")
    print(f"  Expected count   : {EXPECTED_COUNT}")
    print()

    errors = []
    if len(questions) != EXPECTED_COUNT:
        errors.append(f"  ✗ Expected {EXPECTED_COUNT} questions, got {len(questions)}")
    if IMGS_PER_Q and len(img_paths) != EXPECTED_COUNT * IMGS_PER_Q:
        errors.append(f"  ✗ Expected {EXPECTED_COUNT * IMGS_PER_Q} images, got {len(img_paths)}")

    if errors:
        for e in errors:
            print(e)
        print("\nAborting — fix the docx and try again.")
        return False

    print("  ✓ All counts match")
    print()
    print("Questions to upload:")
    for q in questions:
        parts = [f"Q{q['num']:>2}", f"answer={q['answer']}"]
        if IMGS_PER_Q == 2:
            parts.append(f"imgs=img{2*q['num']-1}+img{2*q['num']}")
        elif IMGS_PER_Q == 1:
            parts.append(f"img=img{q['num']}")
        if q.get("answer_options"):
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


# ── Upload ────────────────────────────────────────────────────────────────────

async def upload_question(page, q: dict, img_paths: list[str], q_idx: int, dry_run: bool):
    label = f"Q{q['num']} (answer={q['answer']})"
    print(f"→ Uploading {label} …", end=" ", flush=True)

    if dry_run:
        print("[DRY RUN]")
        return

    await page.goto(ADD_URL)
    await page.wait_for_load_state("networkidle")

    # ── Dropdowns ─────────────────────────────────────────────────────────────
    await page.wait_for_selector("select")
    selects = await page.query_selector_all("select")
    await selects[0].select_option(TEST_TYPE)

    await page.wait_for_function("document.querySelectorAll('select').length >= 3")
    selects = await page.query_selector_all("select")
    await selects[1].select_option(SUBJECT)

    await page.wait_for_function(
        f"Array.from(document.querySelectorAll('select')).some("
        f"  s => Array.from(s.options).some(o => o.text === '{SITE_TOPIC}')"
        f")"
    )
    selects = await page.query_selector_all("select")
    await selects[2].select_option(SITE_TOPIC)
    await page.wait_for_timeout(300)

    # ── Difficulty ────────────────────────────────────────────────────────────
    diff_radio = await page.query_selector(f'input[type="radio"][value="{DIFFICULTY}"]')
    if not await diff_radio.is_checked():
        await diff_radio.click()

    # ── Question Title ────────────────────────────────────────────────────────
    await page.fill('input[name="questionTitle"]', QUESTION_TITLE)

    # ── Question content ──────────────────────────────────────────────────────
    if CONTENT_TYPE == "image":
        # Upload question image (first file input)
        q_img = img_paths[q_idx * IMGS_PER_Q]
        await _inject_file(page, 0, q_img)
        try:
            await page.wait_for_selector('button:has-text("Change Image")', timeout=15000)
        except Exception:
            raise RuntimeError("Image upload timed out — 'Change Image' never appeared")

    else:  # text-based
        # Put the question text into the first rich-text editor
        await _set_editor(page, 0, q["question_text"])
        await page.wait_for_timeout(200)

    # ── Answer options ────────────────────────────────────────────────────────
    if CONTENT_TYPE == "image":
        # Fill A/B/C/D/E labels (options are shown inside the question image)
        for i, lbl in enumerate(["A", "B", "C", "D", "E"]):
            await page.fill(f'input[name="answers.{i}.text"]', lbl)
        answer_index = ["A", "B", "C", "D", "E"].index(q["answer"])
    else:
        # Fill the actual option values parsed from the docx
        for i, opt in enumerate(q["answer_options"]):
            await page.fill(f'input[name="answers.{i}.text"]', opt)
        answer_index = ord(q["answer"]) - ord('A')

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

    # ── Explanation image — for FR (IMGS_PER_Q=2) ────────────────────────────
    if IMGS_PER_Q == 2:
        expl_img = img_paths[q_idx * 2 + 1]
        
        # Scroll to bottom to ensure explanation section is rendered
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(500)
        
        # Inject the explanation image into the second file input
        await _inject_file(page, 1, expl_img)
        
        # Wait for the upload request to complete
        async with page.expect_response(
            lambda r: "upload-image" in r.url and r.status == 200,
            timeout=15000
        ) as response_info:
            pass
        await response_info.value  # Ensure response is fully received
        
        # Extra buffer to let the server associate the image before Finish
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
        ss = f"/tmp/error_q{q['num']}.png"
        await page.screenshot(path=ss)
        raise RuntimeError(f"Submit failed — screenshot at {ss}")


# ── Main ──────────────────────────────────────────────────────────────────────

async def main(args):
    print(f"Parsing {os.path.basename(DOCX_PATH)} …")
    questions = parse_docx(DOCX_PATH)

    img_paths: list[str] = []
    if IMGS_PER_Q:
        print("Extracting images …")
        img_paths = extract_images(DOCX_PATH, IMG_DIR)

    if not preflight_check(questions, img_paths):
        sys.exit(1)

    upload_qs = questions
    if args.start or args.end:
        lo = args.start or 1
        hi = args.end or len(questions)
        upload_qs = [q for q in questions if lo <= q["num"] <= hi]
        print(f"\nFiltered to Q{lo}–Q{hi}: {len(upload_qs)} questions\n")

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
        for i, q in enumerate(upload_qs):
            q_idx = q["num"] - 1
            try:
                await upload_question(page, q, img_paths, q_idx, args.dry_run)
            except Exception as e:
                print(f"✗  ({e})")
                failed.append(q["num"])

            if i < len(upload_qs) - 1 and DELAY_MIN > 0:
                delay = random.randint(DELAY_MIN, DELAY_MAX)
                print(f"   ⏳ waiting {delay}s before next question …", flush=True)
                await asyncio.sleep(delay)

        await browser.close()

    print(f"\n{'='*50}")
    ok = len(upload_qs) - len(failed)
    print(f"Done: {ok}/{len(upload_qs)} uploaded successfully")
    if failed:
        print(f"Failed: Q{', Q'.join(map(str, failed))}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    asyncio.run(main(parser.parse_args()))