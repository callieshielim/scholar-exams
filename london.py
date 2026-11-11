#!/usr/bin/env python3
"""
Upload London Consortium papers to admin.11plusonline.co.uk

One script, five subjects (site "Test type" = London Consortium):

    topic       site "Subject" label            layout used on the site
    ─────────   ─────────────────────────────   ───────────────────────────────────
    maths       Maths                           text (or image) + 5 text options
    nvr         Non Verbal Reasoning            stem image + 5 option-image crops (per-question
                                                 fallback: skipped for manual entry if the
                                                 composite can't be auto-split — see parse_nvr)
    english_vr  English & Verbal Reasoning      text + 5 text options (1–2 correct)
    creative    Creative Comprehension          image (passage) + boxed text + 5 opts
    puzzle      Puzzle & Problem Solving        instruction text + Geogebra + direct answer

Notes
  * English + Verbal Reasoning live in ONE combined docx and one site subject.
    `TOPIC = "english_vr"` uploads BOTH sections in a single run (36 Qs total):
    English Q1-20 followed by Verbal Reasoning, renumbered Q21-36 so every
    question in the run has a unique number.
  * Passage handling:
      - creative : the passage is rendered to a PNG and attached as the question
        image on every question (the CC layout is "Image, boxed text, ...").
      - english  : the "English & Verbal Reasoning" layout has 3 editors
        (0=Passage Text, 1=Question max 100 chars, 2=Explanation), so the
        comprehension passage goes in editor 0 and the stem alone in editor 1.
  * VR (and some layouts) allow 1–2 correct answers; multi-answer keys
    ("A & C" or "A and C") are handled.
  * puzzle uploads auto-submit once the GeoGebra Material ID and Correct
    Answer are both present; a puzzle missing either is left unsubmitted
    (screenshot saved) and reported in the Failed list.

Usage: edit TOPIC and DOCX_PATH below, then run:
    python london.py [--dry-run] [--no-submit] [--start N] [--end N]
"""

import argparse
import asyncio
import base64
import html
import io
import os
import random
import re
import sys
import urllib.request
import zipfile

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from playwright.async_api import async_playwright

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    Image = None

try:
    import numpy as np
except ImportError:
    np = None

# ── USER CONFIG — only change these two lines per paper ──────────────────────
TOPIC     = "puzzle"  # must be a key in SUBJECTS above: maths, nvr, english_vr, creative, puzzle
DOCX_PATH = r"C:\Users\angel\Downloads\London Puzzles\SET_2_FINAL_ANSWER_GUIDE.docx"

# Delay between uploads (seconds) — set both to 0 to disable
DELAY_MIN = 45
DELAY_MAX = 90

LOGIN_URL = "https://admin.11plusonline.co.uk/login"
ADD_URL   = "https://admin.11plusonline.co.uk/question-bank/question/add"
EMAIL     = "admin1@11plusonline.co.uk"
PASSWORD  = "password1"
SHOT_DIR  = os.path.join(os.environ.get("TEMP", "/tmp"), "london_shots")

TEST_TYPE  = "London Consortium"
DIFFICULTY = "easy"

LETTERS = ["A", "B", "C", "D", "E"]
BLIP  = qn("a:blip")
EMBED = qn("r:embed")

# ── Subject registry ─────────────────────────────────────────────────────────
# folder     : sub-folder inside the SET directory holding the docx
# site_subject: exact label in the site's Subject dropdown
# layout     : substring of the layout-card text to click (None -> single/auto)
# expected   : sanity-check question count
SUBJECTS = {
    "maths": {
        "folder": "MATHS",
        "site_subject": "Maths",
        "title": "Maths",
        "layout_text":  "Question text, 5 text-based answer options",
        "layout_image": "Question text, image, 5 text-based answer options",
        "expected": 20,
        "kind": "maths",
    },
    "nvr": {
        "folder": "NVR",
        "site_subject": "Non Verbal Reasoning",
        "title": "Non-Verbal Reasoning",
        "layout_text":  "Question text, 5 text-based answer options",
        "layout_image": "Question text, image, 5 text-based answer options",
        # Used when split_nvr_image() successfully separates the composite
        # into a stem crop + 5 option crops (see parse_nvr / upload_question).
        "layout_image_split": "Question text, image, 5 image-based answer options",
        # Stem-less splits (odd one out): no question image, just 5 option
        # images (file inputs 0-4 = Answer Image A-E).
        "layout_options_only": "Question text, 5 image-based answer options",
        "expected": 20,
        "kind": "nvr",
    },
    "english_vr": {
        "folder": "ENGLISH&VRCOMBINED",
        "site_subject": "English & Verbal Reasoning",
        "title": "English & Verbal Reasoning",
        # Two different site layouts are needed here — they are NOT interchangeable:
        #   layout_text    ("... 1 or 2 correct answers")  -> 2 editors: Question, Explanation.
        #                  No Passage Text field at all. Used for every question WITHOUT a
        #                  reading passage (Q9-20 Spelling/Punctuation/Vocabulary, all of VR).
        #   layout_passage ("... max 2 correct")           -> 3 editors: Passage Text, Question,
        #                  Explanation. Passage Text is REQUIRED by the site's own validation
        #                  ("Passage text is required"), so this layout must only be used for
        #                  questions that actually carry passage_text (Q1-8 comprehension).
        # Picking the wrong one either leaves a mandatory Passage field empty (submit silently
        # fails / times out) or has nowhere to put a real passage.
        "layout_text": "5 text-based answer options (1 or 2 correct answers)",
        "layout_passage": "5 text-based answer options (max 2 correct)",
        "expected": 36,
        "kind": "english_vr",
    },
    "creative": {
        "folder": "CREATIVE COMPREHENSION",
        "site_subject": "Creative Comprehension",
        "title": "Creative Comprehension",
        "layout_text": "Image, boxed text, 5 text-based answer options",
        "expected": 16,
        "kind": "creative",
    },
    "puzzle": {
        "folder": "PUZZLES",
        "site_subject": "Puzzle & Problem Solving",
        "title": "Puzzle & Problem Solving",
        "layout_text": "Question text, instruction text, Geogebra script, direct answer input",
        "expected": 10,
        "kind": "puzzle",
    },
}

if TOPIC not in SUBJECTS:
    print(f"ERROR: Unknown topic '{TOPIC}'. Valid topics: {list(SUBJECTS)}")
    sys.exit(1)

CFG = SUBJECTS[TOPIC]

# NOTE: scoped per-TOPIC (not a shared constant) — every docx numbers its media
# image1.png, image2.png, ... so if two topics' uploads ever run concurrently
# (e.g. two terminals, to parallelize the slow DELAY_MIN/MAX wait between
# questions), a shared folder means the second run's extract_media() wipe+
# repopulate silently overwrites the first run's in-flight image files with a
# DIFFERENT topic's content at the same filename — corrupting whichever
# question is uploaded next in the first run. Scoping by topic makes that
# impossible: concurrent runs for different topics simply can't collide.
IMG_DIR = os.path.join(
    os.environ.get("TEMP", "/tmp"), "london_upload_images_" + re.sub(r"[^A-Za-z0-9]+", "_", TOPIC)
)


# ── Image helpers ─────────────────────────────────────────────────────────────

def extract_media(doc, docx_path, out_dir):
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


def para_images(par, rid2path):
    out = []
    for blip in par._p.findall(".//" + BLIP):
        rid = blip.get(EMBED)
        if rid and rid in rid2path:
            out.append(rid2path[rid])
    return out


def render_passage_image(text, out_path, width=1000, pad=40, font_size=26):
    """Render passage text to a PNG (black on white, word-wrapped)."""
    if Image is None:
        raise RuntimeError("Pillow is not installed — run: pip install pillow")
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
        bold = ImageFont.truetype("arialbd.ttf", font_size + 4)
    except Exception:
        font = ImageFont.load_default()
        bold = font

    max_w = width - 2 * pad
    dummy = Image.new("RGB", (10, 10))
    dd = ImageDraw.Draw(dummy)

    def wrap(line, fnt):
        if not line.strip():
            return [""]
        words, cur, out = line.split(), "", []
        for w in words:
            trial = (cur + " " + w).strip()
            if dd.textlength(trial, font=fnt) <= max_w:
                cur = trial
            else:
                if cur:
                    out.append(cur)
                cur = w
        if cur:
            out.append(cur)
        return out

    line_h = font_size + 10
    rows = []  # (text, font)
    for i, para in enumerate(text.split("\n")):
        fnt = bold if i == 0 else font
        for wl in wrap(para, fnt):
            rows.append((wl, fnt))
        rows.append(("", font))  # paragraph gap

    height = pad * 2 + line_h * max(1, len(rows))
    img = Image.new("RGB", (width, height), "white")
    d = ImageDraw.Draw(img)
    y = pad
    for txt, fnt in rows:
        d.text((pad, y), txt, fill="black", font=fnt)
        y += line_h
    img.save(out_path)
    return out_path


# ── NVR image handling ─────────────────────────────────────────────────────────
# Composite NVR question images are split into a stem crop + 5 option crops
# (see split_nvr_image). Layouts that can't be split are flagged for manual entry.

def _runs_1d(mask):
    runs, in_run, start = [], False, 0
    for i, v in enumerate(mask):
        if v and not in_run:
            start, in_run = i, True
        elif not v and in_run:
            runs.append([start, i - 1])
            in_run = False
    if in_run:
        runs.append([start, len(mask) - 1])
    return runs


def _merge_runs(runs, gap):
    merged = []
    for r in runs:
        if merged and r[0] - merged[-1][1] < gap:
            merged[-1][1] = r[1]
        else:
            merged.append(list(r))
    return merged


_MIN_DIM = 20  # px — guards against nonsensical near-zero crops turning into
                # silent garbage or a PIL "coordinate 'right' is less than
                # 'left'" crash, across every split strategy below.


def _longest_vertical_runs(dark):
    """Per column: length of the longest continuous run of dark pixels."""
    h, w = dark.shape
    best = np.zeros(w, dtype=int)
    for x in np.where(dark.any(axis=0))[0]:
        col = dark[:, x].astype(np.int8)
        # run lengths via diff of padded column
        d = np.diff(np.concatenate(([0], col, [0])))
        starts, ends = np.where(d == 1)[0], np.where(d == -1)[0]
        best[x] = int((ends - starts).max())
    return best


def _find_vertical_divider(arr, min_frac=0.6, margin_frac=0.15):
    """Detect an explicit ruled vertical divider line (e.g. SET 2/SET 4:
    stem panels | 5 lettered option boxes), distinct from box borders.

    A divider is ONE continuous line that is both long (>= min_frac of the
    image height) and clearly longer than every other vertical line in the
    image (by >= margin_frac of the height) - box borders only span their
    own box, while the divider extends above and below the boxes. Measured
    across all 15 London sets: dividers are 0.65-0.83 of the height with
    the next-longest line at <= 0.54, while tall-box layouts without a
    divider (e.g. SET 8/9) have several equally long borders and so fail
    the uniqueness test. (An earlier "dark-pixel fraction > 0.65" rule
    missed SET 4's shorter divider on some images.)

    Returns the divider's centre x, or None if there is no divider.
    """
    h = arr.shape[0]
    best = _longest_vertical_runs(arr < 128)
    top = int(best.max()) if best.size else 0
    if top < min_frac * h:
        return None
    cols = np.where(best >= top - 3)[0]
    # the divider may be a few px wide: all qualifying columns must be adjacent
    if cols[-1] - cols[0] > 6:
        return None
    others = np.concatenate((best[:max(cols[0] - 3, 0)], best[cols[-1] + 4:]))
    if others.size and others.max() > top - margin_frac * h:
        return None
    return int(np.median(cols))


def _split_nvr_image_divider(img, arr, divider_x, pad):
    """Split using an explicit vertical divider (e.g. SET 2 "horizontal
    codes"): example panels | divider | question panel + 5 lettered options.

    The question panel (the shape with the dashed "missing code" boxes) sits
    to the RIGHT of the divider, immediately before option A, so it belongs
    to the stem — NOT to option A. The stem is therefore everything from the
    left edge up to (but excluding) option A's own box, divider included.
    Each option crop is just the ink block sitting directly above its A-E
    label (found per label center, so option spacing need not be even).
    """
    h, w = arr.shape

    # Skip past the (possibly several-px-wide) divider line itself; it must
    # not bleed into the right-hand region or it would merge every row/col.
    right_x0 = divider_x + pad
    right_arr = arr[:, right_x0:]
    if right_arr.shape[1] < _MIN_DIM:
        raise RuntimeError("NVR image: no option content right of divider line")
    ink_rows = (right_arr < 250).any(axis=1)
    row_blocks = _runs_1d(ink_rows)
    if len(row_blocks) < 2:
        raise RuntimeError("NVR image: could not separate the A-E label row (divider layout)")

    label_y0, label_y1 = row_blocks[-1]
    content_y0, content_y1 = row_blocks[0][0], row_blocks[-2][1]

    label_band = right_arr[label_y0:label_y1 + 1, :]
    label_blocks = _merge_runs(_runs_1d((label_band < 250).any(axis=0)), gap=3)
    if len(label_blocks) != 5:
        raise RuntimeError(
            f"NVR image: expected 5 answer labels (divider layout), found {len(label_blocks)}"
        )
    centers = [(b[0] + b[1]) / 2 for b in label_blocks]

    # Column ink blocks of the diagram band (right of divider, above labels).
    content_band = right_arr[content_y0:content_y1 + 1, :]
    band_runs = _runs_1d((content_band < 250).any(axis=0))
    if not band_runs:
        raise RuntimeError("NVR image: no diagram content right of divider line")
    mids = [(centers[i] + centers[i + 1]) / 2 for i in range(4)]

    def match(col_blocks):
        """Block above each label, or an error string. Rejects a block that
        is only PART of its option (other ink in the same label slot) - e.g.
        the middle letter of a text code like "AEG" (SET 5 codes)."""
        opt = []
        for i, c in enumerate(centers):
            hit = [b for b in col_blocks if b[0] <= c <= b[1]]
            if len(hit) != 1:
                return f"could not find option {LETTERS[i]}'s box above its label"
            opt.append(hit[0])
        if any(opt[i] is opt[i + 1] for i in range(4)):
            return "adjacent options merged into one block"
        for i, b in enumerate(opt):
            lo = mids[i - 1] if i > 0 else b[0]
            hi = mids[i] if i < 4 else b[1]
            # stray ink in this label's slot that isn't some other option
            # (a neighbour poking past the midpoint is fine, e.g. SET 7 Q12)
            if any(r[1] >= lo and r[0] <= hi and not any(r is o for o in opt) for r in col_blocks):
                return f"option {LETTERS[i]} is split into several pieces"
        return opt

    col_blocks = _merge_runs(band_runs, gap=3)
    opt_blocks = match(col_blocks)
    if isinstance(opt_blocks, str):
        # Multi-piece options (text codes, scattered shapes): merge pieces
        # with the SMALLEST gap that yields 5 clean options, so a nearby
        # question figure (e.g. SET 5 Q3's test shape, ~40px left of code
        # "CDG") isn't swallowed into option A.
        spacing = min(centers[i + 1] - centers[i] for i in range(4))
        for g in range(4, int(spacing * 0.45) + 1, 2):
            col_blocks = _merge_runs(band_runs, gap=g)
            opt_blocks = match(col_blocks)
            if not isinstance(opt_blocks, str):
                break
    if isinstance(opt_blocks, str):
        raise RuntimeError(f"NVR image: {opt_blocks} (divider layout)")

    # Stem = everything left of option A's box. Two variants:
    #  - SET 2 codes: a question panel (dashed boxes) sits between the
    #    divider and option A -> it belongs to the stem, divider included.
    #  - SET 4 series: option A is the first thing after the divider -> the
    #    stem is just the panels left of the divider (divider excluded).
    a_idx = col_blocks.index(opt_blocks[0])
    if a_idx == 0:
        stem_right = divider_x
        dark = arr < 128
        while stem_right > 0 and dark[:, stem_right - 1].sum() > 0.5 * h:
            stem_right -= 1  # step left off the (possibly multi-px) divider
        opt_left_limit = divider_x + 1
    else:
        prev_right = col_blocks[a_idx - 1][1]
        stem_right = right_x0 + (prev_right + opt_blocks[0][0]) // 2 + 1
        opt_left_limit = stem_right

    stem_region = arr[:, :stem_right]
    stem_ink_cols = np.where((stem_region < 250).any(axis=0))[0]
    stem_ink_rows = np.where((stem_region < 250).any(axis=1))[0]
    if len(stem_ink_cols) == 0 or len(stem_ink_rows) == 0:
        raise RuntimeError("NVR image: no stem content (divider layout)")
    stem_x0 = max(int(stem_ink_cols[0]) - pad, 0)
    stem_x1 = min(int(stem_ink_cols[-1]) + pad, stem_right)
    stem_y0 = max(int(stem_ink_rows[0]) - pad, 0)
    stem_y1 = min(int(stem_ink_rows[-1]) + pad, h)
    if stem_x1 - stem_x0 < _MIN_DIM or stem_y1 - stem_y0 < _MIN_DIM:
        raise RuntimeError("NVR image: degenerate stem crop (divider layout)")
    stem_img = img.crop((stem_x0, stem_y0, stem_x1, stem_y1))

    y0 = max(content_y0 - pad, 0)
    y1 = min(content_y1 + pad, h)
    opt_imgs = []
    for i, (bx0, bx1) in enumerate(opt_blocks):
        x0 = max(right_x0 + bx0 - pad, opt_left_limit)
        x1 = min(right_x0 + bx1 + pad + 1, w)
        if x1 - x0 < _MIN_DIM:
            raise RuntimeError(f"NVR image: computed option {LETTERS[i]} crop is degenerate (divider layout)")
        opt_imgs.append(img.crop((x0, y0, x1, y1)))

    return stem_img, opt_imgs


def _split_nvr_image_mirror(img, arr, pad):
    """Split via the mirrored-label-boundary heuristic (see split_nvr_image's
    docstring): the original strategy, used whenever no explicit divider
    line is present. Assumes stem + options sit side by side, sharing one
    continuous horizontal ink band broken only by a real vertical gap
    before the bottom "A B C D E" label row.
    """
    ink_rows = (arr < 250).any(axis=1)
    row_blocks = _runs_1d(ink_rows)
    if len(row_blocks) < 2:
        raise RuntimeError("NVR image: could not separate the A-E label row from the diagram")

    # Some paper layouts stack the stem ABOVE the options (a large vertical
    # gap between them) instead of side by side, which this left/right split
    # cannot handle correctly — silently mis-cropping the stem if left
    # unchecked. Detect that case by the size of the largest blank-row gap
    # between the non-label content blocks, and bail out so the caller can
    # fall back to a different strategy.
    content_blocks = row_blocks[:-1]
    if len(content_blocks) >= 2:
        max_gap = max(
            content_blocks[i + 1][0] - content_blocks[i][1] - 1
            for i in range(len(content_blocks) - 1)
        )
        if max_gap > max(40, 0.12 * img.height):
            raise RuntimeError(
                "NVR image: stem and options appear vertically stacked, not "
                "side by side - unsupported layout for this splitter"
            )

    label_y0, label_y1 = row_blocks[-1]
    diagram_y0, diagram_y1 = row_blocks[0][0], row_blocks[-2][1]

    label_band = arr[label_y0:label_y1 + 1, :]
    label_blocks = _merge_runs(_runs_1d((label_band < 250).any(axis=0)), gap=3)
    if len(label_blocks) != 5:
        raise RuntimeError(f"NVR image: expected 5 answer labels (A-E), found {len(label_blocks)}")
    centers = [(b[0] + b[1]) / 2 for b in label_blocks]
    spacing = min(centers[i + 1] - centers[i] for i in range(4))
    mids = [(centers[i] + centers[i + 1]) / 2 for i in range(4)]
    stem_mid = 2 * mids[0] - mids[1]

    diagram_band = arr[diagram_y0:diagram_y1 + 1, :]
    band_ink_cols = (diagram_band < 250).any(axis=0)
    if not band_ink_cols.any():
        raise RuntimeError("NVR image: could not analyse diagram content")

    # Snap every cut to the nearest fully blank column so a figure that
    # pokes past the label midpoint (e.g. a long 3D arm) is neither clipped
    # nor leaked into its neighbour's crop.
    def snap(x, name):
        s = _snap_to_blank(band_ink_cols, x, int(spacing * 0.35))
        if s is None:
            raise RuntimeError(f"NVR image: no clean gap at the {name} boundary")
        return s

    stem_right = snap(stem_mid, "stem/A")
    bounds = [snap(m, f"{LETTERS[i]}/{LETTERS[i + 1]}") for i, m in enumerate(mids)]

    y0 = max(diagram_y0 - pad, 0)
    y1 = min(diagram_y1 + pad, img.height)

    stem_cols = np.where(band_ink_cols[:stem_right])[0]
    if len(stem_cols) == 0:
        raise RuntimeError("NVR image: computed stem crop is degenerate (bad boundary detection)")
    stem_x0 = max(int(stem_cols[0]) - pad, 0)
    stem_x1 = min(int(stem_cols[-1]) + 1 + pad, stem_right)
    if stem_x1 - stem_x0 < _MIN_DIM or y1 - y0 < _MIN_DIM:
        raise RuntimeError("NVR image: computed stem crop is degenerate (bad boundary detection)")
    stem_img = img.crop((stem_x0, y0, stem_x1, y1))

    edges = [stem_right] + bounds + [img.width]
    opt_imgs = []
    for i in range(5):
        cols = np.where(band_ink_cols[edges[i]:edges[i + 1]])[0]
        if len(cols) == 0:
            raise RuntimeError(f"NVR image: no figure above label {LETTERS[i]}")
        x0 = max(edges[i] + int(cols[0]) - pad, edges[i])
        x1 = min(edges[i] + int(cols[-1]) + 1 + pad, edges[i + 1])
        if x1 - x0 < _MIN_DIM:
            raise RuntimeError(f"NVR image: computed option {LETTERS[i]} crop is degenerate")
        opt_imgs.append(img.crop((x0, y0, x1, y1)))

    return stem_img, opt_imgs


def _snap_to_blank(ink_cols, x, search):
    """Nearest column to x (within +/- search) with no ink, or None."""
    x = int(round(x))
    for d in range(search + 1):
        for c in (x - d, x + d):
            if 0 <= c < len(ink_cols) and not ink_cols[c]:
                return c
    return None


def _split_nvr_image_stacked(img, arr, pad):
    """Split the stacked layout (e.g. SET 2 "3D complete the shape"):

        row 1 :              [target shape]
        row 2 : [given piece]   [A] [B] [C] [D] [E]
        row 3 :                  A   B   C   D   E     <- labels

    The stem is the target shape PLUS the unlabelled given piece. They
    can't be cropped as one rectangle (options A/B sit between them), so a
    new stem image is composed: given piece bottom-left, target top-right,
    both at their original scale. Each option = the ink block directly above
    its label.
    """
    h, w = arr.shape
    ink = arr < 250
    row_blocks = _runs_1d(ink.any(axis=1))
    if len(row_blocks) < 3:
        raise RuntimeError("NVR image: stacked layout needs target, option and label rows")

    label_y0, label_y1 = row_blocks[-1]
    band_y0, band_y1 = row_blocks[-2]
    target_rows_end = row_blocks[-2][0]  # everything above the option band

    label_blocks = _merge_runs(_runs_1d(ink[label_y0:label_y1 + 1, :].any(axis=0)), gap=3)
    if len(label_blocks) != 5:
        raise RuntimeError(f"NVR image: expected 5 answer labels (stacked layout), found {len(label_blocks)}")
    centers = [(b[0] + b[1]) / 2 for b in label_blocks]

    col_blocks = _merge_runs(_runs_1d(ink[band_y0:band_y1 + 1, :].any(axis=0)), gap=10)
    opt_blocks = []
    for i, c in enumerate(centers):
        hit = [b for b in col_blocks if b[0] <= c <= b[1]]
        if len(hit) != 1:
            raise RuntimeError(f"NVR image: could not find option {LETTERS[i]} above its label (stacked layout)")
        opt_blocks.append(hit[0])
    if any(opt_blocks[i] is opt_blocks[i + 1] for i in range(4)):
        raise RuntimeError("NVR image: adjacent options merged into one block (stacked layout)")

    leftover = [b for b in col_blocks if not any(b is o for o in opt_blocks)]
    if not leftover or any(b[0] > opt_blocks[0][0] for b in leftover):
        raise RuntimeError("NVR image: expected exactly one given piece left of option A (stacked layout)")

    def bbox(y0, y1, x0, x1):
        sub = ink[y0:y1, x0:x1]
        rs, cs = np.where(sub.any(axis=1))[0], np.where(sub.any(axis=0))[0]
        if len(rs) == 0 or len(cs) == 0:
            return None
        return (x0 + int(cs[0]), y0 + int(rs[0]), x0 + int(cs[-1]) + 1, y0 + int(rs[-1]) + 1)

    target_box = bbox(0, target_rows_end, 0, w)
    piece_box = bbox(band_y0, band_y1 + 1, leftover[0][0], leftover[-1][1] + 1)
    if target_box is None or piece_box is None:
        raise RuntimeError("NVR image: missing target shape or given piece (stacked layout)")
    target = img.crop(target_box)
    piece = img.crop(piece_box)
    for part in (target, piece):
        if part.width < _MIN_DIM or part.height < _MIN_DIM:
            raise RuntimeError("NVR image: degenerate target/piece crop (stacked layout)")

    # Compose: piece bottom-left, target above and to its right, with a
    # compact gap (the paper's own row gap is large and wastes space).
    gap_x = max(40, piece.width // 2)
    gap_y = 30
    cw = pad + piece.width + gap_x + target.width + pad
    ch = pad + target.height + gap_y + piece.height + pad
    stem_img = Image.new(img.mode if img.mode in ("RGB", "L") else "RGB", (cw, ch), "white")
    stem_img.paste(target.convert(stem_img.mode), (pad + piece.width + gap_x, pad))
    stem_img.paste(piece.convert(stem_img.mode), (pad, pad + target.height + gap_y))

    y0 = max(band_y0 - pad, 0)
    y1 = min(band_y1 + 1 + pad, h)
    opt_imgs = []
    for i, (bx0, bx1) in enumerate(opt_blocks):
        lo = 0 if i == 0 else (opt_blocks[i - 1][1] + bx0) // 2
        hi = w if i == 4 else (bx1 + opt_blocks[i + 1][0]) // 2
        x0 = max(bx0 - pad, lo)
        x1 = min(bx1 + 1 + pad, hi)
        if x1 - x0 < _MIN_DIM:
            raise RuntimeError(f"NVR image: computed option {LETTERS[i]} crop is degenerate (stacked layout)")
        opt_imgs.append(img.crop((x0, y0, x1, y1)))

    return stem_img, opt_imgs


def _split_nvr_image_options_only(img, arr, pad):
    """Split a stem-less layout (e.g. SET 3 "odd one out"): just 5 figures
    in one row, each directly above its A-E label, nothing else.

    Returns (None, [5 option images]). Each label owns the column slot
    between the midpoints to its neighbours; a figure may be made of several
    disconnected pieces (dots around a frame, etc.) so we crop to the ink
    bounding box inside the slot rather than relying on one ink block.
    Refuses (RuntimeError) if any figure crosses a slot boundary, if the
    content is vertically stacked, or if a slot's ink isn't centred on its
    label (i.e. there's an extra unlabelled stem figure somewhere).
    """
    h, w = arr.shape
    ink = arr < 250
    row_blocks = _runs_1d(ink.any(axis=1))
    if len(row_blocks) < 2:
        raise RuntimeError("NVR image: options-only layout needs a figure row and a label row")

    label_y0, label_y1 = row_blocks[-1]
    content_blocks = row_blocks[:-1]
    if len(content_blocks) >= 2:
        max_gap = max(content_blocks[i + 1][0] - content_blocks[i][1] - 1
                      for i in range(len(content_blocks) - 1))
        if max_gap > max(40, 0.12 * h):
            raise RuntimeError("NVR image: content vertically stacked - not an options-only layout")
    band_y0, band_y1 = content_blocks[0][0], content_blocks[-1][1]

    label_blocks = _merge_runs(_runs_1d(ink[label_y0:label_y1 + 1, :].any(axis=0)), gap=3)
    if len(label_blocks) != 5:
        raise RuntimeError(f"NVR image: expected 5 answer labels (options-only), found {len(label_blocks)}")
    centers = [(b[0] + b[1]) / 2 for b in label_blocks]
    spacing = min(centers[i + 1] - centers[i] for i in range(4))

    band_ink_cols = ink[band_y0:band_y1 + 1, :].any(axis=0)
    # Slot boundaries: the blank column nearest each label midpoint.
    bounds = []
    for i in range(4):
        mid = int((centers[i] + centers[i + 1]) / 2)
        search = int(spacing * 0.25)
        blank = [x for x in range(mid - search, mid + search + 1)
                 if 0 <= x < w and not band_ink_cols[x]]
        if not blank:
            raise RuntimeError(f"NVR image: figures {LETTERS[i]}/{LETTERS[i + 1]} overlap (options-only)")
        bounds.append(min(blank, key=lambda x: abs(x - mid)))
    edges = [0] + bounds + [w]

    y0 = max(band_y0 - pad, 0)
    y1 = min(band_y1 + 1 + pad, h)
    opt_imgs = []
    for i in range(5):
        cols = np.where(band_ink_cols[edges[i]:edges[i + 1]])[0]
        if len(cols) == 0:
            raise RuntimeError(f"NVR image: no figure above label {LETTERS[i]} (options-only)")
        fx0, fx1 = edges[i] + int(cols[0]), edges[i] + int(cols[-1]) + 1
        if abs((fx0 + fx1) / 2 - centers[i]) > spacing * 0.25:
            raise RuntimeError(
                f"NVR image: figure in slot {LETTERS[i]} not centred on its label - "
                "probably has an unlabelled stem figure (options-only)"
            )
        x0 = max(fx0 - pad, edges[i])
        x1 = min(fx1 + pad, edges[i + 1])
        if x1 - x0 < _MIN_DIM:
            raise RuntimeError(f"NVR image: computed option {LETTERS[i]} crop is degenerate (options-only)")
        opt_imgs.append(img.crop((x0, y0, x1, y1)))

    return None, opt_imgs


def split_nvr_image(img, pad=8):
    """Split a composite NVR question image into (stem_image, [5 option images]).
    stem_image is None for stem-less layouts (odd one out).

    Tries these geometrically different layouts, in order:

      1. Divider layout (e.g. SET 2): an explicit ruled vertical line
         separates example panels (left) from the question panel + 5
         lettered option boxes (right); the question panel joins the stem.
         See _find_vertical_divider / _split_nvr_image_divider.
      2. Options-only (e.g. SET 3 odd one out): 5 labelled figures, no stem.
         See _split_nvr_image_options_only.
      3. Mirror layout (e.g. SET 1): no divider line; stem sits directly
         beside the 5 options, with only the bottom "A B C D E" label row
         set apart by a blank-row gap. Option boundaries come from the
         label centers, and the stem/option boundary is derived by
         mirroring the A/B boundary across label A's own center:
         stem_right = 2*bounds[0] - bounds[1]. This assumes options are
         evenly spaced (true in every non-divider paper seen so far) and
         needs no gap-width heuristics at all.

         NOTE: an earlier version picked "the widest ink gap left of option
         A" as the boundary instead. That failed whenever a stem contained
         several sub-elements (e.g. "shape -> shape : shape ->") whose
         internal gaps could coincidentally be as wide as, or wider than,
         the true final stem/options gap — which silently sliced part of
         the stem into option A's crop. The mirrored-boundary approach
         avoids gap-width comparisons entirely.
      4. Stacked layout (e.g. SET 2 3D complete the shape): target above,
         given piece + options below. See _split_nvr_image_stacked.

    Anything else raises RuntimeError so the question is flagged for manual
    entry (see parse_nvr).
    """
    if np is None:
        raise RuntimeError("numpy is not installed - run: pip install numpy")
    arr = np.array(img.convert("L"))

    divider_x = _find_vertical_divider(arr)
    if divider_x is not None:
        return _split_nvr_image_divider(img, arr, divider_x, pad)
    # Options-only goes before mirror: it rejects any image with an
    # unlabelled stem figure (slot-centring check), whereas mirror could
    # mis-crop a stem-less image whose option A is unusually wide.
    errors = []
    for name, fn in (("options-only", _split_nvr_image_options_only),
                     ("mirror", _split_nvr_image_mirror),
                     ("stacked", _split_nvr_image_stacked),
                     ("top-stem", _split_nvr_image_top_stem),
                     ("blank-gap", _split_nvr_image_blank_gap)):
        try:
            return fn(img, arr, pad)
        except RuntimeError as e:
            errors.append(f"{name}: {e}")
    raise RuntimeError("; ".join(errors))


def _split_nvr_image_top_stem(img, arr, pad):
    """Stem is a full-width row ABOVE the options (e.g. SET 5 "combining
    shapes": "shape + shape + shape = ?" on top, 5 labelled options below).

    Cut at the largest blank-row gap between content rows; everything above
    is the stem, and the part below must pass the strict options-only
    splitter (5 labels, one figure centred over each, nothing unlabelled).
    That check is what keeps this from misfiring on the stacked layout,
    where an unlabelled given piece sits in the options row.
    """
    h = arr.shape[0]
    row_blocks = _runs_1d((arr < 250).any(axis=1))
    content = row_blocks[:-1]  # last block = A-E labels
    if len(content) < 2:
        raise RuntimeError("no separate stem row above the options")
    gaps = [(content[i + 1][0] - content[i][1] - 1, i) for i in range(len(content) - 1)]
    gap, i = max(gaps)
    if gap < max(40, 0.12 * h):
        raise RuntimeError("no clear blank gap between stem row and options row")
    cut = content[i][1] + 1 + gap // 2

    _, opt_imgs = _split_nvr_image_options_only(img.crop((0, cut, img.width, h)), arr[cut:], pad)

    top = arr[:cut] < 250
    rs, cs = np.where(top.any(axis=1))[0], np.where(top.any(axis=0))[0]
    x0, y0 = max(int(cs[0]) - pad, 0), max(int(rs[0]) - pad, 0)
    x1, y1 = min(int(cs[-1]) + 1 + pad, img.width), min(int(rs[-1]) + 1 + pad, cut)
    if x1 - x0 < _MIN_DIM or y1 - y0 < _MIN_DIM:
        raise RuntimeError("degenerate stem crop (top-stem layout)")
    return img.crop((x0, y0, x1, y1)), opt_imgs


def _split_nvr_image_blank_gap(img, arr, pad):
    """Last resort for "tall stem on the left" layouts (e.g. SET 8/13 cube
    nets) where the stem reaches down beside the A-E labels, so no blank
    ROW separates the label row and the row-based splitters all fail.

    Treat a fully blank full-height COLUMN gap as a virtual divider and
    reuse the divider splitter. That splitter validates the cut itself (5
    labels + 5 options to the right, stem ink to the left), so a wrong gap
    - between two options, or left of everything - simply raises and the
    next candidate is tried. Widest gaps first.
    """
    ink_cols = (arr < 250).any(axis=0)
    gaps = [g for g in _runs_1d(~ink_cols) if g[0] > 0 and g[1] < len(ink_cols) - 1]
    gaps.sort(key=lambda g: g[1] - g[0], reverse=True)
    last = "no blank column gaps"
    for g0, g1 in gaps:
        if g1 - g0 + 1 < 2 * pad:
            break
        try:
            return _split_nvr_image_divider(img, arr, (g0 + g1) // 2, pad)
        except RuntimeError as e:
            last = str(e)
    raise RuntimeError(f"no blank column gap gives a valid stem/options split ({last})")


# ── Shared text parsing ───────────────────────────────────────────────────────

def _split_stem_options(text):
    """Split a question block ('stem\\nA. ..\\nB. ..') into (stem, [5 options])."""
    opts, order, stem_lines = {}, [], []
    in_opts = False
    for ln in text.split("\n"):
        s = ln.strip()
        om = re.match(r"^([A-E])[.\)\t]\s*(.*)$", s)
        if om and (om.group(1) == "A" or in_opts):
            in_opts = True
            letter = om.group(1)
            if letter not in opts:
                order.append(letter)
            opts[letter] = om.group(2).strip()
        elif not in_opts:
            if s and not re.fullmatch(r"[\]\[\)\(.\s]+", s):
                stem_lines.append(s)
        else:
            if s and order:
                opts[order[-1]] = (opts[order[-1]] + " " + s).strip()
    stem = "\n".join(stem_lines).strip()
    options = [opts.get(l, "") for l in LETTERS]
    return stem, options


def _is_section_header(text):
    """True for sub-section header lines like 'Q1-8 Comprehension' or
    'Q16 New Type - Word Precision' that sit between questions and must
    NOT be glued onto the previous question's last answer option."""
    return bool(re.match(r"^Q\d+(?:\s*[-–—]\s*\d+)?\s+[A-Za-z]", text.strip()))


def _collect_questions(paras, start, end):
    """Group paragraphs paras[start:end] into raw {num, text, images} blocks."""
    raw, cur = [], None
    for par in paras[start:end]:
        t = par.text
        m = re.match(r"^Q(\d+)\.\s*(.*)$", t.strip(), re.S)
        if m:
            cur = {"num": int(m.group(1)), "text": m.group(2), "images": []}
            raw.append(cur)
        elif cur is not None:
            if _is_section_header(t):
                continue
            cur["text"] += "\n" + t
    return raw


def _find(paras, pattern, start=0, maxlen=80):
    for i in range(start, len(paras)):
        t = paras[i].text.strip()
        if re.search(pattern, t, re.IGNORECASE) and len(t) < maxlen:
            return i
    return None


def _key_lines(paras_slice):
    """Flatten paragraphs into individual lines (answer keys pack several
    'Qn.' entries into one paragraph separated by newlines)."""
    lines = []
    for p in paras_slice:
        for ln in p.text.split("\n"):
            s = ln.strip()
            if s:
                lines.append(s)
    return lines


# ── Answer-key parsers (per section, formats differ) ──────────────────────────
#
# The answer letter itself may be written in many equivalent ways across
# different docx authors, e.g. "A", "a", "A)", "A.", "A:", "A -", "A –".
# `_LETTER_TAIL` matches the (optional, single) separator that can follow the
# letter so every parser below tolerates all of these forms.
_LETTER_TAIL = r"\b\s*[.):]?\s*[-–—\u2212]?\s*"


def _key_creative(lines):
    """'Q1. c) he was escaping ... – expl' -> {num: ([idx], expl)}."""
    key = {}
    for t in lines:
        m = re.match(r"^Q(\d+)[.:]\s*([A-Ea-e])" + _LETTER_TAIL, t)
        if not m:
            continue
        num = int(m.group(1))
        idx = LETTERS.index(m.group(2).upper())
        expl = _explanation_after_dash(t)
        key[num] = ([idx], expl)
    return key


def _key_english(lines):
    """'Q1. B. he is distracted ... – expl' -> {num: ([idx], expl)}."""
    key = {}
    for t in lines:
        m = re.match(r"^Q(\d+)[.:]\s*([A-Ea-e])" + _LETTER_TAIL, t)
        if not m:
            continue
        num = int(m.group(1))
        idx = LETTERS.index(m.group(2).upper())
        key[num] = ([idx], _explanation_after_dash(t))
    return key


def _key_vr(lines):
    """'Q1. A & C – ...' / 'Q1. C and E – ...' / 'Q6. C – precise'."""
    key = {}
    for t in lines:
        m = re.match(
            r"^Q(\d+)[.:]\s*([A-Ea-e])(?:\s*(?:&|and)\s*([A-Ea-e]))?" + _LETTER_TAIL, t
        )
        if not m:
            continue
        num = int(m.group(1))
        letters = [m.group(2).upper()]
        if m.group(3):
            letters.append(m.group(3).upper())
        idxs = [LETTERS.index(l) for l in letters]
        key[num] = (idxs, _explanation_after_dash(t))
    return key


def _explanation_after_dash(t):
    m = re.search(r"[—–\u2212-]\s*(.+)$", t)
    return m.group(1).strip() if m else ""


# ── Subject parsers ───────────────────────────────────────────────────────────

def _build_questions(raw, key, cfg, passage_text=None, passage_img=None):
    questions = []
    for r in raw:
        stem, options = _split_stem_options(r["text"])
        correct, expl = key.get(r["num"], (None, ""))
        needs_manual = not correct
        questions.append({
            "num": r["num"],
            "title": cfg["title"],
            "stem": stem,
            "options": options,
            "correct": correct or [0],
            "answer": ",".join(LETTERS[i] for i in (correct or [0])),
            "explanation": expl,
            "images": list(r["images"]),
            "passage_text": passage_text,
            "passage_image": passage_img,
            "needs_manual": needs_manual,
        })
    questions.sort(key=lambda q: q["num"])
    return questions


def _iter_doc_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _table_to_text(tbl):
    lines = []
    for row in tbl.rows:
        cells = [c.text.replace("\n", " ").strip() for c in row.cells]
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def parse_creative(path, cfg):
    doc = docx.Document(path)
    paras = doc.paragraphs
    q_start = _find(paras, r"^questions", maxlen=40)
    a_start = _find(paras, r"answers?\s+and\s+explanation", maxlen=60)
    if q_start is None or a_start is None:
        raise RuntimeError("Creative docx: could not locate QUESTIONS / ANSWERS sections")

    q_el = paras[q_start]._p
    parts = []
    for block in _iter_doc_blocks(doc):
        if isinstance(block, Paragraph):
            if block._p is q_el:
                break
            t = block.text.strip()
            if not t:
                continue
            if re.match(r"^(passage|stimulus)\b", t, re.I):
                continue
            if re.match(r"^creative comprehension\b", t, re.I):
                continue
            parts.append(t)
        else:
            tbl_txt = _table_to_text(block)
            if tbl_txt:
                parts.append(tbl_txt)
    passage = "\n".join(parts)
    passage_img = None
    if passage:
        os.makedirs(IMG_DIR, exist_ok=True)
        passage_img = render_passage_image(passage, os.path.join(IMG_DIR, "creative_passage.png"))

    raw = _collect_questions(paras, q_start, a_start)
    key = _key_creative(_key_lines(paras[a_start + 1:]))
    return _build_questions(raw, key, cfg, passage_img=passage_img)


def parse_english_vr(path, cfg, section):
    """Combined file -> english (Q1-20) or vr (Q1-16) questions."""
    doc = docx.Document(path)
    paras = doc.paragraphs

    vr_hdr   = _find(paras, r"london\s+consortium\s+(vr|verbal\s*reasoning)", maxlen=60)
    eng_ans  = _find(paras, r"answers?\s+and\s+explanation.*english", maxlen=80)
    if eng_ans is None:
        eng_ans = _find(paras, r"answers?\s+and\s+explanation", maxlen=80)
    if vr_hdr is None or eng_ans is None:
        raise RuntimeError("ENGVR docx: could not locate VR header / answer sections")

    # A bare "VR" / "VERBAL REASONING" line marks a section change. Papers
    # differ: some put a second "ANSWERS AND EXPLANATIONS" before the VR
    # key (SET 1), others ONLY a bare "VR" line (SET 2-4). Without this,
    # the VR key lines were read as part of the English key and silently
    # overwrote the English answers (VR "Q1. A and D" -> English Q1 = A).
    VR_LABEL = r"^(vr|verbal\s*reasoning)$"
    vr_ans_candidates = [
        i for i in (
            _find(paras, r"answers?\s+and\s+explanation", start=eng_ans + 1, maxlen=80),
            _find(paras, VR_LABEL, start=eng_ans + 1, maxlen=30),
        ) if i is not None
    ]
    vr_ans = min(vr_ans_candidates) if vr_ans_candidates else None

    # English questions end at the bare VR label that precedes the VR header
    # (if any), so it isn't glued onto the last English option.
    eng_end = vr_hdr
    vr_label_q = _find(paras, VR_LABEL, maxlen=30)
    if vr_label_q is not None and vr_label_q < vr_hdr:
        eng_end = vr_label_q

    if section == "english":
        q_start = _find(paras, r"^questions", maxlen=40) or 0
        raw = _collect_questions(paras, q_start, eng_end)
        key = _key_english(_key_lines(paras[eng_ans + 1:(vr_ans or len(paras))]))
        # passage/poem = between PASSAGE or POEM and QUESTIONS
        p_start = _find(paras, r"^(passage|poem)\b", maxlen=40)
        passage = "\n".join(
            p.text.strip() for p in paras[(p_start + 1 if p_start is not None else 0):q_start]
            if p.text.strip()
        ) if p_start is not None else None
        qs = _build_questions(raw, key, cfg, passage_text=None)
        # Passage applies only to comprehension questions (those before the first
        # subheader such as "Q9-13 Spelling"). Heuristic: Q1-8 get the passage.
        if passage:
            for q in qs:
                if q["num"] <= 8:
                    q["passage_text"] = passage
        return qs

    # section == "vr"
    raw = _collect_questions(paras, vr_hdr, eng_ans)
    key = _key_vr(_key_lines(paras[(vr_ans or eng_ans) + 1:]))
    return _build_questions(raw, key, cfg)


def parse_english_vr_combined(path, cfg):
    """Parse BOTH sections of the combined docx and merge into one run.

    English keeps its original numbering (1..N). VR is renumbered to
    continue right after English (N+1..N+M) so every question in the
    combined run has a unique number, e.g. English 1-20, VR 21-36.
    """
    eng_cfg = dict(cfg, title="English")
    vr_cfg = dict(cfg, title="Verbal Reasoning")
    eng_qs = parse_english_vr(path, eng_cfg, "english")
    vr_qs = parse_english_vr(path, vr_cfg, "vr")

    offset = max((q["num"] for q in eng_qs), default=0)
    for q in vr_qs:
        q["num"] += offset

    combined = eng_qs + vr_qs
    combined.sort(key=lambda q: q["num"])
    return combined


def _key_maths(lines):
    """Answer-key lines -> {num: ([idx], expl)}. Two layouts exist:

      SETs 1-3: 'Q1. C. 615' (answer VALUE on the Q line), explanation on
                the following line(s) with no 'Qn.' marker, e.g.
                '248 + 367 = 615.'  (SET 2 sometimes uses two lines.)
      SETs 4-5: 'Q1. A. 09:38 + 1 hour = 10:38. Adding 47 minutes gives
                11:25.' - the explanation IS the rest of the Q line.

    Rule: if continuation lines follow, they are the explanation; otherwise
    the rest of the Q line is. A short 'value - expl' form (dash with
    spaces within the first 25 chars) is still split on that dash.

    NOTE: an earlier version took "text after the FIRST dash anywhere on the
    line" - which on SET 4/5 grabbed fragments after a minus sign
    ('£64 − £16 = £48' -> '£16 = £48.') or nothing at all.
    """
    key = {}
    i, n = 0, len(lines)
    while i < n:
        t = lines[i]
        m = re.match(r"^Q(\d+)[.:]\s*([A-Ea-e])" + _LETTER_TAIL + r"(.*)$", t)
        if m:
            num = int(m.group(1))
            idx = LETTERS.index(m.group(2).upper())
            rest = m.group(3).strip()
            cont = []
            while i + 1 < n and not re.match(r"^Q\d+\b", lines[i + 1]):
                cont.append(lines[i + 1].strip())
                i += 1
            if cont:
                expl = " ".join(cont)
            else:
                dm = re.match(r"^(.{1,25}?)\s+[-–—\u2212]\s+(.+)$", rest)
                expl = dm.group(2).strip() if dm else rest
                # a bare answer value ('Q17. E. 0') is not an explanation
                if not dm and len(rest) < 15 and "=" not in rest:
                    expl = ""
            key[num] = ([idx], expl)
        i += 1
    return key


def parse_maths(path, cfg):
    doc = docx.Document(path)
    rid2path = extract_media(doc, path, IMG_DIR)
    paras = doc.paragraphs
    a_start = _find(paras, r"answers?\s+and\s+explanation", maxlen=60) or len(paras)

    raw, cur = [], None
    for par in paras[:a_start]:
        t = par.text
        imgs = para_images(par, rid2path)
        m = re.match(r"^Q(\d+)\.\s*(.*)$", t.strip(), re.S)
        if m:
            cur = {"num": int(m.group(1)), "text": m.group(2), "images": []}
            raw.append(cur)
        elif cur is not None:
            if _is_section_header(t):
                continue
            cur["text"] += "\n" + t
        if cur is not None and imgs:
            cur["images"] += imgs

    key = _key_maths(_key_lines(paras[a_start + 1:]))
    return _build_questions(raw, key, cfg)


def parse_nvr(path, cfg):
    """Parse NVR questions, one composite diagram image per question.

    Two upload strategies, tried per-question (papers mix layouts freely):

      1. "image_options" (preferred) — split the composite into a stem crop
         + 5 individual option-letter crops via split_nvr_image(), so each
         answer choice is uploaded as its own image (site layout "Question
         text, image, 5 image-based answer options"). This only works for
         the "stem left, options right, A-E label row at the very bottom"
         layout (see split_nvr_image's docstring for the geometry it
         assumes).
      2. "text_options" (fallback) — used whenever split_nvr_image() can't
         confidently split a given image (e.g. papers where the stem is
         stacked above the options instead of beside them, like SET 2's
         "3D complete the shape"). Falls back to uploading the ENTIRE composite as the stem
         image and using the plain "A"/"B"/.../"E" letters as text options,
         exactly like this parser used to work unconditionally.

    Word sometimes anchors an image to the following question-number
    paragraph, so images are assigned by document order rather than by the
    paragraph that happens to contain the image relationship.
    """
    doc = docx.Document(path)
    rid2path = extract_media(doc, path, IMG_DIR)

    question_nums = []
    image_paths = []
    key = {}
    section_of = {}
    section = ""
    in_key = False
    for p in doc.paragraphs:
        t = p.text.strip()
        imgs = para_images(p, rid2path)

        if re.match(r"^answers?\b", t, re.IGNORECASE):
            in_key = True

        m_label = re.match(r"^(\d+)\.$", t)
        if m_label and not in_key:
            question_nums.append(int(m_label.group(1)))
            section_of[int(m_label.group(1))] = section
        elif t and not in_key and not re.match(r"^\d", t) and len(t) < 60:
            section = t  # sub-section header, e.g. "ODD ONE OUT"

        m_key = re.match(r"^(\d+)[.)]?\s+([A-Ea-e])" + _LETTER_TAIL + r"(.*)$", t)
        if m_key:
            num = int(m_key.group(1))
            expl = re.sub(r"^[-–—]\s*", "", m_key.group(3).strip()).strip()
            key[num] = (m_key.group(2).upper(), expl)

        image_paths.extend(imgs)

    question_nums = sorted(set(question_nums))
    if len(image_paths) != len(question_nums):
        raise RuntimeError(
            f"NVR: found {len(question_nums)} question numbers but "
            f"{len(image_paths)} question images"
        )

    n_split, n_fallback = 0, 0
    questions = []
    for num, image_path in zip(question_nums, image_paths):
        letter, expl = key.get(num, (None, ""))
        base = {
            "num": num,
            "title": cfg["title"],
            # The site's "Question" rich-text field is required and rejects
            # empty submissions ("Question is required"), even though NVR's
            # actual content lives entirely in the image(s). Use a fixed
            # generic instruction so the field is never blank.
            "stem": "Study the image and choose the correct answer.",
            "correct": [LETTERS.index(letter)] if letter else [0],
            "answer": letter or "A",
            "explanation": expl,
            "passage_text": None,
            "passage_image": None,
            "needs_manual": letter is None,
        }

        split_result = None
        try:
            with Image.open(image_path) as im:
                stem_img, opt_imgs = split_nvr_image(im.copy())
            stem_path = None
            if stem_img is not None:
                stem_path = os.path.join(IMG_DIR, f"nvr_q{num}_stem.png")
                stem_img.save(stem_path)
            opt_paths = []
            for letter_i, opt_img in zip(LETTERS, opt_imgs):
                opt_path = os.path.join(IMG_DIR, f"nvr_q{num}_opt{letter_i}.png")
                opt_img.save(opt_path)
                opt_paths.append(opt_path)
            split_result = (stem_path, opt_paths)
        except Exception:
            split_result = None

        if split_result is not None:
            stem_path, opt_paths = split_result
            n_split += 1
            if stem_path is None:
                # Stem-less (odd one out): site layout "Question text, 5
                # image-based answer options" - no question image at all.
                is_odd = "odd one out" in section_of.get(num, "").lower()
                base["stem"] = ("Which one is the odd one out?" if is_odd
                                else "Choose the correct answer.")
            questions.append({
                **base,
                "layout_kind": "image_options",
                "images": [stem_path] if stem_path else [],
                "option_images": opt_paths,
                "options": list(LETTERS),
            })
        else:
            n_fallback += 1
            questions.append({
                **base,
                "layout_kind": "text_options",
                "images": [image_path],
                "option_images": None,
                "options": list(LETTERS),
            })

    print(f"  NVR image split: {n_split} question(s) as image-options, "
          f"{n_fallback} fell back to full-image + text options")
    return questions


GEOGEBRA_DOWNLOAD = "https://www.geogebra.org/material/download/format/file/id/{}"
PUZZLE_QUESTION_HEADS = {"THE QUESTION", "QUESTION", "FINAL TEST ANSWER", "YOUR ANSWER"}
INSTRUCTION_MAX = 100   # the site's Instruction Text box limit


def _svg_texts(svg):
    """Visible <text> strings of an SVG, in document order."""
    out = []
    for t in re.findall(r"<text[^>]*>(.*?)</text>", svg, re.S):
        t = html.unescape(re.sub(r"<[^>]+>", "", t)).strip()
        if t:
            out.append(t)
    return out


def _is_caps_label(t):
    letters = re.sub(r"[^A-Za-z]", "", t)
    return len(letters) >= 3 and letters.isupper()


def fetch_puzzle_text(material_id):
    """Pull the instruction and question wording out of a GeoGebra activity.

    The answer guides don't contain the puzzle wording - it only exists in the
    activity's background SVG (ui_*.svg inside the .ggb zip), laid out as:
        SET n · QUESTION k | Title | <instruction lines> | CAPS LABELS ...
        | THE QUESTION | <question lines> | ANSWER / Enter ... hints
    Returns (instruction, question); either may be "" if not found.
    """
    data = urllib.request.urlopen(GEOGEBRA_DOWNLOAD.format(material_id), timeout=30).read()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        ui = [n for n in z.namelist() if os.path.basename(n).startswith("ui_") and n.endswith(".svg")]
        if not ui:
            return "", ""
        texts = _svg_texts(z.read(ui[0]).decode("utf-8"))

    # instruction: lines after the title, up to the first CAPS label
    instruction = []
    for t in texts[2:]:
        if _is_caps_label(t):
            break
        instruction.append(t)

    # question: lines after the question heading, up to the answer box hints
    question, in_q = [], False
    for t in texts:
        if t.upper() in PUZZLE_QUESTION_HEADS:
            in_q, question = True, []
            continue
        if in_q:
            if _is_caps_label(t) or re.match(r"^(Enter|Type)\b", t):
                break
            question.append(t)
    return " ".join(instruction).strip(), " ".join(question).strip()


def split_instruction(instruction, question):
    """Fit whole sentences of the instruction into the Instruction box; any
    that don't fit go in front of the question so no clue is lost."""
    sentences = re.split(r"(?<=[.!?])\s+", instruction.strip())
    head = []
    while sentences and len(" ".join(head + sentences[:1])) <= INSTRUCTION_MAX:
        head.append(sentences.pop(0))
    if not head:   # one long sentence - nothing fits whole, so move it all
        return "", f"{instruction} {question}".strip()
    return " ".join(head), " ".join(sentences + [question]).strip()


def parse_puzzles(path, cfg):
    """Answer-guide layout (SET_N_FINAL_ANSWER_GUIDE.docx), either as a table
    per question (sets 2+) or plain paragraphs (set 1):

        QUESTION 1  |  Complete the triangle
        Open GeoGebra activity                 <- hyperlink to geogebra.org/m/<id>
        PORTAL ANSWER: 15  —  Enter 15 as a two-digit number.
        Explanation: ...

    The puzzle wording isn't in the guide, so it's fetched from each linked
    GeoGebra activity. Falls back to the older 'Question N: Title' layout.
    """
    doc = docx.Document(path)
    rels = doc.part.rels
    HYPERLINK, RID, P = qn("w:hyperlink"), qn("r:id"), qn("w:p")

    raw, cur = [], None
    for p in doc.element.body.iter(P):   # body paragraphs and table-cell paragraphs, in order
        t = Paragraph(p, doc).text.strip()
        m = re.match(r"^question\s+(\d+)\s*\|\s*(.+)$", t, re.I)
        if m:
            cur = {"num": int(m.group(1)), "title": m.group(2).strip(), "material_id": None,
                   "correct_answer": None, "explanation": ""}
            raw.append(cur)
            continue
        if cur is None:
            continue
        for link in p.iter(HYPERLINK):
            rid = link.get(RID)
            lm = rid in rels and re.search(r"geogebra\.org/m/([A-Za-z0-9]+)", rels[rid].target_ref)
            if lm and not cur["material_id"]:
                cur["material_id"] = lm.group(1)
        am = re.match(r"^portal answer:\s*(-?[\d,]+(?:\.\d+)?)", t, re.I)
        if am:
            cur["correct_answer"] = am.group(1).replace(",", "")
        em = re.match(r"^explanation:\s*(.+)$", t, re.I | re.S)
        if em:
            cur["explanation"] = em.group(1).strip()

    if not raw:
        return _parse_puzzles_legacy(path, cfg)

    questions = []
    for r in raw:
        instruction = question = ""
        if r["material_id"]:
            try:
                instruction, question = fetch_puzzle_text(r["material_id"])
            except Exception as e:
                print(f"  ! Q{r['num']}: could not read GeoGebra {r['material_id']}: {e}")
        questions.append({
            **r,
            "instruction": instruction,
            "question": question,
            "needs_manual": not (r["material_id"] and r["correct_answer"] and instruction and question),
        })
    return questions


def _parse_puzzles_legacy(path, cfg):
    """PUZZLES docx layout: each question is a 'Question N: <Title>' block
    (repeated: once as a bare heading, once with the instruction sentence and
    an 'Open Question N in GeoGebra' hyperlink), followed by a free-form
    'Website Embedding Information' section describing how the interactive
    GeoGebra construction should be checked.

    We extract: title, instruction text, and the linked GeoGebra Material ID
    (from the docx hyperlink relationship). We also make a best-effort guess
    at a scalar "correct answer" by looking for phrases like 'must total 15'
    or 'must equal 45' in the embedding-info text — this is NOT reliable for
    puzzles whose check is positional/multi-condition (no single number), so
    those are left for manual entry (needs_manual=True).

    NOTE: Word equation objects (e.g. the actual arithmetic in Q5/Q6) are not
    extracted as text by python-docx, so instruction text for those may be
    missing symbols/numbers — always review before submitting.
    """
    doc = docx.Document(path)
    rels = doc.part.rels
    HYPERLINK = qn("w:hyperlink")
    RID = qn("r:id")

    para_material = {}
    for i, p in enumerate(doc.paragraphs):
        for link in p._p.findall(".//" + HYPERLINK):
            rid = link.get(RID)
            if rid and rid in rels:
                m = re.search(r"geogebra\.org/m/([A-Za-z0-9]+)", rels[rid].target_ref)
                if m:
                    para_material[i] = m.group(1)
                    break

    raw, cur_num = {}, None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        m = re.match(r"^Question (\d+):\s*(.+)$", t, re.S)
        if m:
            num = int(m.group(1))
            lines = [ln.strip() for ln in t.split("\n") if ln.strip()]
            title_line = re.sub(r"^Question \d+:\s*", "", lines[0]).strip()
            title = re.sub(r"\s*[-–—]\s*GeoGebra$", "", title_line, flags=re.I).strip()
            instruction_lines = [
                ln for ln in lines[1:]
                if not re.search(r"open.*geogebra", ln, re.I) and not re.fullmatch(r"\[.*\]", ln)
            ]
            instruction = " ".join(instruction_lines).strip()
            if num not in raw:
                raw[num] = {"title": title, "instruction": instruction, "material_id": None, "info_text": ""}
            elif instruction and not raw[num]["instruction"]:
                raw[num]["instruction"] = instruction
            cur_num = num
        if i in para_material and cur_num is not None and not raw[cur_num]["material_id"]:
            raw[cur_num]["material_id"] = para_material[i]
        if cur_num is not None and t and not re.match(r"^Question (\d+):", t):
            raw[cur_num]["info_text"] += " " + t

    questions = []
    for num in sorted(raw):
        r = raw[num]
        m = re.search(r"must\s+(?:total|equal)\s+(-?\d+(?:\.\d+)?)", r["info_text"], re.I)
        guessed = m.group(1) if m else None
        questions.append({
            "num": num,
            "title": r["title"],
            "instruction": r["instruction"],
            "material_id": r["material_id"],
            "correct_answer": guessed,
            "needs_manual": r["material_id"] is None or guessed is None,
        })
    return questions


def parse_subject(subject, path):
    cfg = SUBJECTS[subject]
    kind = cfg["kind"]
    if kind == "creative":
        return parse_creative(path, cfg)
    if kind == "puzzle":
        return parse_puzzles(path, cfg)
    if kind == "english_vr":
        return parse_english_vr_combined(path, cfg)
    if kind == "maths":
        return parse_maths(path, cfg)
    if kind == "nvr":
        return parse_nvr(path, cfg)
    raise RuntimeError(f"Unknown subject kind: {kind}")


# ── Preflight ─────────────────────────────────────────────────────────────────

def preflight_check(questions, subject, cfg):
    print("\n" + "=" * 62)
    print("PRE-FLIGHT CHECK")
    print("=" * 62)
    print(f"  Subject (site)   : {cfg['site_subject']}")
    print(f"  Section          : {subject}")
    print(f"  Layout           : {cfg.get('layout_text')}")
    print(f"  Questions parsed : {len(questions)}  (expected {cfg['expected']})")
    n_pass_img = sum(1 for q in questions if q.get("passage_image"))
    n_pass_txt = sum(1 for q in questions if q.get("passage_text"))
    if n_pass_img:
        print(f"  Passage image on : {n_pass_img} question(s)")
    if n_pass_txt:
        print(f"  Passage text on  : {n_pass_txt} question(s)")
    if is_nvr := (cfg.get("kind") == "nvr"):
        n_img_opts = sum(1 for q in questions if q.get("layout_kind") == "image_options")
        n_manual = len(questions) - n_img_opts
        print(f"  NVR layout split : {n_img_opts} auto (image-options), {n_manual} NEED MANUAL ENTRY (could not split image; no site fallback layout exists)")
    print()

    errors, warnings = [], []
    is_puzzle = cfg.get("kind") == "puzzle"
    if len(questions) != cfg["expected"]:
        warnings.append(f"  ! Expected {cfg['expected']} questions, got {len(questions)}")
    for q in questions:
        if is_nvr:
            if not q.get("images") and q.get("layout_kind") != "image_options":
                errors.append(f"  x Q{q['num']} missing stem image")
            if q.get("options") != LETTERS:
                errors.append(f"  x Q{q['num']} has invalid text options: {q.get('options')}")
            if q.get("layout_kind") == "image_options" and len(q.get("option_images") or []) != 5:
                errors.append(f"  x Q{q['num']} image-options layout has {len(q.get('option_images') or [])} option images (expected 5)")
            elif q.get("layout_kind") != "image_options":
                warnings.append(f"  ! Q{q['num']} could not split composite image - will be SKIPPED (needs manual entry)")
        elif is_puzzle:
            if not q.get("material_id"):
                errors.append(f"  x Q{q['num']} missing GeoGebra Material ID")
            if not q.get("correct_answer"):
                warnings.append(f"  ! Q{q['num']} correct answer could not be inferred - enter it manually before submitting")
            if "question" in q and not (q.get("instruction") and q.get("question")):
                warnings.append(f"  ! Q{q['num']} instruction/question text not found in the GeoGebra activity - type it in the browser")
        else:
            if not (q["stem"] or "").strip():
                errors.append(f"  x Q{q['num']} empty question text")
            filled = [o for o in q["options"] if o.strip()]
            if len(filled) != 5:
                errors.append(f"  x Q{q['num']} has {len(filled)} options (expected 5): {q['options']}")
        if q["needs_manual"] and not is_puzzle:
            warnings.append(f"  ! Q{q['num']} answer missing from key - defaulted to A (fix manually)")

    if is_puzzle and any("question" not in q for q in questions):   # legacy layout only
        warnings.append(
            "  ! Puzzle instructions may be missing symbols/numbers that were embedded as Word\n"
            "    equation objects (not plain text) - review each question in the browser before\n"
            "    submitting, and always verify 'Correct Answer' yourself."
        )

    if warnings:
        print("\n".join(warnings) + "\n")
    if errors:
        print("\n".join(errors))
        print("\nAborting - fix the docx and try again.")
        return False

    if any("Expected" in w for w in warnings):
        print("  ! COUNT MISMATCH - check the warnings above before proceeding\n")
    else:
        print("  OK: counts look right\n")
    print("Questions to upload:")
    for q in questions:
        extra = ""
        if q.get("passage_image"):
            extra = " +passageIMG"
        elif q.get("passage_text"):
            extra = " +passageTXT"
        expl = "  expl=y" if q.get("explanation") else "  expl=n"
        if is_nvr:
            layout_tag = "auto" if q.get("layout_kind") == "image_options" else "MANUAL"
            img_name = os.path.basename(q["images"][0]) if q.get("images") else "(none - options only)"
            print(f"    Q{q['num']:>2} ans={q['answer']:<3}{extra}  layout={layout_tag:<8} image={img_name}{expl}")
        elif is_puzzle:
            ans = q.get("correct_answer") or "??? (MANUAL)"
            print(f"    Q{q['num']:>2} title={q['title']!r} materialId={q['material_id']} ans={ans}{expl}")
            if "question" in q:
                instruction, question = split_instruction(q["instruction"], q["question"])
                print(f"          instruction: {instruction}")
                print(f"          question   : {question}")
        else:
            print(f"    Q{q['num']:>2} ans={q['answer']:<3}{extra}  opts={q['options']}{expl}")
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
    print("Logged in\n")


async def select_option_js(page, label, retries=8):
    """Set a <select> by option text purely in JS (survives React remounts)."""
    last_err = None
    for _ in range(retries):
        try:
            await page.wait_for_function(
                """(lab) => Array.from(document.querySelectorAll('select')).some(
                    s => Array.from(s.options).some(
                        o => (o.text||'').trim() === lab || o.value === lab))""",
                arg=label, timeout=15000,
            )
            ok = await page.evaluate(
                """(lab) => {
                    const s = Array.from(document.querySelectorAll('select')).find(sel =>
                        Array.from(sel.options).some(o => (o.text||'').trim()===lab || o.value===lab));
                    if (!s) return false;
                    const opt = Array.from(s.options).find(o => (o.text||'').trim()===lab || o.value===lab);
                    const prev = s.value;
                    const desc = Object.getOwnPropertyDescriptor(window.HTMLSelectElement.prototype, 'value');
                    desc.set.call(s, opt.value);
                    if (s._valueTracker) s._valueTracker.setValue(prev);
                    opt.selected = true;
                    s.dispatchEvent(new Event('input', {bubbles:true}));
                    s.dispatchEvent(new Event('change', {bubbles:true}));
                    return s.value === opt.value;
                }""",
                label,
            )
            if ok:
                await page.wait_for_timeout(500)
                return
        except Exception as e:
            last_err = e
        await page.wait_for_timeout(400)
    raise RuntimeError(f"Could not select '{label}': {last_err}")


async def select_dropdowns(page, site_subject):
    await page.wait_for_selector("select")
    await select_option_js(page, TEST_TYPE)
    await select_option_js(page, site_subject)
    await page.wait_for_timeout(500)


async def select_layout(page, needle):
    """Click the layout card whose description contains `needle`.

    NOTE: this MUST be a real (trusted) Playwright click. An earlier version
    used `element.click()` via page.evaluate() — that dispatches an untrusted
    click event which the layout-card component silently ignores unless the
    card is already the default (first) one. That bug meant the script was
    always stuck on layout card 1 no matter which layout was configured,
    which went unnoticed until a card with different required fields (e.g. a
    mandatory "Passage Text") was the one actually needed instead.
    """
    if not needle:
        return
    card = page.locator(".cursor-pointer", has_text=needle).first
    if await card.count() == 0:
        raise RuntimeError(f"Layout card matching '{needle}' not found")
    await card.click()
    await page.wait_for_timeout(700)


async def set_editor(page, editor_index, text):
    editors = await page.query_selector_all(".rsw-ce")
    if len(editors) <= editor_index:
        raise RuntimeError(f"Editor #{editor_index} not found (found {len(editors)})")
    await page.evaluate(
        """([el, t]) => {
            el.focus(); el.innerText = t;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
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
            inputs[idx].dispatchEvent(new Event('change', {bubbles:true}));
            inputs[idx].dispatchEvent(new Event('input', {bubbles:true}));
            return true;
        }""",
        [file_input_index, img_b64, name, mime],
    )
    if not ok:
        raise RuntimeError(f"File input [{file_input_index}] not found")


async def set_difficulty(page, value):
    await page.evaluate(
        """(val) => { const r = document.querySelector(`input[type=radio][value="${val}"]`);
            if (r && !r.checked) r.click(); }""",
        value,
    )


async def set_correct(page, indices):
    for i in indices:
        sel = f'button[data-testid="answer-{i}-correct"]'
        for _ in range(3):
            state = await page.evaluate(
                """(sel) => { const el = document.querySelector(sel);
                    return el ? el.getAttribute('aria-checked') : 'missing'; }""",
                sel,
            )
            if state == "true":
                break
            if state == "missing":
                raise RuntimeError(f"Correct toggle {sel} not found")
            await page.evaluate("(sel) => document.querySelector(sel).click()", sel)
            await page.wait_for_timeout(250)


async def fill_options(page, options):
    for i, opt in enumerate(options):
        sel = f'input[name="answers.{i}.text"]'
        el = await page.query_selector(sel)
        if el is None:
            raise RuntimeError(f"Answer input {sel} not found")
        await page.fill(sel, opt)


# ── Upload one question ───────────────────────────────────────────────────────

async def upload_image_field(page, file_input_index, img_path):
    """inject_image() + wait for the upload-image network round trip."""
    await inject_image(page, file_input_index, img_path)
    try:
        async with page.expect_response(
            lambda r: "upload-image" in r.url and r.status == 200, timeout=20000
        ):
            pass
    except Exception:
        await page.wait_for_timeout(2000)
    await page.wait_for_timeout(800)


async def upload_puzzle_question(page, q, cfg, no_submit):
    print(f"-> Q{q['num']} ({q['title']}) ...", end=" ", flush=True)

    await page.goto(ADD_URL)
    await page.wait_for_load_state("networkidle")

    await select_dropdowns(page, cfg["site_subject"])
    await select_layout(page, cfg.get("layout_text"))
    await set_difficulty(page, DIFFICULTY)

    await page.fill('input[name="questionTitle"]', q["title"][:120])
    await page.fill('input[name="geogebraContent"]', q.get("material_id") or "")
    if q.get("correct_answer"):
        await page.fill('input[name="correctAnswer"]', q["correct_answer"])

    # rich-text editors, in DOM order: [0]=Instruction Text, [1]=Question
    if "question" in q:
        instruction, question = split_instruction(q["instruction"], q["question"])
        if instruction:
            await set_editor(page, 0, instruction)
            await page.wait_for_timeout(150)
        if question:
            await set_editor(page, 1, question)
            await page.wait_for_timeout(150)
    elif q.get("instruction"):   # legacy layout: instruction only
        await set_editor(page, 0, q["instruction"][:INSTRUCTION_MAX])
        await page.wait_for_timeout(150)
        await set_editor(page, 1, q["instruction"])
        await page.wait_for_timeout(150)
    if q.get("explanation"):
        # fill the Explanation editor only if this layout has one
        if len(await page.query_selector_all(".rsw-ce")) > 2:
            await set_editor(page, 2, q["explanation"])
            await page.wait_for_timeout(150)
        else:
            print("(no explanation box on this layout - skipped)", end=" ")

    missing = []
    if not q.get("material_id"):
        missing.append("Material ID")
    if not q.get("correct_answer"):
        missing.append("Correct Answer")
    if q.get("needs_manual"):
        missing.append("manual check")

    if no_submit or missing:
        os.makedirs(SHOT_DIR, exist_ok=True)
        shot = os.path.join(SHOT_DIR, f"q{q['num']}.png")
        await page.screenshot(path=shot, full_page=True)
        if no_submit:
            print(f"[NO-SUBMIT] screenshot -> {shot}")
            return
        # don't submit an incomplete puzzle - report it so it lands in the Failed list
        raise RuntimeError(f"not submitted, still needs: {', '.join(missing)} (screenshot {shot})")

    await click_finish(page, q)


async def upload_question(page, q, cfg, no_submit):
    if cfg.get("kind") == "puzzle":
        return await upload_puzzle_question(page, q, cfg, no_submit)

    # NVR questions whose composite image couldn't be split into a stem +
    # 5 option crops have nowhere to go: the site has no "text-based answer
    # options" layout for NVR (only image-based-options layouts exist), so
    # there is no automated path for them. Skip cleanly and flag for manual
    # entry instead of failing on a missing layout card.
    if cfg.get("kind") == "nvr" and q.get("layout_kind") != "image_options":
        raise RuntimeError(
            "needs manual entry - could not split composite image into "
            "stem + 5 options, and NVR has no text-based-options layout on "
            f"the site to fall back to (image={os.path.basename(q['images'][0])})"
        )

    print(f"-> Q{q['num']} (answer={q['answer']}) ...", end=" ", flush=True)

    await page.goto(ADD_URL)
    await page.wait_for_load_state("networkidle")

    await select_dropdowns(page, cfg["site_subject"])

    has_image = bool(q.get("passage_image") or q.get("images"))
    has_passage = bool(q.get("passage_text"))
    is_nvr_image_options = cfg.get("kind") == "nvr" and q.get("layout_kind") == "image_options"
    is_nvr_options_only = is_nvr_image_options and not q.get("images")
    if cfg.get("kind") == "english_vr":
        # layout_passage has a REQUIRED "Passage Text" editor (must only be
        # used when we actually have passage text); layout_text has no
        # Passage editor at all (Question is editor 0 there).
        layout = cfg["layout_passage"] if has_passage else cfg["layout_text"]
    elif is_nvr_options_only:
        layout = cfg["layout_options_only"]
    elif is_nvr_image_options:
        layout = cfg["layout_image_split"]
    else:
        layout = cfg.get("layout_image") if (has_image and cfg.get("layout_image")) else cfg.get("layout_text")
    await select_layout(page, layout)
    await set_difficulty(page, DIFFICULTY)

    await page.fill('input[name="questionTitle"]', q["title"][:120])

    # image (passage for creative, or maths/NVR content image)
    img_path = q.get("passage_image") or (q["images"][0] if q.get("images") else None)
    if img_path:
        await upload_image_field(page, 0, img_path)

    # NVR image-options layouts: upload each option's own crop into its file
    # input — indices 1-5 right after the stem image, or 0-4 when the layout
    # has no stem image (options-only / odd one out).
    if is_nvr_image_options:
        first = 0 if is_nvr_options_only else 1
        for i, opt_path in enumerate(q["option_images"]):
            await upload_image_field(page, first + i, opt_path)

    # question / boxed text editor.
    # english_vr layout_passage has 3 editors: 0=Passage Text, 1=Question
    # (100 chars max), 2=Explanation. layout_text (no passage) has only 2:
    # 0=Question, 1=Explanation. All other subjects' layouts have 2:
    # 0=Question, 1=Explanation.
    if cfg.get("kind") == "english_vr":
        if has_passage:
            await set_editor(page, 0, q["passage_text"])
            await page.wait_for_timeout(150)
            question_idx, explanation_idx = 1, 2
        else:
            question_idx, explanation_idx = 0, 1
        if (q["stem"] or "").strip():
            await set_editor(page, question_idx, q["stem"])
            await page.wait_for_timeout(150)
    else:
        explanation_idx = 1
        if (q["stem"] or "").strip():
            await set_editor(page, 0, q["stem"])
            await page.wait_for_timeout(150)

    # NVR image-options layout has no text answer inputs — options are the
    # images just uploaded above. Every other layout (incl. the NVR
    # text-options fallback) fills the plain-text answer boxes.
    if not is_nvr_image_options:
        await fill_options(page, q["options"])

    if q["correct"]:
        await set_correct(page, q["correct"])

    if q.get("explanation"):
        try:
            await set_editor(page, explanation_idx, q["explanation"])
        except Exception:
            pass

    if no_submit:
        os.makedirs(SHOT_DIR, exist_ok=True)
        shot = os.path.join(SHOT_DIR, f"q{q['num']}.png")
        await page.screenshot(path=shot, full_page=True)
        print(f"[NO-SUBMIT] screenshot -> {shot}")
        return

    await click_finish(page, q)


async def click_finish(page, q):
    try:
        btn = page.locator("button", has_text="Finish").last
        await btn.scroll_into_view_if_needed()
        await btn.click(timeout=5000)
    except Exception:
        await page.evaluate(
            """() => { const b = [...document.querySelectorAll('button')]
                .find(b => (b.innerText||'').trim() === 'Finish'); if (b) b.click(); }"""
        )
    try:
        await page.wait_for_function(
            "document.body.innerText.includes('successfully')"
            " || !window.location.href.includes('/add')",
            timeout=20000,
        )
        print("OK")
    except Exception:
        os.makedirs(SHOT_DIR, exist_ok=True)
        shot = os.path.join(SHOT_DIR, f"error_q{q['num']}.png")
        await page.screenshot(path=shot, full_page=True)
        raise RuntimeError(f"Submit failed — screenshot {shot}")


# ── Main ──────────────────────────────────────────────────────────────────────

async def run_uploads(questions, cfg, args):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        async def on_response(resp):
            if "upload-image" in resp.url:
                print(f"    [net] upload-image -> {resp.status}", flush=True)
        page.on("response", on_response)

        await login(page)

        failed = []
        for i, q in enumerate(questions):
            try:
                await upload_question(page, q, cfg, args.no_submit)
            except Exception as e:
                print(f"x  ({e})")
                failed.append(q["num"])
            if i < len(questions) - 1 and DELAY_MIN > 0 and not args.no_submit:
                delay = random.randint(DELAY_MIN, DELAY_MAX)
                print(f"   waiting {delay}s before next question ...", flush=True)
                await asyncio.sleep(delay)

        await browser.close()

    print(f"\n{'=' * 60}")
    ok = len(questions) - len(failed)
    print(f"Done: {ok}/{len(questions)} processed")
    if failed:
        print(f"Failed: Q{', Q'.join(map(str, failed))}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-submit", action="store_true")
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    args = parser.parse_args()

    if not os.path.isfile(DOCX_PATH):
        raise SystemExit(f"DOCX_PATH not found: {DOCX_PATH}")

    print(f"Parsing {os.path.basename(DOCX_PATH)}  (topic={TOPIC}) ...")

    questions = parse_subject(TOPIC, DOCX_PATH)

    if not preflight_check(questions, TOPIC, CFG):
        sys.exit(1)

    if args.start or args.end:
        lo = args.start or 1
        hi = args.end or max(q["num"] for q in questions)
        questions = [q for q in questions if lo <= q["num"] <= hi]
        print(f"\nFiltered to Q{lo}-Q{hi}: {len(questions)} questions\n")

    if args.dry_run:
        print("\n[DRY RUN — no browser, no uploads]\n")
        return

    asyncio.run(run_uploads(questions, CFG, args))


if __name__ == "__main__":
    main()
