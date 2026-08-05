"""
docx_to_json.py
---------------
Phase 1: Convert a DOCX file into structured JSON (one entry per topic/scheme).

Steps:
  1. Parse DOCX — walk paragraphs + tables in document order
  2. Detect topic boundaries (bold / heading paragraphs)
  3. Extract every paragraph and table under each topic (nothing skipped)
  4. Send each topic's raw text to Ollama → get slide JSON
  5. Save per-topic JSON files + one combined all_topics.json

Usage:
  python docx_to_json.py --file uploads/your.docx --out output/json
  python docx_to_json.py --file uploads/your.docx --out output/json --ollama http://host:11434 --model deepseek-r1:14b
"""

import argparse
import json
import logging
import os
import re
import sys
import time

import requests
from docx import Document as DocxDocument

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("docx_to_json.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("docx_to_json")

# ── Defaults ───────────────────────────────────────────────────────────────────
DEFAULT_OLLAMA_URL = "http://172.23.198.77:11434"
DEFAULT_MODEL      = "deepseek-r1:14b"

SYSTEM_PROMPT = """You are an expert banking trainer creating PowerPoint presentations from financial scheme documents.

=============================================
SECTION A — PPT BULLET RULES (STRICT CONTENT EXTRACTION)
=============================================

1. EVERY SINGLE piece of information from the source document MUST become a bullet point. No exceptions.
   - Every line → a bullet.
   - Every sub-point (a, b, c, d, e, f...) → a separate bullet each.
   - Every nested point (i, ii, iii, iv...) → a separate bullet each.
   - Every row in a table → a bullet (or multiple bullets if the row has multiple pieces of info).
   - Every note, remark, condition, exception, clause → a separate bullet.

2. COUNT AND VERIFY: Before generating output, count the total information points in the source for each parameter/section. Your bullets for that slide MUST match or exceed that count.
   - If a parameter has sub-points a) through f) → minimum 6 bullets.
   - If a sub-point has nested points i) through iv) → those are 4 additional bullets.
   - If the source has 10 points, you MUST have at least 10 bullets. NEVER fewer.

3. BULLETS = EXACT SOURCE CONTENT:
   - Bullets must contain ONLY what is explicitly written in the source document.
   - Do NOT rephrase, rewrite, or humanize the bullet text. Keep it as close to source wording as possible.
   - Do NOT combine two separate lines into one bullet.
   - Do NOT summarize multiple conditions into one bullet.
   - Do NOT drop footnotes, remarks, parenthetical text, or "minor" details.
   - Include full text of conditions including "i.e." explanations and parenthetical qualifications.

4. PRESERVE ALL DATA EXACTLY:
   - All numbers, percentages, dates, monetary values, interest rates, names, references, circular numbers, CIC scores — EXACTLY as they appear.
   - Do NOT round numbers, do NOT paraphrase amounts, do NOT skip reference numbers.
   - Preserve parenthetical text. Example: "(For acceptance of lower CIC score, guidelines issued by IRMD shall be referred)" must be included.

5. ONE slide per parameter/section. Slide TITLE = exact parameter name from the source.
   - If a parameter has too many bullets for one slide, split into multiple slides with (1/2), (2/2) etc.
   - Even when splitting, NO information can be lost — all bullets must appear across the split slides.

6. VERIFICATION STEP: After generating all slides, go back to source and check EACH parameter:
   - Count sub-points in source vs your bullets. If bullets < source points → ADD the missing ones.
   - Check: Did I include every (a), (b), (c)...? Every (i), (ii), (iii)...? Every note/condition at the end?

7. First slide = title slide. Last slide = thank_you slide.

BULLET MISTAKES TO AVOID:
- Source has points a) through f) → you output only 4 bullets. WRONG. Output all 6.
- Source has nested i) through iv) → you merge into parent. WRONG. Each gets its own bullet.
- Source says "Revenue should have commenced i.e. the unit must have acquired customer(s)" → you write "Revenue should have commenced." WRONG. Include full text with "i.e." part.
- Source mentions "Promoter/s CIC Score: Minimum 700 (For acceptance of lower CIC score, guidelines issued by IRMD shall be referred)" → you write "CIC Score: Minimum 700." WRONG. Include full parenthetical.
- Source has a final minor point → you skip it. WRONG. Include it.
- Two conditions exist → you merge into one bullet. WRONG. Keep separate.

=============================================
SECTION B — VOICEOVER RULES (HUMAN-LIKE EXPLANATION)
=============================================

RULE 1 — EXPLAIN EVERY BULLET INDIVIDUALLY:
- Take each bullet point ONE BY ONE in order.
- For each bullet, deliver a separate explanation. Do NOT merge multiple bullets into one paragraph or summary.
- If a slide has 8 bullets, the voiceover must have 8 distinct explanations.
- NEVER skip a bullet in the voiceover.

RULE 2 — ZERO INFORMATION LOSS IN VOICEOVER:
- Every single word, number, name, condition, percentage, amount, date in the bullet MUST appear in the voiceover.
- Do NOT drop any detail. Do NOT shorten.
- Example:
  - Bullet: "Turnover: Not to exceed ₹100 crore in any Financial Year post incorporation."
  - WRONG voiceover: "There's a turnover limit for eligibility." (lost amount, lost "post incorporation")
  - RIGHT voiceover: "Now coming to turnover — it should not exceed ₹100 crore in any financial year after incorporation. So in any given year post incorporation, if the turnover crosses ₹100 crore, the entity won't qualify."

RULE 3 — HUMAN-LIKE EXPLANATION (NOT READING, NOT SUMMARIZING):
- Do NOT read the bullet word-for-word in the voiceover.
- Do NOT summarize or compress the bullet in the voiceover.
- REPHRASE it naturally as a trainer would explain — while keeping ALL information intact.
- Same information, natural spoken delivery, zero loss.
- Examples:
  - Bullet: "Age of startup: Maximum 10 years from the date of incorporation/registration."
  - READING (bad): "Age of startup is maximum 10 years from the date of incorporation or registration."
  - SUMMARIZING (bad): "There's an age limit for startups."
  - EXPLAINING (good): "Now regarding the age of the startup — it should not be more than 10 years old from the date it was incorporated or registered. So if a startup was registered more than 10 years ago, it won't be eligible under this scheme."

  - Bullet: "Private Limited Company, Registered Partnership or Limited Liability Partnership firm registered & recognized as Start-up by DPIIT, Government of India."
  - READING (bad): "Private Limited Company, Registered Partnership or LLP registered and recognized as Start-up by DPIIT, Government of India."
  - SUMMARIZING (bad): "The entity must be legally recognized by DPIIT."
  - EXPLAINING (good): "The entity type should be either a Private Limited Company, a Registered Partnership, or a Limited Liability Partnership firm. And importantly, it must be registered and recognized as a Start-up by DPIIT, which is the Department for Promotion of Industry and Internal Trade under the Government of India."

  - Bullet: "Financing to Start-ups in MSME, Agriculture and Allied Services."
  - READING (bad): "Financing to Start-ups in MSME, Agriculture and Allied Services."
  - SUMMARIZING (bad): "The scheme covers MSMEs and agriculture."
  - EXPLAINING (good): "This scheme provides financing specifically to start-ups operating in three areas — MSME sector, Agriculture, and Allied Services. So the scope is limited to these domains."

RULE 4 — TONE AND DELIVERY:
- Speak as a trainer addressing bank officers in a training session.
- Use "you" to address the learner naturally.
- Use transitions between bullets: "Now the next point is...", "Coming to the next condition...", "Here's another important aspect...", "Let's look at the next one..."
- Keep it conversational — like explaining to a colleague, not reading a document.

RULE 5 — VOICEOVER MUST NEVER:
- Merge all bullets into one summary paragraph.
- Say "That's straightforward" or "Let's move on" without explaining the bullet.
- Be empty.
- Skip any bullet.
- Drop numbers, names, amounts, or conditions.
- Add information NOT in the bullet (no external examples, no assumptions, no invented benefits).
- Use vague phrases like "there are certain criteria" or "multiple options are available."

RULE 6 — STRUCTURE:
- Title slide voiceover: Welcome learners, state the scheme name, briefly set expectation.
- Content slides voiceover: Explain each bullet one by one with natural transitions.
- Thank you slide voiceover: Wrap up naturally, remind to refer source document.
- Voiceover length: 25 to 35 words PER bullet. A slide with 4 bullets = ~100-140 words. A slide with 8 bullets = ~200-280 words. Scale with bullet count — NEVER cap at a fixed word limit.

=============================================
OUTPUT FORMAT — Return ONLY valid JSON:
=============================================
{
  "slides": [
    {"type": "title", "title": "...", "subtitle": "Source Document", "voiceover": "..."},
    {"type": "section_title", "title": "...", "subtitle": "...", "voiceover": "..."},
    {"type": "content", "title": "...", "bullets": ["...", "..."], "voiceover": "..."},
    {"type": "thank_you", "title": "Thank You", "subtitle": "For more details refer source document", "voiceover": "..."}
  ]
}"""

# ── DOCX Parsing ───────────────────────────────────────────────────────────────

def _table_to_text(table) -> str:
    """Convert a DOCX table to pipe-delimited text, deduplicating merged cells."""
    lines = []
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        deduped = []
        for c in cells:
            if not deduped or c != deduped[-1]:
                deduped.append(c)
        line = " | ".join(deduped)
        if line.strip():
            lines.append(line)
    return "\n".join(lines)


def _is_topic_heading(para) -> bool:
    """Return True if this paragraph looks like a top-level topic/scheme title."""
    text = para.text.strip()
    if not text or len(text) < 4 or len(text) > 300:
        return False
    if para.style.name.startswith("Heading"):
        return True
    if para.runs and all(r.bold for r in para.runs if r.text.strip()):
        return True
    return False


def parse_docx_topics(docx_path: str) -> list[dict]:
    """
    Walk the DOCX body in document order (paragraphs + tables).
    Split content at every topic-heading boundary.

    Returns:
        [{"num": 1, "title": "...", "raw_text": "..."}, ...]
    """
    logger.info(f"Parsing DOCX: {docx_path}")
    doc = DocxDocument(docx_path)

    # Build ordered list of (kind, obj) preserving document order
    para_map  = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    ordered = []
    for child in doc.element.body:
        if child in para_map:
            ordered.append(("para", para_map[child]))
        elif child in table_map:
            ordered.append(("table", table_map[child]))

    logger.info(f"  Total elements in body: {len(ordered)}")

    # Identify topic boundary indices
    boundaries = []
    for i, (kind, obj) in enumerate(ordered):
        if kind == "para" and _is_topic_heading(obj):
            boundaries.append((i, obj.text.strip()))

    logger.info(f"  Detected {len(boundaries)} topic headings")
    for idx, (pos, title) in enumerate(boundaries):
        logger.info(f"    [{idx+1}] pos={pos}  title={title[:80]}")

    if not boundaries:
        logger.warning("No topic headings found — treating entire document as one topic")
        all_text = []
        for kind, obj in ordered:
            if kind == "para":
                t = obj.text.strip()
                if t:
                    all_text.append(t)
            else:
                t = _table_to_text(obj)
                if t:
                    all_text.append(t)
        return [{"num": 1, "title": "Document", "raw_text": "\n".join(all_text)}]

    topics = []
    for idx, (start_pos, title) in enumerate(boundaries):
        end_pos = boundaries[idx + 1][0] if idx + 1 < len(boundaries) else len(ordered)
        segment = ordered[start_pos:end_pos]

        parts = []
        for kind, obj in segment:
            if kind == "para":
                t = obj.text.strip()
                if t:
                    parts.append(t)
            else:
                t = _table_to_text(obj)
                if t:
                    parts.append(t)

        raw_text = "\n".join(parts)
        char_count = len(raw_text)
        logger.info(f"  Topic [{idx+1}] '{title[:60]}' — {len(segment)} elements, {char_count} chars")

        if char_count < 20:
            logger.warning(f"    ⚠ Very short content ({char_count} chars) — still including")

        topics.append({
            "num": idx + 1,
            "title": title,
            "raw_text": raw_text,
        })

    return topics


# ── Ollama ─────────────────────────────────────────────────────────────────────

def call_ollama(prompt: str, ollama_url: str, model: str, retries: int = 2) -> str | None:
    url = f"{ollama_url}/api/generate"
    payload = {
        "model": model,
        "prompt": prompt,
        "system": SYSTEM_PROMPT,
        "stream": False,
        "options": {"temperature": 0.2, "num_ctx": 131072, "num_predict": 65536},
        "format": "json",
    }
    for attempt in range(1, retries + 2):
        try:
            resp = requests.post(url, json=payload, timeout=1200)
            resp.raise_for_status()
            return resp.json().get("response", "")
        except Exception as e:
            logger.error(f"Ollama call failed (attempt {attempt}): {e}")
            if attempt <= retries:
                wait = attempt * 5
                logger.info(f"  Retrying in {wait}s...")
                time.sleep(wait)
    return None


def parse_json_response(text: str) -> dict | None:
    """Try to extract valid JSON from LLM response, with truncation repair."""
    if not text:
        return None
    # Direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # Extract first {...} block
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    # Repair truncated JSON: close any open arrays/objects
    logger.warning("JSON parse failed — attempting truncation repair")
    snippet = re.search(r"\{[\s\S]*", text)
    if snippet:
        partial = snippet.group().rstrip(", \t\n")
        # Close dangling string
        if partial.count('"') % 2 != 0:
            partial += '"'
        # Close open array/object by counting brackets
        open_sq  = partial.count('[') - partial.count(']')
        open_cur = partial.count('{') - partial.count('}')
        partial += ']' * max(open_sq, 0) + '}' * max(open_cur, 0)
        try:
            result = json.loads(partial)
            logger.warning("Truncation repair succeeded")
            return result
        except Exception:
            pass
    logger.error("Could not parse JSON from LLM response")
    logger.debug(f"Raw response (first 500): {text[:500]}")
    return None


# ── Post-processing ────────────────────────────────────────────────────────────

def _clean_bullets(bullets: list) -> list:
    """Deduplicate bullets; only drop pure filler with no real content alongside it."""
    FILLER_ONLY = re.compile(
        r'^refer\s+scheme\s+circular\s+for\s+additional\s+guidelines\.?$', re.IGNORECASE
    )
    seen, deduped = set(), []
    for b in bullets:
        b = b.strip()
        if not b:
            continue
        if FILLER_ONLY.match(b):
            continue
        k = b.lower()
        if k not in seen:
            seen.add(k)
            deduped.append(b)
    return deduped


def _build_voiceover(bullets: list, title: str = "", ollama_url: str = "", model: str = "") -> str:
    """Call Ollama to generate natural trainer-style voiceover from exact bullets."""
    if not bullets:
        return ""
    if ollama_url and model:
        bullets_text = "\n".join(f"- {b}" for b in bullets)
        prompt = (
            f"You are a banking trainer narrating a PowerPoint slide to bank officers.\n\n"
            f"Slide Title: {title}\n\nBullets on this slide:\n{bullets_text}\n\n"
            f"Write a natural spoken narration for this slide.\n"
            f"RULES:\n"
            f"- Cover EVERY bullet point individually in order. Do NOT skip any bullet.\n"
            f"- Use EXACT values, numbers, names, conditions from each bullet. Do not drop or shorten any detail.\n"
            f"- Explain each bullet naturally as a trainer — do not just read it word for word.\n"
            f"- Use transitions: 'First...', 'Next...', 'Coming to...', 'Also...', 'And finally...'.\n"
            f"- Word count: approximately 25-35 words PER bullet (so {len(bullets) * 25}-{len(bullets) * 35} words total).\n"
            f"- Output ONLY the narration text. No labels, no JSON, no bullet markers."
        )
        try:
            resp = requests.post(
                f"{ollama_url}/api/generate",
                json={"model": model, "prompt": prompt, "stream": False,
                      "options": {"temperature": 0.3, "num_ctx": 8192, "num_predict": 1024}},
                timeout=120,
            )
            resp.raise_for_status()
            text = resp.json().get("response", "").strip()
            if text:
                return text
        except Exception as e:
            logger.warning(f"Voiceover regeneration failed: {e}")
    parts = []
    for i, b in enumerate(bullets):
        b = b.strip().rstrip(".")
        if i == 0:
            parts.append(b + ".")
        elif i == len(bullets) - 1:
            parts.append("And finally, " + b[0].lower() + b[1:] + ".")
        else:
            parts.append("Also, " + b[0].lower() + b[1:] + ".")
    return " ".join(parts)


def enforce_8_bullets(data: dict, ollama_url: str = "", model: str = "") -> dict:
    """
    Post-processing:
    - Any content slide with no bullets has its title and voiceover carried forward
      and merged into the next content slide that has bullets (even across section_title boundaries).
    - Club consecutive content slides with bullets until >= 8 bullets total.
    - If clubbed bullets > 8, split into (1/N), (2/N), etc.
    - No filler padding — only real content.
    """
    slides = data.get("slides", [])

    # Step 1: clean bullets on all content slides
    cleaned = []
    for slide in slides:
        if slide.get("type") == "content":
            cleaned.append({**slide, "bullets": _clean_bullets(slide.get("bullets", []))})
        else:
            cleaned.append(slide)

    # Step 0: normalise bad types + drop redundant section_title slides
    def _titles_similar(a: str, b: str) -> bool:
        a, b = a.strip().lower(), b.strip().lower()
        if a == b or a in b or b in a:
            return True
        wa, wb = set(a.split()), set(b.split())
        shorter = wa if len(wa) <= len(wb) else wb
        return bool(shorter) and len(wa & wb) / len(shorter) >= 0.5

    normalised = []
    for slide in cleaned:
        if slide.get("type") not in ("title", "thank_you", "section_title", "content"):
            slide = {**slide, "type": "section_title"}
        normalised.append(slide)

    deduped = []
    for idx, slide in enumerate(normalised):
        if slide.get("type") == "section_title":
            next_any = next((s for s in normalised[idx+1:]
                             if s.get("type") in ("content", "section_title", "thank_you")), None)
            if next_any and next_any.get("type") == "content" and _titles_similar(
                slide.get("title", ""), next_any.get("title", "")
            ):
                logger.info(f"  Dropping duplicate section_title: '{slide['title']}'")
                continue
            if next_any and next_any.get("type") != "content":
                next_content = next((s for s in normalised[idx+1:] if s.get("type") == "content"), None)
                if next_content:
                    vo = slide.get("voiceover", "")
                    if vo:
                        next_content["voiceover"] = (vo + " " + next_content.get("voiceover", "")).strip()
                logger.info(f"  Dropping standalone section_title: '{slide['title']}'")
                continue
        deduped.append(slide)
    carried_input = deduped

    # Step 2: carry forward empty content slide titles to the next content slide with bullets
    pending_titles    = []
    pending_voiceover = ""
    carried = []
    for slide in carried_input:
        if slide.get("type") == "content" and not slide.get("bullets"):
            pending_titles.append(slide["title"])
            if not pending_voiceover:
                pending_voiceover = slide.get("voiceover", "")
            logger.info(f"  Carrying forward empty slide: '{slide['title']}'")
            continue
        if slide.get("type") == "content" and pending_titles:
            merged_title = " and ".join(pending_titles + [slide["title"]])
            merged_vo    = " ".join(filter(None, [pending_voiceover, slide.get("voiceover", "")]))
            logger.info(f"  Merged into: '{merged_title}'")
            carried.append({**slide, "title": merged_title, "voiceover": merged_vo})
            pending_titles    = []
            pending_voiceover = ""
        else:
            if slide.get("type") != "content" and pending_titles:
                logger.info(f"  Dropping orphan empty slides: {pending_titles}")
                pending_titles    = []
                pending_voiceover = ""
            carried.append(slide)

    # Step 3: club consecutive content slides until bullets >= 5
    out = []
    i   = 0
    while i < len(carried):
        slide = carried[i]
        if slide.get("type") != "content":
            out.append(slide)
            i += 1
            continue

        clubbed_titles    = [slide["title"]]
        clubbed_bullets   = list(slide["bullets"])
        clubbed_voiceover = slide.get("voiceover", "")

        while len(clubbed_bullets) < 5 and i + 1 < len(carried):
            nxt = carried[i + 1]
            if nxt.get("type") != "content":
                break
            clubbed_titles.append(nxt["title"])
            clubbed_bullets.extend(nxt["bullets"])
            clubbed_voiceover = (clubbed_voiceover + " " + nxt.get("voiceover", "")).strip()
            i += 1

        merged_title = " and ".join(clubbed_titles)
        if len(clubbed_titles) > 1:
            logger.info(f"  Clubbed: {' + '.join(clubbed_titles)} ({len(clubbed_bullets)} bullets)")
            clubbed_voiceover = _build_voiceover(clubbed_bullets, merged_title, ollama_url, model)

        if len(clubbed_bullets) <= 8:
            out.append({"type": "content", "title": merged_title,
                        "bullets": clubbed_bullets, "voiceover": clubbed_voiceover})
        else:
            chunks = [clubbed_bullets[j:j+8] for j in range(0, len(clubbed_bullets), 8)]
            n = len(chunks)
            for pi, chunk in enumerate(chunks, 1):
                out.append({
                    "type":      "content",
                    "title":     f"{merged_title} ({pi}/{n})" if n > 1 else merged_title,
                    "bullets":   chunk,
                    "voiceover": _build_voiceover(chunk, merged_title, ollama_url, model),
                })
        i += 1

    data["slides"] = [s for s in out if s.get("type") != "content" or s.get("bullets")]
    return data


# ── Per-topic processing ───────────────────────────────────────────────────────

# Max chars per chunk sent to LLM (~100K chars ≈ ~25K tokens, well within 131072 ctx)
_CHUNK_CHARS = 100_000


def _make_prompt(title: str, raw_text: str, part: int = 0, total: int = 1) -> str:
    part_note = f" (Part {part} of {total} — continue from previous slides, DO NOT repeat slides already generated)" if total > 1 else ""
    return f"""Scheme / Topic Name: {title}{part_note}

Full Content:
---
{raw_text}
---

CRITICAL INSTRUCTIONS:
- Generate slides ONLY for the scheme/topic named above.
- READ EVERY LINE of the content above. Do NOT skip, summarize, or omit ANY row, cell, paragraph, note, condition, or remark.
- Create ONE slide per parameter/row (Objective, Eligibility, Purpose, Quantum, Margin, Security, Interest Rate, Repayment, etc.).
- Slide TITLE = exact parameter name as written in the document.
- BULLETS: Copy the exact text from the document for each parameter. Every piece of information must appear as a bullet. If a parameter has 1 sentence, 1 bullet. If it has 5 sentences, 5 bullets. NEVER invent, expand, pad, or skip.
- Every table row = at least one bullet. Every paragraph = at least one bullet. Every note/remark/condition = a separate bullet.
- If ANY information from the source is missing from the slides, that is an error.
- VOICEOVER: Say each bullet naturally in trainer language — same content, conversational tone. Cover every bullet. Use transitions: "First...", "Next...", "Also...", "Moving on...", "And finally...". 60-100 words per content slide.
- NEVER repeat the same bullet text more than once in the same slide.
- Respond ONLY in valid JSON matching the OUTPUT FORMAT in your instructions."""


def topic_to_json(topic: dict, ollama_url: str, model: str) -> dict | None:
    title    = topic["title"]
    raw_text = topic["raw_text"]

    # Split into chunks if content is very large
    if len(raw_text) <= _CHUNK_CHARS:
        chunks = [raw_text]
    else:
        # Split on newline boundaries to avoid cutting mid-sentence
        lines  = raw_text.split("\n")
        chunks, current = [], []
        current_len = 0
        for line in lines:
            if current_len + len(line) + 1 > _CHUNK_CHARS and current:
                chunks.append("\n".join(current))
                current, current_len = [], 0
            current.append(line)
            current_len += len(line) + 1
        if current:
            chunks.append("\n".join(current))
        logger.info(f"  Large topic split into {len(chunks)} chunks ({len(raw_text)} total chars)")

    all_slides = []
    total = len(chunks)

    for part_idx, chunk in enumerate(chunks, 1):
        prompt = _make_prompt(title, chunk, part_idx, total)
        logger.info(f"  Calling Ollama for: {title[:60]} chunk {part_idx}/{total} ({len(chunk)} chars)")
        t0 = time.time()
        response = call_ollama(prompt, ollama_url, model)
        elapsed = time.time() - t0
        logger.info(f"  Ollama responded in {elapsed:.1f}s")

        if not response:
            logger.error(f"  No response from Ollama for: {title} chunk {part_idx}")
            continue

        data = parse_json_response(response)
        if not data or "slides" not in data:
            logger.error(f"  Invalid JSON from Ollama for: {title} chunk {part_idx}")
            continue

        # Strip title/thank_you from mid-chunks — keep only content/section_title
        slides = data["slides"]
        if total > 1 and part_idx < total:
            slides = [s for s in slides if s.get("type") not in ("title", "thank_you")]
        if total > 1 and part_idx > 1:
            slides = [s for s in slides if s.get("type") not in ("title",)]

        all_slides.extend(slides)
        logger.info(f"  Chunk {part_idx}/{total}: {len(slides)} slides collected")

    if not all_slides:
        logger.error(f"  No slides produced for: {title}")
        return None

    combined = {"slides": all_slides}
    slide_count_before = len(all_slides)
    combined = enforce_8_bullets(combined, ollama_url, model)
    slide_count_after  = len(combined["slides"])

    logger.info(
        f"  ✅ {title[:60]} → {slide_count_before} slides raw → {slide_count_after} after enforcement"
    )
    return combined


# ── Main ───────────────────────────────────────────────────────────────────────

def run(docx_path: str, out_dir: str, ollama_url: str, model: str):
    os.makedirs(out_dir, exist_ok=True)

    logger.info("=" * 70)
    logger.info(f"DOCX  : {docx_path}")
    logger.info(f"Output: {out_dir}")
    logger.info(f"Ollama: {ollama_url}  Model: {model}")
    logger.info("=" * 70)

    # Step 1 — Parse
    topics = parse_docx_topics(docx_path)
    logger.info(f"\nTotal topics found: {len(topics)}\n")

    # Save raw topic index
    index_path = os.path.join(out_dir, "topics_index.json")
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(
            [{"num": t["num"], "title": t["title"], "chars": len(t["raw_text"])} for t in topics],
            f, ensure_ascii=False, indent=2,
        )
    logger.info(f"Saved topic index → {index_path}")

    # Step 2 — Convert each topic to JSON via Ollama
    all_results = []
    failed      = []

    for topic in topics:
        num   = topic["num"]
        title = topic["title"]
        logger.info(f"\n[{num}/{len(topics)}] Processing: {title[:70]}")

        result = topic_to_json(topic, ollama_url, model)

        if result:
            result["_meta"] = {"num": num, "title": title, "chars": len(topic["raw_text"])}
            # Save individual file
            safe = re.sub(r'[<>:"/\\|?*]', '', title)[:60].strip()
            safe = safe.encode("ascii", "ignore").decode("ascii").strip()
            fname = os.path.join(out_dir, f"{num:03d}_{safe}.json")
            with open(fname, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            logger.info(f"  Saved → {fname}")
            all_results.append(result)
        else:
            logger.error(f"  ❌ FAILED: {title}")
            failed.append({"num": num, "title": title})

    # Step 3 — Save combined file
    combined_path = os.path.join(out_dir, "all_topics.json")
    with open(combined_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    logger.info(f"\nSaved combined JSON → {combined_path}")

    # Summary
    logger.info("\n" + "=" * 70)
    logger.info(f"DONE: {len(all_results)} succeeded, {len(failed)} failed")
    if failed:
        logger.warning("Failed topics:")
        for t in failed:
            logger.warning(f"  [{t['num']}] {t['title']}")
    logger.info("=" * 70)

    return all_results, failed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert DOCX topics to structured JSON via Ollama")
    parser.add_argument("--file",   required=True,                   help="Path to input DOCX file")
    parser.add_argument("--out",    default="output/json",           help="Output directory (default: output/json)")
    parser.add_argument("--ollama", default=DEFAULT_OLLAMA_URL,      help="Ollama base URL")
    parser.add_argument("--model",  default=DEFAULT_MODEL,           help="Ollama model name")
    args = parser.parse_args()

    run(args.file, args.out, args.ollama, args.model)
