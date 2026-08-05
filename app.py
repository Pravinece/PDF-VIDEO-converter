"""
FastAPI App - PPT Generator
- Upload DOCX/PDF from frontend
- Custom system prompt input
- Generates PPT files via Ollama
- Returns downloadable .pptx

Run: pip install -r requirements.txt
     python app.py
Open: http://localhost:8000
"""
import sys
import os
import json
import logging
import re
import io
import shutil
import tempfile
import time
import requests as http_requests
from functools import lru_cache

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("ppt_generator")

sys.stdout.reconfigure(encoding='utf-8')

import asyncio
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from PyPDF2 import PdfReader
from docx import Document as DocxDocument
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ─── App Setup ───────────────────────────────────────────────────────────────
app = FastAPI(title="PPT Generator", version="2.0")
templates = Jinja2Templates(directory="templates")

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ─── Defaults ────────────────────────────────────────────────────────────────
DEFAULT_OLLAMA_URL = "http://172.23.198.77:11434"
DEFAULT_MODEL = "deepseek-r1:14b"
DEFAULT_PDF = r"c:\Users\a1334\Desktop\genarate_ppt\Gist of Agriculture Lending Schemes as on 10102025.pdf"

DEFAULT_SYSTEM_PROMPT = """You are an expert banking trainer creating PowerPoint presentations from financial scheme documents.

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
}
"""

# ─── Colors ──────────────────────────────────────────────────────────────────
DARK_RED = RGBColor(0x9E, 0x17, 0x3B)
GOLD = RGBColor(0xFF, 0xC0, 0x00)
AMBER = RGBColor(0xFA, 0xB9, 0x0C)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
LIGHT_GRAY = RGBColor(0x55, 0x55, 0x55)


# ─── Ollama ──────────────────────────────────────────────────────────────────
def call_ollama(prompt: str, system_prompt: str, ollama_url: str, model: str, retries: int = 2) -> str | None:
    url = f"{ollama_url}/api/generate"
    payload = {
        "model": model,
        "prompt": prompt,
        "system": system_prompt,
        "stream": False,
        "options": {"temperature": 0.2, "num_ctx": 32768, "num_predict": 16384},
        "format": "json",
    }
    for attempt in range(1, retries + 2):
        try:
            resp = http_requests.post(url, json=payload, timeout=1200)
            resp.raise_for_status()
            return resp.json().get("response", "")
        except Exception as e:
            logger.error(f"Ollama call failed (attempt {attempt}): {e}")
            if attempt <= retries:
                wait = attempt * 5
                logger.info(f"  Retrying in {wait}s...")
                time.sleep(wait)
    return None


def _clean_bullets(bullets: list) -> list:
    FILLER_ONLY = re.compile(r'^refer\s+scheme\s+circular\s+for\s+additional\s+guidelines\.?$', re.IGNORECASE)
    seen, deduped = set(), []
    for b in bullets:
        b = b.strip()
        if not b or FILLER_ONLY.match(b):
            continue
        k = b.lower()
        if k not in seen:
            seen.add(k)
            deduped.append(b)
    return deduped


def _build_voiceover_app(bullets: list, title: str = "") -> str:
    """Fallback voiceover builder used when slides are split post-processing."""
    if not bullets:
        return ""
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


def _enforce_8_bullets(result: dict) -> dict:
    slides = list(result.get("slides", []))

    cleaned = [
        {**s, "bullets": _clean_bullets(s.get("bullets", []))} if s.get("type") == "content" else s
        for s in slides
    ]

    def _titles_similar(a: str, b: str) -> bool:
        a, b = a.strip().lower(), b.strip().lower()
        if a == b or a in b or b in a:
            return True
        wa, wb = set(a.split()), set(b.split())
        shorter = wa if len(wa) <= len(wb) else wb
        return bool(shorter) and len(wa & wb) / len(shorter) >= 0.5

    normalised = [
        {**s, "type": "section_title"} if s.get("type") not in ("title", "thank_you", "section_title", "content") else s
        for s in cleaned
    ]

    deduped = []
    for idx, slide in enumerate(normalised):
        if slide.get("type") == "section_title":
            next_any = next((s for s in normalised[idx+1:]
                             if s.get("type") in ("content", "section_title", "thank_you")), None)
            if next_any and next_any.get("type") == "content" and _titles_similar(
                slide.get("title", ""), next_any.get("title", "")
            ):
                continue
            if next_any and (next_any.get("type") != "content" or not _titles_similar(
                slide.get("title", ""), next_any.get("title", "")
            )):
                next_content = next((s for s in normalised[idx+1:] if s.get("type") == "content"), None)
                if next_content:
                    vo = slide.get("voiceover", "")
                    if vo:
                        next_content["voiceover"] = (vo + " " + next_content.get("voiceover", "")).strip()
                continue
        deduped.append(slide)

    pending_titles, pending_voiceover = [], ""
    carried = []
    for slide in deduped:
        if slide.get("type") == "content" and not slide.get("bullets"):
            pending_titles.append(slide["title"])
            if not pending_voiceover:
                pending_voiceover = slide.get("voiceover", "")
            continue
        if slide.get("type") == "content" and pending_titles:
            merged_title = " and ".join(pending_titles + [slide["title"]])
            merged_vo = " ".join(filter(None, [pending_voiceover, slide.get("voiceover", "")]))
            carried.append({**slide, "title": merged_title, "voiceover": merged_vo})
            pending_titles, pending_voiceover = [], ""
        else:
            if slide.get("type") != "content" and pending_titles:
                pending_titles, pending_voiceover = [], ""
            carried.append(slide)

    out = []
    i = 0
    while i < len(carried):
        slide = carried[i]
        if slide.get("type") != "content":
            out.append(slide)
            i += 1
            continue

        clubbed_titles = [slide["title"]]
        clubbed_bullets = list(slide["bullets"])
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
        if len(clubbed_bullets) <= 8:
            out.append({"type": "content", "title": merged_title,
                        "bullets": clubbed_bullets, "voiceover": clubbed_voiceover})
        else:
            chunks = [clubbed_bullets[j:j+8] for j in range(0, len(clubbed_bullets), 8)]
            n = len(chunks)
            for pi, chunk in enumerate(chunks, 1):
                chunk_title = f"{merged_title} ({pi}/{n})" if n > 1 else merged_title
                chunk_vo = _build_voiceover_app(chunk, chunk_title)
                out.append({
                    "type": "content",
                    "title": chunk_title,
                    "bullets": chunk,
                    "voiceover": chunk_vo,
                })
        i += 1

    result["slides"] = out
    return result


_CHUNK_CHARS = 100_000


def _parse_json_response(text: str) -> dict | None:
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    # Repair truncated JSON
    logger.warning("JSON parse failed — attempting truncation repair")
    snippet = re.search(r'\{[\s\S]*', text)
    if snippet:
        partial = snippet.group().rstrip(", \t\n")
        if partial.count('"') % 2 != 0:
            partial += '"'
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
    return None


def _make_prompt(title: str, raw_text: str, part: int = 1, total: int = 1) -> str:
    part_note = f" (Part {part} of {total} — continue from previous slides, DO NOT repeat slides already generated)" if total > 1 else ""
    return f"""Scheme Name: {title}{part_note}

Full Document Content:
---
{raw_text}
---

CRITICAL INSTRUCTIONS:
- From the full document above, locate ONLY the section for the scheme named: "{title}".
- Include ALL content: main parameter table AND all sub-sections/appendices.
- Do NOT include content from any other scheme.
- Title slide = scheme name. For each sub-section/appendix add a section_title slide.
- ONE slide per table parameter (Objective, Eligibility, Purpose, Quantum, Margin, etc.).
- Slide TITLE = exact parameter name.
- BULLETS: Copy exact text. Every piece of info = a bullet. NEVER skip, merge, or invent.
- If content is too large for one slide, SPLIT: e.g. "Collateral Security (1/2)", "Collateral Security (2/2)".
- NEVER repeat the same bullet text more than once in a slide.
- Respond ONLY in valid JSON."""


def generate_slides(title: str, raw_content: str, system_prompt: str, ollama_url: str, model: str) -> dict | None:
    # Split into chunks if content is very large
    if len(raw_content) <= _CHUNK_CHARS:
        chunks = [raw_content]
    else:
        lines = raw_content.split("\n")
        chunks, current, current_len = [], [], 0
        for line in lines:
            if current_len + len(line) + 1 > _CHUNK_CHARS and current:
                chunks.append("\n".join(current))
                current, current_len = [], 0
            current.append(line)
            current_len += len(line) + 1
        if current:
            chunks.append("\n".join(current))
        logger.info(f"Large content split into {len(chunks)} chunks ({len(raw_content)} total chars)")

    all_slides = []
    total = len(chunks)

    for part_idx, chunk in enumerate(chunks, 1):
        prompt = _make_prompt(title, chunk, part_idx, total)
        start = time.time()
        response_text = call_ollama(prompt, system_prompt, ollama_url, model)
        logger.info(f"LLM chunk {part_idx}/{total} took {time.time() - start:.2f}s for: {title}")

        if not response_text:
            logger.error(f"No response for chunk {part_idx}/{total}")
            continue

        result = _parse_json_response(response_text)
        if not result or "slides" not in result:
            logger.error(f"Failed to parse JSON for chunk {part_idx}/{total}")
            continue

        slides = result["slides"]
        if total > 1 and part_idx < total:
            slides = [s for s in slides if s.get("type") not in ("title", "thank_you")]
        if total > 1 and part_idx > 1:
            slides = [s for s in slides if s.get("type") != "title"]
        all_slides.extend(slides)

    if not all_slides:
        logger.error(f"No slides produced for: {title}")
        return None

    result = {"slides": all_slides}
    result = _enforce_8_bullets(result)
    return result


# ─── PDF helpers ─────────────────────────────────────────────────────────────
@lru_cache(maxsize=8)
def _read_pdf(pdf_path: str) -> list[str]:
    start = time.time()
    reader = PdfReader(pdf_path)
    pages = []
    for pg in reader.pages:
        text = pg.extract_text()
        pages.append(text if text else "")
    logger.info(f"PDF extraction took {time.time() - start:.2f}s ({len(pages)} pages)")
    return pages


def extract_pages(pdf_path: str, start_page: int, end_page: int) -> str:
    pages = _read_pdf(pdf_path)
    return "\n".join(pages[start_page:end_page])


def get_pdf_page_count(pdf_path: str) -> int:
    reader = PdfReader(pdf_path)
    return len(reader.pages)


# ─── DOCX helpers ────────────────────────────────────────────────────────────
def extract_docx_text(docx_path: str) -> str:
    doc = DocxDocument(docx_path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def _table_to_text(table) -> str:
    """Convert a docx table to structured text preserving all cell content."""
    lines = []
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        # Deduplicate merged cells
        seen = []
        for c in cells:
            if not seen or c != seen[-1]:
                seen.append(c)
        line = " | ".join(seen)
        if line.strip():
            lines.append(line)
    return "\n".join(lines)


def extract_docx_section_by_title(docx_path: str, title: str) -> str:
    """Extract the section matching the given scheme title from a DOCX.
    Reads both paragraphs AND tables in document order."""
    doc = DocxDocument(docx_path)
    title_lower = title.lower().strip()

    def similarity(a: str, b: str) -> float:
        a, b = a.lower(), b.lower()
        words = [w for w in a.split() if len(w) > 3]
        if not words:
            return 0.0
        return sum(1 for w in words if w in b) / len(words)

    def is_section_heading(p) -> bool:
        text = p.text.strip()
        if not text or len(text) > 300:
            return False
        return p.style.name.startswith('Heading') or (
            p.runs and all(r.bold for r in p.runs if r.text.strip())
        )

    # Build an ordered list of (element_type, element) from the XML body
    # so we can walk paragraphs and tables in document order
    body = doc.element.body
    ordered: list[tuple[str, object]] = []
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in body:
        if child in para_map:
            ordered.append(("para", para_map[child]))
        elif child in table_map:
            ordered.append(("table", table_map[child]))

    # Find start index — element most similar to title
    best_idx = -1
    best_score = 0.35
    for i, (etype, elem) in enumerate(ordered):
        if etype != "para":
            continue
        text = elem.text.strip()
        if not text:
            continue
        score = similarity(title_lower, text.lower())
        if score > best_score:
            best_score = score
            best_idx = i

    if best_idx == -1:
        logger.warning(f"Section '{title}' not found in DOCX, returning full text+tables")
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for t in doc.tables:
            parts.append(_table_to_text(t))
        return "\n".join(parts)

    # Collect everything from best_idx until the next unrelated section heading
    collected = []
    for i, (etype, elem) in enumerate(ordered[best_idx:], start=best_idx):
        if etype == "para":
            text = elem.text.strip()
            if not text:
                continue
            # Stop at the next section heading that is NOT part of this scheme
            if i > best_idx and is_section_heading(elem):
                if similarity(title_lower, text.lower()) < 0.25:
                    break
            collected.append(text)
        elif etype == "table":
            table_text = _table_to_text(elem)
            if table_text:
                collected.append(table_text)

    result = "\n".join(collected)
    logger.info(f"Extracted section '{title}': {len(result)} chars")
    return result if len(result) > 100 else "\n".join(
        [p.text.strip() for p in doc.paragraphs if p.text.strip()] +
        [_table_to_text(t) for t in doc.tables]
    )


@lru_cache(maxsize=8)
def _read_docx_pages(docx_path: str) -> list[str]:
    """Split DOCX into pages using lastRenderedPageBreak markers. Includes table content."""
    import zipfile
    from xml.etree import ElementTree as ET

    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    pages: list[list[str]] = []
    current: list[str] = []

    def _extract_para_text(para_elem) -> tuple[bool, str]:
        """Returns (has_page_break, text) for a paragraph element."""
        parts = []
        page_break = False
        for run in para_elem.iter(f"{{{ns}}}r"):
            for child in run:
                if child.tag == f"{{{ns}}}lastRenderedPageBreak":
                    page_break = True
                elif child.tag == f"{{{ns}}}t":
                    parts.append(child.text or "")
        return page_break, "".join(parts).strip()

    def _extract_table_text(tbl_elem) -> str:
        """Extract all text from a table element as pipe-delimited rows."""
        rows = []
        for row in tbl_elem.findall(f".//{{{ns}}}tr"):
            cells = []
            prev = None
            for cell in row.findall(f"{{{ns}}}tc"):
                cell_text = " ".join(
                    t.text or ""
                    for t in cell.iter(f"{{{ns}}}t")
                ).strip()
                if cell_text != prev:  # deduplicate merged cells
                    cells.append(cell_text)
                    prev = cell_text
            line = " | ".join(cells)
            if line.strip():
                rows.append(line)
        return "\n".join(rows)

    # Walk direct children of <w:body> to preserve document order
    body = root.find(f"{{{ns}}}body")
    if body is None:
        return []

    for child in body:
        tag = child.tag
        if tag == f"{{{ns}}}p":
            page_break, text = _extract_para_text(child)
            if page_break:
                if current:
                    pages.append(current)
                current = []
            if text:
                current.append(text)
        elif tag == f"{{{ns}}}tbl":
            table_text = _extract_table_text(child)
            if table_text:
                current.append(table_text)

    if current:
        pages.append(current)

    return ["\n".join(p) for p in pages if any(t.strip() for t in p)]


def extract_docx_pages(docx_path: str, start_page: int, end_page: int) -> str:
    pages = _read_docx_pages(docx_path)
    return "\n".join(pages[start_page:end_page])


def get_docx_page_count(docx_path: str) -> int:
    return len(_read_docx_pages(docx_path))


def extract_docx_topics(docx_path: str) -> list[str]:
    """Return all scheme/topic titles found in a DOCX by scanning bold/heading paragraphs."""
    doc = DocxDocument(docx_path)
    titles = []
    seen = set()
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or len(text) < 5 or len(text) > 250:
            continue
        is_heading = p.style.name.startswith('Heading')
        is_bold = p.runs and all(r.bold for r in p.runs if r.text.strip())
        if (is_heading or is_bold) and text.lower() not in seen:
            seen.add(text.lower())
            titles.append(text)
    return titles


def extract_uploaded_file_text(filename: str, title: str = None, start_page: int = None, end_page: int = None) -> str | None:
    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        return None
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".docx":
        if start_page is not None and end_page is not None:
            return extract_docx_pages(filepath, start_page, end_page)
        if title:
            return extract_docx_section_by_title(filepath, title)
        return extract_docx_text(filepath)
    elif ext == ".pdf":
        if start_page is not None and end_page is not None:
            return extract_pages(filepath, start_page, end_page)
        pages = _read_pdf(filepath)
        return "\n".join(pages)
    return None


# ─── PPT Builder ─────────────────────────────────────────────────────────────
def add_header_bar(slide, title_text: str):
    rect = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(530328))
    rect.fill.solid()
    rect.fill.fore_color.rgb = DARK_RED
    rect.line.fill.background()
    tf = rect.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = GOLD


def build_ppt(slides_data: list) -> Presentation:
    start = time.time()
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)

    for sd in slides_data:
        logger.info(f"Processing slide: {sd.get('title', 'Untitled')}")
        stype = sd.get("type", "content")

        # ── Title slide ───────────────────────────────────────────────────────
        if stype == "title":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = DARK_RED
            bg.line.fill.background()

            inner = slide.shapes.add_shape(1, Emu(914400), Emu(541895), Emu(10348101), Emu(5820806))
            inner.fill.solid()
            inner.fill.fore_color.rgb = AMBER
            inner.line.fill.background()

            tb = slide.shapes.add_textbox(Emu(1500000), Emu(1600000), Emu(9200000), Emu(2400000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = sd.get("title", "")
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = DARK_RED

            tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(3800000), Emu(9200000), Emu(2200000))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = sd.get("subtitle", "")
            r2.font.size = Pt(15)
            r2.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

        # ── Thank You slide ───────────────────────────────────────────────────
        elif stype == "thank_you":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = DARK_RED
            bg.line.fill.background()

            inner = slide.shapes.add_shape(1, Emu(914400), Emu(541895), Emu(10348101), Emu(5820806))
            inner.fill.solid()
            inner.fill.fore_color.rgb = AMBER
            inner.line.fill.background()

            tb = slide.shapes.add_textbox(Emu(1500000), Emu(2000000), Emu(9200000), Emu(1500000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "Thank You"
            r.font.size = Pt(44)
            r.font.bold = True
            r.font.color.rgb = DARK_RED

            tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(3600000), Emu(9200000), Emu(1800000))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = sd.get("subtitle", "")
            r2.font.size = Pt(16)
            r2.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

        # ── Content slide ─────────────────────────────────────────────────────
        elif stype == "section_title":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = DARK_RED
            bg.line.fill.background()

            inner = slide.shapes.add_shape(1, Emu(914400), Emu(541895), Emu(10348101), Emu(5820806))
            inner.fill.solid()
            inner.fill.fore_color.rgb = AMBER
            inner.line.fill.background()

            tb = slide.shapes.add_textbox(Emu(1500000), Emu(2000000), Emu(9200000), Emu(2000000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = sd.get("title", "")
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = DARK_RED

            sub = sd.get("subtitle", "")
            if sub:
                tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(4100000), Emu(9200000), Emu(1200000))
                tf2 = tb2.text_frame
                tf2.word_wrap = True
                p2 = tf2.paragraphs[0]
                p2.alignment = PP_ALIGN.CENTER
                r2 = p2.add_run()
                r2.text = sub
                r2.font.size = Pt(16)
                r2.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

        else:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_header_bar(slide, sd.get("title", ""))

            # Auto-fit font size based on bullet count to prevent overflow
            bullet_count = len(sd.get("bullets", []))
            font_size = Pt(18) if bullet_count <= 5 else Pt(15) if bullet_count <= 7 else Pt(13)

            tb = slide.shapes.add_textbox(Emu(600000), Emu(700000), Emu(10992000), Emu(5800000))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.auto_size = None  # let word_wrap handle it

            for i, bullet in enumerate(sd.get("bullets", [])):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                indent = 0
                text = bullet
                if bullet.startswith("   "):
                    indent = 1
                    text = bullet.strip()
                p.level = indent
                p.space_after = Pt(4)
                r = p.add_run()
                r.text = ("\u2022 " if indent == 0 else "\u2013 ") + text
                r.font.size = font_size if indent == 0 else Pt(font_size.pt - 2)
                r.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

    logger.info(f"PPT generation took {time.time() - start:.2f}s ({len(slides_data)} slides)")
    return prs


# ─── Video Generator helpers ─────────────────────────────────────────────────
DEFAULT_TTS_GENERATE_URL = "http://172.23.198.77:8001/tts"


def pptx_to_images(pptx_path: str, output_dir: str) -> list[str]:
    """Export PPTX slides as PNG images using PowerPoint COM (Windows only)."""
    import comtypes.client
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
    presentation.SaveAs(output_dir, 18)  # 18 = ppSaveAsPNG
    presentation.Close()
    powerpoint.Quit()

    img_folder = output_dir
    if not os.path.exists(os.path.join(img_folder, "Slide1.PNG")):
        for item in os.listdir(output_dir):
            candidate = os.path.join(output_dir, item)
            if os.path.isdir(candidate):
                img_folder = candidate
                break

    images = sorted(
        [os.path.join(img_folder, f) for f in os.listdir(img_folder) if f.lower().endswith(".png")],
        key=lambda x: int(re.search(r'(\d+)', os.path.basename(x)).group()),
    )
    return images


# ─── Pronunciation Dictionary ────────────────────────────────────────────────
# Maps Hindi/Indian terms to phonetic English so TTS pronounces them correctly.
# Add new entries as needed. Keys are case-insensitive during matching.
PRONUNCIATION_MAP = {
    # Government schemes & Hindi words
    "Pradhan Mantri": "Pruh-dhaan Mun-tree",
    "Mudra": "Mood-raa",
    "Yojna": "Yoj-naa",
    "Yojana": "Yoj-naa",
    "Kisan": "Kis-aan",
    "Samman": "Sum-maan",
    "Nidhi": "Nid-hee",
    "Fasal": "Fuh-sul",
    "Bima": "Bee-maa",
    "Suraksha": "Su-ruk-shaa",
    "Jeevan": "Jee-vun",
    "Jyoti": "Jyo-tee",
    "Jan Dhan": "Jun Dhun",
    "Swamitva": "Swaa-mit-vuh",
    "Awas": "Aa-waas",
    "Ujjwala": "Uj-waa-laa",
    "Ayushman": "Aa-yush-maan",
    "Bharat": "Bhaa-rut",
    "Atmanirbhar": "Aat-muh-nir-bhar",
    "Rozgar": "Roz-gaar",
    "Udyam": "Ud-yum",
    "Shishu": "Shee-shoo",
    "Kishore": "Kish-ore",
    "Tarun": "Tuh-roon",
    "Stree Shakti": "Stree Shuk-tee",
    "Mahila": "Muh-hee-laa",
    "Swavalamban": "Swuh-vuh-lum-bun",
    "Samriddhi": "Sum-rid-dhee",
    "Sukanya": "Su-kun-yaa",
    "Vikas": "Vik-aas",
    "Swarojgar": "Swuh-roj-gaar",
    "Unnati": "Un-nuh-tee",
    "Sahay": "Suh-haay",
    "Tatkal": "Tut-kaal",
    "Aadhaar": "Aa-dhaar",
    "Seva": "Say-vaa",
    # Banking abbreviations
    "PNB": "P N B",
    "MSME": "M S M E",
    "DPIIT": "D P I I T",
    "CGTMSE": "C G T M S E",
    "IRMD": "I R M D",
    "NBFC": "N B F C",
    "NPA": "N P A",
    "KYC": "K Y C",
    "EMI": "E M I",
    "ROI": "R O I",
    "CIC": "C I C",
    "CIBIL": "Sih-bil",
    # Currency & units
    "Lakh": "Laakh",
    "Crore": "Kuh-roar",
    "₹": "Rupees ",
}


def _apply_pronunciation_fixes(text: str) -> str:
    """Replace Hindi/Indian terms with phonetic equivalents for better TTS pronunciation."""
    for term, phonetic in PRONUNCIATION_MAP.items():
        # Case-insensitive replacement preserving sentence flow
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        text = pattern.sub(phonetic, text)
    return text


def _clean_tts_text(text: str) -> str:
    """Strip voiceover labels, fix pronunciation, and normalize characters before sending to TTS."""
    # Remove common label prefixes like "Voiceover:", "Narration:", "Voiceover/Narration:"
    text = re.sub(r'^(voiceover|narration|voiceover\s*/\s*narration)\s*:\s*', '', text, flags=re.IGNORECASE)
    # Replace slash with " or " so TTS doesn't read it as a symbol
    text = text.replace('/', ' or ')
    # Apply pronunciation fixes for Hindi/Indian terms
    text = _apply_pronunciation_fixes(text)
    return text.strip()


def generate_tts_audio(text: str, tts_url: str, output_path: str) -> str:
    start = time.time()
    resp = http_requests.post(tts_url, json={"text": text}, timeout=300)
    resp.raise_for_status()
    with open(output_path, "wb") as f:
        f.write(resp.content)
    logger.info(f"TTS audio generation took {time.time() - start:.2f}s")
    return output_path


def get_audio_duration(audio_path: str) -> float:
    import subprocess
    result = subprocess.run(
        ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
         '-of', 'default=noprint_wrappers=1:nokey=1', audio_path],
        capture_output=True, text=True,
    )
    output = result.stdout.strip()
    if not output:
        raise ValueError(f"Could not get duration for: {audio_path}")
    return float(output)


def build_video(images: list[str], audio_files: list[str], output_path: str) -> str:
    tmp_dir = os.path.dirname(output_path)
    concat_file = os.path.join(tmp_dir, "concat.txt")

    with open(concat_file, "w") as f:
        for img, audio in zip(images, audio_files):
            duration = get_audio_duration(audio)
            f.write(f"file '{img}'\n")
            f.write(f"duration {duration}\n")
        f.write(f"file '{images[-1]}'\n")

    audio_concat_file = os.path.join(tmp_dir, "audio_concat.txt")
    with open(audio_concat_file, "w") as f:
        for audio in audio_files:
            f.write(f"file '{audio}'\n")

    merged_audio = os.path.join(tmp_dir, "merged_audio.wav")
    os.system(f'ffmpeg -y -f concat -safe 0 -i "{audio_concat_file}" -c copy "{merged_audio}"')
    os.system(
        f'ffmpeg -y -f concat -safe 0 -i "{concat_file}" '
        f'-i "{merged_audio}" '
        f'-vf "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2" '
        f'-c:v libx264 -pix_fmt yuv420p -c:a aac -shortest "{output_path}"'
    )
    return output_path


# ─── Routes ──────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(name="index.html", request=request)


@app.post("/api/upload-pdf")
async def upload_pdf(file: UploadFile = File(...)):
    """Upload PDF or DOCX and return metadata."""
    filepath = os.path.join(UPLOAD_DIR, file.filename)
    with open(filepath, "wb") as f:
        shutil.copyfileobj(file.file, f)
    ext = os.path.splitext(file.filename)[1].lower()
    if ext == ".docx":
        page_count = get_docx_page_count(filepath)
    else:
        page_count = get_pdf_page_count(filepath)
    return {"status": "ok", "filename": file.filename, "filepath": filepath, "pages": page_count}


@app.get("/api/list-uploads")
async def list_uploads():
    """List uploaded DOCX and PDF files."""
    files = [
        f for f in os.listdir(UPLOAD_DIR)
        if os.path.splitext(f)[1].lower() in (".pdf", ".docx", ".pptx")
    ]
    return {"files": files}


@app.get("/api/extract-topics")
async def extract_topics(filename: str):
    """Return topic/scheme titles found in an uploaded DOCX."""
    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        return JSONResponse(status_code=404, content={"status": "error", "message": "File not found"})
    ext = os.path.splitext(filename)[1].lower()
    if ext != ".docx":
        return JSONResponse(status_code=400, content={"status": "error", "message": "Only DOCX supported for topic extraction"})
    titles = extract_docx_topics(filepath)
    return {"status": "ok", "topics": [{"num": i + 1, "title": t} for i, t in enumerate(titles)]}


@app.post("/api/test-connection")
async def test_connection(request: Request):
    data = await request.json()
    ollama_url = data.get("ollama_url", DEFAULT_OLLAMA_URL)
    try:
        r = http_requests.get(f"{ollama_url}/api/tags", timeout=5)
        models = [m["name"] for m in r.json().get("models", [])]
        return {"status": "ok", "models": models}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e)})


@app.post("/api/scheme-to-ppt")
async def scheme_to_ppt(request: Request):
    """
    Single API: scheme_name + filename → extract → Ollama JSON → PPT download.
    Body: { filename, scheme_name, ollama_url (optional), model (optional) }
    Returns: downloadable .pptx
    """
    data        = await request.json()
    filename    = data.get("filename")
    scheme_name = data.get("scheme_name")
    ollama_url  = data.get("ollama_url", DEFAULT_OLLAMA_URL)
    model       = data.get("model", DEFAULT_MODEL)
    out_dir     = os.path.join(os.path.dirname(__file__), "output", "json")

    if not filename or not scheme_name:
        return JSONResponse(status_code=400, content={"status": "error", "message": "filename and scheme_name are required"})

    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        return JSONResponse(status_code=404, content={"status": "error", "message": "File not found"})

    raw_content = extract_uploaded_file_text(filename, title=scheme_name)
    if not raw_content or len(raw_content) < 20:
        return JSONResponse(status_code=400, content={"status": "error", "message": f"Could not extract content for: {scheme_name}"})

    logger.info(f"[scheme-to-ppt] '{scheme_name}' | {len(raw_content)} chars | model: {model}")
    result = generate_slides(scheme_name, raw_content, DEFAULT_SYSTEM_PROMPT, ollama_url, model)
    if not result or "slides" not in result:
        return JSONResponse(status_code=500, content={"status": "error", "message": "Ollama failed to generate slides"})

    result["scheme_title"] = scheme_name

    os.makedirs(out_dir, exist_ok=True)
    safe = re.sub(r'[<>:"/\\|?*]', '', scheme_name)[:60].strip()
    safe = safe.encode("ascii", "ignore").decode("ascii").strip()
    json_file = os.path.join(out_dir, f"{safe}.json")
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    logger.info(f"[scheme-to-ppt] JSON saved → {json_file}")

    slides = [
        s for s in result["slides"]
        if s.get("type") in ["title", "thank_you", "section_title"]
        or (s.get("bullets") and len(s["bullets"]) > 0)
    ]
    prs    = build_ppt(slides)
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    fname = f"{safe}.pptx"
    logger.info(f"[scheme-to-ppt] Done → {len(slides)} slides → {fname}")

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{fname}"',
            "X-Slide-Count":       str(len(slides)),
            "X-Json-File":         json_file,
        },
    )


@app.post("/api/generate")
async def generate_ppt(request: Request):
    data = await request.json()
    topic = data.get("topic")
    ollama_url = data.get("ollama_url", DEFAULT_OLLAMA_URL)
    model = data.get("model", DEFAULT_MODEL)
    pdf_path = data.get("pdf_path", DEFAULT_PDF)
    system_prompt = data.get("system_prompt") or DEFAULT_SYSTEM_PROMPT
    uploaded_filename = data.get("uploaded_filename")

    title = topic["title"]

    # Extract FULL file content - LLM will locate scheme by title
    if uploaded_filename:
        raw_content = extract_uploaded_file_text(uploaded_filename, title=title)
        if not raw_content:
            return JSONResponse(
                status_code=400,
                content={"status": "error", "message": f"Could not read uploaded file: {uploaded_filename}"},
            )
    else:
        pages = _read_pdf(pdf_path)
        raw_content = "\n".join(pages)

    if len(raw_content) < 50:
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": "Content too short from source file"},
        )

    logger.info(f"Generating slides for: {title} (full doc, title-based extraction) | Model: {model}")
    result = generate_slides(title, raw_content, system_prompt, ollama_url, model)

    if not result or "slides" not in result:
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": "Failed to generate slides from Ollama"},
        )

    # Filter slides that are completely empty
    slides = [
        s for s in result["slides"]
        if s.get("type") in ["title", "thank_you", "section_title"]
        or (s.get("bullets") and len(s["bullets"]) > 0)
    ]

    prs = build_ppt(slides)
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    safe_title = re.sub(r'[<>:"/\\|?*]', '', title)[:80].strip()
    safe_title = safe_title.encode('ascii', 'ignore').decode('ascii').strip()
    filename = f"{topic['num']:02d}_{safe_title}.pptx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/generate-video")
async def generate_video(request: Request):
    """Generate PPT, convert to images, add TTS audio, export as video."""
    data = await request.json()
    topic = data.get("topic")
    ollama_url = data.get("ollama_url", DEFAULT_OLLAMA_URL)
    model = data.get("model", DEFAULT_MODEL)
    pdf_path = data.get("pdf_path", DEFAULT_PDF)
    system_prompt = data.get("system_prompt") or DEFAULT_SYSTEM_PROMPT
    tts_url = DEFAULT_TTS_GENERATE_URL

    title = topic["title"]
    uploaded_filename = data.get("uploaded_filename")

    if uploaded_filename:
        raw_content = extract_uploaded_file_text(uploaded_filename, title=title)
        if not raw_content:
            return JSONResponse(status_code=400, content={"status": "error", "message": f"Could not read uploaded file: {uploaded_filename}"})
    else:
        pages = _read_pdf(pdf_path)
        raw_content = "\n".join(pages)

    if len(raw_content) < 50:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Content too short"})

    logger.info(f"Generating video for: {title}")
    result = generate_slides(title, raw_content, system_prompt, ollama_url, model)
    if not result or "slides" not in result:
        return JSONResponse(status_code=500, content={"status": "error", "message": "Failed to generate slides"})

    slides = [
        s for s in result["slides"]
        if s.get("type") in ["title", "thank_you"]
        or (s.get("bullets") and len(s["bullets"]) > 0)
    ]

    tmp_dir = tempfile.mkdtemp(prefix="ppt_video_")
    try:
        prs = build_ppt(slides)
        pptx_path = os.path.join(tmp_dir, "presentation.pptx")
        prs.save(pptx_path)

        img_dir = os.path.join(tmp_dir, "images")
        os.makedirs(img_dir, exist_ok=True)
        images = pptx_to_images(os.path.abspath(pptx_path), os.path.abspath(img_dir))
        logger.info(f"Exported {len(images)} slide images")

        audio_dir = os.path.join(tmp_dir, "audio")
        os.makedirs(audio_dir, exist_ok=True)
        audio_files = []
        for i, sd in enumerate(slides):
            vo = _clean_tts_text(sd.get("voiceover", "") or sd.get("title", ""))
            audio_path = os.path.join(audio_dir, f"slide_{i:03d}.wav")
            generate_tts_audio(vo, tts_url, audio_path)
            audio_files.append(audio_path)
            logger.info(f"Generated audio for slide {i + 1}")

        video_path = os.path.join(tmp_dir, "output.mp4")
        build_video(images, audio_files, video_path)

        if not os.path.exists(video_path):
            return JSONResponse(status_code=500, content={"status": "error", "message": "Video generation failed"})

        with open(video_path, "rb") as f:
            video_bytes = f.read()

        safe_title = re.sub(r'[<>:"/\\|?*]', '', title)[:80].strip()
        safe_title = safe_title.encode('ascii', 'ignore').decode('ascii').strip()
        filename = f"{topic['num']:02d}_{safe_title}.mp4"

        return StreamingResponse(
            io.BytesIO(video_bytes),
            media_type="video/mp4",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post("/api/pptx-to-video")
async def pptx_to_video(request: Request):
    """
    Accept an already-built PPTX (with voiceover in slide notes),
    generate TTS audio per slide, and return an MP4 video.
    Body: { "filename": "uploaded.pptx", "tts_url": "..." (optional) }
    """
    data = await request.json()
    filename = data.get("filename")
    tts_url = data.get("tts_url", DEFAULT_TTS_GENERATE_URL)

    if not filename:
        return JSONResponse(status_code=400, content={"status": "error", "message": "filename is required"})

    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        return JSONResponse(status_code=404, content={"status": "error", "message": "PPTX file not found in uploads"})

    if not filename.lower().endswith(".pptx"):
        return JSONResponse(status_code=400, content={"status": "error", "message": "Only .pptx files are supported"})

    # Read voiceover text from slide notes
    prs = Presentation(filepath)
    slide_notes = []
    for slide in prs.slides:
        notes_frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
        text = notes_frame.text.strip() if notes_frame else ""
        slide_notes.append(text)

    if not any(slide_notes):
        return JSONResponse(status_code=400, content={"status": "error", "message": "No voiceover notes found in the PPTX"})

    tmp_dir = tempfile.mkdtemp(prefix="pptx_video_")
    try:
        img_dir = os.path.join(tmp_dir, "images")
        os.makedirs(img_dir, exist_ok=True)
        images = pptx_to_images(os.path.abspath(filepath), os.path.abspath(img_dir))
        logger.info(f"Exported {len(images)} slide images from {filename}")

        if len(images) != len(slide_notes):
            logger.warning(f"Image count ({len(images)}) != slide notes count ({len(slide_notes)}), aligning to min")
            count = min(len(images), len(slide_notes))
            images = images[:count]
            slide_notes = slide_notes[:count]

        audio_dir = os.path.join(tmp_dir, "audio")
        os.makedirs(audio_dir, exist_ok=True)
        audio_files = []
        for i, vo in enumerate(slide_notes):
            vo = _clean_tts_text(vo) if vo else f"Slide {i + 1}"
            audio_path = os.path.join(audio_dir, f"slide_{i:03d}.wav")
            generate_tts_audio(vo, tts_url, audio_path)
            audio_files.append(audio_path)
            logger.info(f"Generated audio for slide {i + 1}/{len(slide_notes)}")

        video_path = os.path.join(tmp_dir, "output.mp4")
        build_video(images, audio_files, video_path)

        if not os.path.exists(video_path):
            return JSONResponse(status_code=500, content={"status": "error", "message": "Video generation failed"})

        with open(video_path, "rb") as f:
            video_bytes = f.read()

        safe_name = re.sub(r'[<>:"/\\|?*]', '', os.path.splitext(filename)[0])[:80].strip()
        safe_name = safe_name.encode('ascii', 'ignore').decode('ascii').strip()
        out_filename = f"{safe_name}.mp4"

        return StreamingResponse(
            io.BytesIO(video_bytes),
            media_type="video/mp4",
            headers={"Content-Disposition": f'attachment; filename="{out_filename}"'},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ─── Run ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    print("=" * 60)
    print("  PPT Generator - FastAPI Server v2.0")
    print("  Open: http://localhost:8000")
    print(f"  TTS Generate: {DEFAULT_TTS_GENERATE_URL}")
    print("  Features: PDF/DOCX Upload | Custom Prompt | PPT Download")
    print("=" * 60)
    uvicorn.run(app, host="0.0.0.0", port=8000)
