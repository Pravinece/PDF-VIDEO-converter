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
import copy
import time
import unicodedata
import requests as http_requests
from functools import lru_cache

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("ppt_generator")

sys.stdout.reconfigure(encoding='utf-8')

import asyncio
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from typing import List
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
    - For English output, keep bullet wording as close to the source as possible. For non-English output, translate every bullet into the selected language while preserving its exact meaning and details.
   - Do NOT combine two separate lines into one bullet.
   - Do NOT summarize multiple conditions into one bullet.
   - Do NOT drop footnotes, remarks, parenthetical text, or "minor" details.
   - Include full text of conditions including "i.e." explanations and parenthetical qualifications.

4. PRESERVE ALL DATA EXACTLY:
   - All numbers, percentages, dates, monetary values, interest rates, names, references, circular numbers, CIC scores — EXACTLY as they appear.
   - Do NOT round numbers, do NOT paraphrase amounts, do NOT skip reference numbers.
   - Preserve parenthetical text. Example: "(For acceptance of lower CIC score, guidelines issued by IRMD shall be referred)" must be included.

5. ONE slide per parameter/section. Slide TITLE must preserve the exact parameter meaning from the source and be translated into the selected language when non-English output is requested.
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

# ─── Language selection ──────────────────────────────────────────────────────
# Everything the LLM produces (slide TITLES/headings, BULLET points, and VOICEOVER
# notes) is generated in the chosen language. English is the default. The directive
# below is appended to whatever system prompt is
# used so the behaviour is consistent across every generation endpoint.
DEFAULT_LANGUAGE = "english"
SUPPORTED_LANGUAGES = {"english", "hindi", "tamil", "telugu", "malayalam"}
_LANGUAGE_NAMES = {
    "english": "English",
    "hindi": "Hindi",
    "tamil": "Tamil",
    "telugu": "Telugu",
    "malayalam": "Malayalam",
}
_LANGUAGE_SCRIPTS = {
    "english": "Latin script",
    "hindi": "Devanagari script",
    "tamil": "Tamil script",
    "telugu": "Telugu script",
    "malayalam": "Malayalam script",
}


def _normalize_language(language: str | None) -> str:
    """Return a canonical supported language key. Unknown/empty → default."""
    if not language:
        return DEFAULT_LANGUAGE
    key = str(language).strip().lower()
    aliases = {
        "en": "english", "eng": "english", "english": "english",
        "hi": "hindi", "hin": "hindi", "hindi": "hindi",
        "हिंदी": "hindi", "हिन्दी": "hindi",
        "ta": "tamil", "tam": "tamil", "tamil": "tamil", "தமிழ்": "tamil",
        "te": "telugu", "tel": "telugu", "telugu": "telugu", "తెలుగు": "telugu",
        "ml": "malayalam", "mal": "malayalam", "malayalam": "malayalam", "മലയാളം": "malayalam",
    }
    return aliases.get(key, DEFAULT_LANGUAGE if key not in SUPPORTED_LANGUAGES else key)


_HINDI_DIRECTIVE = """

=============================================
LANGUAGE REQUIREMENT — OUTPUT IN HINDI (हिंदी)
=============================================
- Produce ALL human-readable text in HINDI using Devanagari script: every slide
  TITLE/heading, every BULLET point, every subtitle, and the entire VOICEOVER.
- Keep numbers, percentages, dates, monetary values, scheme names, abbreviations
  (e.g. MSME, DPIIT, CIC, EMI), and proper nouns intact — write them as they appear
  (you may transliterate common Hindi terms naturally).
- Do NOT output English sentences for titles, bullets, or voiceover — translate the
  meaning into natural, fluent Hindi while preserving every fact and detail.
- The JSON structure, keys, and the "type" values ("title", "section_title",
  "content", "thank_you") MUST stay in English. ONLY the human-readable VALUES
  (title, subtitle, bullets, voiceover) are in Hindi.
"""

_ENGLISH_DIRECTIVE = """

=============================================
LANGUAGE REQUIREMENT — OUTPUT IN ENGLISH
=============================================
- Produce ALL human-readable text (titles, bullets, subtitles, voiceover) in English.
"""

_INDIAN_LANGUAGE_DIRECTIVES = {
    "tamil": """
LANGUAGE REQUIREMENT — OUTPUT IN TAMIL (தமிழ்)
- Write all human-readable slide titles, headings, bullets, subtitles, and voiceover in fluent Tamil using Tamil script.
- Preserve numbers, percentages, dates, monetary values, scheme names, abbreviations, and proper nouns accurately.
- Keep JSON keys and slide type values in English; translate only human-readable values.
""",
    "telugu": """
LANGUAGE REQUIREMENT — OUTPUT IN TELUGU (తెలుగు)
- Write all human-readable slide titles, headings, bullets, subtitles, and voiceover in fluent Telugu using Telugu script.
- Preserve numbers, percentages, dates, monetary values, scheme names, abbreviations, and proper nouns accurately.
- Keep JSON keys and slide type values in English; translate only human-readable values.
""",
    "malayalam": """
LANGUAGE REQUIREMENT — OUTPUT IN MALAYALAM (മലയാളം)
- Write all human-readable slide titles, headings, bullets, subtitles, and voiceover in fluent Malayalam using Malayalam script.
- Preserve numbers, percentages, dates, monetary values, scheme names, abbreviations, and proper nouns accurately.
- Keep JSON keys and slide type values in English; translate only human-readable values.
""",
}


def _language_directive(language: str | None) -> str:
    """Return the system-prompt suffix that forces the LLM to output in `language`."""
    lang = _normalize_language(language)
    if lang == "hindi":
        return _HINDI_DIRECTIVE
    return _INDIAN_LANGUAGE_DIRECTIVES.get(lang, _ENGLISH_DIRECTIVE)


def _apply_language(system_prompt: str, language: str | None) -> str:
    """Append the language directive to a system prompt."""
    return (system_prompt or "") + _language_directive(language)


# ─── Colors ──────────────────────────────────────────────────────────────────
DARK_RED   =  RGBColor(0xFF, 0xCB, 0x08)   # #121212 near-black
GOLD       =  RGBColor(0x12, 0x12, 0x12)# #FFCB08 yellow
AMBER      = RGBColor(0xE5, 0xE5, 0xE5)  # #E5E5E5 light gray panel
DARK_GRAY  = RGBColor(0x12, 0x12, 0x12)  # #121212 near-black body text
LIGHT_GRAY = RGBColor(0xE5, 0xE5, 0xE5)  # #E5E5E5 light gray

# Explicit theme colors for the consistent yellow-header / white-text-box design.
HEADER_YELLOW = RGBColor(0xFF, 0xCB, 0x08)  # #FFCB08 — the consistent header/title band
HEADER_TEXT   = RGBColor(0x12, 0x12, 0x12)  # near-black title text (readable on yellow)
BOX_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)  # white text-box fill
BOX_BORDER    = RGBColor(0xFF, 0xFF, 0xFF)  # white text-box border
BOX_TEXT      = RGBColor(0x12, 0x12, 0x12)  # dark, readable body text in the box


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


def _build_voiceover_app(bullets: list, title: str = "", language: str | None = None) -> str:
    """Fallback voiceover builder used when slides are split post-processing."""
    if not bullets:
        return ""
    lang = _normalize_language(language)
    if lang != "english":
        return _build_narration_local(title, bullets, lang)
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


def _enforce_8_bullets(result: dict, language: str | None = None) -> dict:
    lang = _normalize_language(language)
    conjunction = {
        "hindi": " और ",
        "tamil": " மற்றும் ",
        "telugu": " మరియు ",
        "malayalam": " കൂടാതെ ",
    }.get(lang, " and ")
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
            merged_title = conjunction.join(pending_titles + [slide["title"]])
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

        merged_title = conjunction.join(clubbed_titles)
        if len(clubbed_bullets) <= 8:
            out.append({"type": "content", "title": merged_title,
                        "bullets": clubbed_bullets, "voiceover": clubbed_voiceover})
        else:
            chunks = [clubbed_bullets[j:j+8] for j in range(0, len(clubbed_bullets), 8)]
            n = len(chunks)
            for pi, chunk in enumerate(chunks, 1):
                chunk_title = f"{merged_title} ({pi}/{n})" if n > 1 else merged_title
                chunk_vo = _build_voiceover_app(chunk, chunk_title, lang)
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


def _make_prompt(title: str, raw_text: str, part: int = 1, total: int = 1,
                 language: str | None = None) -> str:
    part_note = f" (Part {part} of {total} — continue from previous slides, DO NOT repeat slides already generated)" if total > 1 else ""
    lang = _normalize_language(language)
    language_note = (
        f"- Write every human-readable value ONLY in fluent {_LANGUAGE_NAMES[lang]} using {_LANGUAGE_SCRIPTS[lang]}.\n"
        "- Translate all source prose into that language; do not copy sentences in another language.\n"
        "- Keep only numbers, abbreviations, scheme names, and proper nouns in their original form when needed.\n"
        if lang != "english" else
        "- Write every human-readable value in English.\n"
    )
    return f"""Scheme Name: {title}{part_note}

Full Document Content:
---
{raw_text}
---

CRITICAL INSTRUCTIONS:
- From the full document above, locate ONLY the section for the scheme named: "{title}".
- Include ALL content: main parameter table AND all sub-sections/appendices.
- Do NOT include content from any other scheme.
- Title slide = scheme name (translate it unless it is a proper name). For each sub-section/appendix add a section_title slide.
- ONE slide per table parameter (Objective, Eligibility, Purpose, Quantum, Margin, etc.).
- Translate each parameter name into the selected output language without changing its meaning.
- Preserve every source fact and detail. Translate source wording when required by the output language; never omit, merge, or invent information.
- If content is too large for one slide, SPLIT: e.g. "Collateral Security (1/2)", "Collateral Security (2/2)".
- NEVER repeat the same bullet text more than once in a slide.
{language_note}
- Respond ONLY in valid JSON."""


def generate_slides(title: str, raw_content: str, system_prompt: str, ollama_url: str, model: str,
                    language: str | None = None) -> dict | None:
    # Force the LLM to emit titles/bullets/voiceover in the requested language.
    system_prompt = _apply_language(system_prompt, language)
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
        prompt = _make_prompt(title, chunk, part_idx, total, language=language)
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
    result = _enforce_8_bullets(result, language=language)
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


# ─── Structured DOCX extraction (headings/subheadings + embedded images) ─────
_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REL_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _looks_like_subheading(text: str, next_text: str) -> bool:
    """Heuristic: a short line that introduces a block (acts as a subheading) rather
    than being a full body sentence. Matches lines like 'Tenure Eligibility Check:' or
    'Parent and Child Loan dedupe' that are followed by detail paragraphs."""
    if not text or len(text) > 90:
        return False
    words = text.split()
    if len(words) > 12:
        return False
    # Ends with a colon → almost always a subheading label.
    if text.rstrip().endswith(":"):
        return True
    # Short, title-ish line (no trailing period, has following detail content).
    if not text.rstrip().endswith(".") and next_text and len(next_text) > len(text):
        # Avoid treating a normal short sentence as a heading: require it be
        # reasonably "label-like" (few words, mostly capitalized-ish start).
        if len(words) <= 10:
            return True
    return False


def extract_docx_sections_with_images(docx_path: str):
    """Walk a DOCX in document order and produce a list of sections. Each section is:
        {
          "title":  <heading/subheading text>,
          "level":  "heading" | "subheading",
          "body":   [list of body text lines],
          "images": [list of extracted image file paths for this section],
        }
    Embedded images are written to a temp dir and attributed to the section whose
    heading/subheading most recently preceded them. Sections with no images get [].
    Returns (doc_title, sections, image_tmp_dir).
    """
    from xml.etree import ElementTree as ET
    import zipfile

    # ── Pull media bytes + the rId → media-path relationship map ──────────────
    media_bytes: dict[str, bytes] = {}
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                media_bytes[name] = z.read(name)
        try:
            rels_xml = z.read("word/_rels/document.xml.rels")
        except KeyError:
            rels_xml = b""
        doc_xml = z.read("word/document.xml")

    rid_to_target: dict[str, str] = {}
    if rels_xml:
        rels_root = ET.fromstring(rels_xml)
        for rel in rels_root:
            rid = rel.get("Id")
            target = rel.get("Target", "")
            if rid and target:
                # Normalise to the zip path form 'word/media/imageN.png'
                rid_to_target[rid] = "word/" + target.lstrip("/") if not target.startswith("word/") else target

    # ── Write media to a temp dir so python-pptx can embed them ───────────────
    img_tmp = tempfile.mkdtemp(prefix="docx_embed_img_")
    saved_media: dict[str, str] = {}  # zip-path → local file path
    for zip_path, data in media_bytes.items():
        local = os.path.join(img_tmp, os.path.basename(zip_path))
        with open(local, "wb") as f:
            f.write(data)
        saved_media[zip_path] = local

    def _para_text(p_elem) -> str:
        return "".join(t.text or "" for t in p_elem.iter(f"{{{_WORD_NS}}}t")).strip()

    def _para_is_bold(p_elem) -> bool:
        runs = list(p_elem.iter(f"{{{_WORD_NS}}}r"))
        texted = [r for r in runs if "".join(t.text or "" for t in r.iter(f"{{{_WORD_NS}}}t")).strip()]
        if not texted:
            return False
        for r in texted:
            rpr = r.find(f"{{{_WORD_NS}}}rPr")
            if rpr is None or rpr.find(f"{{{_WORD_NS}}}b") is None:
                return False
        return True

    def _para_style(p_elem) -> str:
        ppr = p_elem.find(f"{{{_WORD_NS}}}pPr")
        if ppr is None:
            return ""
        pstyle = ppr.find(f"{{{_WORD_NS}}}pStyle")
        return pstyle.get(f"{{{_WORD_NS}}}val", "") if pstyle is not None else ""

    def _para_images(p_elem) -> list[str]:
        paths = []
        for blip in p_elem.iter(f"{{{_DRAW_NS}}}blip"):
            embed = blip.get(f"{{{_REL_NS}}}embed")
            if not embed:
                continue
            target = rid_to_target.get(embed)
            if target and target in saved_media:
                paths.append(saved_media[target])
        return paths

    root = ET.fromstring(doc_xml)
    body = root.find(f"{{{_WORD_NS}}}body")
    if body is None:
        return "", [], img_tmp

    # First pass: collect ordered paragraph records (text, style, bold, images).
    records = []
    for child in body:
        if child.tag != f"{{{_WORD_NS}}}p":
            continue
        records.append({
            "text":   _para_text(child),
            "style":  _para_style(child),
            "bold":   _para_is_bold(child),
            "images": _para_images(child),
        })

    def _is_heading(rec) -> bool:
        return rec["style"].lower().startswith("heading") or (bool(rec["text"]) and rec["bold"])

    # Document title = first heading (or first non-empty line).
    doc_title = ""
    for rec in records:
        if rec["text"]:
            doc_title = rec["text"]
            if _is_heading(rec):
                break

    # Second pass: group into sections by heading / subheading.
    sections: list[dict] = []
    current = None
    prev_was_head = False  # the line right after a heading/subheading is its body

    def _next_text(i):
        for j in range(i + 1, len(records)):
            if records[j]["text"]:
                return records[j]["text"]
        return ""

    for i, rec in enumerate(records):
        text = rec["text"]
        imgs = rec["images"]

        is_head = _is_heading(rec) and text
        # A line is only a subheading if it looks like one AND it does not
        # immediately follow another heading/subheading (that line is body —
        # this prevents repeated body lines like 'Applicable for ...' right
        # under a real subheading from being mistaken for subheadings).
        is_sub  = (not is_head) and text and current is not None and \
                  not prev_was_head and _looks_like_subheading(text, _next_text(i))

        if is_head or is_sub:
            current = {
                "title": text,
                "level": "heading" if is_head else "subheading",
                "body": [],
                "images": list(imgs),
            }
            sections.append(current)
            prev_was_head = True
            continue

        # Body paragraph / image-only paragraph.
        if current is None:
            # Content before any heading → start an implicit section.
            current = {"title": doc_title or "Overview", "level": "heading",
                       "body": [], "images": []}
            sections.append(current)
        if text:
            current["body"].append(text)
        if imgs:
            current["images"].extend(imgs)
        if text:
            prev_was_head = False  # a real body line resets the guard

    # Drop a leading section that only repeats the document title with no content.
    sections = [s for s in sections if s["body"] or s["images"] or s["level"] == "subheading"]

    logger.info(f"[docx-sections] '{os.path.basename(docx_path)}': {len(sections)} sections; "
                f"images per section: "
                + ", ".join(f"{s['title'][:25]}={len(s['images'])}" for s in sections))
    return doc_title, sections, img_tmp


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


# ─── Content-relevant image fetching ─────────────────────────────────────────
# Images are fetched from Openverse (Creative Commons, no API key required) using
# keywords derived from the slide's own content. An image is embedded ONLY when a
# relevant match is found; otherwise the slide stays text-only. The whole feature
# is best-effort: any network/parse failure silently falls back to text-only so
# existing behaviour is never broken.

OPENVERSE_API_URL = "https://api.openverse.org/v1/images/"
IMAGE_FETCH_ENABLED = True          # master switch for auto-images
_IMAGE_HTTP_HEADERS = {"User-Agent": "ppt-generator/2.0 (slide illustration)"}
_IMAGE_CACHE_DIR = os.path.join(UPLOAD_DIR, "_auto_images")

# Generic / boilerplate words that carry no visual meaning — excluded from queries
# and from relevance checks so we never match on them alone.
_IMAGE_STOPWORDS = {
    "the", "and", "for", "with", "from", "that", "this", "shall", "must", "not",
    "any", "all", "per", "will", "may", "can", "are", "has", "have", "been", "its",
    "a", "an", "of", "to", "in", "on", "or", "as", "by", "is", "be", "at", "it",
    "scheme", "schemes", "section", "part", "point", "points", "slide", "title",
    "subtitle", "note", "notes", "refer", "guidelines", "guideline", "circular",
    "parameter", "parameters", "condition", "conditions", "clause", "detail",
    "details", "etc", "page", "thank", "you", "objective", "eligibility", "purpose",
    "quantum", "margin", "amount", "value", "crore", "lakh", "rupees", "rs", "min",
    "max", "minimum", "maximum", "year", "years", "date", "no", "yes", "applicable",
    "following", "above", "below", "same", "other", "such", "which", "who", "whom",
    "financial", "finance", "financing", "loan", "credit", "bank", "banking",
    "borrower", "applicant", "entity", "firm", "company", "limited", "under",
    "including", "include", "included", "based", "given", "required", "requirement",
}

# Maps common scheme/content keywords to clean, highly-illustratable search terms.
# This improves relevance (and avoids odd matches) for recurring banking topics.
_IMAGE_TOPIC_HINTS = {
    "agriculture": "agriculture farming field",
    "agricultural": "agriculture farming field",
    "farmer": "farmer agriculture field",
    "farming": "agriculture farming field",
    "crop": "crop field harvest",
    "kisan": "indian farmer agriculture",
    "dairy": "dairy cattle farm",
    "cattle": "cattle livestock farm",
    "livestock": "livestock farm animals",
    "poultry": "poultry farm chickens",
    "fishery": "fishery fish pond",
    "fisheries": "fishery fish pond",
    "irrigation": "irrigation water field",
    "tractor": "tractor farming field",
    "startup": "startup business team office",
    "msme": "small business manufacturing workshop",
    "manufacturing": "manufacturing factory industry",
    "industry": "industry factory manufacturing",
    "business": "business office team",
    "export": "export shipping cargo containers",
    "import": "import shipping cargo port",
    "housing": "housing residential building home",
    "home": "house home building",
    "education": "education student classroom",
    "student": "student education study",
    "health": "healthcare hospital medical",
    "medical": "healthcare hospital medical",
    "hospital": "hospital healthcare medical",
    "solar": "solar panels renewable energy",
    "energy": "renewable energy power",
    "vehicle": "vehicle car automobile",
    "transport": "transport logistics truck",
    "women": "women entrepreneur business",
    "mahila": "indian women entrepreneur",
    "digital": "digital technology computer",
    "technology": "technology computer digital",
    "insurance": "insurance protection document",
    "pension": "retirement senior pension",
    "gold": "gold jewellery ornaments",
    "property": "property real estate building",
    "warehouse": "warehouse storage logistics",
    "food": "food processing production",
    "textile": "textile fabric manufacturing",
    "water": "water conservation river",
}

# Minimum relevance score (shared meaningful keywords) required to accept an image.
_IMAGE_RELEVANCE_MIN = 1


def _image_keywords(text: str) -> list[str]:
    """Extract meaningful, lowercase keywords from text (stopwords/numbers removed)."""
    words = re.findall(r"[A-Za-z]{3,}", text or "")
    out, seen = [], set()
    for w in words:
        wl = w.lower()
        if wl in _IMAGE_STOPWORDS or wl in seen:
            continue
        seen.add(wl)
        out.append(wl)
    return out


def _build_image_query(title: str, bullets: list) -> tuple[str, set]:
    """Build an image search query + the set of meaningful keywords for relevance.
    Prefers mapped topic hints (cleaner, more illustratable terms) when present."""
    title_kw = _image_keywords(title)
    bullet_kw = _image_keywords(" ".join(bullets or []))
    all_kw = title_kw + bullet_kw
    keyword_set = set(all_kw)

    # If any known topic hint appears, use its curated search phrase for best relevance.
    for kw in all_kw:
        if kw in _IMAGE_TOPIC_HINTS:
            hint = _IMAGE_TOPIC_HINTS[kw]
            keyword_set.update(_image_keywords(hint))
            return hint, keyword_set

    # Otherwise build a focused query from the most informative keywords
    # (title keywords first, then the strongest bullet keywords).
    query_terms = title_kw[:4] or bullet_kw[:4]
    if not query_terms:
        return "", set()
    return " ".join(query_terms), keyword_set


def _is_relevant_image(item: dict, keyword_set: set) -> bool:
    """Return True only if the candidate image shares meaningful keywords with the
    slide content. Prevents random/unrelated images from being used."""
    if not keyword_set:
        return False
    meta = " ".join(filter(None, [
        item.get("title", ""),
        " ".join(item.get("tags_text", "") if isinstance(item.get("tags_text"), str) else ""),
        " ".join(t.get("name", "") for t in (item.get("tags") or []) if isinstance(t, dict)),
        item.get("source", ""),
    ]))
    meta_kw = set(_image_keywords(meta))
    return len(keyword_set & meta_kw) >= _IMAGE_RELEVANCE_MIN


def _download_image(url: str, dest_path: str) -> bool:
    """Download an image and verify it is a valid, decodable raster. Returns success."""
    try:
        resp = http_requests.get(url, headers=_IMAGE_HTTP_HEADERS, timeout=25)
        resp.raise_for_status()
        with open(dest_path, "wb") as f:
            f.write(resp.content)
        # Validate it is a real image python-pptx can embed (jpeg/png/gif/bmp).
        from PIL import Image
        with Image.open(dest_path) as im:
            fmt = (im.format or "").upper()
            if fmt not in ("JPEG", "PNG", "GIF", "BMP"):
                # Convert anything else (e.g. WEBP) to PNG for pptx compatibility.
                converted = os.path.splitext(dest_path)[0] + ".png"
                im.convert("RGB").save(converted, "PNG")
                if converted != dest_path:
                    os.replace(converted, dest_path)
        return True
    except Exception as e:
        logger.info(f"[auto-image] download/validate failed for {url}: {e}")
        try:
            if os.path.exists(dest_path):
                os.remove(dest_path)
        except OSError:
            pass
        return False


@lru_cache(maxsize=256)
def _fetch_relevant_image(query: str, keyword_key: str) -> str | None:
    """Search Openverse for `query` and return a local path to a RELEVANT image,
    or None if nothing relevant is found. Cached per (query, keywords).
    `keyword_key` is a sorted, comma-joined keyword string (hashable for the cache)."""
    if not IMAGE_FETCH_ENABLED or not query:
        return None
    keyword_set = set(k for k in keyword_key.split(",") if k)
    os.makedirs(_IMAGE_CACHE_DIR, exist_ok=True)
    try:
        resp = http_requests.get(
            OPENVERSE_API_URL,
            params={"q": query, "page_size": 8, "license_type": "all",
                    "mature": "false", "aspect_ratio": "wide,square"},
            headers=_IMAGE_HTTP_HEADERS,
            timeout=20,
        )
        resp.raise_for_status()
        results = resp.json().get("results", []) or []
    except Exception as e:
        logger.info(f"[auto-image] search failed for '{query}': {e}")
        return None

    for idx, item in enumerate(results):
        img_url = item.get("url")
        if not img_url:
            continue
        if not _is_relevant_image(item, keyword_set):
            continue
        safe_q = re.sub(r"[^A-Za-z0-9]+", "_", query)[:40].strip("_") or "img"
        dest = os.path.join(_IMAGE_CACHE_DIR, f"{safe_q}_{idx}.img")
        if _download_image(img_url, dest):
            logger.info(f"[auto-image] matched '{query}' → {item.get('title', '')[:60]}")
            return dest

    logger.info(f"[auto-image] no relevant image for '{query}' — slide stays text-only")
    return None


def get_slide_image(title: str, bullets: list) -> str | None:
    """Return a local path to a content-relevant image for this slide, or None.
    Only non-random, keyword-matched images are returned."""
    if not IMAGE_FETCH_ENABLED:
        return None
    query, keyword_set = _build_image_query(title, bullets)
    if not query:
        return None
    keyword_key = ",".join(sorted(keyword_set))
    try:
        return _fetch_relevant_image(query, keyword_key)
    except Exception as e:
        logger.info(f"[auto-image] unexpected failure: {e}")
        return None


# Minimum shared keywords required for an UPLOADED image's filename to be
# considered relevant to a slide's content.
_UPLOAD_RELEVANCE_MIN = 1


def _filename_keywords(path: str) -> list[str]:
    """Derive meaningful keywords from an uploaded image's filename.
    e.g. 'solar-panel_subsidy_2024.jpg' -> ['solar', 'panel', 'subsidy']."""
    base = os.path.splitext(os.path.basename(path or ""))[0]
    # Split on non-letters so separators/digits are dropped, then filter stopwords.
    return _image_keywords(re.sub(r"[^A-Za-z]+", " ", base))


# EXIF tag ids we read for a human title/description/keywords.
_EXIF_IMAGE_DESCRIPTION = 0x010E
_EXIF_XP_TITLE          = 0x9C9B
_EXIF_XP_COMMENT        = 0x9C9C
_EXIF_XP_SUBJECT        = 0x9C9F  # not standard but seen; harmless if absent
_EXIF_XP_KEYWORDS       = 0x9C9E
# Sidecar caption file extensions searched next to an image (same base name).
_CAPTION_SIDECAR_EXTS = (".txt", ".caption", ".cap", ".md", ".json")


def _decode_xp_tag(value) -> str:
    """Windows XP* EXIF tags are UTF-16LE byte arrays/tuples. Decode defensively."""
    try:
        if isinstance(value, (bytes, bytearray)):
            return bytes(value).decode("utf-16-le", errors="ignore").replace("\x00", "").strip()
        if isinstance(value, (tuple, list)):
            return bytes(bytearray(value)).decode("utf-16-le", errors="ignore").replace("\x00", "").strip()
        if isinstance(value, str):
            return value.replace("\x00", "").strip()
    except Exception:
        pass
    return ""


def _image_embedded_text(path: str) -> str:
    """Best-effort extraction of an image's TITLE / CAPTION / DESCRIPTION / KEYWORDS
    from its embedded metadata (EXIF, PNG text chunks, XMP). Returns a single text
    blob (may be empty). Never raises — any failure yields ''.

    Note: true pixel-level 'visual content' recognition needs a vision model which is
    not available here; this captures the descriptive signals that authoring tools
    embed (image title, caption, subject, keywords) plus XMP dc:* fields.
    """
    parts: list[str] = []
    try:
        from PIL import Image
        with Image.open(path) as im:
            # 1) EXIF description + Windows XP title/subject/comment/keywords.
            try:
                exif = im.getexif()
                for tag in (_EXIF_IMAGE_DESCRIPTION,):
                    v = exif.get(tag)
                    if isinstance(v, str) and v.strip():
                        parts.append(v.strip())
                for tag in (_EXIF_XP_TITLE, _EXIF_XP_SUBJECT, _EXIF_XP_COMMENT, _EXIF_XP_KEYWORDS):
                    dv = _decode_xp_tag(exif.get(tag))
                    if dv:
                        parts.append(dv)
            except Exception:
                pass

            # 2) PNG / generic text chunks exposed via Image.info
            #    (keys like Title, Description, Comment, Keywords, Author).
            info = getattr(im, "info", {}) or {}
            for key in ("Title", "title", "Description", "description",
                        "Comment", "comment", "Keywords", "keywords",
                        "Subject", "subject", "Caption", "caption"):
                val = info.get(key)
                if isinstance(val, str) and val.strip():
                    parts.append(val.strip())

            # 3) XMP packet (JPEG/TIFF/PNG) — pull dc:title / dc:description / dc:subject.
            xmp = info.get("xmp")
            if isinstance(xmp, (bytes, bytearray)):
                try:
                    xmp = xmp.decode("utf-8", errors="ignore")
                except Exception:
                    xmp = ""
            if isinstance(xmp, str) and xmp:
                for m in re.findall(r"<dc:(?:title|description|subject)[\s\S]*?</dc:(?:title|description|subject)>", xmp):
                    # Strip tags, keep inner text (rdf:li values etc.).
                    text = re.sub(r"<[^>]+>", " ", m)
                    text = re.sub(r"\s+", " ", text).strip()
                    if text:
                        parts.append(text)
    except Exception:
        return ""
    return " ".join(parts).strip()


def _image_sidecar_caption(path: str) -> str:
    """Read a caption from a sidecar file sharing the image's base name
    (e.g. 'kisan_card.jpg' → 'kisan_card.txt'). Returns '' if none found.
    For .json sidecars, concatenates common caption/title/description/keywords fields."""
    base = os.path.splitext(path or "")[0]
    for ext in _CAPTION_SIDECAR_EXTS:
        cand = base + ext
        if not os.path.isfile(cand):
            continue
        try:
            with open(cand, "r", encoding="utf-8", errors="ignore") as f:
                raw = f.read().strip()
        except Exception:
            continue
        if not raw:
            continue
        if ext == ".json":
            try:
                data = json.loads(raw)
                fields = []
                if isinstance(data, dict):
                    for k in ("title", "caption", "description", "alt", "keywords", "tags", "subject"):
                        v = data.get(k)
                        if isinstance(v, str):
                            fields.append(v)
                        elif isinstance(v, (list, tuple)):
                            fields.extend(str(x) for x in v)
                return " ".join(fields).strip()
            except Exception:
                return raw
        return raw
    return ""


@lru_cache(maxsize=256)
def _image_description_keywords(path: str) -> tuple:
    """Combined, de-duplicated keyword set for an uploaded image, derived from ALL
    available signals: FILENAME + embedded TITLE/CAPTION/DESCRIPTION/KEYWORDS metadata
    + any sidecar caption file. Cached per path. Returned as a tuple (hashable)."""
    text_blobs = [
        " ".join(_filename_keywords(path)),
        _image_embedded_text(path),
        _image_sidecar_caption(path),
    ]
    combined = " ".join(b for b in text_blobs if b)
    kws = _image_keywords(combined)
    # Keep filename keywords even if short, so a purely-named image still matches.
    for fk in _filename_keywords(path):
        if fk not in kws:
            kws.append(fk)
    return tuple(kws)


def match_uploaded_images_to_slides(content_slides: list, image_paths: list) -> list:
    """Assign each uploaded image to the single most relevant content slide, based on
    overlap between the image's FILENAME keywords and the slide's content keywords.


    Returns a list parallel to `content_slides`: each entry is an image path (str) if a
    relevant image was matched to that slide, else None (slide stays text-only).

    Matching signals per image (combined): FILE NAME + embedded TITLE / CAPTION /
    DESCRIPTION / KEYWORDS metadata (EXIF, PNG text, XMP) + any sidecar caption file.
    Each image is scored against every slide's title + bullet keywords; the strongest
    (image, slide) pairs are assigned greedily so each image lands on its most relevant
    slide.

    - Each image is used at most once (assigned to its best-matching slide).
    - Each slide receives at most one image.
    - Only matches with score >= _UPLOAD_RELEVANCE_MIN are accepted; unmatched images
      are skipped and unmatched slides remain text-only.
    - After assignment, a verification pass logs the relevance of every placement and
      flags any image left unplaced, so placement can be reviewed before finalizing.
    """
    assigned = [None] * len(content_slides)
    if not image_paths or not content_slides:
        return assigned

    # Pre-compute slide keyword sets (title weighted higher than bullets).
    slide_title_kw = [set(_image_keywords(sd.get("title", ""))) for sd in content_slides]
    slide_body_kw = [set(_image_keywords(" ".join(sd.get("bullets", [])))) for sd in content_slides]

    # Pre-compute each image's combined keyword set (filename + title + caption + meta)
    # and remember which keywords came from the filename/title for logging.
    img_kw = []
    for path in image_paths:
        kws = set(_image_description_keywords(path))
        img_kw.append(kws)

    # Build all (score, img_idx, slide_idx) candidate pairs. A title match is worth
    # more than a body match so an image lands on the slide whose TOPIC it matches.
    candidates = []
    for img_i, ikw in enumerate(img_kw):
        if not ikw:
            continue
        for s_i in range(len(content_slides)):
            title_hits = len(ikw & slide_title_kw[s_i])
            body_hits = len(ikw & slide_body_kw[s_i])
            score = title_hits * 2 + body_hits
            if title_hits + body_hits >= _UPLOAD_RELEVANCE_MIN:
                candidates.append((score, title_hits + body_hits, img_i, s_i))

    # Strongest matches win first (by weighted score, then by raw keyword overlap).
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    used_imgs, filled_slides = set(), set()
    match_detail = {}  # img_i -> (slide_i, overlap_keywords)
    for score, raw, img_i, s_i in candidates:
        if img_i in used_imgs or s_i in filled_slides:
            continue
        assigned[s_i] = image_paths[img_i]
        used_imgs.add(img_i)
        filled_slides.add(s_i)
        overlap = sorted((img_kw[img_i] & (slide_title_kw[s_i] | slide_body_kw[s_i])))
        match_detail[img_i] = (s_i, overlap)
        logger.info(f"[upload-image] matched '{os.path.basename(image_paths[img_i])}' "
                    f"→ slide '{content_slides[s_i].get('title','')[:50]}' "
                    f"(score={score}, on: {', '.join(overlap) or 'n/a'})")

    # ── Verification pass: review relevance & placement across ALL slides ──────
    placed = len(used_imgs)
    unplaced = [os.path.basename(p) for i, p in enumerate(image_paths) if i not in used_imgs]
    logger.info(f"[upload-image][verify] {placed}/{len(image_paths)} image(s) placed on "
                f"their most-relevant slide; {len(filled_slides)}/{len(content_slides)} "
                f"content slide(s) received an image.")
    if unplaced:
        logger.warning(f"[upload-image][verify] {len(unplaced)} image(s) had no keyword "
                       f"overlap with any slide and were NOT placed (left for review): {unplaced}")
    # Sanity: confirm each placed image is genuinely relevant (non-empty overlap).
    for img_i, (s_i, overlap) in match_detail.items():
        if not overlap:
            logger.warning(f"[upload-image][verify] placement of "
                           f"'{os.path.basename(image_paths[img_i])}' has weak relevance.")
    return assigned


# Shape-name marker used to flag the single most important image on a slide.
# The animation builder looks for this prefix to give the image a cinematic
# Zoom-In / Scale-Up entrance plus a subtle Ken Burns motion.
MAIN_IMAGE_NAME_PREFIX = "KBMainImage"

# Role markers used so the animation builder can orchestrate a professional,
# click-driven build on each content slide:
#   • the HEADER and the FIRST bullet appear together when the slide opens,
#   • every remaining bullet is revealed one-by-one on each subsequent click.
HEADER_NAME_PREFIX  = "KBHeader"   # the slide heading shape
BULLETS_NAME_PREFIX = "KBBullets"  # the multi-paragraph bullets text box

# When a slide has TWO (or more) images we show them one-at-a-time as full-slide
# BACKGROUND visuals: the first image is the main background; each later image is
# tagged with this prefix so the animation builder reveals it (fade/zoom) on the next
# click, covering the previous image. Never two images shown at once.
IMAGE_SEQ_NAME_PREFIX = "KBBgSeq"  # a secondary full-slide background image


def _place_image_fit(slide, img_path: str, left: int, top: int, box_w: int, box_h: int,
                     is_main: bool = False):
    """Place an image inside the (left, top, box_w, box_h) region in EMU, preserving
    its original aspect ratio (no stretching/distortion) and centring it in the box.
    Scales down to fit; never scales a small image beyond its natural proportions
    relative to the box (always fits within the box).

    When `is_main` is True the created picture is tagged (via its shape name) so the
    animation pass can single it out for the strongest cinematic emphasis.
    Returns the created picture shape on success, else None."""
    if not img_path or not os.path.isfile(img_path):
        return None
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            iw, ih = im.size
        if iw <= 0 or ih <= 0:
            return None

        # Scale factor that fits the image within the box while keeping aspect ratio.
        scale = min(box_w / iw, box_h / ih)
        draw_w = int(iw * scale)
        draw_h = int(ih * scale)

        # Centre within the reserved box.
        pic_left = int(left + (box_w - draw_w) / 2)
        pic_top = int(top + (box_h - draw_h) / 2)

        pic = slide.shapes.add_picture(img_path, Emu(pic_left), Emu(pic_top),
                                       width=Emu(draw_w), height=Emu(draw_h))
        if is_main:
            # Tag the shape so _apply_slide_animation can target it specifically.
            pic.name = f"{MAIN_IMAGE_NAME_PREFIX}_{pic.shape_id}"
        return pic
    except Exception as e:
        logger.info(f"[auto-image] placement failed for {img_path}: {e}")
        return None


def _place_image_cover(slide, img_path: str, left: int, top: int, box_w: int, box_h: int,
                       is_main: bool = False, seq_index: int | None = None):
    """Place an image so it COVERS the whole (left, top, box_w, box_h) region — i.e. a
    full-bleed background. Aspect ratio is preserved (NO stretching); the overflowing
    edges are cropped symmetrically via the picture's native crop so the subject stays
    centred and image quality is untouched.

    Tagging:
      • is_main=True              → tagged as the primary background (zoom-in + Ken Burns)
      • seq_index is not None      → tagged as a sequential background (revealed on click)
    Returns the created picture shape, or None on failure."""
    if not img_path or not os.path.isfile(img_path):
        return None
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            iw, ih = im.size
        if iw <= 0 or ih <= 0:
            return None

        # Insert at full box size; cropping (below) trims the overflow so the image
        # fills the box exactly while keeping its aspect ratio.
        pic = slide.shapes.add_picture(img_path, Emu(left), Emu(top),
                                       width=Emu(box_w), height=Emu(box_h))

        # Compute symmetric crop fractions so the *visible* area matches the box's
        # aspect ratio (center-crop). box_ar vs img_ar decides which axis is cropped.
        box_ar = box_w / box_h
        img_ar = iw / ih
        if img_ar > box_ar:
            # Image is wider than the box → crop left/right.
            crop = (1.0 - box_ar / img_ar) / 2.0
            pic.crop_left = crop
            pic.crop_right = crop
        elif img_ar < box_ar:
            # Image is taller than the box → crop top/bottom.
            crop = (1.0 - img_ar / box_ar) / 2.0
            pic.crop_top = crop
            pic.crop_bottom = crop

        if is_main:
            pic.name = f"{MAIN_IMAGE_NAME_PREFIX}_{pic.shape_id}"
        elif seq_index is not None:
            pic.name = f"{IMAGE_SEQ_NAME_PREFIX}{seq_index}_{pic.shape_id}"
        return pic
    except Exception as e:
        logger.info(f"[auto-image] cover placement failed for {img_path}: {e}")
        return None


def _set_fill_alpha(shape, rgb: RGBColor, alpha_pct: int):
    """Give a shape a solid fill with transparency. `alpha_pct` is the OPACITY in
    percent (0 = fully transparent, 100 = fully opaque). python-pptx has no public
    alpha API, so we patch the <a:srgbClr> with an <a:alpha> child directly."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()
    # Clamp and convert opacity% → DrawingML alpha (0..100000).
    alpha_pct = max(0, min(100, alpha_pct))
    alpha_val = int(alpha_pct * 1000)
    srgb = shape.fill.fore_color._xFill.find(qn("a:srgbClr"))
    if srgb is not None:
        for existing in srgb.findall(qn("a:alpha")):
            srgb.remove(existing)
        alpha_el = etree.SubElement(srgb, qn("a:alpha"))
        alpha_el.set("val", str(alpha_val))


def _add_image_overlay_scrim(slide, left: int, top: int, w: int, h: int,
                             opacity_pct: int = 45):
    """Add a dark, semi-transparent rectangle over an image region so overlaid text
    stays readable. Returns the scrim shape (so the caller can order/tag it)."""
    scrim = slide.shapes.add_shape(1, Emu(left), Emu(top), Emu(w), Emu(h))
    _set_fill_alpha(scrim, RGBColor(0x12, 0x12, 0x12), opacity_pct)
    return scrim


def _render_overlay_content_slide(slide, title: str, bullets: list, img_path: str,
                                  slide_w: int, slide_h: int):
    """Render a content slide where the image is LARGE and PROMINENT and the text is
    overlaid directly on top of it, over a dark transparent scrim for readability.

    Layout order (back → front): big image → dark scrim → header → bullets.
    The header and bullets are REAL, EDITABLE text frames. Shapes are tagged so the
    animation builder can: zoom the image in, show header + first bullet together,
    then reveal each remaining bullet on click.
    Returns dict with the created shapes (image, scrim, header, bullets)."""
    PAD = 400000

    # ── 1) Big, prominent image filling most of the slide ────────────────────
    # The image is the hero: it occupies the full slide area (minus a small margin),
    # tagged as the main image so it gets the zoom-in + Ken Burns treatment.
    img_left, img_top = PAD, PAD
    img_w = slide_w - 2 * PAD
    img_h = slide_h - 2 * PAD
    pic = None
    if img_path:
        pic = _place_image_fit(slide, img_path, img_left, img_top, img_w, img_h,
                               is_main=True)

    # Use the actual placed-image box for the scrim/text so the overlay sits on the
    # image (not on empty slide). Fall back to the reserved region if placement failed.
    if pic is not None:
        box_left, box_top = int(pic.left), int(pic.top)
        box_w, box_h = int(pic.width), int(pic.height)
    else:
        box_left, box_top, box_w, box_h = img_left, img_top, img_w, img_h

    # ── 2) Dark transparent scrim over the image for text readability ─────────
    scrim = _add_image_overlay_scrim(slide, box_left, box_top, box_w, box_h,
                                     opacity_pct=45)

    # ── 3) Header (overlaid, near the top of the image) ───────────────────────
    inner_pad = 350000
    header_h = 900000
    htb = slide.shapes.add_textbox(Emu(box_left + inner_pad), Emu(box_top + inner_pad),
                                   Emu(box_w - 2 * inner_pad), Emu(header_h))
    htf = htb.text_frame
    htf.word_wrap = True
    hp = htf.paragraphs[0]
    hp.alignment = PP_ALIGN.LEFT
    hr = hp.add_run()
    hr.text = title
    hr.font.size = Pt(30)
    hr.font.bold = True
    hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    htb.name = f"{HEADER_NAME_PREFIX}_{htb.shape_id}"

    # ── 4) Bullets (overlaid, below the header) ───────────────────────────────
    bullets_top = box_top + inner_pad + header_h
    bullets_h = box_h - (bullets_top - box_top) - inner_pad
    btb = slide.shapes.add_textbox(Emu(box_left + inner_pad), Emu(bullets_top),
                                   Emu(box_w - 2 * inner_pad), Emu(bullets_h))
    btf = btb.text_frame
    btf.word_wrap = True
    btf.auto_size = None

    count = max(1, len(bullets))
    font_size = Pt(20) if count <= 5 else Pt(17) if count <= 7 else Pt(14)
    for i, bullet in enumerate(bullets):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        indent = 0
        text = bullet
        if bullet.startswith("   "):
            indent = 1
            text = bullet.strip()
        p.level = indent
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = ("\u2022 " if indent == 0 else "\u2013 ") + text
        r.font.size = font_size if indent == 0 else Pt(font_size.pt - 2)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    btb.name = f"{BULLETS_NAME_PREFIX}_{btb.shape_id}"

    return {"image": pic, "scrim": scrim, "header": htb, "bullets": btb}


def _render_bg_sequence_slide(slide, title: str, bullets: list, images: list,
                              slide_w: int, slide_h: int):
    """Render a slide whose MULTIPLE images are shown ONE AT A TIME as full-slide
    background visuals. The first image is the primary background (zoom-in on open);
    every later image is stacked full-bleed on top but hidden until a click, when it
    fades/zooms in to replace the previous image. A dark transparent scrim and the
    header + bullets sit ABOVE all images so text stays readable over whichever image
    is currently showing. Never shows two images at once.

    Z-order (back → front): image1 → image2 … → scrim → header → bullets.
    Returns a dict of the created shapes.
    """
    imgs = [p for p in (images or []) if p and os.path.isfile(p)]
    if not imgs:
        # Nothing to show as background → fall back to a plain overlay (no image).
        return _render_overlay_content_slide(slide, title, bullets, None, slide_w, slide_h)

    # ── 1) Full-slide background images, stacked. First = main, rest = sequence ──
    bg_pics = []
    for idx, p in enumerate(imgs):
        if idx == 0:
            pic = _place_image_cover(slide, p, 0, 0, slide_w, slide_h, is_main=True)
        else:
            pic = _place_image_cover(slide, p, 0, 0, slide_w, slide_h, seq_index=idx)
        if pic is not None:
            bg_pics.append(pic)

    # ── 2) Dark transparent scrim over the full slide for text readability ────
    scrim = _add_image_overlay_scrim(slide, 0, 0, slide_w, slide_h, opacity_pct=45)

    # ── 3) Header (overlaid, top) ─────────────────────────────────────────────
    inner_pad = 500000
    header_h = 950000
    htb = slide.shapes.add_textbox(Emu(inner_pad), Emu(inner_pad),
                                   Emu(slide_w - 2 * inner_pad), Emu(header_h))
    htf = htb.text_frame
    htf.word_wrap = True
    hp = htf.paragraphs[0]
    hp.alignment = PP_ALIGN.LEFT
    hr = hp.add_run()
    hr.text = title
    hr.font.size = Pt(32)
    hr.font.bold = True
    hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    htb.name = f"{HEADER_NAME_PREFIX}_{htb.shape_id}"

    # ── 4) Bullets (overlaid, below header) ───────────────────────────────────
    bullets_top = inner_pad + header_h
    bullets_h = slide_h - bullets_top - inner_pad
    btb = slide.shapes.add_textbox(Emu(inner_pad), Emu(bullets_top),
                                   Emu(slide_w - 2 * inner_pad), Emu(bullets_h))
    btf = btb.text_frame
    btf.word_wrap = True
    btf.auto_size = None
    count = max(1, len(bullets))
    font_size = Pt(20) if count <= 5 else Pt(17) if count <= 7 else Pt(14)
    for i, bullet in enumerate(bullets):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        indent = 0
        text = bullet
        if bullet.startswith("   "):
            indent = 1
            text = bullet.strip()
        p.level = indent
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = ("\u2022 " if indent == 0 else "\u2013 ") + text
        r.font.size = font_size if indent == 0 else Pt(font_size.pt - 2)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    btb.name = f"{BULLETS_NAME_PREFIX}_{btb.shape_id}"

    return {"backgrounds": bg_pics, "scrim": scrim, "header": htb, "bullets": btb}


# ─── PPT Builder ─────────────────────────────────────────────────────────────
def add_header_bar(slide, title_text: str, slide_w: int = 12192000, tag: bool = False):
    """Add the consistent YELLOW header band with the slide title in dark, bold,
    readable text. Used on every slide for a uniform look. When `tag` is True the
    header is marked so the animation builder reveals it on slide open."""
    rect = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(slide_w), Emu(530328))
    rect.fill.solid()
    rect.fill.fore_color.rgb = HEADER_YELLOW
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
    run.font.color.rgb = HEADER_TEXT
    if tag:
        rect.name = f"{HEADER_NAME_PREFIX}_{rect.shape_id}"
    return rect


def _render_image_top_textbox_slide(slide, title: str, bullets: list, img_path,
                                    slide_w: int, slide_h: int):
    """Standard content layout for ANY slide that has both image(s) and text:

        ┌─────────────────────────────────────────┐
        │           YELLOW HEADER (title)          │
        ├─────────────────────────────────────────┤
        │  IMAGE fills the whole body as a         │
        │  full-bleed BACKGROUND.                   │
        │  ╭─────────────────────────────────────╮ │
        │  │  BLACK box · padding                │ │   ← text sits in a black box
        │  │  WHITE readable bullet text         │ │     overlaid on the background
        │  ╰─────────────────────────────────────╯ │
        └─────────────────────────────────────────┘

    `img_path` may be a single path (str) or a list of paths. The first image is used
    as a full-slide COVER background (tagged as the main image for the zoom-in + Ken
    Burns treatment); any additional images are stacked full-bleed on top and revealed
    one-at-a-time on click (never shown simultaneously). The bullet text lives in a
    BLACK rectangular box with WHITE text and inner padding, overlaid on the image.
    Shapes are tagged so the animation builder reveals the header + first bullet
    together, then each remaining bullet on click.
    """
    # Normalise to a list of existing image paths.
    if isinstance(img_path, (list, tuple)):
        imgs = [p for p in img_path if p and os.path.isfile(p)]
    else:
        imgs = [img_path] if (img_path and os.path.isfile(img_path)) else []

    HEADER_H = 530328
    MARGIN   = 500000          # outer slide margin for the text box
    BOX_PAD  = 260000          # inner padding inside the black text box

    # 1) Image(s) as a FULL-SLIDE background (cover). First = main (zoom-in); the
    #    rest are sequential backgrounds revealed on click.
    pics = []
    for idx, p in enumerate(imgs):
        if idx == 0:
            pic = _place_image_cover(slide, p, 0, 0, slide_w, slide_h, is_main=True)
        else:
            pic = _place_image_cover(slide, p, 0, 0, slide_w, slide_h, seq_index=idx)
        if pic is not None:
            pics.append(pic)

    # 2) Consistent yellow header with the title (tagged for the entrance). Added
    #    AFTER the background so it sits above the image.
    add_header_bar(slide, title, slide_w=slide_w, tag=True)

    # 3) Black text box overlaid on the background, below the header. It occupies the
    #    lower portion of the slide so the image stays visible above it.
    box_left   = MARGIN
    box_w      = slide_w - 2 * MARGIN
    box_top    = HEADER_H + int((slide_h - HEADER_H) * 0.42)
    box_h      = slide_h - MARGIN - box_top
    box = slide.shapes.add_shape(1, Emu(box_left), Emu(box_top),
                                 Emu(box_w), Emu(box_h))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)   # black box
    box.line.fill.background()
    # The box itself is static (always visible) so the text has a stable backdrop;
    # the TEXT (added as a separate shape) is what animates in.

    # 4) Bullet text inside the black box with proper padding (separate, tagged
    #    shape). Text is WHITE for contrast against the black box.
    tb = slide.shapes.add_textbox(Emu(box_left + BOX_PAD), Emu(box_top + BOX_PAD),
                                  Emu(box_w - 2 * BOX_PAD), Emu(box_h - 2 * BOX_PAD))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf.vertical_anchor = MSO_ANCHOR.TOP
    except Exception:
        pass

    count = max(1, len(bullets))
    font_size = Pt(18) if count <= 4 else Pt(15) if count <= 6 else Pt(13)
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        indent = 0
        text = bullet
        if bullet.startswith("   "):
            indent = 1
            text = bullet.strip()
        p.level = indent
        p.space_after = Pt(6)
        r = p.add_run()
        r.text = ("\u2022 " if indent == 0 else "\u2013 ") + text
        r.font.size = font_size if indent == 0 else Pt(font_size.pt - 2)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)   # white text
    tb.name = f"{BULLETS_NAME_PREFIX}_{tb.shape_id}"

    return {"images": pics, "image": pics[0] if pics else None, "box": box, "bullets": tb}


# ─── Slide animations (injected directly into PPTX XML) ──────────────────────
# python-pptx has no animation API, so we append a <p:timing> block to each
# slide's XML. This produces real PowerPoint entrance animations (Fade) that
# play in the slideshow and in exported video — no PowerPoint/COM needed.
from pptx.oxml.ns import qn
from lxml import etree

ADD_PPTX_ANIMATIONS = True   # set False to generate decks without animations
ADD_PPTX_TRANSITIONS = True  # set False to generate decks without slide transitions

# Timing constants (milliseconds) tuned for smooth, professional motion.
_IMG_ENTRANCE_DUR = 700      # main image Zoom-In / Scale-Up entrance
_TEXT_FADE_DUR = 500         # text / other shapes fade-in
_KENBURNS_DUR = 9000         # slow continuous Ken Burns zoom+pan per image
_KENBURNS_ZOOM_TO = 112000   # end scale (112%) → subtle zoom (base 100000 = 100%)


def _is_main_image_shape(shp) -> bool:
    """True when a shape was tagged as the slide's primary image."""
    try:
        return bool(shp.name) and shp.name.startswith(MAIN_IMAGE_NAME_PREFIX)
    except Exception:
        return False


def _entrance_fade(sid: int, start_delay: str, dur: int) -> str:
    """A simple Fade-In entrance (presetID 10) for a single shape."""
    return f"""
          <p:par>
            <p:cTn id="PLACEHOLDER" fill="hold">
              <p:stCondLst><p:cond delay="{start_delay}"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="PLACEHOLDER" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="PLACEHOLDER" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="afterEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                            <p:animEffect transition="in" filter="fade">
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="{dur}"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              </p:cBhvr>
                            </p:animEffect>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>"""


def _entrance_fade_para(sid: int, para_idx: int, start_delay: str, dur: int) -> str:
    """A Fade-In entrance (presetID 10) targeting a SINGLE paragraph (`para_idx`) of a
    text shape, so bullets can be revealed one-by-one. `start_delay` is "0" (appears
    with the slide) or "indefinite" (waits for a click)."""
    return f"""
          <p:par>
            <p:cTn id="PLACEHOLDER" fill="hold">
              <p:stCondLst><p:cond delay="{start_delay}"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="PLACEHOLDER" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="PLACEHOLDER" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="{sid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                            <p:animEffect transition="in" filter="fade">
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="{dur}"/>
                                <p:tgtEl><p:spTgt spid="{sid}"><p:txEl><p:pRg st="{para_idx}" end="{para_idx}"/></p:txEl></p:spTgt></p:tgtEl>
                              </p:cBhvr>
                            </p:animEffect>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>"""


def _entrance_zoom_image(sid: int, start_delay: str, dur: int) -> str:
    """A cinematic Zoom-In / Scale-Up entrance for the main image: the picture
    fades in while scaling up from 70% to 100% (presetID 23 = 'Zoom' entrance),
    immediately followed by a slow continuous Ken Burns zoom (scale to ~112%) that
    keeps the image feeling dynamic without distracting motion."""
    return f"""
          <p:par>
            <p:cTn id="PLACEHOLDER" fill="hold">
              <p:stCondLst><p:cond delay="{start_delay}"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="PLACEHOLDER" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="PLACEHOLDER" presetID="23" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:set>
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="1" fill="hold">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                </p:cTn>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:to><p:strVal val="visible"/></p:to>
                            </p:set>
                            <p:anim calcmode="lin" valueType="num">
                              <p:cBhvr additive="base">
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:tavLst>
                                <p:tav tm="0"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                                <p:tav tm="100000"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                              </p:tavLst>
                            </p:anim>
                            <p:anim calcmode="lin" valueType="num">
                              <p:cBhvr additive="base">
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:tavLst>
                                <p:tav tm="0"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                                <p:tav tm="100000"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                              </p:tavLst>
                            </p:anim>
                            <p:animScale>
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              </p:cBhvr>
                              <p:by x="100000" y="100000"/>
                              <p:from x="70000" y="70000"/>
                              <p:to x="100000" y="100000"/>
                            </p:animScale>
                            <p:animEffect transition="in" filter="fade">
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="{dur}"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              </p:cBhvr>
                            </p:animEffect>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>"""


def _kenburns_after(sid: int) -> str:
    """A subtle, slow Ken Burns emphasis: continuous zoom from 100% to ~112% with a
    gentle pan, started right after the entrance so the image feels cinematic and
    alive. Uses an emphasis Grow/Scale (presetClass='emph') that auto-advances."""
    zoom = _KENBURNS_ZOOM_TO
    dur = _KENBURNS_DUR
    return f"""
          <p:par>
            <p:cTn id="PLACEHOLDER" fill="hold">
              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="PLACEHOLDER" fill="hold">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="PLACEHOLDER" presetID="6" presetClass="emph" presetSubtype="0" fill="hold" nodeType="afterEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:animScale>
                              <p:cBhvr>
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                              </p:cBhvr>
                              <p:from x="100000" y="100000"/>
                              <p:to x="{zoom}" y="{zoom}"/>
                            </p:animScale>
                            <p:anim calcmode="lin" valueType="num">
                              <p:cBhvr additive="base">
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:tavLst>
                                <p:tav tm="0"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                                <p:tav tm="100000"><p:val><p:strVal val="0.52"/></p:val></p:tav>
                              </p:tavLst>
                            </p:anim>
                            <p:anim calcmode="lin" valueType="num">
                              <p:cBhvr additive="base">
                                <p:cTn id="PLACEHOLDER" dur="{dur}" fill="hold"/>
                                <p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>
                              </p:cBhvr>
                              <p:tavLst>
                                <p:tav tm="0"><p:val><p:strVal val="0.5"/></p:val></p:tav>
                                <p:tav tm="100000"><p:val><p:strVal val="0.48"/></p:val></p:tav>
                              </p:tavLst>
                            </p:anim>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>"""


def _build_timing_xml(steps: list[dict]) -> str:
    """Return a <p:timing> block from an ordered list of animation `steps`.

    Each step is a dict:
      {"kind": "image"|"image2"|"text"|"para", "sid": int, "delay": "0"|"indefinite",
       "para": <int, only for kind == "para">}
    - delay "0"       → the effect plays automatically when the slide opens
                        (used for the image, the header, and the FIRST bullet so they
                        appear together).
    - delay "indefinite" → the effect waits for a mouse click (used for every
                        remaining bullet, revealing them one-by-one).
    The main image additionally gets a slow Ken Burns emphasis after its entrance.
    """
    if not steps:
        return ""

    parts = []
    kenburns_parts = []
    sids = []
    for step in steps:
        sid = step["sid"]
        sids.append(sid)
        delay = step.get("delay", "indefinite")
        kind = step.get("kind")
        if kind == "image":
            parts.append(_entrance_zoom_image(sid, delay, _IMG_ENTRANCE_DUR))
            kenburns_parts.append(_kenburns_after(sid))
        elif kind == "image2":
            parts.append(_entrance_zoom_image(sid, delay, _IMG_ENTRANCE_DUR))
        elif kind == "para":
            parts.append(_entrance_fade_para(sid, step["para"], delay, _TEXT_FADE_DUR))
        else:  # "text"
            parts.append(_entrance_fade(sid, delay, _TEXT_FADE_DUR))

    seq_body = "".join(parts)
    kb_body = "".join(kenburns_parts)
    # Build list: one <p:bldP> per shape. For the bullets shape we use build="p"
    # so each paragraph is an independently-animatable build step.
    seen_sids = []
    for s in sids:
        if s not in seen_sids:
            seen_sids.append(s)
    build_entries = []
    para_sids = {step["sid"] for step in steps if step.get("kind") == "para"}
    for s in seen_sids:
        if s in para_sids:
            build_entries.append(f'<p:bldP spid="{s}" grpId="0" build="p" animBg="1"/>')
        else:
            build_entries.append(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>')
    build_targets = "".join(build_entries)

    xml = f"""<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>{seq_body}{kb_body}
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  <p:bldLst>{build_targets}</p:bldLst>
</p:timing>"""
    return xml


def _renumber_ctn_ids(timing_el, start: int = 3) -> None:
    """Assign unique sequential ids to every <p:cTn> that uses the PLACEHOLDER id.

    PowerPoint requires every timing node (cTn) to have a unique id. We leave
    ids 1 and 2 for tmRoot/mainSeq and renumber everything else from `start`.
    """
    counter = start
    for ctn in timing_el.iter(qn("p:cTn")):
        if ctn.get("id") == "PLACEHOLDER":
            ctn.set("id", str(counter))
            counter += 1


def _count_paragraphs(shp) -> int:
    """Number of paragraphs in a text shape (0 if it has no text frame)."""
    try:
        if shp.has_text_frame:
            return len(shp.text_frame.paragraphs)
    except Exception:
        pass
    return 0


def _apply_slide_animation(slide) -> None:
    """Inject the orchestrated, click-driven build timeline into a slide's XML.

    Order of reveal:
      1. The main image zooms in and begins a slow Ken Burns (on slide open).
      2. The HEADER and the FIRST bullet appear together (on slide open).
      3. Each remaining bullet fades in one-by-one on every subsequent click.
    Other pictures/text (if any) fade in with-slide as supporting elements.
    """
    header_shp = None
    bullets_shp = None
    main_img_shp = None
    seq_imgs = []      # (seq_index, shape) for full-slide background sequence images
    other_imgs = []
    other_text = []

    for shp in slide.shapes:
        try:
            sid = shp.shape_id
        except Exception:
            continue
        name = getattr(shp, "name", "") or ""
        if name.startswith(MAIN_IMAGE_NAME_PREFIX):
            main_img_shp = shp
        elif name.startswith(IMAGE_SEQ_NAME_PREFIX):
            # Name is like 'KBBgSeq2_<sid>' → recover the sequence index for ordering.
            try:
                after = name[len(IMAGE_SEQ_NAME_PREFIX):].split("_", 1)[0]
                order = int(after)
            except Exception:
                order = 999
            seq_imgs.append((order, shp))
        elif name.startswith(HEADER_NAME_PREFIX):
            header_shp = shp
        elif name.startswith(BULLETS_NAME_PREFIX):
            bullets_shp = shp
        else:
            is_pic = False
            try:
                is_pic = shp.shape_type == 13  # MSO_SHAPE_TYPE.PICTURE
            except Exception:
                is_pic = False
            if is_pic:
                other_imgs.append(shp)
            # Non-tagged shapes (background rectangles, scrims, plain header bars,
            # side textboxes) are intentionally NOT animated so they're present from
            # the start — this keeps layouts (and the dark scrim) stable.

    steps: list[dict] = []

    # 1) Main image: zoom-in on open (+ Ken Burns handled in builder).
    if main_img_shp is not None:
        steps.append({"kind": "image", "sid": main_img_shp.shape_id, "delay": "0"})
    for pic in other_imgs:
        steps.append({"kind": "image2", "sid": pic.shape_id, "delay": "0"})

    # 2) Header on open.
    if header_shp is not None:
        steps.append({"kind": "text", "sid": header_shp.shape_id, "delay": "0"})

    # 3) Bullets: paragraph 0 on open (with the header), the rest on click.
    if bullets_shp is not None:
        n_paras = _count_paragraphs(bullets_shp)
        for p_idx in range(n_paras):
            steps.append({
                "kind": "para",
                "sid": bullets_shp.shape_id,
                "para": p_idx,
                "delay": "0" if p_idx == 0 else "indefinite",
            })

    # 4) Sequential full-slide background images: revealed ONE AT A TIME on click,
    #    AFTER the bullets, each with a smooth zoom+fade so it covers the previous
    #    image. (Never shown simultaneously — each waits for its own click.)
    for _order, pic in sorted(seq_imgs, key=lambda t: t[0]):
        steps.append({"kind": "image2", "sid": pic.shape_id, "delay": "indefinite"})

    # Fallback: if this slide has no tagged header/bullets (e.g. title / section /
    # thank-you slides, or the old side-by-side layout). Reveal every element
    # AUTOMATICALLY on slide open (no clicks) so these intro/outro slides present
    # fully at once. A full-slide background rectangle is skipped (left static) so we
    # don't fade the whole slide.
    if header_shp is None and bullets_shp is None:
        try:
            sw = slide.part.package.presentation_part.presentation.slide_width
            sh = slide.part.package.presentation_part.presentation.slide_height
        except Exception:
            sw, sh = 12192000, 6858000
        for shp in slide.shapes:
            try:
                sid = shp.shape_id
            except Exception:
                continue
            # Skip a shape that spans (nearly) the whole slide — the background.
            try:
                if int(shp.width) >= int(sw) - 1000 and int(shp.height) >= int(sh) - 1000:
                    continue
            except Exception:
                pass
            is_pic = False
            try:
                is_pic = shp.shape_type == 13
            except Exception:
                is_pic = False
            if _is_main_image_shape(shp):
                steps.append({"kind": "image", "sid": sid, "delay": "0"})
            elif is_pic:
                steps.append({"kind": "image2", "sid": sid, "delay": "0"})
            else:
                steps.append({"kind": "text", "sid": sid, "delay": "0"})

    if not steps:
        return

    xml = _build_timing_xml(steps)
    if not xml:
        return

    timing_el = etree.fromstring(xml.encode("utf-8"))
    _renumber_ctn_ids(timing_el)

    sld = slide._element  # <p:sld>
    # Remove any pre-existing timing, then append ours as the last child.
    for existing in sld.findall(qn("p:timing")):
        sld.remove(existing)
    sld.append(timing_el)


def _apply_slide_transition(slide, duration_ms: int = 900) -> None:
    """Inject a smooth zoom slide transition into a slide's XML.

    Uses PowerPoint's zoom-family transition so moving to the next slide feels like a
    gentle push/zoom toward the upcoming slide's main image. We emit it inside an
    <mc:AlternateContent> block: the modern (p14) 'warp'/zoom effect for newer
    PowerPoint, with a classic fade fallback for older viewers. Timing is kept
    consistent across slides for a professional, non-jarring flow.
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

    sld = slide._element  # <p:sld>

    # Remove any existing transition so we don't stack duplicates.
    for existing in sld.findall(qn("p:transition")):
        sld.remove(existing)

    # advTm / spd kept consistent; a zoom-in reveals the next slide's focal image.
    xml = f"""<mc:AlternateContent xmlns:mc="{MC_NS}">
  <mc:Choice xmlns:p14="{P14_NS}" Requires="p14">
    <p:transition xmlns:p="{P_NS}" spd="slow" p14:dur="{duration_ms}">
      <p14:warp dir="in"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition xmlns:p="{P_NS}" spd="slow">
      <p:fade/>
    </p:transition>
  </mc:Fallback>
</mc:AlternateContent>"""

    trans_el = etree.fromstring(xml.encode("utf-8"))

    # <p:transition> must appear before <p:timing> in the schema order. Insert it
    # right before any existing timing element (append otherwise).
    timing = sld.find(qn("p:timing"))
    if timing is not None:
        timing.addprevious(trans_el)
    else:
        sld.append(trans_el)


def _verify_image_text_layout(prs) -> list:
    """Verify that EVERY slide containing both image(s) and text follows the standard
    layout: the image is a FULL-SLIDE background and the text sits in a SEPARATE BLACK
    box (with white text) overlaid on the lower portion of the slide. Returns a list of
    violation strings (empty = all good). Logs a summary. Runs automatically before a
    deck is finalized.
    """
    BLACK = RGBColor(0x00, 0x00, 0x00)
    violations = []
    checked = 0
    for i, slide in enumerate(prs.slides, 1):
        pics = []
        bullets_shp = None
        black_boxes = []
        for shp in slide.shapes:
            name = getattr(shp, "name", "") or ""
            # Picture?
            is_pic = False
            try:
                is_pic = shp.shape_type == 13
            except Exception:
                is_pic = False
            if is_pic or name.startswith(MAIN_IMAGE_NAME_PREFIX) or name.startswith(IMAGE_SEQ_NAME_PREFIX):
                pics.append(shp)
                continue
            if name.startswith(BULLETS_NAME_PREFIX):
                bullets_shp = shp
            # Black rectangle (the text box backdrop)?
            try:
                if shp.shape_type == 1:  # AUTO_SHAPE (rectangle)
                    if (shp.fill.type is not None
                            and shp.fill.fore_color.rgb == BLACK):
                        black_boxes.append(shp)
            except Exception:
                pass

        has_image = len(pics) > 0
        has_text = bullets_shp is not None
        if not (has_image and has_text):
            continue  # layout rule only applies to image+text slides

        checked += 1

        # 1) A black text box must exist.
        if not black_boxes:
            violations.append(f"slide {i}: image+text but NO black text box")
            continue

        # 2) The bullet text must sit inside the black box region (overlap it).
        box = min(black_boxes, key=lambda b: int(b.top))  # the topmost black box
        box_top, box_bottom = int(box.top), int(box.top) + int(box.height)
        if not (box_top <= int(bullets_shp.top) <= box_bottom):
            violations.append(
                f"slide {i}: bullet text not inside the black box "
                f"(text.top={int(bullets_shp.top)}, box=[{box_top},{box_bottom}])"
            )

    if violations:
        logger.error(f"[layout-verify] {len(violations)} violation(s) across "
                     f"{checked} image+text slide(s):")
        for v in violations:
            logger.error(f"[layout-verify]   - {v}")
    else:
        logger.info(f"[layout-verify] OK — all {checked} image+text slide(s) use "
                    f"image-background / black-text-box layout")
    return violations


def _apply_motion_to_presentation(prs) -> None:
    """Apply cinematic entrance animations and smooth zoom slide transitions to every
    slide in `prs`. Safe to call on any Presentation built by any of the builders.
    Transition is applied first (schema order: transition before timing), then the
    animation timeline is injected. Any per-slide failure is logged and skipped so a
    single bad slide never aborts the whole deck.

    Before returning, verifies that every image+text slide follows the standard
    image-top / white-text-box-below layout and logs any violation.
    """
    for slide in prs.slides:
        if ADD_PPTX_TRANSITIONS:
            try:
                _apply_slide_transition(slide)
            except Exception:
                logger.exception("Failed to apply transition to a slide; continuing")
        if ADD_PPTX_ANIMATIONS:
            try:
                _apply_slide_animation(slide)
            except Exception:
                logger.exception("Failed to apply animation to a slide; continuing")

    # Finalize-time check: enforce the consistent image+text layout across the deck.
    try:
        _verify_image_text_layout(prs)
    except Exception:
        logger.exception("[layout-verify] verification pass failed; continuing")


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
            # Clean white background for a consistent, professional look.
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = BOX_WHITE
            bg.line.fill.background()

            # Consistent YELLOW header band carrying the title.
            add_header_bar(slide, sd.get("title", ""), slide_w=12192000)

            # Title (large) + subtitle centred in the white body.
            tb = slide.shapes.add_textbox(Emu(1500000), Emu(2400000), Emu(9200000), Emu(1800000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = sd.get("title", "")
            r.font.size = Pt(34)
            r.font.bold = True
            r.font.color.rgb = DARK_GRAY

            tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(4300000), Emu(9200000), Emu(1600000))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = sd.get("subtitle", "")
            r2.font.size = Pt(16)
            r2.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

        # ── Thank You slide ───────────────────────────────────────────────────
        elif stype == "thank_you":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = BOX_WHITE
            bg.line.fill.background()

            # Consistent YELLOW header band.
            add_header_bar(slide, sd.get("title", "Thank You"), slide_w=12192000)

            tb = slide.shapes.add_textbox(Emu(1500000), Emu(2400000), Emu(9200000), Emu(1500000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "Thank You"
            r.font.size = Pt(44)
            r.font.bold = True
            r.font.color.rgb = DARK_GRAY

            tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(4100000), Emu(9200000), Emu(1600000))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = sd.get("subtitle", "")
            r2.font.size = Pt(16)
            r2.font.color.rgb = DARK_GRAY

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

        # ── Section title slide ────────────────────────────────────────────────
        elif stype == "section_title":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(12192000), Emu(6858000))
            bg.fill.solid()
            bg.fill.fore_color.rgb = BOX_WHITE
            bg.line.fill.background()

            # Consistent YELLOW header band.
            add_header_bar(slide, sd.get("title", ""), slide_w=12192000)

            tb = slide.shapes.add_textbox(Emu(1500000), Emu(2600000), Emu(9200000), Emu(1800000))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = sd.get("title", "")
            r.font.size = Pt(32)
            r.font.bold = True
            r.font.color.rgb = DARK_GRAY

            sub = sd.get("subtitle", "")
            if sub:
                tb2 = slide.shapes.add_textbox(Emu(1500000), Emu(4400000), Emu(9200000), Emu(1200000))
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

            bullets = sd.get("bullets", [])

            # ── Try to find a content-relevant image for this slide ───────────
            # Only a keyword-matched (non-random) image is returned; otherwise None,
            # in which case the slide remains text-only (original behaviour).
            img_path = get_slide_image(sd.get("title", ""), bullets)

            SLIDE_W, SLIDE_H, HEADER_H = 12192000, 6858000, 530328

            if img_path:
                # ── Image on TOP, text in a separate white box BELOW ──────────
                # Consistent yellow header; image in the top region; dark bullet text
                # inside a white, bordered, padded box directly beneath the image.
                # The box never overlaps the image — clean, aligned, professional.
                _render_image_top_textbox_slide(
                    slide, sd.get("title", ""), bullets, img_path, SLIDE_W, SLIDE_H
                )
            else:
                # ── Text-only layout (yellow header + full-width editable bullets) ─
                add_header_bar(slide, sd.get("title", ""), slide_w=SLIDE_W, tag=True)
                PAD = 300000
                text_top = 700000
                text_h = SLIDE_H - text_top - PAD
                text_left = 600000
                text_w = 10992000

                bullet_count = len(bullets)
                font_size = Pt(18) if bullet_count <= 5 else Pt(15) if bullet_count <= 7 else Pt(13)

                tb = slide.shapes.add_textbox(Emu(text_left), Emu(text_top), Emu(text_w), Emu(text_h))
                tf = tb.text_frame
                tf.word_wrap = True
                tf.auto_size = None  # let word_wrap handle it

                for i, bullet in enumerate(bullets):
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
                # Tag so the text-only slide still gets the one-by-one bullet build.
                tb.name = f"{BULLETS_NAME_PREFIX}_{tb.shape_id}"

            slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

    _apply_motion_to_presentation(prs)

    logger.info(f"PPT generation took {time.time() - start:.2f}s ({len(slides_data)} slides)")
    return prs


# ─── Video Generator helpers ─────────────────────────────────────────────────
DEFAULT_TTS_GENERATE_URL = "http://172.23.198.242:8001/"


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
    return templates.TemplateResponse(name="ppt_studio.html", request=request)


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
    language    = data.get("language", DEFAULT_LANGUAGE)
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
    result = generate_slides(scheme_name, raw_content, DEFAULT_SYSTEM_PROMPT, ollama_url, model, language=language)
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
    language = data.get("language", DEFAULT_LANGUAGE)

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
    result = generate_slides(title, raw_content, system_prompt, ollama_url, model, language=language)

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
    language = data.get("language", DEFAULT_LANGUAGE)

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
    result = generate_slides(title, raw_content, system_prompt, ollama_url, model, language=language)
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


# ─── PPT Builder with Images ─────────────────────────────────────────────────
IMAGES_PER_SLIDE = 5   # how many images to show in the grid on each content slide
BULLETS_PER_SLIDE = 5  # target bullets per content slide


def _split_into_image_slides(slides_data: list) -> list:
    """Re-chunk content slides so each has exactly BULLETS_PER_SLIDE bullets.
    Title / thank_you / section_title slides pass through unchanged."""
    out = []
    for sd in slides_data:
        if sd.get("type") != "content":
            out.append(sd)
            continue
        bullets = sd.get("bullets", [])
        if not bullets:
            continue
        # Split into chunks of BULLETS_PER_SLIDE
        chunks = [bullets[i:i + BULLETS_PER_SLIDE] for i in range(0, len(bullets), BULLETS_PER_SLIDE)]
        n = len(chunks)
        for pi, chunk in enumerate(chunks, 1):
            title = f"{sd['title']} ({pi}/{n})" if n > 1 else sd["title"]
            out.append({
                "type": "content",
                "title": title,
                "bullets": chunk,
                "voiceover": sd.get("voiceover", ""),
            })
    return out


def _embed_single_image(slide, img_path: str, slide_w: int, slide_h: int, header_h: int):
    """Embed one image on the right 44% of the slide, vertically centred."""
    if not img_path or not os.path.isfile(img_path):
        return

    padding  = 150000
    area_left = int(slide_w * 0.56)
    area_top  = header_h + padding
    area_w    = slide_w - area_left - padding
    area_h    = slide_h - area_top - padding

    try:
        pic = slide.shapes.add_picture(img_path, Emu(area_left), Emu(area_top), height=Emu(area_h))
        if pic.width > area_w:
            pic._element.getparent().remove(pic._element)
            pic = slide.shapes.add_picture(img_path, Emu(area_left), Emu(area_top), width=Emu(area_w))
        pic.left = Emu(int(area_left + (area_w - pic.width) / 2))
        pic.top  = Emu(int(area_top  + (area_h - pic.height) / 2))
    except Exception as e:
        logger.error(f"Failed to embed image {img_path}: {e}")


def build_ppt_with_images(slides_data: list, image_paths: list) -> Presentation:
    """Build PPT with exactly len(image_paths) content slides (one image per slide).
    Title / thank_you / section_title slides are excluded from image assignment."""
    import copy
    SLIDE_W  = 12192000
    SLIDE_H  = 6858000
    HEADER_H = 530328

    slides_data = _split_into_image_slides(slides_data)

    # ── Relevance matching: assign each uploaded image to its best-matching slide ─
    # Slide count follows the CONTENT (not the image count). Images are matched to
    # slides by filename-vs-content keyword overlap; only relevant images are placed,
    # unmatched images are skipped, and unmatched slides stay text-only.
    content_slides_ref = [s for s in slides_data if s.get("type") == "content"]
    matched = match_uploaded_images_to_slides(content_slides_ref, image_paths)
    # Map each content slide object's id → its matched image path (or None).
    slide_image_map = {id(s): matched[i] for i, s in enumerate(content_slides_ref)}

    start = time.time()
    prs = Presentation()
    prs.slide_width  = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)

    for sd in slides_data:
        logger.info(f"Processing slide (with images): {sd.get('title', 'Untitled')}")
        stype = sd.get("type", "content")

        # ── Title / Thank-you / Section-title: reuse standard build_ppt ───────
        if stype in ("title", "thank_you", "section_title"):
            mini = build_ppt([sd])
            xml_slide = mini.slides[0]._element
            new_slide = prs.slides.add_slide(prs.slide_layouts[6])
            sp_tree = new_slide.shapes._spTree
            for child in list(sp_tree):
                sp_tree.remove(child)
            for child in list(xml_slide.spTree):
                sp_tree.append(copy.deepcopy(child))
            new_slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")
            continue

        # ── Content slide ─────────────────────────────────────────────────────
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        img_for_slide = slide_image_map.get(id(sd))

        if img_for_slide:
            # Image on TOP, text in a separate white box BELOW (not overlaid).
            # Consistent yellow header; dark bullet text in a white bordered,
            # padded box directly beneath the image.
            _render_image_top_textbox_slide(
                slide, sd.get("title", ""), sd.get("bullets", []),
                img_for_slide, SLIDE_W, SLIDE_H
            )
        else:
            # Text-only slide: yellow header bar + full-width editable bullets.
            add_header_bar(slide, sd.get("title", ""), slide_w=SLIDE_W, tag=True)

            text_w = SLIDE_W - 1200000
            font_size = Pt(15)
            tb = slide.shapes.add_textbox(Emu(300000), Emu(HEADER_H + 150000),
                                          Emu(text_w), Emu(SLIDE_H - HEADER_H - 250000))
            btf = tb.text_frame
            btf.word_wrap = True
            for i, bullet in enumerate(sd.get("bullets", [])):
                p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
                indent = 1 if bullet.startswith("   ") else 0
                text = bullet.strip()
                p.level = indent
                p.space_after = Pt(6)
                r = p.add_run()
                r.text = ("\u2022 " if indent == 0 else "\u2013 ") + text
                r.font.size = font_size if indent == 0 else Pt(font_size.pt - 2)
                r.font.color.rgb = DARK_GRAY
            tb.name = f"{BULLETS_NAME_PREFIX}_{tb.shape_id}"

        slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "")

    _apply_motion_to_presentation(prs)

    logger.info(f"PPT-with-images generation took {time.time() - start:.2f}s ({len(slides_data)} slides)")
    return prs


def _place_images_in_region(slide, img_paths: list, left: int, top: int,
                            box_w: int, box_h: int):
    """Place 1..N images inside a region, each preserving its aspect ratio with no
    overlap. One image = centred & fit. Multiple images = stacked vertically (or in a
    2-column grid for 3+), each fit within its own cell. Returns count placed."""
    paths = [p for p in (img_paths or []) if p and os.path.isfile(p)]
    if not paths:
        return 0

    n = len(paths)
    gap = 120000  # EMU gap between images

    if n == 1:
        _place_image_fit(slide, paths[0], left, top, box_w, box_h, is_main=True)
        return 1

    if n == 2:
        cell_h = int((box_h - gap) / 2)
        for i, p in enumerate(paths):
            # The first (top) image is treated as the primary/most important one.
            _place_image_fit(slide, p, left, top + i * (cell_h + gap), box_w, cell_h,
                             is_main=(i == 0))
        return 2

    # 3+ images → 2-column grid.
    cols = 2
    rows = (n + cols - 1) // cols
    cell_w = int((box_w - gap * (cols - 1)) / cols)
    cell_h = int((box_h - gap * (rows - 1)) / rows)
    for idx, p in enumerate(paths):
        r, c = divmod(idx, cols)
        cx = left + c * (cell_w + gap)
        cy = top + r * (cell_h + gap)
        _place_image_fit(slide, p, cx, cy, cell_w, cell_h, is_main=(idx == 0))
    return n


def _split_sentences(text: str) -> list:
    """Split a block of text into individual sentences — one per full stop (or ? / !).
    Each sentence becomes its own bullet. Guards against breaking on common
    abbreviations (e.g. 'i.e.', 'e.g.', 'etc.') and decimal numbers (e.g. '3.5')."""
    if not text or not text.strip():
        return []

    protected = text
    # Protect abbreviations and decimals by temporarily masking their dots.
    abbreviations = [
        "i.e.", "e.g.", "etc.", "vs.", "viz.", "no.", "nos.", "mr.", "mrs.",
        "ms.", "dr.", "sr.", "jr.", "govt.", "pvt.", "ltd.", "fig.", "ref.",
        "approx.", "a.m.", "p.m.",
    ]
    for ab in abbreviations:
        # Mask the dots in each abbreviation occurrence WITHOUT altering its case.
        protected = re.sub(
            re.escape(ab),
            lambda m: m.group(0).replace(".", "<DOT>"),
            protected,
            flags=re.IGNORECASE,
        )
    # Protect decimals like 3.5 or 10.25
    protected = re.sub(r"(\d)\.(\d)", r"\1<DOT>\2", protected)

    # Split after . ? or ! when followed by whitespace (or end of string).
    parts = re.split(r"(?<=[.?!])\s+", protected)

    sentences = []
    for part in parts:
        s = part.replace("<DOT>", ".").strip()
        if s:
            sentences.append(s)
    return sentences


def _build_greeting(doc_title: str, sections: list, language: str | None = None) -> str:
    """Compose a welcome message that previews what the session will cover, built
    from the section headings/subheadings (so it reflects the actual content)."""
    topics = []
    seen = set()
    for s in sections:
        t = (s.get("title") or "").strip().rstrip(":").strip()
        if not t:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        topics.append(t)
    lang = _normalize_language(language)
    topic_line = "; ".join(topics)
    if lang == "hindi":
        if topics:
            return (f"नमस्कार! {doc_title} पर इस सत्र में, हम इन विषयों के बारे में सीखेंगे: "
                    f"{topic_line}।")
        return f"नमस्कार! इस सत्र में हम {doc_title} के बारे में सीखेंगे।"
    if lang == "tamil":
        if topics:
            return f"வணக்கம்! {doc_title} குறித்த இந்த அமர்வில், நாம் கற்கவிருக்கும் தலைப்புகள்: {topic_line}."
        return f"வணக்கம்! இந்த அமர்வில் {doc_title} பற்றி அறிந்துகொள்வோம்."
    if lang == "telugu":
        if topics:
            return f"నమస్కారం! {doc_title} పై ఈ సెషన్‌లో మనం నేర్చుకోబోయే అంశాలు: {topic_line}."
        return f"నమస్కారం! ఈ సెషన్‌లో మనం {doc_title} గురించి నేర్చుకుంటాము."
    if lang == "malayalam":
        if topics:
            return f"നമസ്കാരം! {doc_title} എന്ന ഈ സെഷനിൽ നമ്മൾ പഠിക്കുന്ന വിഷയങ്ങൾ: {topic_line}."
        return f"നമസ്കാരം! ഈ സെഷനിൽ നമ്മൾ {doc_title} നെക്കുറിച്ച് പഠിക്കും."
    if topics:
        return (f"Welcome! In this session on {doc_title}, we will learn about: "
                f"{topic_line}.")
    return f"Welcome! In this session we will learn about {doc_title}."


def _overlay_logo(slide, logo_path: str, slide_w: int, slide_h: int):
    """Place a logo in the top-right corner, preserving aspect ratio. Best-effort."""
    if not logo_path or not os.path.isfile(logo_path):
        return
    try:
        from PIL import Image
        with Image.open(logo_path) as im:
            iw, ih = im.size
        if iw <= 0 or ih <= 0:
            return
        # Target box: ~16% of slide width, ~12% of slide height, top-right with margin.
        margin = 200000
        box_w = int(slide_w * 0.16)
        box_h = int(slide_h * 0.12)
        scale = min(box_w / iw, box_h / ih)
        draw_w, draw_h = int(iw * scale), int(ih * scale)
        left = slide_w - margin - draw_w
        top = margin
        slide.shapes.add_picture(logo_path, Emu(left), Emu(top),
                                 width=Emu(draw_w), height=Emu(draw_h))
    except Exception as e:
        logger.info(f"[logo] placement failed for {logo_path}: {e}")


def _copy_built_slide(prs, slide_dict: dict):
    """Render a title/thank_you/section_title slide via build_ppt and copy its shapes
    into `prs`. Returns the new slide so callers can overlay extras (e.g. a logo)."""
    mini = build_ppt([slide_dict])
    xml_slide = mini.slides[0]._element
    new_slide = prs.slides.add_slide(prs.slide_layouts[6])
    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree):
        sp_tree.remove(child)
    for child in list(xml_slide.spTree):
        sp_tree.append(copy.deepcopy(child))
    new_slide.notes_slide.notes_text_frame.text = slide_dict.get("voiceover", "")
    return new_slide


# ─── Narration (voiceover) generators ─────────────────────────────────────────
_NARRATION_SYSTEM_PROMPT = """\
You are a professional trainer explaining PowerPoint slides to an audience.
You will be given a SLIDE TITLE and a list of BULLET POINTS.

RULES:
- Explain each bullet point ONE BY ONE in order, as a trainer naturally would.
- DO NOT repeat the bullet text word-for-word. Rephrase and explain it conversationally.
- Keep every fact, number, name, date, and detail from the bullet. ZERO information loss.
- Use natural transitions: "Now let's look at…", "Coming to the next point…", "Here's an important aspect…" etc.
- Speak in second person ("you") as if addressing trainees.
- Do NOT add information that is not in the bullets.
- Do NOT prefix with "Voiceover:" or "Narration:" — just output the spoken text.
- Target ~25-35 words per bullet. Scale with bullet count.

Return ONLY the narration text (plain text, no JSON, no markdown).
"""

_NARRATION_SYSTEM_PROMPT_HINDI = """\
आप एक पेशेवर प्रशिक्षक हैं जो दर्शकों को PowerPoint स्लाइड समझा रहे हैं।
आपको एक स्लाइड शीर्षक (SLIDE TITLE) और बुलेट पॉइंट्स (BULLET POINTS) दिए जाएंगे।

नियम:
- हर बुलेट पॉइंट को एक-एक करके क्रम में समझाएं, जैसे एक प्रशिक्षक स्वाभाविक रूप से समझाता है।
- बुलेट के शब्दों को हूबहू न दोहराएं। उन्हें अपने शब्दों में, बातचीत की शैली में समझाएं।
- हर तथ्य, संख्या, नाम, तारीख और विवरण बरकरार रखें। कोई जानकारी न छूटे।
- स्वाभाविक जोड़ने वाले वाक्यांशों का उपयोग करें: "अब देखते हैं…", "अगला महत्वपूर्ण बिंदु…" आदि।
- "आप" का प्रयोग करते हुए दर्शकों को संबोधित करें।
- बुलेट में न दी गई कोई जानकारी न जोड़ें।
- "Voiceover:" या "Narration:" जैसा कोई उपसर्ग न लगाएं — केवल बोला जाने वाला पाठ दें।
- पूरा उत्तर हिंदी (देवनागरी) में दें। संख्याएं, संक्षिप्ताक्षर और विशेष नाम यथावत रखें।

केवल कथन (narration) पाठ लौटाएं (सादा पाठ, कोई JSON या markdown नहीं)।
"""


def _generate_narration_llm(title: str, bullets: list,
                            ollama_url: str, model: str,
                            language: str | None = None) -> str | None:
    """Ask the LLM to produce a trainer-style narration for one slide, in `language`."""
    bullet_text = "\n".join(f"- {b}" for b in bullets)
    prompt = f"SLIDE TITLE: {title}\n\nBULLET POINTS:\n{bullet_text}\n\nGenerate the spoken narration for this slide."
    lang = _normalize_language(language)
    if lang == "hindi":
        system = _NARRATION_SYSTEM_PROMPT_HINDI
    elif lang in SUPPORTED_LANGUAGES and lang != "english":
        system = (_NARRATION_SYSTEM_PROMPT + f"\n\nOUTPUT LANGUAGE REQUIREMENT:\n"
                  f"- Write all spoken narration in fluent {_LANGUAGE_NAMES[lang]} using "
                  f"{_LANGUAGE_SCRIPTS[lang]}.\n"
                  "- Preserve numbers, names, dates, abbreviations, and every fact from the bullets.")
    else:
        system = _NARRATION_SYSTEM_PROMPT
    try:
        url = f"{ollama_url}/api/generate"
        payload = {
            "model": model,
            "prompt": prompt,
            "system": system,
            "stream": False,
            "options": {"temperature": 0.4, "num_ctx": 4096, "num_predict": 2048},
        }
        resp = http_requests.post(url, json=payload, timeout=120)
        resp.raise_for_status()
        text = resp.json().get("response", "").strip()
        # Strip the <think>…</think> block some models produce (e.g. deepseek-r1)
        text = re.sub(r"<think>[\s\S]*?</think>", "", text).strip()
        if text and len(text) > 20:
            return text
    except Exception as e:
        logger.info(f"[narration] LLM failed for '{title[:30]}': {e}")
    return None


# Transition phrases for the local fallback narrator, rotated per bullet.
_NARRATION_TRANSITIONS = [
    "Let's start with this — ",
    "Now, the next point here is — ",
    "Coming to the next one — ",
    "Another important aspect is — ",
    "Here's something to note — ",
    "Moving on — ",
    "Next up — ",
    "Let's also look at this — ",
    "Now here's the thing — ",
    "And importantly — ",
]

# Hindi transition phrases for the local fallback narrator, rotated per bullet.
_NARRATION_TRANSITIONS_HINDI = [
    "सबसे पहले — ",
    "अब अगला बिंदु है — ",
    "अगली बात यह है कि — ",
    "एक और महत्वपूर्ण पहलू है — ",
    "ध्यान देने योग्य बात — ",
    "आगे बढ़ते हैं — ",
    "इसके बाद — ",
    "इसे भी देखते हैं — ",
    "अब बात यह है कि — ",
    "और महत्वपूर्ण रूप से — ",
]

_NARRATION_TRANSITIONS_INDIAN = {
    "tamil": ["முதலில் — ", "அடுத்ததாக — ", "இப்போது அடுத்த கருத்து — ", "மற்றொரு முக்கிய அம்சம் — "],
    "telugu": ["మొదటగా — ", "తర్వాతి అంశం — ", "ఇప్పుడు తదుపరి విషయం — ", "మరో ముఖ్యమైన అంశం — "],
    "malayalam": ["ആദ്യം — ", "അടുത്തതായി — ", "ഇനി അടുത്ത വിഷയം — ", "മറ്റൊരു പ്രധാന കാര്യം — "],
}


def _build_narration_local(title: str, bullets: list, language: str | None = None) -> str:
    """Build a human-ish narration locally (no LLM) by rephrasing each bullet with
    varied transitions so the notes are NOT identical to the slide text."""
    if not bullets:
        return ""
    lang = _normalize_language(language)
    is_hindi = lang == "hindi"
    transitions = (_NARRATION_TRANSITIONS_HINDI if is_hindi else
                   _NARRATION_TRANSITIONS_INDIAN.get(lang, _NARRATION_TRANSITIONS))
    parts = []
    n = len(transitions)
    for i, b in enumerate(bullets):
        b = b.strip().rstrip(".")
        if not b:
            continue
        transition = transitions[i % n]
        if lang == "english":
            # Lower-case the first letter after the transition for natural flow,
            # but preserve acronyms (all-caps words). (English only.)
            first_char = b[0]
            if len(b) > 1 and first_char.isupper() and b[1].islower():
                b = first_char.lower() + b[1:]
        parts.append(f"{transition}{b}.")
    if not parts:
        return ""
    clean_title = title.strip().rstrip(":.-").strip()
    if is_hindi:
        topic = clean_title or "इस खंड"
        opening = f"आइए {topic} के मुख्य बिंदुओं पर चर्चा करते हैं। "
        closing = " तो ये थे इस खंड के मुख्य बिंदु।"
    elif lang == "tamil":
        topic = clean_title or "இந்தப் பகுதி"
        opening = f"{topic} குறித்த முக்கிய அம்சங்களைப் பார்ப்போம். "
        closing = " இவை இந்தப் பகுதியின் முக்கிய அம்சங்கள்."
    elif lang == "telugu":
        topic = clean_title or "ఈ విభాగం"
        opening = f"{topic}లోని ముఖ్యమైన అంశాలను పరిశీలిద్దాం. "
        closing = " ఇవి ఈ విభాగంలోని ప్రధాన అంశాలు."
    elif lang == "malayalam":
        topic = clean_title or "ഈ ഭാഗം"
        opening = f"{topic} ലെ പ്രധാന കാര്യങ്ങൾ നോക്കാം. "
        closing = " ഇവയാണ് ഈ ഭാഗത്തിലെ പ്രധാന കാര്യങ്ങൾ."
    else:
        topic = clean_title or "this section"
        opening = f"Let's go through the key points on {topic}. "
        closing = " So those are the main points for this section."
    return opening + " ".join(parts) + closing


def generate_slide_narration(title: str, bullets: list,
                             ollama_url: str = None, model: str = None,
                             language: str | None = None) -> str:
    """Generate a human-like narration for one slide's notes. Tries the LLM first;
    falls back to a local rephraser if the LLM is unavailable or fails."""
    if ollama_url and model:
        llm_result = _generate_narration_llm(title, bullets, ollama_url, model, language=language)
        if llm_result:
            return llm_result
    return _build_narration_local(title, bullets, language=language)


# ─── On-slide text translation (titles + bullets) ────────────────────────────
_TRANSLATE_SYSTEM_PROMPT = """\
You are a precise translator. Translate the given lines into natural, fluent {language}
using {script}.

STRICT RULES:
- Return a JSON object: {{"lines": ["<translated line 1>", "<translated line 2>", ...]}} with EXACTLY
  the same number of lines, in the same order, as the input.
- Translate the MEANING faithfully — do not add, drop, merge, or reorder lines.
- Keep numbers, percentages, dates, monetary values, scheme names, proper nouns and
  abbreviations (e.g. MSME, DPIIT, CIC, EMI, KYC) intact; you may transliterate common
  terms naturally.
- Translate every source sentence; never copy an English sentence unchanged.
- Use only {language} script for translated prose. Do not mix in words from any other language,
  and do not add labels, explanations, markdown, or narration wrappers.
- Keep Latin text only for names, abbreviations, and numbers present in the source line,
  or exact business terms explicitly listed in the translation request.
- Return ONLY the JSON object, nothing else.
"""


class LanguageTranslationError(RuntimeError):
    """Raised when translated output cannot be verified as the requested language."""


_LANGUAGE_SCRIPT_RANGES = {
    "hindi": ((0x0900, 0x097F),),
    "tamil": ((0x0B80, 0x0BFF),),
    "telugu": ((0x0C00, 0x0C7F),),
    "malayalam": ((0x0D00, 0x0D7F),),
}


def _language_text_is_valid(text: str, language: str, source_text: str = "",
                            allowed_terms: list[str] | None = None) -> bool:
    """Reject mixed-script output and unchanged English prose for Indian languages."""
    lang = _normalize_language(language)
    if lang == "english":
        return True

    script_ranges = _LANGUAGE_SCRIPT_RANGES.get(lang, ())
    native_letters = 0
    letter_count = 0
    source_latin = {word.casefold() for word in re.findall(r"[A-Za-z]+", source_text)}
    allowed_latin = {
        word.casefold()
        for term in (allowed_terms or [])
        for word in re.findall(r"[A-Za-z]+", str(term))
    }

    def is_source_token(token: str) -> bool:
        token = token.casefold()
        if token in source_latin or token in allowed_latin:
            return True
        if token.endswith("'s") and token[:-2] in source_latin | allowed_latin:
            return True
        if token.endswith("ies") and token[:-3] + "y" in source_latin | allowed_latin:
            return True
        if token.endswith("es") and token[:-2] in source_latin | allowed_latin:
            return True
        if token.endswith("s") and token[:-1] in source_latin | allowed_latin:
            return True
        return False

    for char in text:
        if not unicodedata.category(char).startswith("L"):
            continue
        letter_count += 1
        codepoint = ord(char)
        if any(start <= codepoint <= end for start, end in script_ranges):
            native_letters += 1
        elif "LATIN" not in unicodedata.name(char, ""):
            return False

    latin_words = re.findall(r"[A-Za-z]+", text)
    allowed_latin_letters = 0
    for word in latin_words:
        if not is_source_token(word):
            return False
        allowed_latin_letters += len(word)

    source_words = re.findall(r"[A-Za-z]{3,}", source_text)
    has_source_prose = any(not word.isupper() for word in source_words)
    if has_source_prose and native_letters == 0:
        return False
    translated_letters = letter_count - allowed_latin_letters
    if has_source_prose and translated_letters and native_letters / translated_letters < 0.2:
        return False
    return True


def _translation_fallback_model(ollama_url: str, selected_model: str) -> str | None:
    """Choose an installed instruction model suited to multilingual translation."""
    try:
        response = http_requests.get(f"{ollama_url}/api/tags", timeout=8)
        response.raise_for_status()
        available = {item.get("name") for item in response.json().get("models", [])}
    except Exception as e:
        logger.info(f"[translate] could not list fallback models: {e}")
        return None

    preferred = (
        "qwen2.5:32b", "qwen2.5:14b", "qwen2.5:7b", "qwen2.5:3b",
        "deepseek-r1:32b", "llama3.1:latest", "llama3.2:latest", "llama3:latest",
    )
    return next((name for name in preferred if name in available and name != selected_model), None)


def _translate_lines(lines: list, ollama_url: str | None, model: str | None,
                     language: str | None, allowed_terms: list[str] | None = None) -> list:
    """Translate lines and reject output that does not match the requested language."""
    lang = _normalize_language(language)
    if lang not in SUPPORTED_LANGUAGES or lang == "english":
        return lines
    clean = [str(l) for l in (lines or [])]
    if not clean:
        return clean

    def request_translation(batch: list[str], model_override: str | None = None) -> list[str] | None:
        numbered = "\n".join(f"{i + 1}. {line}" for i, line in enumerate(batch))
        prompt = (f"Translate each of these {len(batch)} lines into {_LANGUAGE_NAMES[lang]}. "
                  f'Return exactly {len(batch)} translated strings in JSON under "lines":\n\n'
                  f"{numbered}")
        terms = [str(term).strip() for term in (allowed_terms or []) if str(term).strip()]
        if terms:
            prompt += ("\n\nKeep these business terms exactly as written in English when they "
                       f"appear: {', '.join(terms)}.")
        url = f"{ollama_url}/api/generate"
        payload = {
            "model": model_override or model,
            "prompt": prompt,
            "system": _TRANSLATE_SYSTEM_PROMPT.format(
                language=_LANGUAGE_NAMES[lang], script=_LANGUAGE_SCRIPTS[lang]),
            "stream": False,
            "options": {"temperature": 0.2, "num_ctx": 8192, "num_predict": 4096},
            "format": "json",
        }
        resp = http_requests.post(url, json=payload, timeout=180)
        resp.raise_for_status()
        text = resp.json().get("response", "") or ""
        text = re.sub(r"<think>[\s\S]*?</think>", "", text).strip()
        parsed = _parse_json_response(text) or {}
        out = parsed.get("lines")
        return [str(item).strip() for item in out] if isinstance(out, list) and len(out) == len(batch) else None

    if not ollama_url or not model:
        if all(_language_text_is_valid(line, lang, line, allowed_terms) for line in clean):
            return clean
        raise LanguageTranslationError(f"Translation to {_LANGUAGE_NAMES[lang]} requires a configured Ollama model.")

    try:
        translated = request_translation(clean)
    except Exception as e:
        logger.warning(f"[translate] request failed for {lang}: {e}")
        translated = None

    if translated is None:
        invalid_indices = list(range(len(clean)))
        translated = list(clean)
    else:
        invalid_indices = [
            i for i, value in enumerate(translated)
            if not value or not _language_text_is_valid(value, lang, clean[i], allowed_terms)
        ]

    fallback_model = _translation_fallback_model(ollama_url, model) if invalid_indices else None
    if fallback_model:
        logger.warning(f"[translate] retrying {len(invalid_indices)} line(s) with fallback model {fallback_model}")
        try:
            fallback_batch = request_translation(
                [clean[index] for index in invalid_indices], fallback_model
            )
        except Exception as e:
            logger.warning(f"[translate] fallback model {fallback_model} failed: {e}")
            fallback_batch = None
        if fallback_batch:
            for index, value in zip(invalid_indices, fallback_batch):
                if value and _language_text_is_valid(value, lang, clean[index], allowed_terms):
                    translated[index] = value

    invalid_indices = [
        i for i, value in enumerate(translated)
        if not value or not _language_text_is_valid(value, lang, clean[i], allowed_terms)
    ]
    for index in invalid_indices:
        retry_models = [fallback_model] if fallback_model else []
        if model not in retry_models:
            retry_models.append(model)
        for retry_model in retry_models:
            try:
                retry = request_translation([clean[index]], retry_model)
            except Exception as e:
                logger.warning(f"[translate] {retry_model} retry failed for line {index + 1}: {e}")
                retry = None
            if retry and retry[0] and _language_text_is_valid(
                retry[0], lang, clean[index], allowed_terms
            ):
                translated[index] = retry[0]
                break
        else:
            raise LanguageTranslationError(
                f"Could not produce verified {_LANGUAGE_NAMES[lang]} text for line {index + 1}. "
                "Check that an available multilingual model is installed and try again."
            )

    return translated


def _translate_text(text: str, ollama_url: str | None, model: str | None,
                    language: str | None, allowed_terms: list[str] | None = None) -> str:
    """Translate a single string into `language`. Safe fallback to the original."""
    if not text or _normalize_language(language) == "english":
        return text
    return _translate_lines([text], ollama_url, model, language, allowed_terms)[0]


def build_ppt_from_sections(doc_title: str, sections: list, logo_path: str = None,
                            ollama_url: str = None, model: str = None,
                            language: str | None = None) -> Presentation:
    """Build a PPT directly from structured DOCX sections (heading/subheading + body
    + embedded images). Each section becomes ONE content slide whose HEADER is the
    section title (so subheadings appear as slide headers). The section's embedded
    images are placed on that slide, preserving aspect ratio and avoiding text overlap.
    Sections without images are rendered text-only.

    A greeting (title) slide is added at the start and a Thank You slide at the end.
    If `logo_path` is provided, the logo is overlaid on BOTH the first and last slides.
    If `ollama_url` and `model` are provided, narration is generated via LLM per slide;
    otherwise a local rephraser produces the narration so notes differ from slide text.
    `language` controls the greeting, thank-you, slide text and voiceover language.
    """
    SLIDE_W, SLIDE_H, HEADER_H = 12192000, 6858000, 530328
    PAD = 300000
    lang = _normalize_language(language)
    is_hindi = lang == "hindi"
    is_non_english = lang != "english"

    start = time.time()
    prs = Presentation()
    prs.slide_width  = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)

    # ── Greeting / title slide ────────────────────────────────────────────────
    # On-screen subtitle is a short tagline; the spoken narration (notes) is a fuller,
    # natural welcome so the notes are NOT identical to what's shown on the slide.
    # Translate the document title and section names so the title slide matches.
    greeting_doc_title = doc_title
    greeting_sections = sections
    if is_non_english:
        greeting_doc_title = _translate_text(doc_title or "this topic", ollama_url, model, language)
        sec_titles = [s.get("title", "") for s in sections]
        translated_titles = _translate_lines(sec_titles, ollama_url, model, language)
        greeting_sections = [{"title": translated_titles[i]} for i in range(len(sections))]

    greeting_narration = _build_greeting(greeting_doc_title or "this topic",
                                         greeting_sections, language=language)
    greeting_subtitles = {
        "english": "An overview of the key topics covered in this session",
        "hindi": "इस सत्र में शामिल मुख्य विषयों का अवलोकन",
        "tamil": "இந்த அமர்வில் இடம்பெறும் முக்கிய தலைப்புகளின் கண்ணோட்டம்",
        "telugu": "ఈ సెషన్‌లోని ముఖ్యమైన అంశాల సమీక్ష",
        "malayalam": "ഈ സെഷനിലെ പ്രധാന വിഷയങ്ങളുടെ ഒരു അവലോകനം",
    }
    greeting_subtitle = greeting_subtitles[lang]
    first_slide = _copy_built_slide(prs, {
        "type": "title",
        "title": greeting_doc_title or {"english": "Welcome", "hindi": "स्वागत है",
                         "tamil": "வரவேற்கிறோம்", "telugu": "స్వాగతం",
                         "malayalam": "സ്വാഗതം"}[lang],
        "subtitle": greeting_subtitle,
        "voiceover": greeting_narration,
    })
    if logo_path:
        _overlay_logo(first_slide, logo_path, SLIDE_W, SLIDE_H)

    for sec in sections:
        title = sec.get("title", "").strip()
        body  = sec.get("body", []) or []
        images = [p for p in (sec.get("images", []) or []) if p and os.path.isfile(p)]

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Body bullets (auto-fit font to amount of text). Split each body paragraph
        # into individual sentences (one per full stop) so every distinct point gets
        # its own bullet — this also enables the one-by-one, click-driven reveal.
        bullets = []
        for line in body:
            sentences = _split_sentences(line)
            if sentences:
                bullets.extend(sentences)
            elif line.strip():
                bullets.append(line.strip())

        # DOCX text keeps its source language, so translate titles and bullets here.
        if is_non_english:
            title = _translate_text(title, ollama_url, model, language)
            bullets = _translate_lines(bullets, ollama_url, model, language)

        has_img = bool(images)
        valid_imgs = [p for p in images if p and os.path.isfile(p)]

        if valid_imgs:
            # ── IMAGE(S) ON TOP, text in a white box BELOW ────────────────────
            # Consistent layout for EVERY image+text slide regardless of image count:
            # yellow header; image(s) in the top region (multiple shown one-at-a-time
            # on click); dark bullet text in a white, bordered, padded box directly
            # beneath the image — never overlaid.
            _render_image_top_textbox_slide(slide, title, bullets, valid_imgs,
                                            SLIDE_W, SLIDE_H)
        else:
            # ── Text-only slide: yellow header bar + full-width editable bullets ─
            add_header_bar(slide, title, slide_w=SLIDE_W, tag=True)

            text_left = 600000
            text_w = SLIDE_W - 1200000
            bullet_count = max(1, len(bullets))
            font_size = Pt(18) if bullet_count <= 5 else Pt(15) if bullet_count <= 7 else Pt(13)
            if bullet_count > 10:
                font_size = Pt(11)

            tb = slide.shapes.add_textbox(Emu(text_left), Emu(HEADER_H + 150000),
                                          Emu(text_w), Emu(SLIDE_H - HEADER_H - 250000))
            btf = tb.text_frame
            btf.word_wrap = True
            if bullets:
                for i, sentence in enumerate(bullets):
                    p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
                    p.space_after = Pt(4)
                    r = p.add_run()
                    r.text = "\u2022 " + sentence
                    r.font.size = font_size
                    r.font.color.rgb = DARK_GRAY
            tb.name = f"{BULLETS_NAME_PREFIX}_{tb.shape_id}"

        # Voiceover narration: a human-like explanation, NOT identical to the slide text.
        narration = generate_slide_narration(title, bullets,
                                             ollama_url=ollama_url, model=model,
                                             language=language)
        if is_non_english and not _language_text_is_valid(
            narration, lang, " ".join([title, *bullets])
        ):
            narration = _build_narration_local(title, bullets, language=lang)
        slide.notes_slide.notes_text_frame.text = narration

    # ── Thank You slide ───────────────────────────────────────────────────────
    thankyou_slides = {
        "english": {"title": "Thank You", "subtitle": "For more details refer the source document",
                    "voiceover": "That brings us to the end of this session. Thank you for your attention. For any finer details, please refer to the source document we discussed."},
        "hindi": {"title": "धन्यवाद", "subtitle": "अधिक जानकारी के लिए स्रोत दस्तावेज़ देखें",
                  "voiceover": "इसी के साथ हम इस सत्र के अंत पर पहुंचते हैं। आपके ध्यान के लिए धन्यवाद। किसी भी बारीक विवरण के लिए, कृपया हमारे द्वारा चर्चा किए गए स्रोत दस्तावेज़ को देखें।"},
        "tamil": {"title": "நன்றி", "subtitle": "மேலும் விவரங்களுக்கு மூல ஆவணத்தைப் பார்க்கவும்",
                  "voiceover": "இதனுடன் இந்த அமர்வு நிறைவடைகிறது. உங்கள் கவனத்திற்கு நன்றி. கூடுதல் விவரங்களுக்கு, நாம் பார்த்த மூல ஆவணத்தைப் பார்க்கவும்."},
        "telugu": {"title": "ధన్యవాదాలు", "subtitle": "మరిన్ని వివరాలకు మూల పత్రాన్ని చూడండి",
                   "voiceover": "ఇంతటితో ఈ సెషన్ ముగుస్తుంది. మీ శ్రద్ధకు ధన్యవాదాలు. మరిన్ని వివరాల కోసం మనం చర్చించిన మూల పత్రాన్ని చూడండి."},
        "malayalam": {"title": "നന്ദി", "subtitle": "കൂടുതൽ വിവരങ്ങൾക്ക് ഉറവിട രേഖ കാണുക",
                      "voiceover": "ഇതോടെ ഈ സെഷൻ അവസാനിക്കുന്നു. ശ്രദ്ധിച്ചതിന് നന്ദി. കൂടുതൽ വിശദാംശങ്ങൾക്ക് നമ്മൾ ചർച്ച ചെയ്ത ഉറവിട രേഖ കാണുക."},
    }
    thankyou_slide = {"type": "thank_you", **thankyou_slides[lang]}
    last_slide = _copy_built_slide(prs, thankyou_slide)
    if logo_path:
        _overlay_logo(last_slide, logo_path, SLIDE_W, SLIDE_H)

    _apply_motion_to_presentation(prs)

    logger.info(f"PPT-from-sections took {time.time() - start:.2f}s "
                f"({len(sections)} content slides + greeting + thank-you"
                f"{', logo' if logo_path else ''})")
    return prs


# ─── New endpoint: DOCX + Images → single PPT ────────────────────────────────
@app.post("/api/docx-image-to-ppt")
async def docx_image_to_ppt(
    docx_file:   UploadFile = File(...),
    images:      List[UploadFile] = File(default=[]),
    logo:        UploadFile = File(default=None),
    scheme_name: str = Form(default=""),
    ollama_url:  str = Form(default=DEFAULT_OLLAMA_URL),
    model:       str = Form(default=DEFAULT_MODEL),
    num_slides:  int = Form(default=5),
    imgs_per_slide: int = Form(default=5),
    embedded_images: bool = Form(default=True),
    language:    str = Form(default=DEFAULT_LANGUAGE),
):
    """
    Accept one DOCX + one or more images.
    Extract content → Ollama → build single PPT with images embedded.
    num_slides      : target number of content slides (default 5)
    imgs_per_slide  : images to show per content slide (default 5)
    Returns downloadable .pptx
    """
    global IMAGES_PER_SLIDE, BULLETS_PER_SLIDE
    IMAGES_PER_SLIDE = max(1, imgs_per_slide)
    BULLETS_PER_SLIDE = max(1, num_slides)  # reuse as bullets-per-chunk to control slide count
    # ── Save DOCX ─────────────────────────────────────────────────────────────
    docx_path = os.path.join(UPLOAD_DIR, docx_file.filename)
    with open(docx_path, "wb") as f:
        shutil.copyfileobj(docx_file.file, f)
    logger.info(f"[docx-image-to-ppt] DOCX saved: {docx_path}")

    # ── Save images ───────────────────────────────────────────────────────────
    img_dir = os.path.join(UPLOAD_DIR, "_img_tmp")
    os.makedirs(img_dir, exist_ok=True)
    saved_images = []
    for img in images:
        if not img.filename:
            continue
        ext = os.path.splitext(img.filename)[1].lower()
        if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"):
            continue
        img_path = os.path.join(img_dir, img.filename)
        with open(img_path, "wb") as f:
            shutil.copyfileobj(img.file, f)
        saved_images.append(img_path)
        logger.info(f"[docx-image-to-ppt] Image saved: {img_path}")

    # ── Save logo (if uploaded) ───────────────────────────────────────────────
    logo_path = None
    if logo and logo.filename:
        ext = os.path.splitext(logo.filename)[1].lower()
        if ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"):
            logo_path = os.path.join(img_dir, "logo_" + logo.filename)
            with open(logo_path, "wb") as f:
                shutil.copyfileobj(logo.file, f)
            logger.info(f"[docx-image-to-ppt] Logo saved: {logo_path}")

    # ── Embedded-image mode: use images that live INSIDE the DOCX ─────────────
    # Each heading/subheading becomes a slide header; the section's embedded images
    # are attached to that slide (text-only when a section has no image). This runs
    # only when the DOCX actually contains embedded images in its sections.
    if embedded_images:
        try:
            doc_title, sections, img_tmp = extract_docx_sections_with_images(docx_path)
        except Exception as e:
            logger.info(f"[docx-image-to-ppt] section extraction failed, "
                        f"falling back to LLM flow: {e}")
            sections, img_tmp = [], None

        has_embedded = any(s.get("images") for s in sections)
        if has_embedded:
            logger.info(f"[docx-image-to-ppt] embedded-image mode: {len(sections)} sections, "
                        f"{sum(len(s['images']) for s in sections)} embedded images")
            try:
                try:
                    prs = build_ppt_from_sections(doc_title, sections, logo_path=logo_path,
                                                  ollama_url=ollama_url, model=model,
                                                  language=language)
                except LanguageTranslationError as e:
                    logger.warning(f"[docx-image-to-ppt] translation rejected: {e}")
                    return JSONResponse(status_code=502, content={
                        "status": "error",
                        "message": str(e),
                    })
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)

                base = scheme_name.strip() or doc_title or os.path.splitext(docx_file.filename)[0]
                safe = re.sub(r'[<>:"/\\|?*]', '', base)[:80].strip()
                safe = safe.encode("ascii", "ignore").decode("ascii").strip() or "presentation"
                fname = f"{safe}.pptx"
                content_count = len(sections) + (1 if doc_title else 0)
                logger.info(f"[docx-image-to-ppt] Done (embedded) → {content_count} slides → {fname}")

                return StreamingResponse(
                    buffer,
                    media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    headers={
                        "Content-Disposition": f'attachment; filename="{fname}"',
                        "X-Slide-Count": str(content_count),
                        "X-Image-Mode": "embedded",
                    },
                )
            finally:
                if img_tmp:
                    shutil.rmtree(img_tmp, ignore_errors=True)
        else:
            logger.info("[docx-image-to-ppt] no embedded images found — using LLM flow")
            if img_tmp:
                shutil.rmtree(img_tmp, ignore_errors=True)

    # ── Extract text ──────────────────────────────────────────────────────────
    title = scheme_name.strip() or os.path.splitext(docx_file.filename)[0]
    raw_content = extract_docx_section_by_title(docx_path, title) if scheme_name.strip() else extract_docx_text(docx_path)

    if not raw_content or len(raw_content) < 20:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Could not extract content from DOCX"})

    logger.info(f"[docx-image-to-ppt] Extracted {len(raw_content)} chars for '{title}'")

    # ── Generate slides via Ollama ────────────────────────────────────────────
    result = generate_slides(title, raw_content, DEFAULT_SYSTEM_PROMPT, ollama_url, model, language=language)
    if not result or "slides" not in result:
        return JSONResponse(status_code=500, content={"status": "error", "message": "Ollama failed to generate slides"})

    slides = [
        s for s in result["slides"]
        if s.get("type") in ("title", "thank_you", "section_title")
        or (s.get("bullets") and len(s["bullets"]) > 0)
    ]

    # ── Build PPT with images ─────────────────────────────────────────────────
    prs    = build_ppt_with_images(slides, saved_images)
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    safe = re.sub(r'[<>:"/\\|?*]', '', title)[:80].strip()
    safe = safe.encode("ascii", "ignore").decode("ascii").strip() or "presentation"
    fname = f"{safe}.pptx"
    logger.info(f"[docx-image-to-ppt] Done → {len(slides)} slides → {fname}")

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{fname}"',
            "X-Slide-Count": str(len(slides)),
        },
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ░░  PPT GENERATOR STUDIO  ░░  (independent feature — additive only)
# ───────────────────────────────────────────────────────────────────────────────
# A standalone editable PPT workflow served at /generator. It REUSES the existing
# Ollama integration (call_ollama / generate_slides), DOCX extraction helpers and
# the python-pptx rendering primitives already defined above. None of the existing
# routes, functions, globals, or index.html are modified. All new endpoints live
# under the /api/studio/* namespace and all new state lives in STUDIO_* dirs.
# ═══════════════════════════════════════════════════════════════════════════════
import hashlib
import glob as _glob
from fastapi.responses import FileResponse

STUDIO_IMG_DIR = os.path.join(UPLOAD_DIR, "_studio_img")
STUDIO_TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "ppt_templates")
STUDIO_THUMB_DIR = os.path.join(UPLOAD_DIR, "_studio_thumbs")
os.makedirs(STUDIO_IMG_DIR, exist_ok=True)
os.makedirs(STUDIO_TEMPLATE_DIR, exist_ok=True)
os.makedirs(STUDIO_THUMB_DIR, exist_ok=True)

# ── Visual templates (layout themes) ──────────────────────────────────────────
# Each theme is a pure colour palette applied at render time. We do NOT mutate the
# module-level colour constants used by the other routes; the studio builder reads
# the theme explicitly. Any .pptx dropped into ppt_templates/ is also exposed as a
# selectable template whose slide master/background is reused on export.
STUDIO_THEMES = {
    "classic_yellow": {
        "label": "Classic Yellow",
        "header": "FFCB08", "header_text": "121212",
        "body_bg": "FFFFFF", "body_text": "121212", "accent": "FFCB08",
    },
    "royal_maroon": {
        "label": "Royal Maroon",
        "header": "9E173B", "header_text": "FFFFFF",
        "body_bg": "FFFFFF", "body_text": "222222", "accent": "FFC000",
    },
    "ocean_blue": {
        "label": "Ocean Blue",
        "header": "0F4C81", "header_text": "FFFFFF",
        "body_bg": "F5F9FF", "body_text": "16293B", "accent": "4FA3E3",
    },
    "forest_green": {
        "label": "Forest Green",
        "header": "1E6E4F", "header_text": "FFFFFF",
        "body_bg": "F3FBF6", "body_text": "15332A", "accent": "3FBF84",
    },
    "slate_dark": {
        "label": "Slate Dark",
        "header": "1F2937", "header_text": "F9FAFB",
        "body_bg": "111827", "body_text": "E5E7EB", "accent": "F59E0B",
    },
}

# ── Default Hindi glossary: business terms kept in English inside Hindi slides ──
STUDIO_DEFAULT_GLOSSARY = [
    "Child Loan", "Parent Loan", "Personal Loan", "EMI", "Interest Rate",
    "Home Loan", "Credit Score", "Processing Fee", "Tenure", "Collateral",
]

STUDIO_LANGUAGES = [
    {"key": "english", "label": "English"},
    {"key": "hindi", "label": "Hindi (हिंदी)"},
    {"key": "tamil", "label": "Tamil (தமிழ்)"},
    {"key": "telugu", "label": "Telugu (తెలుగు)"},
    {"key": "malayalam", "label": "Malayalam (മലയാളം)"},
    {"key": "kannada", "label": "Kannada (ಕನ್ನಡ)"},
    {"key": "marathi", "label": "Marathi (मराठी)"},
    {"key": "bengali", "label": "Bengali (বাংলা)"},
]

_STUDIO_LANG_DIRECTIVES = {
    "tamil": "Tamil (தமிழ்) using Tamil script",
    "telugu": "Telugu (తెలుగు) using Telugu script",
    "kannada": "Kannada (ಕನ್ನಡ) using Kannada script",
    "marathi": "Marathi (मराठी) using Devanagari script",
    "bengali": "Bengali (বাংলা) using Bengali script",
}


def _studio_language_prompt(system_prompt: str, language: str, glossary: list) -> str:
    """Append a language directive for the studio flow.

    Supported languages reuse the project's existing directives (via _apply_language).
    Other Indian languages get a generic directive. For NON-English output a glossary
    of business terms is injected so they stay in English inside the translated text.
    """
    lang = (language or "english").strip().lower()
    if lang in SUPPORTED_LANGUAGES:
        prompt = _apply_language(system_prompt, lang)
    else:
        directive = _STUDIO_LANG_DIRECTIVES.get(
            lang, f"{language} (use the native script of {language})")
        prompt = (system_prompt or "") + f"""

=============================================
LANGUAGE REQUIREMENT — OUTPUT IN {language.upper()}
=============================================
- Produce ALL human-readable text (slide titles, bullets, subtitles, voiceover) in {directive}.
- Keep numbers, percentages, dates, monetary values, scheme names and abbreviations intact.
- The JSON structure, keys, and "type" values MUST stay in English. Only the human
  readable VALUES are translated.
"""
    glossary = [g.strip() for g in (glossary or []) if g and g.strip()]
    if lang != "english" and glossary:
        terms = ", ".join(f'"{g}"' for g in glossary)
        prompt += f"""

=============================================
GLOSSARY — KEEP THESE BUSINESS TERMS IN ENGLISH
=============================================
- The following business/product terms MUST remain written in English (Latin script)
  EXACTLY as given, even inside otherwise-translated titles, bullets and voiceover: {terms}.
- Do NOT translate or transliterate these terms. Keep their English spelling verbatim.
"""
    return prompt


def _studio_apply_glossary(slides: list, language: str, glossary: list) -> list:
    """Post-process guard: if the LLM translated a glossary term that has an obvious
    Devanagari/other-script variant we cannot easily reverse, we at least ensure the
    exact English term is preserved where it already appears. This is a light-touch
    normaliser that guarantees casing of glossary terms in bullets/titles.
    """
    lang = (language or "english").strip().lower()
    glossary = [g.strip() for g in (glossary or []) if g and g.strip()]
    if lang == "english" or not glossary:
        return slides
    # Normalise case-insensitive English occurrences to the canonical glossary spelling.
    patterns = [(re.compile(re.escape(g), re.IGNORECASE), g) for g in glossary]

    def _fix(text):
        if not isinstance(text, str):
            return text
        for pat, canon in patterns:
            text = pat.sub(canon, text)
        return text

    for s in slides:
        if "title" in s:
            s["title"] = _fix(s.get("title", ""))
        if "subtitle" in s:
            s["subtitle"] = _fix(s.get("subtitle", ""))
        if "voiceover" in s:
            s["voiceover"] = _fix(s.get("voiceover", ""))
        if isinstance(s.get("bullets"), list):
            s["bullets"] = [_fix(b) for b in s["bullets"]]
    return slides


def _studio_save_image(upload: UploadFile, prefix: str = "") -> dict | None:
    """Persist an uploaded image to STUDIO_IMG_DIR under a content-stable id and
    return {id, filename, url}. Returns None for unsupported/empty uploads."""
    if not upload or not upload.filename:
        return None
    ext = os.path.splitext(upload.filename)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"):
        return None
    data = upload.file.read()
    if not data:
        return None
    digest = hashlib.md5(data).hexdigest()[:16]
    img_id = f"{prefix}{digest}{ext}"
    path = os.path.join(STUDIO_IMG_DIR, img_id)
    with open(path, "wb") as f:
        f.write(data)
    return {"id": img_id, "filename": upload.filename, "url": f"/api/studio/image/{img_id}"}


def _studio_resolve_image(image_id: str) -> str | None:
    """Map a studio image id OR a known embedded temp path back to a readable file."""
    if not image_id:
        return None
    cand = os.path.join(STUDIO_IMG_DIR, os.path.basename(image_id))
    if os.path.isfile(cand):
        return cand
    # Also allow absolute paths that live within the uploads tree (embedded images).
    norm = os.path.normpath(image_id)
    if os.path.isfile(norm) and os.path.commonpath([norm, UPLOAD_DIR]) == os.path.normpath(UPLOAD_DIR):
        return norm
    return None


def _studio_build_pptx(slides, theme_key: str, logo_id: str | None):
    """Build a Presentation from the (edited) studio slide model so the OUTPUT matches
    the project's reference deck (…_with_images.pptx) exactly.

    It REUSES the same proven renderers the embedded-image flow uses:
      • title / section_title / thank_you  →  _copy_built_slide  (yellow header band,
        white body, centred title + subtitle, logo overlay on first & last slides)
      • content WITH image(s)              →  _render_image_top_textbox_slide
        (full-bleed image background + yellow header + bullet text box) — the exact
        'image on top, text box below/over' layout seen in the reference file
      • content WITHOUT image              →  yellow header + full-width bullet textbox
        (identical to build_ppt's text-only branch)
    Finally _apply_motion_to_presentation adds the same entrance/transition timing as
    the reference. `theme_key` selects a header palette; 'classic_yellow' reproduces
    the reference's #FFCB08 header verbatim. The logo is overlaid via _overlay_logo.
    """
    SLIDE_W, SLIDE_H, HEADER_H = 12192000, 6858000, 530328
    theme = STUDIO_THEMES.get(theme_key, STUDIO_THEMES["classic_yellow"])
    header_rgb = RGBColor.from_string(theme["header"])
    header_txt_rgb = RGBColor.from_string(theme["header_text"])
    is_classic = (theme_key in (None, "classic_yellow")) or theme_key not in STUDIO_THEMES

    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)

    logo_path = _studio_resolve_image(logo_id) if logo_id else None
    n = len(slides)

    def _recolor_header(slide):
        """Re-tint the yellow header band + title to the chosen theme (non-classic)."""
        if is_classic:
            return
        for shp in slide.shapes:
            nm = getattr(shp, "name", "") or ""
            if nm.startswith("KBHeader") or nm in ("Rectangle 2",):
                try:
                    shp.fill.solid()
                    shp.fill.fore_color.rgb = header_rgb
                except Exception:
                    pass
                if shp.has_text_frame:
                    for p in shp.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.color.rgb = header_txt_rgb

    for idx, sd in enumerate(slides):
        stype = sd.get("type", "content")

        if stype in ("title", "section_title", "thank_you"):
            # Reuse the exact title/section/thank-you renderer (via build_ppt).
            slide = _copy_built_slide(prs, {
                "type": stype,
                "title": sd.get("title", ""),
                "subtitle": sd.get("subtitle", ""),
                "voiceover": sd.get("voiceover", ""),
            })
            _recolor_header(slide)
            if logo_path and (idx == 0 or idx == n - 1):
                try:
                    _overlay_logo(slide, logo_path, SLIDE_W, SLIDE_H)
                except Exception as e:
                    logger.warning(f"[studio] logo overlay failed: {e}")
            continue

        # ── Content slide ──────────────────────────────────────────────────────
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title = sd.get("title", "")
        bullets = list(sd.get("bullets", []) or [])
        # Support a single image id ("image") or multiple ("images"). All are passed
        # to the renderer which stacks extra screenshots as click-revealed sequence
        # images (matching the reference deck's multi-image slides).
        img_ids = []
        if sd.get("images"):
            img_ids = list(sd.get("images"))
        elif sd.get("image"):
            img_ids = [sd.get("image")]
        img_paths = [p for p in (_studio_resolve_image(i) for i in img_ids)
                     if p and os.path.isfile(p)]

        if img_paths:
            # Exact reference layout: full-bleed image(s) + yellow header + text box.
            _render_image_top_textbox_slide(slide, title, bullets, img_paths,
                                            SLIDE_W, SLIDE_H)
        else:
            # Text-only: yellow header bar + full-width editable bullets (build_ppt parity).
            add_header_bar(slide, title, slide_w=SLIDE_W, tag=True)
            text_left = 600000
            text_top = HEADER_H + 150000
            text_w = SLIDE_W - 1200000
            text_h = SLIDE_H - text_top - 300000
            bc = max(1, len(bullets))
            font_size = Pt(18) if bc <= 5 else Pt(15) if bc <= 7 else Pt(13)
            if bc > 10:
                font_size = Pt(11)
            tb = slide.shapes.add_textbox(Emu(text_left), Emu(text_top), Emu(text_w), Emu(text_h))
            tf = tb.text_frame
            tf.word_wrap = True
            for i, bullet in enumerate(bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_after = Pt(4)
                r = p.add_run()
                r.text = "\u2022 " + str(bullet)
                r.font.size = font_size
                r.font.color.rgb = DARK_GRAY
            tb.name = f"{BULLETS_NAME_PREFIX}_{tb.shape_id}"

        _recolor_header(slide)
        slide.notes_slide.notes_text_frame.text = sd.get("voiceover", "") or ""

    # Same entrance/transition animations as the reference deck.
    try:
        _apply_motion_to_presentation(prs)
    except Exception as e:
        logger.warning(f"[studio] motion apply failed: {e}")

    return prs


# ─── Studio: page ─────────────────────────────────────────────────────────────
@app.get("/generator", response_class=HTMLResponse)
async def studio_page(request: Request):
    """Standalone PPT Generator Studio UI (independent of index.html)."""
    return templates.TemplateResponse(name="ppt_studio.html", request=request)


# ─── Studio: config (languages, themes, default glossary) ──────────────────────
@app.get("/api/studio/config")
async def studio_config():
    return {
        "languages": STUDIO_LANGUAGES,
        "default_glossary": STUDIO_DEFAULT_GLOSSARY,
        "default_ollama_url": DEFAULT_OLLAMA_URL,
        "default_model": DEFAULT_MODEL,
    }


# ─── Studio: Ollama models ─────────────────────────────────────────────────────
@app.get("/api/studio/models")
async def studio_models(ollama_url: str = DEFAULT_OLLAMA_URL):
    try:
        r = http_requests.get(f"{ollama_url}/api/tags", timeout=6)
        r.raise_for_status()
        models = [m["name"] for m in r.json().get("models", [])]
        return {"status": "ok", "models": models}
    except Exception as e:
        return JSONResponse(status_code=500, content={"status": "error", "message": str(e),
                                                       "models": [DEFAULT_MODEL]})


# ─── Studio: templates with thumbnails ─────────────────────────────────────────
@app.get("/api/studio/templates")
async def studio_templates():
    """Return the colour-theme templates plus any .pptx dropped in ppt_templates/."""
    out = []
    for key, t in STUDIO_THEMES.items():
        out.append({
            "id": key, "label": t["label"], "kind": "theme",
            "colors": {"header": f"#{t['header']}", "header_text": f"#{t['header_text']}",
                       "body_bg": f"#{t['body_bg']}", "body_text": f"#{t['body_text']}",
                       "accent": f"#{t['accent']}"},
            "thumb_url": f"/api/studio/template-thumb/{key}",
        })
    for pptx in sorted(_glob.glob(os.path.join(STUDIO_TEMPLATE_DIR, "*.pptx"))):
        name = os.path.splitext(os.path.basename(pptx))[0]
        out.append({"id": f"file:{os.path.basename(pptx)}", "label": name, "kind": "file",
                    "thumb_url": f"/api/studio/template-thumb/file:{os.path.basename(pptx)}"})
    return {"templates": out}


@app.get("/api/studio/template-thumb/{template_id:path}")
async def studio_template_thumb(template_id: str):
    """Generate a lightweight SVG thumbnail for a theme (no external deps)."""
    theme = STUDIO_THEMES.get(template_id, STUDIO_THEMES["classic_yellow"])
    svg = f"""<svg xmlns='http://www.w3.org/2000/svg' width='240' height='135' viewBox='0 0 240 135'>
  <rect width='240' height='135' fill='#{theme['body_bg']}' stroke='#d0d0d0'/>
  <rect width='240' height='34' fill='#{theme['header']}'/>
  <text x='12' y='22' font-family='Segoe UI, sans-serif' font-size='14' font-weight='700' fill='#{theme['header_text']}'>Title</text>
  <rect x='14' y='48' width='150' height='8' rx='4' fill='#{theme['body_text']}' opacity='0.75'/>
  <rect x='14' y='64' width='180' height='8' rx='4' fill='#{theme['body_text']}' opacity='0.55'/>
  <rect x='14' y='80' width='120' height='8' rx='4' fill='#{theme['body_text']}' opacity='0.55'/>
  <rect x='14' y='96' width='160' height='8' rx='4' fill='#{theme['body_text']}' opacity='0.4'/>
  <circle cx='210' cy='70' r='18' fill='#{theme['accent']}'/>
</svg>"""
    return StreamingResponse(io.BytesIO(svg.encode("utf-8")), media_type="image/svg+xml")


# ─── Studio: upload DOCX → extracted sections (initial editable model seed) ────
@app.post("/api/studio/upload-docx")
async def studio_upload_docx(file: UploadFile = File(...)):
    """Save a DOCX and return a lightweight outline (title + section headings) plus
    any embedded images as studio image ids, so the UI can seed/preview content."""
    ext = os.path.splitext(file.filename)[1].lower()
    if ext != ".docx":
        return JSONResponse(status_code=400, content={"status": "error", "message": "Only .docx is supported"})
    docx_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(docx_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    outline, embedded = [], []
    try:
        doc_title, sections, img_tmp = extract_docx_sections_with_images(docx_path)
        try:
            for sec in sections:
                outline.append({"title": sec.get("title", ""),
                                 "level": sec.get("level", "heading"),
                                 "has_images": bool(sec.get("images"))})
                for ip in sec.get("images", []) or []:
                    if ip and os.path.isfile(ip):
                        with open(ip, "rb") as fh:
                            data = fh.read()
                        digest = hashlib.md5(data).hexdigest()[:16]
                        dest = os.path.join(STUDIO_IMG_DIR, f"embed_{digest}.png")
                        if not os.path.isfile(dest):
                            with open(dest, "wb") as fh:
                                fh.write(data)
                        embedded.append({"id": f"embed_{digest}.png",
                                         "url": f"/api/studio/image/embed_{digest}.png",
                                         "section": sec.get("title", "")})
        finally:
            if img_tmp:
                shutil.rmtree(img_tmp, ignore_errors=True)
    except Exception as e:
        logger.warning(f"[studio] docx outline extraction failed: {e}")
        doc_title = os.path.splitext(file.filename)[0]

    return {"status": "ok", "filename": file.filename, "doc_title": doc_title or "",
            "outline": outline, "embedded_images": embedded}


# ─── Studio: upload image / logo ───────────────────────────────────────────────
@app.post("/api/studio/upload-image")
async def studio_upload_image(file: UploadFile = File(...), kind: str = Form(default="image")):
    prefix = "logo_" if kind == "logo" else ""
    saved = _studio_save_image(file, prefix=prefix)
    if not saved:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Unsupported or empty image"})
    saved["status"] = "ok"
    return saved


# ─── Studio: serve an image for preview ────────────────────────────────────────
@app.get("/api/studio/image/{image_id}")
async def studio_image(image_id: str):
    path = _studio_resolve_image(image_id)
    if not path:
        return JSONResponse(status_code=404, content={"status": "error", "message": "Image not found"})
    return FileResponse(path)


def _studio_sections_to_model(doc_title: str, sections: list, language: str,
                              glossary: list, ollama_url: str, model: str) -> list:
    """Convert DOCX sections (heading/subheading + body + embedded images) into the
    editable slide model — the SAME mapping build_ppt_from_sections uses, so every
    section becomes one content slide and every embedded screenshot is covered.

    Produces: [greeting title] + [one content slide per section w/ its images] +
    [thank-you]. Body text is split into one-bullet-per-sentence. For non-English
    languages the section titles/bullets/greeting are translated (glossary terms kept
    in English). Each slide carries LLM/local narration in its voiceover.
    Embedded images are copied into STUDIO_IMG_DIR so the UI can preview & re-export.
    """
    lang = _normalize_language(language)
    is_hindi = lang == "hindi"
    is_non_english = lang in SUPPORTED_LANGUAGES and lang != "english"
    other_lang = (language or "english").strip().lower() not in ("english",)
    gl = [g.strip() for g in (glossary or []) if g and g.strip()]

    def _tr_text(t):
        if not t or not is_non_english:
            return t
        return _translate_text(t, ollama_url, model, language, gl)

    def _tr_lines(lines):
        if not lines or not is_non_english:
            return lines
        return _translate_lines(lines, ollama_url, model, language, gl)

    def _store_image(src_path):
        """Copy an embedded image into STUDIO_IMG_DIR, return its studio id."""
        try:
            with open(src_path, "rb") as fh:
                data = fh.read()
        except Exception:
            return None
        digest = hashlib.md5(data).hexdigest()[:16]
        img_id = f"embed_{digest}.png"
        dest = os.path.join(STUDIO_IMG_DIR, img_id)
        if not os.path.isfile(dest):
            with open(dest, "wb") as fh:
                fh.write(data)
        return img_id

    model_slides = []

    # ── Greeting / title slide ──────────────────────────────────────────────
    greeting_doc_title = _tr_text(doc_title) if is_non_english else doc_title
    greeting_sections = sections
    if is_non_english:
        sec_titles = [s.get("title", "") for s in sections]
        tt = _tr_lines(sec_titles)
        greeting_sections = [{"title": tt[i]} for i in range(len(sections))]
    greeting_vo = _build_greeting(greeting_doc_title or doc_title, greeting_sections,
                                  language=language)
    greeting_sub = {
        "english": "An overview of the key topics covered in this session",
        "hindi": "इस सत्र में शामिल मुख्य विषयों का अवलोकन",
        "tamil": "இந்த அமர்வில் இடம்பெறும் முக்கிய தலைப்புகளின் கண்ணோட்டம்",
        "telugu": "ఈ సెషన్‌లోని ముఖ్యమైన అంశాల సమీక్ష",
        "malayalam": "ഈ സെഷനിലെ പ്രധാന വിഷയങ്ങളുടെ ഒരു അവലോകനം",
    }[lang]
    model_slides.append({
        "type": "title",
        "title": greeting_doc_title or doc_title or {
            "english": "Welcome", "hindi": "स्वागत है", "tamil": "வரவேற்கிறோம்",
            "telugu": "స్వాగతం", "malayalam": "സ്വാഗതം",
        }[lang],
        "subtitle": greeting_sub,
        "voiceover": greeting_vo,
    })

    # ── One content slide per section ────────────────────────────────────────
    for sec in sections:
        title = (sec.get("title") or "").strip()
        body = sec.get("body", []) or []
        imgs = [p for p in (sec.get("images", []) or []) if p and os.path.isfile(p)]

        bullets = []
        for line in body:
            sent = _split_sentences(line)
            if sent:
                bullets.extend(sent)
            elif line.strip():
                bullets.append(line.strip())

        if is_non_english:
            title = _tr_text(title)
            bullets = _tr_lines(bullets)

        image_ids = [i for i in (_store_image(p) for p in imgs) if i]

        narration = generate_slide_narration(title, bullets, ollama_url=ollama_url,
                                              model=model, language=language)
        if is_non_english and not _language_text_is_valid(
            narration, lang, " ".join([title, *bullets]), gl
        ):
            narration = _build_narration_local(title, bullets, language=lang)
        slide = {
            "type": "content",
            "title": title,
            "bullets": bullets,
            "voiceover": narration,
            "image": image_ids[0] if image_ids else None,
        }
        # Keep every screenshot for this section available in the UI (first shown,
        # the rest selectable via the image picker / can be split into new slides).
        if len(image_ids) > 1:
            slide["images"] = image_ids
        model_slides.append(slide)

    # ── Thank-you slide ──────────────────────────────────────────────────────
    thankyou_slides = {
        "english": {"title": "Thank You", "subtitle": "For more details refer the source document",
                    "voiceover": "That brings us to the end of this session. Thank you for your attention. For any finer details, please refer to the source document."},
        "hindi": {"title": "धन्यवाद", "subtitle": "अधिक जानकारी के लिए स्रोत दस्तावेज़ देखें",
                  "voiceover": "इसी के साथ हम इस सत्र के अंत पर पहुंचते हैं। आपके ध्यान के लिए धन्यवाद। किसी भी बारीक विवरण के लिए, कृपया स्रोत दस्तावेज़ देखें."},
        "tamil": {"title": "நன்றி", "subtitle": "மேலும் விவரங்களுக்கு மூல ஆவணத்தைப் பார்க்கவும்",
                  "voiceover": "இதனுடன் இந்த அமர்வு நிறைவடைகிறது. உங்கள் கவனத்திற்கு நன்றி. கூடுதல் விவரங்களுக்கு, நாம் பார்த்த மூல ஆவணத்தைப் பார்க்கவும்."},
        "telugu": {"title": "ధన్యవాదాలు", "subtitle": "మరిన్ని వివరాలకు మూల పత్రాన్ని చూడండి",
                   "voiceover": "ఇంతటితో ఈ సెషన్ ముగుస్తుంది. మీ శ్రద్ధకు ధన్యవాదాలు. మరిన్ని వివరాల కోసం మనం చర్చించిన మూల పత్రాన్ని చూడండి."},
        "malayalam": {"title": "നന്ദി", "subtitle": "കൂടുതൽ വിവരങ്ങൾക്ക് ഉറവിട രേഖ കാണുക",
                      "voiceover": "ഇതോടെ ഈ സെഷൻ അവസാനിക്കുന്നു. ശ്രദ്ധിച്ചതിന് നന്ദി. കൂടുതൽ വിശദാംശങ്ങൾക്ക് നമ്മൾ ചർച്ച ചെയ്ത ഉറവിട രേഖ കാണുക."},
    }
    model_slides.append({"type": "thank_you", **thankyou_slides[lang]})

    # Normalise glossary casing for non-English output.
    if other_lang and gl:
        model_slides = _studio_apply_glossary(model_slides, language, gl)
    return model_slides


# ─── Studio: generate preview slide model (reuses generate_slides) ─────────────
@app.post("/api/studio/preview")
async def studio_preview(request: Request):
    """Generate the editable slide model from the uploaded DOCX + options.

    PRIMARY PATH (matches the old 7-slide deck): if the DOCX has embedded images,
    build ONE content slide per section (heading/subheading) with that section's
    embedded screenshots attached — plus a greeting title slide and a thank-you
    slide. This mirrors build_ppt_from_sections so EVERY section/screenshot is
    covered. FALLBACK: when the DOCX has no embedded images, use the LLM text flow.

    Body: { filename, scheme_name?, ollama_url?, model?, num_slides?, language?,
            glossary?: [..], theme?, logo_id?, images?: [studio image ids] }
    Returns: { status, slides: [...] } using the shared slide data model.
    """
    data = await request.json()
    filename = data.get("filename")
    scheme_name = (data.get("scheme_name") or "").strip()
    ollama_url = data.get("ollama_url") or DEFAULT_OLLAMA_URL
    model = data.get("model") or DEFAULT_MODEL
    language = data.get("language") or DEFAULT_LANGUAGE
    glossary = data.get("glossary") or []
    num_slides = int(data.get("num_slides") or 0)
    images = data.get("images") or []
    use_llm = bool(data.get("use_llm", False))  # force LLM text flow if requested

    if not filename:
        return JSONResponse(status_code=400, content={"status": "error", "message": "filename is required"})
    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.isfile(filepath):
        return JSONResponse(status_code=404, content={"status": "error", "message": "Uploaded DOCX not found"})

    title = scheme_name or os.path.splitext(filename)[0]

    # ── PRIMARY: section + embedded-image flow (full coverage, every screenshot) ──
    if not use_llm:
        try:
            doc_title, sections, img_tmp = extract_docx_sections_with_images(filepath)
        except Exception as e:
            logger.warning(f"[studio] section extraction failed: {e}")
            sections, img_tmp, doc_title = [], None, title
        try:
            has_embedded = any(s.get("images") for s in sections)
            if sections and has_embedded:
                try:
                    model_slides = _studio_sections_to_model(
                        doc_title or title, sections, language, glossary,
                        ollama_url, model)
                except LanguageTranslationError as e:
                    logger.warning(f"[studio] translation rejected: {e}")
                    return JSONResponse(status_code=502, content={
                        "status": "error",
                        "message": str(e),
                    })
                return {"status": "ok", "title": doc_title or title,
                        "slides": model_slides, "mode": "sections"}
        finally:
            if img_tmp:
                shutil.rmtree(img_tmp, ignore_errors=True)

    # ── FALLBACK: LLM text flow (no embedded images, or explicitly requested) ────
    raw_content = (extract_docx_section_by_title(filepath, title)
                   if scheme_name else extract_docx_text(filepath))
    if not raw_content or len(raw_content) < 20:
        raw_content = extract_docx_text(filepath)
    if not raw_content or len(raw_content) < 20:
        return JSONResponse(status_code=400, content={"status": "error", "message": "Could not extract content from DOCX"})

    sys_prompt = _studio_language_prompt(DEFAULT_SYSTEM_PROMPT, language, glossary)
    # generate_slides also appends its own language directive for english/hindi; the
    # extra glossary/other-language block above is harmless and additive.
    result = generate_slides(title, raw_content, sys_prompt, ollama_url, model, language=language)
    if not result or "slides" not in result:
        return JSONResponse(status_code=502, content={"status": "error", "message": "Ollama failed to generate slides. Check the Ollama URL/model."})

    slides = [s for s in result["slides"]
              if s.get("type") in ("title", "thank_you", "section_title")
              or (s.get("bullets") and len(s["bullets"]) > 0)]
    slides = _studio_apply_glossary(slides, language, glossary)

    # Attach provided images round-robin to content slides (user can re-assign in UI).
    content_idx = [i for i, s in enumerate(slides) if s.get("type") == "content"]
    for n, ci in enumerate(content_idx):
        if images:
            slides[ci].setdefault("image", images[n % len(images)])

    # Soft-trim to requested number of CONTENT slides (keep title + thank_you).
    if num_slides and len(content_idx) > num_slides:
        keep = set(content_idx[:num_slides])
        slides = [s for i, s in enumerate(slides)
                  if s.get("type") != "content" or i in keep]

    return {"status": "ok", "title": title, "slides": slides}


# ─── Studio: export the edited slide model → PPTX ──────────────────────────────
@app.post("/api/studio/export")
async def studio_export(request: Request):
    """Export the CURRENT edited slide model to PPTX. Does not re-run the LLM; it
    renders exactly what the user sees in the browser preview, preserving template,
    images, logo, language text and ordering."""
    data = await request.json()
    slides = data.get("slides") or []
    theme = data.get("theme") or "classic_yellow"
    logo_id = data.get("logo_id")
    title = (data.get("title") or "presentation").strip()

    if not slides:
        return JSONResponse(status_code=400, content={"status": "error", "message": "No slides to export"})

    try:
        prs = _studio_build_pptx(slides, theme, logo_id)
    except Exception as e:
        logger.exception("[studio] export failed")
        return JSONResponse(status_code=500, content={"status": "error", "message": f"Export failed: {e}"})

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    safe = re.sub(r'[<>:"/\\|?*]', '', title)[:80].strip()
    safe = safe.encode("ascii", "ignore").decode("ascii").strip() or "presentation"
    fname = f"{safe}.pptx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{fname}"',
                 "X-Slide-Count": str(len(slides))},
    )


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
