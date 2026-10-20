"""
pptx_to_video/app.py
--------------------
Standalone FastAPI app — PPTX → Video only.

- Accepts a local PPTX file path
- Reads voiceover from slide notes
- Generates TTS audio per slide
- Stitches into MP4 via ffmpeg

Run:
    python app.py
Open: http://localhost:8100
"""
import os
import re
import io
import shutil
import tempfile
import time
import logging
import asyncio
import sys

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import comtypes
import comtypes.client
from pptx import Presentation

# ─── Setup ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("pptx_to_video")

sys.stdout.reconfigure(encoding="utf-8")

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

app = FastAPI(title="PPTX to Video", version="1.0")
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(__file__), "templates"))

TTS_API_URL = "https://app.migotoai.com/llma/speech/tts"
SLIDE_DURATION = 60      # fallback seconds per slide if no TTS duration available
ANIMATION_BUFFER = 3.0   # extra seconds added per slide so animations play at natural speed

# Lead-in before narration: each slide is shown first, THEN its voiceover starts
# after this delay (seconds). The slide always appears before the voice, and the
# gap never exceeds this value. Applied identically to every slide.
SLIDE_LEAD_IN = 1.0

# ─── Voice identity (single source of truth) ─────────────────────────────────
# Azure TTS voice parameters — change here only, applied to every slide.
VOICE_PROFILE = {
    "voice_id": "hi-IN-AaravNeural",   # Azure Neural voice
    "style":    "",                      # speaking style (leave blank for default)
    "rate":     "0%",                    # speaking rate
    "pitch":    "0Hz",                   # pitch adjustment
}

# Target integrated loudness (LUFS) applied to every slide's audio so the
# perceived volume is uniform across the whole presentation.
AUDIO_LOUDNESS_LUFS = -16.0
AUDIO_LOUDNESS_TP   = -1.5    # true-peak ceiling (dBTP)
AUDIO_LOUDNESS_LRA  = 11.0    # loudness range

# ─── Entrance animation settings ─────────────────────────────────────────────
# Applied via PowerPoint COM to shapes/paragraphs on each slide before export,
# so python-pptx-generated decks (which have no animations) still animate in the video.
ADD_ANIMATIONS = True    # set False to disable auto-applied entrance animations

# PowerPoint MsoAnimEffect / trigger / level enum values (COM constants)
_MSO_ANIM_EFFECT_FADE   = 10    # msoAnimEffectFade
_MSO_ANIM_EFFECT_FLY_IN = 2     # msoAnimEffectFly
_MSO_ANIM_TRIGGER_WITH_PREVIOUS  = 2  # msoAnimTriggerWithPrevious
_MSO_ANIM_TRIGGER_AFTER_PREVIOUS = 3  # msoAnimTriggerAfterPrevious
# Level controls how text within a placeholder is grouped into build steps.
_MSO_ANIM_LEVEL_NONE          = 0     # msoAnimateLevelNone (animate shape as one unit)
_MSO_ANIM_TEXT_BY_FIRST_LEVEL = 1     # msoAnimateTextByFirstLevel (one step per top-level bullet)

ANIM_EFFECT = _MSO_ANIM_EFFECT_FADE   # default entrance effect for body text
ANIM_STEP_DELAY = 0.4                 # seconds between successive animated items


# ─── Pronunciation Dictionary ────────────────────────────────────────────────
# Maps Hindi/Indian terms to Devanagari so TTS pronounces them natively.
PRONUNCIATION_MAP = {
    # Government scheme full names (longer phrases first)
    "Pradhan Mantri Mudra Yojna": "प्रधान मंत्री मुद्रा योजना",
    "Pradhan Mantri Mudra Yojana": "प्रधान मंत्री मुद्रा योजना",
    "Pradhan Mantri Kisan Samman Nidhi": "प्रधान मंत्री किसान सम्मान निधि",
    "Pradhan Mantri Fasal Bima Yojana": "प्रधान मंत्री फसल बीमा योजना",
    "Pradhan Mantri Suraksha Bima Yojana": "प्रधान मंत्री सुरक्षा बीमा योजना",
    "Pradhan Mantri Jeevan Jyoti Bima Yojana": "प्रधान मंत्री जीवन ज्योति बीमा योजना",
    "Pradhan Mantri Jan Dhan Yojana": "प्रधान मंत्री जन धन योजना",
    "Pradhan Mantri Awas Yojana": "प्रधान मंत्री आवास योजना",
    "Pradhan Mantri Ujjwala Yojana": "प्रधान मंत्री उज्ज्वला योजना",
    "Ayushman Bharat": "आयुष्मान भारत",
    "Atmanirbhar Bharat": "आत्मनिर्भर भारत",
    "Stree Shakti": "स्त्री शक्ति",
    "Sukanya Samriddhi": "सुकन्या समृद्धि",
    "Samriddhi": "समृद्धि",
    "Jan Dhan": "जन धन",
    "Udyog Aadhaar": "उद्योग आधार",
    # Individual Hindi words
    "Pradhan Mantri": "प्रधान मंत्री",
    "Mantri": "मंत्री",
    "Pradhan": "प्रधान",
    "Mudra": "मुद्रा",
    "Yojna": "योजना",
    "Yojana": "योजना",
    "Kisan": "किसान",
    "Samman": "सम्मान",
    "Nidhi": "निधि",
    "Fasal": "फसल",
    "Bima": "बीमा",
    "Suraksha": "सुरक्षा",
    "Jeevan": "जीवन",
    "Jyoti": "ज्योति",
    "Swamitva": "स्वामित्व",
    "Awas": "आवास",
    "Ujjwala": "उज्ज्वला",
    "Ayushman": "आयुष्मान",
    "Bharat": "भारत",
    "Atmanirbhar": "आत्मनिर्भर",
    "Rozgar": "रोजगार",
    "Udyam": "उद्यम",
    "Udyog": "उद्योग",
    "Aadhaar": "आधार",
    "Shishu": "शिशु",
    "Kishor": "किशोर",
    "Kishore": "किशोर",
    "Tarun": "तरुण",
    "Shakti": "शक्ति",
    "Mahila": "महिला",
    "Swavalamban": "स्वावलंबन",
    "Samriddhi": "समृद्धि",
    "Sampatti": "सम्पत्ति",
    "sampatti": "सम्पत्ति",
    "Sampathi": "सम्पत्ति",
    "sampathi": "सम्पत्ति",
    "Sukanya": "सुकन्या",
    "Vikas": "विकास",
    "Swarojgar": "स्वरोजगार",
    "Unnati": "उन्नति",
    "Sahay": "सहाय",
    "Tatkal": "तत्काल",
    "Seva": "सेवा",
    "Dhan": "धन",
    "Sampark": "सम्पर्क",
    "Sahkar": "सहकार",
    "Laghu": "लघु",
    "Udyami": "उद्यमी",
    "Rin": "ऋण",
    "Karj": "कर्ज",
    "Buniyad": "बुनियाद",
    "Samridhi": "समृद्धि",
    "Swabhiman": "स्वाभिमान",
    "Samagra": "समग्र",
    "Rashtriya": "राष्ट्रीय",
    "Grameen": "ग्रामीण",
    "Kshetra": "क्षेत्र",
    "Yuvak": "युवक",
    "Swasthya": "स्वास्थ्य",
    "Nagar": "नगर",
    "Panchayat": "पंचायत",
    "Mandal": "मंडल",
    "Kendriya": "केंद्रीय",
    "Sarkar": "सरकार",
    "Sarkari": "सरकारी",
    # Banking abbreviations
    # English words with pronunciation issues
    "premium": "preemeeam",
    "restructuring": "reestrakchuring",
    "Restructuring": "Reestrakchuring",
    "guarantees": "garantees",
    "Guarantees": "Garantees",
    "Guarantee": "Garantee",
    "guarantee": "garantee",
    "ensuring": "inshuring",
    "approved": "approovd",
    "loans": "lones",
    "Loans": "Lones",
    "loan": "lone",
    "medical": "medikal",
    "insurance": "inshurance",
    "remittances": "remittanses",
    "remit": "ree-mit",
    "principal": "prin-si-pal",
    "expiry": "ex-pie-ree",
    "considering": "cunsiddering",
    "40,000": "forty thousand",
    "75,000": "seventy five thousand",
    "-1": "minus one",
    "Griha Vatika": "Gree-ha  Vatikha",
    "Matritva": "Maathritva",
    "elderly": "el-der-lee",
    "saraswati": "sara-swa-tee",
    "Udaan": "oo-dahn",
    "UDAAN": "oo-dahn",
    "15,000":"fifteen thousand",
    "20,000": "twenty thousand",
    "25,000": "twenty five thousand",
    "50,000": "fifty thousand",
    "75,000": "seventy five thousand",
    "3,00,000": "three lakh",
    "Required" : "reck-wired",
    "VAHAN": "vah-han",
    "25,000 thousand": "twenty five thousand",
    "EDULOANM": "EDU-LOAN-menu",
    "Sahyog": "sha-yog",
    # Banking abbreviations
    "PNB": "P N B",
    "PMMY": "P M M Y",
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
    "CIBIL": "C I B I L",
    "DIGI MSME": "Digi M S M E",
    "MSMEs": "M S M E s",
    "NRLM": "N R L M",
    "PMEGP": "P M E G P",
    "KVIC": "K V I C",
    "DIC": "D I C",
    "NABARD": "N A B A R D",
    "SIDBI": "S I D B I",
    "RBI": "R B I",
    "GRIHA": "G R I H A",
    "HUF": "H U F",
    "SHLOT": "S H L O T",
    "ZOCAC": "Z O C A C",
    "CHCAC": "C H C A C",
    "CA": "C A",
    "ICWA": "I C W A",
    "CFA": "C F A",
    "IIMs": "I I M",
    "IITs": "I I T",
    "IISc": "I I S c",
    "NIFT": "N I F T",
    "NID": "N I D",
    "BRE": "B R E",
    "5.00 Cr": "five crore",
    "1.00 Cr": "one crore",
    "5.00": "five",
    "1.00": "one",
    "1.30:1": "one point three zero is to 1",
    "1.20:1": "one point two zero is to 1",
    "1cr": "one crore",
    "A-1": "A one",
    "e-BG": "e B G",
    "UIN": "U I N",
    "Rs.": "Rupees",
    "Portfolio": "Port-folio", 
    "i.e.": "that is",
    "IB": "i b",
    "PRIME": "prime",
    "X" : "tenth",
    "VIII" : "eighth",
    "IX" : "ninth",
    "XII" : "twelfth",
    "XI" : "eleventh",
    "LIC" : "L I C",
    "PAN" : "pan",
    "Voters": "voters",
    "EKAA": "E K A A",
    
}


def _indian_amount_to_words(amount_str: str) -> str:
    """Convert Indian-format number string (e.g. '2,99,99,999') to English words."""
    n = int(amount_str.replace(",", ""))
    ones = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
            "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
            "seventeen", "eighteen", "nineteen"]
    tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]

    def two_digits(n):
        if n < 20: return ones[n]
        return (tens[n // 10] + (" " + ones[n % 10] if n % 10 else "")).strip()

    def three_digits(n):
        if n >= 100:
            return ones[n // 100] + " hundred" + (" " + two_digits(n % 100) if n % 100 else "")
        return two_digits(n)

    parts = []
    crore    = n // 10_000_000; n %= 10_000_000
    lakh     = n // 100_000;    n %= 100_000
    thousand = n // 1000;       n %= 1000
    if crore:    parts.append(two_digits(crore) + " crore")
    if lakh:     parts.append(two_digits(lakh) + " lakh")
    if thousand: parts.append(two_digits(thousand) + " thousand")
    if n:        parts.append(three_digits(n))
    return "Rupees " + " ".join(parts) + " only"


def _apply_pronunciation_fixes(text: str) -> str:
    """Replace Hindi/Indian terms with phonetic equivalents for better TTS pronunciation."""
    # Handle Rs. amounts FIRST before any other substitution
    text = re.sub(r'Rs\.?\s*([\d,]+)(?:/-\.?)?', lambda m: _indian_amount_to_words(m.group(1)), text)
    for term, phonetic in PRONUNCIATION_MAP.items():
        pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
        text = pattern.sub(phonetic, text)
    return text


# ─── Strip markdown / formatting symbols before TTS ──────────────────────────
def _strip_formatting_symbols(text: str) -> str:
    """Remove markdown and formatting symbols so the TTS never reads them aloud.

    Handles the common markdown/formatting characters (asterisks, hashtags,
    underscores, backticks, tildes, etc.) while PRESERVING punctuation and
    characters that later cleaning steps rely on for natural speech:
      - decimals / percentages (handled by expand_decimal)
      - "/" between words or digits (handled → " or ")
      - "Rs." amounts (handled → words)
    So this runs FIRST, touching only pure formatting noise.
    """
    # Markdown emphasis / headings / code / lists — remove the symbols, keep text.
    #   **bold**, *italic*, __under__, _under_, `code`, ~~strike~~, ## heading
    text = re.sub(r'[*_~`#]+', ' ', text)        # asterisks, underscores, tildes, backticks, hashtags
    text = re.sub(r'^\s*[-+•]\s+', '', text, flags=re.MULTILINE)   # bullet markers at line start
    text = re.sub(r'>{1,}', ' ', text)           # blockquote markers
    text = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', text)  # [label](url) → label
    text = re.sub(r'[\[\]]', ' ', text)          # leftover square brackets
    # Collapse the whitespace introduced by removals.
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text


# ─── Clean notes text before TTS ─────────────────────────────────────────────
def _clean_tts_text(text: str) -> str:
    # Remove markdown/formatting symbols FIRST so they are never spoken and so
    # they cannot interfere with the pronunciation / number rules below.
    text = _strip_formatting_symbols(text)

    # Strip label prefixes like "Voiceover:", "Narration:", "Voiceover/Narration:"
    text = re.sub(r'^(voiceover|narration|voiceover\s*/\s*narration|vo)\s*:\s*', '', text, flags=re.IGNORECASE)

    # Expand decimals with optional % before touching slashes
    def expand_decimal(m):
        integer_part = m.group(1)
        decimal_part = m.group(2)
        suffix       = m.group(3)
        integer_word = "zero" if integer_part == "0" else integer_part
        digit_map    = ["zero","one","two","three","four","five","six","seven","eight","nine"]
        digit_words  = " ".join(digit_map[int(d)] for d in decimal_part)
        result = f"{integer_word} point {digit_words}"
        if suffix == "%":
            result += " percent"
        return result

    # Apply pronunciation fixes FIRST (handles Rs. amounts before decimal expander runs)
    text = _apply_pronunciation_fixes(text)

    text = re.sub(r'(?<![\d])(\d+)\.(\d+)(%?)', expand_decimal, text)

    # Replace slash between non-digits with " or "
    text = re.sub(r'(?<![\d])/(?![\d])', ' or ', text)
    # Replace slash between digits with " or "
    text = re.sub(r'(\d)/(\d)', lambda m: f"{m.group(1)} or {m.group(2)}", text)

    return text.strip()


# ─── PPTX → animated MP4 via PowerPoint COM ─────────────────────────────────
COM_CALL_REJECTED = -2147418111

def _is_com_call_rejected(exc: BaseException) -> bool:
    if isinstance(exc, OSError):
        return exc.winerror == COM_CALL_REJECTED or getattr(exc, "hresult", None) == COM_CALL_REJECTED
    return getattr(exc, "hresult", None) == COM_CALL_REJECTED or (isinstance(exc.args, tuple) and exc.args and exc.args[0] == COM_CALL_REJECTED)


def _retry_com_call(func, max_attempts: int = 20, initial_delay: float = 5.0):
    delay = initial_delay
    for attempt in range(1, max_attempts + 1):
        try:
            return func()
        except Exception as exc:
            if _is_com_call_rejected(exc) and attempt < max_attempts:
                logger.warning("PowerPoint COM call rejected, retrying %d/%d after %.1fs", attempt, max_attempts, delay)
                time.sleep(delay)
                delay = min(delay * 1.5, 30.0)
                continue
            # Final failure: capture diagnostic info to help debugging
            try:
                import subprocess
                proc = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE"], capture_output=True, text=True)
                logger.error("PowerPoint processes:\n%s", proc.stdout)
            except Exception:
                logger.exception("Failed to list processes for diagnostic purposes")
            raise


def _apply_entrance_animations(slide) -> None:
    """Add staggered entrance animations to the shapes on a single slide via COM.

    - The first text shape (usually the title) fades in on its own.
    - Each subsequent body/content placeholder animates in paragraph-by-paragraph
      (one bullet at a time) using the Level parameter, timed After Previous.

    Any per-shape failure is swallowed so a problematic shape never aborts the
    whole export; the slide simply keeps fewer animations.

    PowerPoint COM reference:
        Sequence.AddEffect(Shape, effectId, Level, trigger, Index) -> Effect
        Level = msoAnimateLevelNone (0) animates the shape as one unit.
        Level = msoAnimateTextByFirstLevel (1) reveals each 1st-level paragraph.
    """
    try:
        sequence = slide.TimeLine.MainSequence
    except Exception:
        logger.warning("Slide has no animation timeline; skipping animations")
        return

    first_effect = True
    for shape in slide.Shapes:
        try:
            has_text = bool(getattr(shape, "HasTextFrame", False)) and shape.TextFrame.HasText
        except Exception:
            has_text = False

        if not has_text:
            continue

        try:
            is_title = first_effect
            # Title: one fade-in unit. Body: reveal per top-level paragraph.
            level   = _MSO_ANIM_LEVEL_NONE if is_title else _MSO_ANIM_TEXT_BY_FIRST_LEVEL
            effect_id = _MSO_ANIM_EFFECT_FADE if is_title else ANIM_EFFECT
            trigger = (_MSO_ANIM_TRIGGER_WITH_PREVIOUS
                       if first_effect else _MSO_ANIM_TRIGGER_AFTER_PREVIOUS)

            effect = sequence.AddEffect(
                Shape=shape,
                effectId=effect_id,
                Level=level,
                trigger=trigger,
            )

            # Natural pacing: short duration + a small stagger between items.
            try:
                effect.Timing.Duration = 0.5
                effect.Timing.TriggerDelayTime = 0.0 if first_effect else ANIM_STEP_DELAY
            except Exception:
                pass

            first_effect = False
        except Exception:
            logger.exception("Failed to add entrance animation to a shape; continuing")
            continue


def _pptx_to_video_impl(pptx_path: str, output_mp4: str, slide_durations: list[float] | None = None) -> str:
    # Kill any stuck PowerPoint processes from previous runs
    import subprocess as _sp
    _sp.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    time.sleep(5)

    powerpoint = None
    presentation = None
    try:
        powerpoint = _retry_com_call(lambda: comtypes.client.CreateObject("PowerPoint.Application"))
        powerpoint.Visible = 1
        powerpoint.DisplayAlerts = False
        time.sleep(1)

        # Open editable (ReadOnly=False) so entrance animations can be injected
        # before export. Still open without a visible window.
        presentation = _retry_com_call(lambda: powerpoint.Presentations.Open(pptx_path, ReadOnly=False, WithWindow=False))
        time.sleep(8)

        slides = _retry_com_call(lambda: list(enumerate(presentation.Slides)))
        for idx, slide in slides:
            base = slide_durations[idx] if slide_durations and idx < len(slide_durations) else SLIDE_DURATION
            slide.SlideShowTransition.AdvanceOnTime = True
            # Match the final segment length: lead-in (slide shown before voice)
            # + spoken audio + animation buffer. Keeps the exported slide timings
            # in lockstep with the per-slide audio built in build_video().
            slide.SlideShowTransition.AdvanceTime = SLIDE_LEAD_IN + base + ANIMATION_BUFFER
            if ADD_ANIMATIONS:
                _retry_com_call(lambda s=slide: _apply_entrance_animations(s))

        _retry_com_call(lambda: presentation.CreateVideo(
            output_mp4,
            UseTimingsAndNarrations=True,
            DefaultSlideDuration=SLIDE_DURATION,
            VertResolution=1080,
            FramesPerSecond=30,
            Quality=100,
        ))

        # Wait until file is stable (export complete)
        for _ in range(600):
            time.sleep(1)
            if os.path.exists(output_mp4) and os.path.getsize(output_mp4) > 0:
                size_before = os.path.getsize(output_mp4)
                time.sleep(2)
                if os.path.getsize(output_mp4) == size_before:
                    break

        logger.info(f"PowerPoint video export done: {output_mp4}")
        return output_mp4
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass


def pptx_to_video_file(pptx_path: str, output_mp4: str, slide_durations: list[float] | None = None) -> str:
    """Run PowerPoint COM export inside a dedicated STA thread.

    Some Office COM servers require a Single-Threaded Apartment (STA). This
    helper spawns a worker thread, initializes COM as STA (using pythoncom if
    available, otherwise comtypes), runs the implementation and then
    uninitializes COM before returning.
    """
    import threading

    # Prefer pythoncom if available because it exposes CoInitializeEx constants
    try:
        import pythoncom
        have_pythoncom = True
    except Exception:
        pythoncom = None
        have_pythoncom = False

    result = {}

    def _worker():
        try:
            if have_pythoncom:
                pythoncom.CoInitialize()
            else:
                comtypes.CoInitialize()
            result['value'] = _pptx_to_video_impl(pptx_path, output_mp4, slide_durations)
        except Exception as e:
            result['exc'] = e
        finally:
            try:
                if have_pythoncom:
                    pythoncom.CoUninitialize()
                else:
                    comtypes.CoUninitialize()
            except Exception:
                pass

    th = threading.Thread(target=_worker, daemon=False)
    th.start()
    th.join()

    if 'exc' in result:
        raise result['exc']
    return result.get('value')


# ─── TTS audio ───────────────────────────────────────────────────────────────
from concurrent.futures import ThreadPoolExecutor, as_completed


def generate_tts_audio(text: str, output_path: str) -> str:
    import httpx
    start = time.time()
    with httpx.Client(timeout=120) as client:
        r = client.post(
            TTS_API_URL,
            headers={
                "accept": "*/*",
                "content-type": "application/x-www-form-urlencoded",
                "origin": "https://app.migotoai.com",
                "referer": "https://app.migotoai.com/llma/docs",
            },
            data={
                "message":  text,
                "voice_id": VOICE_PROFILE["voice_id"],
                "style":    VOICE_PROFILE["style"],
                "rate":     VOICE_PROFILE["rate"],
                "pitch":    VOICE_PROFILE["pitch"],
            },
        )
        r.raise_for_status()
        with open(output_path, "wb") as f:
            f.write(r.content)
    logger.info(f"TTS took {time.time() - start:.2f}s → {os.path.basename(output_path)}")
    return output_path


def generate_all_tts_parallel(texts: list[str], audio_paths: list[str], max_workers: int = 4) -> list[str]:
    """Generate TTS for all slides in parallel using a thread pool."""
    results = [None] * len(texts)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(generate_tts_audio, text, path): idx
            for idx, (text, path) in enumerate(zip(texts, audio_paths))
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            results[idx] = future.result()
    return results


# ─── Audio duration ───────────────────────────────────────────────────────────
def get_audio_duration(audio_path: str) -> float:
    import subprocess
    import wave
    # Try wave module first (works for .wav files)
    try:
        with wave.open(audio_path, 'rb') as wf:
            frames = wf.getnframes()
            rate = wf.getframerate()
            return frames / float(rate)
    except Exception:
        pass
    # Fallback: use ffmpeg to get duration
    from imageio_ffmpeg import get_ffmpeg_exe
    ffmpeg = get_ffmpeg_exe()
    result = subprocess.run(
        [ffmpeg, "-i", audio_path, "-f", "null", "-"],
        capture_output=True, text=True,
    )
    # Parse duration from stderr output
    import re as _re
    match = _re.search(r'Duration:\s*(\d+):(\d+):(\d+\.\d+)', result.stderr)
    if match:
        h, m, s = float(match.group(1)), float(match.group(2)), float(match.group(3))
        return h * 3600 + m * 60 + s
    raise ValueError(f"Could not get duration for: {audio_path}")


# ─── Build final MP4: slice animated video per slide + replace audio ───────────
def build_video(pptx_video: str, audio_files: list[str], slide_durations: list[float], output_path: str):
    import subprocess
    from imageio_ffmpeg import get_ffmpeg_exe
    ffmpeg  = get_ffmpeg_exe()
    tmp_dir = os.path.dirname(output_path)

    segment_files = []
    start_time = 0.0
    for i, audio in enumerate(audio_files):
        tts_duration  = slide_durations[i]
        # Slide is shown first; narration starts SLIDE_LEAD_IN seconds later.
        # Total segment length = lead-in silence + spoken audio + trailing
        # animation buffer, so the full voiceover plays and nothing is clipped.
        total_duration = SLIDE_LEAD_IN + tts_duration + ANIMATION_BUFFER
        segment_path  = os.path.join(tmp_dir, f"seg_{i:03d}.mp4")
        # Audio pipeline per slide (identical for every slide):
        #   1. loudnorm  → uniform perceived volume across the whole video
        #   2. adelay    → prepend SLIDE_LEAD_IN of silence so the voiceover
        #                  starts AFTER the slide appears (never before)
        #   3. apad      → append silence so video/animations finish cleanly
        padded_audio = os.path.join(tmp_dir, f"padded_{i:03d}.wav")
        lead_in_ms = int(round(SLIDE_LEAD_IN * 1000))
        loudnorm_filter = (
            f"loudnorm=I={AUDIO_LOUDNESS_LUFS}:TP={AUDIO_LOUDNESS_TP}:LRA={AUDIO_LOUDNESS_LRA},"
            f"adelay={lead_in_ms}:all=1,"
            f"apad=pad_dur={ANIMATION_BUFFER}"
        )
        subprocess.run(
            [ffmpeg, "-y", "-i", audio,
             "-af", loudnorm_filter,
             "-ar", "48000", "-ac", "2",
             "-t", str(total_duration), padded_audio],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        subprocess.run(
            [ffmpeg, "-y",
             "-ss", str(start_time), "-t", str(total_duration), "-i", pptx_video,
             "-i", padded_audio,
             "-map", "0:v", "-map", "1:a",
             "-c:v", "libx264", "-preset", "ultrafast", "-pix_fmt", "yuv420p",
             "-c:a", "aac", "-b:a", "192k", "-ar", "48000", "-ac", "2",
             "-shortest", segment_path],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        segment_files.append(segment_path)
        start_time += total_duration

    concat_file = os.path.join(tmp_dir, "concat.txt")
    with open(concat_file, "w") as f:
        for seg in segment_files:
            f.write(f"file '{seg}'\n")

    subprocess.run(
        [ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", concat_file, "-c", "copy", output_path],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )


# ─── Routes ──────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")


@app.post("/api/pptx-to-video")
async def pptx_to_video(file: UploadFile = File(...)):
    """
    Accepts any PowerPoint file upload (any extension).
    Reads voiceover from slide notes → TTS → MP4.
    """
    original_name = file.filename or "presentation"
    safe_name     = re.sub(r'[<>:"/\\|?*]', '', os.path.splitext(original_name)[0])[:80].strip()
    # Remove any non-ASCII characters from filename for HTTP header compatibility
    safe_name = safe_name.encode('ascii', 'ignore').decode('ascii').strip() or "presentation"

    tmp_dir = tempfile.mkdtemp(prefix="pptx_video_")
    try:
        # Save uploaded file — always use .pptx extension so python-pptx & PowerPoint COM work
        pptx_path = os.path.join(tmp_dir, "input.pptx")
        with open(pptx_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Read slide notes
        # Validate the file is a valid PPTX (ZIP with [Content_Types].xml)
        import zipfile
        if not zipfile.is_zipfile(pptx_path):
            return JSONResponse(status_code=400, content={"status": "error", "message": "Invalid file. Please upload a valid .pptx file."})

        prs = Presentation(pptx_path)
        slide_notes = []
        for slide in prs.slides:
            notes_frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
            text = notes_frame.text.strip() if notes_frame else ""
            slide_notes.append(text)

        total_slides = len(slide_notes)
        logger.info(f"Processing {total_slides} slides from: {original_name}")

        # Step 1: notes → TTS audio in PARALLEL (biggest speed gain)
        audio_dir = os.path.join(tmp_dir, "audio")
        os.makedirs(audio_dir, exist_ok=True)

        tts_texts = []
        audio_paths_list = []
        for i, vo in enumerate(slide_notes):
            vo = _clean_tts_text(vo) if vo.strip() else f"Slide {i + 1}"
            audio_path = os.path.join(audio_dir, f"slide_{i:03d}.wav")
            tts_texts.append(vo)
            audio_paths_list.append(audio_path)

        logger.info(f"Starting parallel TTS for {total_slides} slides...")
        tts_start = time.time()
        generate_all_tts_parallel(tts_texts, audio_paths_list, max_workers=4)
        logger.info(f"All TTS done in {time.time() - tts_start:.1f}s (parallel)")

        audio_files = audio_paths_list
        slide_durations = []
        for audio_path in audio_files:
            duration = get_audio_duration(audio_path)
            slide_durations.append(duration)
        logger.info(f"Slide durations: {[f'{d:.1f}s' for d in slide_durations]}")

        # Step 2: PPTX → animated MP4 with per-slide timings matching TTS durations
        anim_video = os.path.join(tmp_dir, "animated.mp4")
        pptx_to_video_file(
            os.path.abspath(pptx_path),
            os.path.abspath(anim_video),
            slide_durations,
        )

        # Step 3: slice animated video per slide + replace with TTS audio
        video_path = os.path.join(tmp_dir, "output.mp4")
        build_video(anim_video, audio_files, slide_durations, video_path)

        if not os.path.exists(video_path):
            return JSONResponse(status_code=500, content={"status": "error", "message": "Video generation failed"})

        with open(video_path, "rb") as f:
            video_bytes = f.read()

        return StreamingResponse(
            io.BytesIO(video_bytes),
            media_type="video/mp4",
            headers={"Content-Disposition": f'attachment; filename="{safe_name}.mp4"'},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ─── Run ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("  PPTX to Video — http://localhost:8100")
    print(f"  TTS: {TTS_API_URL}")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8100)