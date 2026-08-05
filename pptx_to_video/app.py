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

TTS_GRADIO_URL = "http://172.23.198.242:8001/"
SLIDE_DURATION = 60      # fallback seconds per slide if no TTS duration available
ANIMATION_BUFFER = 2.0   # extra seconds added per slide so animations play at natural speed


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
    "CHCAC": "C H C A C"
}


def _apply_pronunciation_fixes(text: str) -> str:
    """Replace Hindi/Indian terms with phonetic equivalents for better TTS pronunciation."""
    for term, phonetic in PRONUNCIATION_MAP.items():
        pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
        text = pattern.sub(phonetic, text)
    return text


# ─── Clean notes text before TTS ─────────────────────────────────────────────
def _clean_tts_text(text: str) -> str:
    # Strip label prefixes like "Voiceover:", "Narration:", "Voiceover/Narration:"
    text = re.sub(r'^(voiceover|narration|voiceover\s*/\s*narration)\s*:\s*', '', text, flags=re.IGNORECASE)

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

    text = re.sub(r'(?<![\d])(\d+)\.(\d+)(%?)', expand_decimal, text)

    # Replace slash between non-digits with " or "
    text = re.sub(r'(?<![\d])/(?![\d])', ' or ', text)
    # Replace slash between digits with " or "
    text = re.sub(r'(\d)/(\d)', lambda m: f"{m.group(1)} or {m.group(2)}", text)

    # Apply pronunciation fixes for Hindi/Indian terms
    text = _apply_pronunciation_fixes(text)

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

        presentation = _retry_com_call(lambda: powerpoint.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False))
        time.sleep(8)

        slides = _retry_com_call(lambda: list(enumerate(presentation.Slides)))
        for idx, slide in slides:
            base = slide_durations[idx] if slide_durations and idx < len(slide_durations) else SLIDE_DURATION
            slide.SlideShowTransition.AdvanceOnTime = True
            slide.SlideShowTransition.AdvanceTime = base + ANIMATION_BUFFER

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

_tts_client = None

def _get_tts_client():
    """Reuse a single Gradio client connection."""
    global _tts_client
    if _tts_client is None:
        from gradio_client import Client
        _tts_client = Client(TTS_GRADIO_URL)
    return _tts_client


def generate_tts_audio(text: str, output_path: str) -> str:
    import httpx
    from gradio_client import Client
    start = time.time()
    # Each thread needs its own client for parallel calls
    client = Client(TTS_GRADIO_URL)
    result = client.predict(
        text=text,
        lang="English",
        ns=32,
        gs=2,
        dn=True,
        sp=0.9,
        du=0,
        pp=True,
        po=True,
        param_9="Female / 女",
        param_10="Middle-aged / 中年",
        param_11="Auto",
        param_12="Auto",
        param_13="Indian Accent / 印度口音",
        param_14="Auto",
        api_name="/_design_fn"
    )
    temp_audio_path = result[0] if isinstance(result, (tuple, list)) else result
    if isinstance(temp_audio_path, str) and temp_audio_path.startswith("http"):
        with httpx.Client(timeout=60) as hclient:
            r = hclient.get(temp_audio_path)
            r.raise_for_status()
            with open(output_path, "wb") as f:
                f.write(r.content)
    else:
        shutil.copy(temp_audio_path, output_path)
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
        total_duration = tts_duration + ANIMATION_BUFFER
        segment_path  = os.path.join(tmp_dir, f"seg_{i:03d}.mp4")
        # Use total_duration for video so animations play fully; audio is padded with silence to match
        padded_audio = os.path.join(tmp_dir, f"padded_{i:03d}.wav")
        subprocess.run(
            [ffmpeg, "-y", "-i", audio,
             "-af", f"apad=pad_dur={ANIMATION_BUFFER}",
             "-t", str(total_duration), padded_audio],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        subprocess.run(
            [ffmpeg, "-y",
             "-ss", str(start_time), "-t", str(total_duration), "-i", pptx_video,
             "-i", padded_audio,
             "-map", "0:v", "-map", "1:a",
             "-c:v", "libx264", "-preset", "ultrafast", "-pix_fmt", "yuv420p",
             "-c:a", "aac", "-shortest", segment_path],
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
    print(f"  TTS: {TTS_GRADIO_URL}")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8100)