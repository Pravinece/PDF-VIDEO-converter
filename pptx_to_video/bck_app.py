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

TTS_GRADIO_URL = "https://f3ebfce6867a004011.gradio.live/"
SLIDE_DURATION = 10      # fallback seconds per slide if no TTS duration available
ANIMATION_BUFFER = 3.0   # extra seconds added per slide so animations play at natural speed


# ─── Clean notes text before TTS ─────────────────────────────────────────────
def _clean_tts_text(text: str) -> str:
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

    text = re.sub(r'(?<![\d])(\d+)\.(\d+)(%?)', expand_decimal, text)

    # Replace slash between non-digits with " or "
    text = re.sub(r'(?<![\d])/(?![\d])', ' or ', text)
    # Replace slash between digits with " or "
    text = re.sub(r'(\d)/(\d)', lambda m: f"{m.group(1)} or {m.group(2)}", text)

    return text.strip()


# ─── PPTX → animated MP4 via PowerPoint COM ─────────────────────────────────
def pptx_to_video_file(pptx_path: str, output_mp4: str, slide_durations: list[float] | None = None) -> str:
    comtypes.CoInitialize()
    powerpoint   = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    time.sleep(3)
    presentation = powerpoint.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
    time.sleep(2)

    for idx, slide in enumerate(presentation.Slides):
        base = slide_durations[idx] if slide_durations and idx < len(slide_durations) else SLIDE_DURATION
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceTime   = base + ANIMATION_BUFFER

    presentation.CreateVideo(output_mp4, UseTimingsAndNarrations=True, DefaultSlideDuration=SLIDE_DURATION, VertResolution=1080, FramesPerSecond=30, Quality=100)

    # Wait until file is stable (export complete)
    for _ in range(600):
        time.sleep(1)
        if os.path.exists(output_mp4) and os.path.getsize(output_mp4) > 0:
            size_before = os.path.getsize(output_mp4)
            time.sleep(2)
            if os.path.getsize(output_mp4) == size_before:
                break

    presentation.Close()
    powerpoint.Quit()
    comtypes.CoUninitialize()
    logger.info(f"PowerPoint video export done: {output_mp4}")
    return output_mp4


# ─── TTS audio ───────────────────────────────────────────────────────────────
def generate_tts_audio(text: str, output_path: str) -> str:
    import httpx
    from gradio_client import Client
    start = time.time()
    client = Client(TTS_GRADIO_URL)
    result = client.predict(
        text=text,
        lang="English",
        ns=32,
        gs=2,
        dn=True,
        sp=0.95,
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
    # If result is a remote URL, download it manually
    if isinstance(temp_audio_path, str) and temp_audio_path.startswith("http"):
        with httpx.Client(timeout=60) as hclient:
            r = hclient.get(temp_audio_path)
            r.raise_for_status()
            with open(output_path, "wb") as f:
                f.write(r.content)
    else:
        shutil.copy(temp_audio_path, output_path)
    logger.info(f"TTS took {time.time() - start:.2f}s")
    return output_path


# ─── Audio duration ───────────────────────────────────────────────────────────
def get_audio_duration(audio_path: str) -> float:
    from pydub import AudioSegment
    return len(AudioSegment.from_file(audio_path)) / 1000.0


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
        padded_audio = os.path.join(tmp_dir, f"padded_{i:03d}.mp3")
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

    tmp_dir = tempfile.mkdtemp(prefix="pptx_video_")
    try:
        # Save uploaded file — always use .pptx extension so python-pptx & PowerPoint COM work
        pptx_path = os.path.join(tmp_dir, "input.pptx")
        with open(pptx_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Read slide notes
        prs = Presentation(pptx_path)
        slide_notes = []
        for slide in prs.slides:
            notes_frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
            text = notes_frame.text.strip() if notes_frame else ""
            slide_notes.append(text)

        total_slides = len(slide_notes)
        logger.info(f"Processing {total_slides} slides from: {original_name}")

        # Step 1: notes → TTS audio (must happen first to get per-slide durations)
        audio_dir = os.path.join(tmp_dir, "audio")
        os.makedirs(audio_dir, exist_ok=True)
        audio_files = []
        slide_durations = []

        for i, vo in enumerate(slide_notes):
            vo = _clean_tts_text(vo) if vo.strip() else f"Slide {i + 1}"
            audio_path = os.path.join(audio_dir, f"slide_{i:03d}.mp3")
            generate_tts_audio(vo, audio_path)
            audio_files.append(audio_path)
            duration = get_audio_duration(audio_path)
            slide_durations.append(duration)
            logger.info(f"Audio: slide {i + 1}/{total_slides} — {duration:.2f}s")

        # Step 2: PPTX → animated MP4 with per-slide timings matching TTS durations
        anim_video = os.path.join(tmp_dir, "animated.mp4")
        pptx_to_video_file(os.path.abspath(pptx_path), os.path.abspath(anim_video), slide_durations)

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