# PPT Generation & PPTX to Video

A FastAPI-based tool that converts DOCX/PDF documents into PowerPoint presentations using an Ollama LLM, and converts PPTX files into MP4 videos with AI-generated voiceover.

---

## Project Structure

```
ppt_generation/
├── pptx_to_video/
│   ├── app.py          ← PPTX to Video app (run this for video generation)
│   └── templates/
│       └── index.html
├── templates/
│   └── index.html
├── uploads/            ← uploaded files go here (auto-created)
├── app.py              ← PPT Generator app (DOCX/PDF → PPTX)
├── requirements.txt
└── .gitignore
```

---

## Requirements

- Python 3.11+
- Microsoft PowerPoint installed (required for video export via COM)
- [Ollama](https://ollama.com/) server running with a model (e.g. `deepseek-r1:14b`)
- ffmpeg (for video stitching)
- A TTS Gradio server running

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Running the Apps

### 1. PPTX to Video — `pptx_to_video/app.py`

> Use this app to convert an existing `.pptx` file (with voiceover in slide notes) into an MP4 video.

```bash
cd pptx_to_video
python app.py
```

Open: [http://localhost:8100](http://localhost:8100)

**What it does:**
- Upload a `.pptx` file
- Reads voiceover text from each slide's notes
- Generates TTS audio per slide (in parallel)
- Uses PowerPoint COM automation to export animated video
- Stitches slides + audio into a final `.mp4`

**Requirements for video export:**
- Microsoft PowerPoint must be installed
- No other PowerPoint windows should be open when running
- TTS server must be reachable at the configured `TTS_GRADIO_URL`

---

### 2. PPT Generator — `app.py`

> Use this app to generate a `.pptx` from a DOCX or PDF document using an LLM.

```bash
python app.py
```

Open: [http://localhost:8000](http://localhost:8000)

**What it does:**
- Upload a DOCX or PDF file
- Select a scheme/topic from the document
- Sends content to Ollama LLM to generate structured slide JSON
- Builds and downloads a `.pptx` file with voiceover in slide notes

---

## Configuration

Edit the following defaults at the top of each `app.py`:

| Variable | File | Description |
|---|---|---|
| `TTS_GRADIO_URL` | `pptx_to_video/app.py` | URL of the TTS Gradio server |
| `DEFAULT_OLLAMA_URL` | `app.py` | URL of the Ollama server |
| `DEFAULT_MODEL` | `app.py` | Ollama model name |

---

## Common Errors

**`COMError: Call was rejected by callee`**
PowerPoint was busy or not ready. Fix: close all PowerPoint windows and retry.

**`PdfReadError: EOF marker not found`**
The uploaded PDF is corrupted. Use a valid PDF file.

**`ModuleNotFoundError`**
Run `pip install -r requirements.txt` again.
