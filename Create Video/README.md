# Create Video

Generates a narrated MP4 video from a PowerPoint (.pptx) file. Each slide becomes a video segment with text-to-speech narration derived from speaker notes.

## How It Works

1. **Extracts slides** — Reads embedded images and speaker notes from each slide.
2. **External screenshots** — If a speaker note line contains a file path to a `.png`/`.jpg`, that image is used instead of the embedded one.
3. **TTS narration** — Generates speech audio from the speaker notes using [edge-tts](https://pypi.org/project/edge-tts/) (Microsoft Edge neural voices).
4. **Assembles video** — Combines images and audio into a 1920×1080 MP4 using [moviepy](https://pypi.org/project/moviepy/).

## Prerequisites

- Python 3.9+
- [FFmpeg](https://ffmpeg.org/download.html) installed and on your PATH

## Installation

```bash
pip install python-pptx Pillow edge-tts moviepy
```

## Configuration

When you launch the script, a GUI window appears where you can configure:

| Field            | Description                              |
|------------------|------------------------------------------|
| PowerPoint file  | Path to the source `.pptx` file          |
| Build directory  | Working directory for intermediate files |
| Output video     | Path for the final `.mp4` output         |
| TTS Voice        | edge-tts voice name (default: `en-US-GuyNeural`) |

Each field is pre-filled with a default value. Use the **Browse** / **Save As** buttons or type a path directly.

## Usage

```bash
python create_video.py
```

Or use the included batch file:

```bash
run.bat
```

The GUI provides a **Create Video** button, a progress bar, and a real-time log panel showing each step.

## Speaker Notes Format

- Lines matching a local image path (e.g. `C:\Screenshots\step1.png`) are used as the slide image.
- All other lines become the narration text for that slide.
- Slides with no narration display for 3 seconds.

## Output

The script produces an MP4 video at 1920×1080, 24 fps, with H.264 video and AAC audio.
