# Create Video

Generates a narrated MP4 video from a PowerPoint (.pptx) file. Each slide becomes a video segment with text-to-speech narration derived from speaker notes. Can also export audio only.

## How It Works

1. **Extracts slides** — Reads embedded images and speaker notes from each slide.
2. **External screenshots** — If a speaker note line contains a file path to a `.png`/`.jpg`, that image is used instead of the embedded one.
3. **TTS narration** — Generates speech audio from the speaker notes using [edge-tts](https://pypi.org/project/edge-tts/) (Microsoft Edge neural voices).
4. **Assembles video** — Combines images and audio into a 1920×1080 MP4 using [moviepy](https://pypi.org/project/moviepy/).

## Prerequisites

- Python 3.9+
- [FFmpeg](https://ffmpeg.org/download.html) installed and on your PATH (or install `imageio-ffmpeg` as a fallback)

## Installation

```bash
pip install python-pptx Pillow edge-tts moviepy
```

Optionally, if FFmpeg is not on your PATH:

```bash
pip install imageio-ffmpeg
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

## Output Modes

The GUI provides three output modes:

| Mode                          | Description                                                        |
|-------------------------------|--------------------------------------------------------------------|
| **Create Video**              | Full MP4 video with images and narration (default)                 |
| **Export Audio (separate)**   | One MP3 file per slide, saved to the build directory               |
| **Export Audio (single track)** | All slide narrations combined into a single `combined-audio.mp3` |

## Usage

```bash
python create_video.py
```

Or use the included batch file:

```bash
run.bat
```

The GUI provides a **Run** button, a progress bar, and a real-time log panel showing each step.

## Speaker Notes Format

- Lines matching a local image path (e.g. `C:\Screenshots\step1.png`) are used as the slide image instead of the embedded one.
- All other lines become the narration text for that slide.
- Slides with no narration display for 3 seconds of silence.

## Output

- **Video mode** — MP4 at 1920×1080, 24 fps, H.264 video and AAC audio.
- **Audio modes** — MP3 files in the build directory.
