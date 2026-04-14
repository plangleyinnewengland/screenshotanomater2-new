# Create Storyboard

A simple tool that converts a folder of PNG screenshots into a widescreen PowerPoint storyboard presentation.

## What It Does

- Prompts you to select a folder containing PNG screenshots
- Creates a 16:9 widescreen PowerPoint presentation (`storyboard.pptx`)
- Adds each PNG as a full-slide image (sorted alphabetically)
- Includes the source file path in the speaker notes of each slide
- Saves the output in the same folder as the screenshots

## Requirements

- Python 3.x
- `python-pptx` library

Install dependencies:

```
pip install python-pptx
```

## Usage

**Option 1:** Double-click `run.bat`

**Option 2:** Run from the command line:

```
python create_storyboard.py
```

A folder picker dialog will appear — select the folder with your screenshots, and the storyboard will be generated automatically.

## Output

The generated `storyboard.pptx` file is saved in the selected screenshots folder.
