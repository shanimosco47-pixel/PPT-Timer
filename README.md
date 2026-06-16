# PPT Timer

A PowerPoint countdown timer that auto-installs itself onto every slide.

## How to use

1. Save your presentation as `.pptm` (macro-enabled)
2. Open the VBA editor (`Alt+F11`)
3. Paste `ThisPresentation.cls` into the `ThisPresentation` module
4. Insert a new module (`Insert → Module`) and paste `Module1.bas`
5. Save and close the editor

**Or** — open the pre-built `.pptm` template and paste your slides into it.

## Recipient workflow

1. Open `.pptm` → click **Enable Content**
2. Paste your slides
3. Press `F5` → click `▶ Start` → enter minutes → timer runs on every slide

## What happens automatically

- Opening the file adds a `countdown` label and `▶ Start` button to every slide
- Starting the slideshow catches any slides added after the file was opened
- The timer updates all slides every second
- At zero the display flashes red, then stays red at `00:00`

## Files

- `outputs/vba/ThisPresentation.cls` — lifecycle hooks
- `outputs/vba/Module1.bas` — SetupShapes, StartCountdown, RunTimer
