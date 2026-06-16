# VBA Source Files

Import or paste these source files into a macro-enabled PowerPoint deck (`.pptm`). Do not commit or upload the deck itself if it contains private presentation content.

## Files

| File | How to use it |
|------|---------------|
| `Module1.bas` | Import into a standard VBA module, or create a new module and paste the contents. |
| `ThisPresentation.cls` | Do **not** import this as a new class. Open the existing `ThisPresentation` object under **Microsoft PowerPoint Objects** and paste the event procedures there. |

## Important: ThisPresentation

PowerPoint already has a `ThisPresentation` object in every VBA project. If you import `ThisPresentation.cls` as a separate class file, PowerPoint will not use it for presentation events.

Use it like this:

1. Press `Alt+F11`.
2. In the Project tree, expand your presentation.
3. Expand **Microsoft PowerPoint Objects**.
4. Double-click the existing **ThisPresentation** object.
5. Paste the code from `ThisPresentation.cls` into that code window.

## Current Behavior

- No visible Start button is used.
- When slideshow mode starts, the countdown prompt appears once.
- The prompt default is read from the existing `countdown` timer text, so changing `05:00` to `09:00` before pressing F5 makes the popup default to `9` minutes.
- The selected timer value is applied to all slides.
- During playback, the timer keeps running across slide changes.
- Fast slide navigation updates the current slide immediately, while the timer syncs all slide timer shapes once per second.
- When the slideshow ends, the timer state is reset so the next slideshow run prompts again.

## Setup Workflow

1. Save the PowerPoint file as `.pptm`.
2. Open the VBA editor with `Alt+F11`.
3. Import or paste `Module1.bas` into a standard module.
4. Paste the event procedures from `ThisPresentation.cls` into the deck's existing `ThisPresentation` object.
5. Run `SetupShapes` once, or close/reopen the deck with macros enabled.
6. Press `F5`, enter the countdown duration once, and present normally.

## Simpler Manual Baseline

If you do not want to use `ThisPresentation` events, you can use only `Module1.bas`:

1. Import `Module1.bas`.
2. Run `SetupShapes` manually from `Alt+F8` once.
3. Start the slideshow.

This baseline is simpler, but it does not provide the automatic one-time prompt on slideshow start unless the `ThisPresentation` event code is also installed.
