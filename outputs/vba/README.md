# VBA Source Files

Import these source files into a macro-enabled PowerPoint deck (`.pptm`). Do not commit or upload the deck itself if it contains private presentation content.

## Files

| File | Where to paste/import |
|------|------------------------|
| `Module1.bas` | Insert -> Module, then paste/import the contents |
| `ThisPresentation.cls` | Paste the event procedures into the existing `ThisPresentation` object |

## Current behavior

- No visible Start button is used.
- When slideshow mode starts, the countdown prompt appears once.
- The prompt default is read from the existing `countdown` timer text, so changing `05:00` to `09:00` before pressing F5 makes the popup default to `9` minutes.
- The selected timer value is applied to all slides.
- During playback, the timer keeps running across slide changes.
- Fast slide navigation updates the current slide immediately, while the timer syncs all slide timer shapes once per second.
- When the slideshow ends, the timer state is reset so the next slideshow run prompts again.

## Recipient workflow

1. Save the PowerPoint file as `.pptm`.
2. Open the VBA editor with `Alt+F11`.
3. Import or paste `Module1.bas` into a standard module.
4. Paste the `ThisPresentation.cls` event procedures into the deck's existing `ThisPresentation` object.
5. Run `SetupShapes` once, or close/reopen the deck with macros enabled.
6. Press `F5`, enter the countdown duration once, and present normally.
