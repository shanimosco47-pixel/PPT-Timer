# VBA Source Files

Two files. Import both into the PowerPoint VBA editor (`Alt+F11`).

## Files

| File | Where to paste |
|------|---------------|
| `ThisPresentation.cls` | `ThisPresentation` module (double-click it in the Project tree) |
| `Module1.bas` | Insert → Module, paste contents |

## How it works

- **On open / slideshow start** — `SetupShapes` runs automatically and adds a `countdown` textbox + `▶ Start` button to every slide that doesn't already have them.
- **During slideshow** — click `▶ Start`, enter minutes, timer counts down on every slide simultaneously.
- **At zero** — display flashes red six times, stays red at `00:00`.

## Recipient workflow

1. Open `.pptm` → click **Enable Content**
2. Paste their slides
3. Press `F5` → click `▶ Start` → enter minutes → done
