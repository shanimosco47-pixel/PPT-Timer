# Handoff: PowerPoint Win32 Timer Architecture

## Purpose

This document is the handoff package for continuing the PowerPoint VBA / Win32 timer implementation from another computer.

It explains:
- what has already been completed,
- where the source code lives,
- what decisions have already been made,
- and the exact next steps to continue safely.

## Repository

GitHub repository:

- [peer-review-response-implementation-plan-vba](https://github.com/shanimosco47-pixel/peer-review-response-implementation-plan-vba)

Repository status at handoff:

- Remote exists on GitHub
- Visibility: `PRIVATE`
- Default branch: `main`
- Source code has already been committed and pushed

## What Was Already Approved

The architecture received review approval with one implementation caveat:

- The Win32 timer callback must not touch the PowerPoint Object Model.
- UI updates must execute only on PowerPoint's normal execution path.
- Cleanup must happen from multiple lifecycle paths.
- Missing or renamed shapes must fail safely.
- Timer ownership must remain centralized.
- `Application.OnTime` was removed because it was not verified for PowerPoint.
- The remaining open technical item is validation of the final polling mechanism inside PowerPoint.

## What Has Already Been Implemented

The VBA source modules have already been written and organized.

Main files:

- `README.md`
- `project_active_session.md`
- `outputs/vba/basLogger.bas`
- `outputs/vba/basWin32API.bas`
- `outputs/vba/basTimerBootstrap.bas`
- `outputs/vba/clsTimerState.cls`
- `outputs/vba/clsTimerController.cls`
- `outputs/vba/clsUIUpdater.cls`
- `outputs/vba/clsLifecycle.cls`
- `outputs/vba/basTestHelpers.bas`
- `outputs/vba/basTimerTests.bas`
- `outputs/vba/ACTIVEX_TIMER_SNIPPET.txt`
- `outputs/vba/README.md`

## What Each Module Does

- `basWin32API.bas`
  Contains only the Win32 declarations and the isolated callback entrypoint.

- `clsTimerState.cls`
  Stores timer state such as timer ID, active flag, end time, and tick flag.

- `clsTimerController.cls`
  Central controller for startup, tick processing, stop logic, and cleanup.

- `clsUIUpdater.cls`
  The only module allowed to touch PowerPoint shapes and slideshow UI state.

- `clsLifecycle.cls`
  Hooks PowerPoint lifecycle events such as slideshow start/end and presentation close.

- `basTimerBootstrap.bas`
  Public entrypoints such as `StartCountdownTimer`, `PollTimerTick`, `StopCountdownTimer`, and `ShutdownTimerSystem`.

- `basLogger.bas`
  Writes debug telemetry to `%TEMP%\PPT_Timer_Debug.log`.

- `basTimerTests.bas`
  Contains lightweight VBA test macros for non-UI behavior.

## Important Technical Constraints

Do not violate these rules while continuing implementation:

- Never put a breakpoint inside `TimerProc`.
- Never access `SlideShowWindow`, `View`, `Slide`, `Shape`, or any PowerPoint object from the Win32 callback.
- Keep all PowerPoint UI access inside the PowerPoint-safe path only.
- Keep cleanup idempotent and safe to call more than once.
- Treat the `DoEvents` loop as a fallback experiment, not the preferred production path.

## Current Logging Behavior

Runtime logs are written to:

- `%TEMP%\PPT_Timer_Debug.log`

The system already logs major events such as:

- timer startup,
- callback tick detection,
- UI update attempts,
- missing shape warnings,
- and cleanup events.

## Exact Next Steps on the Other Computer

### 1. Clone the repository

Clone this repository locally from:

- `https://github.com/shanimosco47-pixel/peer-review-response-implementation-plan-vba`

If GitHub authentication is required, sign in with the account that has access to the private repo.

### 2. Open the repository root

After cloning, open the project folder and review:

- `README.md`
- `project_active_session.md`
- `outputs/vba/README.md`

These three files give the fastest orientation.

### 3. Open the target PowerPoint `.pptm`

Open the presentation that should host the timer logic.

If the `.pptm` file is not yet under source control, decide where it will live and make sure the repository still remains the source of truth for the exported `.bas` and `.cls` files.

### 4. Import the VBA modules

Into the PowerPoint VBA editor, import the files from `outputs/vba/` in this order:

1. `basLogger.bas`
2. `clsTimerState.cls`
3. `clsUIUpdater.cls`
4. `clsTimerController.cls`
5. `clsLifecycle.cls`
6. `basWin32API.bas`
7. `basTimerBootstrap.bas`
8. `basTestHelpers.bas`
9. `basTimerTests.bas`

### 5. Add the countdown shape

Add a shape named exactly:

- `CountdownTimer`

Recommended location:

- the relevant slideshow slide, or
- the slide master if the timer must persist visually across slides

### 6. Choose and wire the polling mechanism

Preferred candidate:

- ActiveX timer-based polling on the PowerPoint UI side

Reference:

- `outputs/vba/ACTIVEX_TIMER_SNIPPET.txt`

Hook the chosen PowerPoint-safe mechanism so that it calls:

- `PollTimerTick`

at a 1000 ms interval.

### 7. Start with isolated validation

Before trying full production behavior:

- start a short timer such as `StartCountdownTimer 30`
- verify that the callback only sets the flag
- verify that UI updates happen through `PollTimerTick`
- verify that text updates appear in the `CountdownTimer` shape
- verify that no runtime error interrupts the slideshow

### 8. Run stress validation inside PowerPoint

Test while:

- advancing slides,
- going backward,
- playing embedded media if present,
- using presenter view if relevant,
- and ending the slideshow in multiple ways

Expected behavior:

- no crash,
- no orphaned timer,
- no visible host instability,
- and safe silent failure if the shape is missing

### 9. Verify cleanup paths

Specifically confirm that cleanup happens when:

- slideshow ends normally,
- presentation closes,
- manual stop is called,
- and runtime errors are trapped

### 10. Export updated modules back into the repository

If changes are made inside the VBA editor:

- export the updated `.bas` and `.cls` files back into the repository
- overwrite only after reviewing the content
- commit and push the exported source files back to GitHub

Do not leave the PowerPoint file as the only place where the code exists.

## Recommended First Commands for the Other Computer

After cloning the repo:

1. open `README.md`
2. open `outputs/vba/README.md`
3. import the modules into the `.pptm`
4. wire the polling mechanism
5. test with `StartCountdownTimer 30`
6. inspect `%TEMP%\PPT_Timer_Debug.log`

## Known Limitation

What has not been verified yet:

- full runtime behavior inside a live PowerPoint slideshow host
- the final production polling mechanism choice

This means the codebase is implementation-ready, but host validation still needs to be completed on the PowerPoint machine.

## If You Need To Resume Work in Codex

Start by telling Codex:

"Read `project_active_session.md` and `outputs/HANDOFF.md`, then help me continue the PowerPoint Win32 timer integration inside the `.pptm`."

That will give the next session the shortest path back into context.
