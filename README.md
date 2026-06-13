# PowerPoint Win32 Timer Architecture

This repository contains the VBA source files and implementation notes for a PowerPoint countdown timer built with a Win32 `SetTimer` / `KillTimer` architecture.

## Repository Contents

- `outputs/vba/` - import-ready `.bas` and `.cls` source modules
- `outputs/vba/README.md` - VBA import order and runtime wiring notes
- `outputs/vba/ACTIVEX_TIMER_SNIPPET.txt` - example UI-thread polling hookup
- `project_active_session.md` - session memory and implementation context

## Key Design Rules

- The Win32 callback never touches the PowerPoint Object Model.
- UI updates are centralized and run on PowerPoint's normal execution path.
- Cleanup is defensive and triggered from multiple lifecycle paths.
- Logging is written to `%TEMP%\PPT_Timer_Debug.log`.

## Quick Start

1. Import the files from `outputs/vba/` into the PowerPoint VBA project.
2. Add a shape named `CountdownTimer` to the relevant slide or slide master.
3. Start the timer with `StartCountdownTimer 300`.
4. Connect the validated polling mechanism so it calls `PollTimerTick` every second.

## Status

The VBA source has been generated and organized for import. Final runtime validation still needs to happen inside PowerPoint, especially for the chosen polling mechanism.
