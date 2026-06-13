# PowerPoint Win32 Timer VBA Modules

Files in this folder are import-ready VBA source modules for the approved timer architecture.

These files are meant to be imported into a private PowerPoint `.pptm`. The presentation file itself is intentionally kept out of this public repository.

Suggested import order:

1. `basLogger.bas`
2. `clsTimerState.cls`
3. `clsUIUpdater.cls`
4. `clsTimerController.cls`
5. `clsLifecycle.cls`
6. `basWin32API.bas`
7. `basTimerBootstrap.bas`
8. `basTestHelpers.bas`
9. `basTimerTests.bas`

Minimal runtime flow:

1. Import all modules into the PowerPoint VBA project.
2. Add a shape named `CountdownTimer` to the relevant slideshow slide or slide master.
3. Call `StartCountdownTimer 300` to start a 5-minute timer.
4. Wire the validated UI-thread polling mechanism to call `PollTimerTick` every 1000 ms.
5. Use `StopCountdownTimer` or `ShutdownTimerSystem` for manual cleanup if needed.
6. After making changes inside PowerPoint, export the updated modules back into this folder and commit them to GitHub.

Notes:

- `TimerProc` does not touch the PowerPoint Object Model.
- UI updates are centralized in `clsUIUpdater`.
- Logging defaults to `%TEMP%\PPT_Timer_Debug.log`.
- `ExperimentalRunPollingLoop` is included only as a validation fallback, not the preferred production polling strategy.
