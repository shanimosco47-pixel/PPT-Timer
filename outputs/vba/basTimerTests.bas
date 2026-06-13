Attribute VB_Name = "basTimerTests"
Option Explicit

Public Sub RunAllTimerTests()
    TestTimerStateDefaults
    TestTimerStateStartMarksActive
    TestTimerStateStopClearsActivity
    TestTimerTickFlagRoundTrip
    TestFormatRemainingTime
End Sub

Public Sub TestTimerStateDefaults()
    Dim state As clsTimerState
    Set state = New clsTimerState

    AssertFalse state.TimerActive, "Timer should be inactive by default."
    AssertFalse state.TimerTickFlag, "Tick flag should be clear by default."
    AssertEqualsLong state.TimerID, 0, "Timer ID should default to zero."
End Sub

Public Sub TestTimerStateStartMarksActive()
    Dim state As clsTimerState
    Set state = New clsTimerState

    state.StartTimerState 1001, DateAdd("s", 30, Now)

    AssertTrue state.TimerActive, "Start should mark timer active."
    AssertEqualsLong state.TimerID, 1001, "Start should store timer ID."
    AssertDateSet state.TargetEndTime, "Start should set target end time."
End Sub

Public Sub TestTimerStateStopClearsActivity()
    Dim state As clsTimerState
    Set state = New clsTimerState

    state.StartTimerState 1001, DateAdd("s", 30, Now)
    state.StopTimerState

    AssertFalse state.TimerActive, "Stop should clear active state."
    AssertFalse state.TimerTickFlag, "Stop should clear tick flag."
    AssertEqualsLong state.TimerID, 0, "Stop should clear timer ID."
End Sub

Public Sub TestTimerTickFlagRoundTrip()
    Dim state As clsTimerState
    Set state = New clsTimerState

    state.SetTickFlag
    AssertTrue state.TimerTickFlag, "SetTickFlag should set the flag."

    state.ClearTickFlag
    AssertFalse state.TimerTickFlag, "ClearTickFlag should clear the flag."
End Sub

Public Sub TestFormatRemainingTime()
    Dim updater As clsUIUpdater
    Set updater = New clsUIUpdater

    AssertEqualsString updater.FormatRemainingTime(DateSerial(2026, 6, 13) + TimeSerial(10, 1, 30), DateSerial(2026, 6, 13) + TimeSerial(10, 0, 0)), "01:30", "Updater should format sub-hour durations as mm:ss."
    AssertEqualsString updater.FormatRemainingTime(DateSerial(2026, 6, 13) + TimeSerial(11, 5, 9), DateSerial(2026, 6, 13) + TimeSerial(10, 0, 0)), "01:05:09", "Updater should format hour-plus durations as hh:mm:ss."
    AssertEqualsString updater.FormatRemainingTime(DateSerial(2026, 6, 13) + TimeSerial(10, 0, 0), DateSerial(2026, 6, 13) + TimeSerial(10, 0, 10)), "00:00", "Updater should floor negative durations at zero."
End Sub
