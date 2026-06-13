VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTimerController"
Option Explicit

Private mState As clsTimerState
Private mUIUpdater As clsUIUpdater
Private mPollingActive As Boolean

Public Sub Initialize(Optional ByVal logFilePath As String = "", Optional ByVal targetShapeName As String = "CountdownTimer")
    Set mState = New clsTimerState
    Set mUIUpdater = New clsUIUpdater
    mUIUpdater.TargetShapeName = targetShapeName
    ConfigureLogFilePath logFilePath
End Sub

Public Property Get State() As clsTimerState
    Set State = mState
End Property

Public Property Get IsPollingActive() As Boolean
    IsPollingActive = mPollingActive
End Property

Public Sub StartCountdown(ByVal durationSeconds As Long)
    Dim timerID As LongPtr
    Dim targetEndTime As Date

    On Error GoTo ErrorHandler

    If mState Is Nothing Then
        Initialize
    End If

    StopCountdown "restart"

    targetEndTime = DateAdd("s", durationSeconds, Now)
    timerID = StartWin32Timer(1000)
    If timerID = 0 Then
        Err.Raise vbObjectError + 3000, "clsTimerController.StartCountdown", "SetTimer returned 0."
    End If

    mState.StartTimerState timerID, targetEndTime
    mPollingActive = True

    WriteLog "[SUCCESS] SetTimer initialized. ID=" & CStr(timerID)
    WriteLog "[STATE] Countdown started for " & CStr(durationSeconds) & " seconds."
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] StartCountdown failed: " & Err.Number & " - " & Err.Description
    StopCountdown "start failure"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub StopCountdown(Optional ByVal reason As String = "")
    On Error Resume Next

    If Not mState Is Nothing Then
        If mState.TimerID <> 0 Then
            StopWin32Timer mState.TimerID
            WriteLog "[CLEANUP] KillTimer executed. Reason=" & reason
        End If

        mState.StopTimerState
    End If

    mPollingActive = False
End Sub

Public Sub MarkCallbackTick()
    On Error Resume Next

    If Not mState Is Nothing Then
        mState.SetTickFlag
    End If
End Sub

Public Sub ProcessPendingTick(Optional ByVal slideShowWindow As SlideShowWindow)
    On Error GoTo ErrorHandler

    If mState Is Nothing Then
        Exit Sub
    End If

    If Not mPollingActive Then
        Exit Sub
    End If

    If Not mState.TimerActive Then
        Exit Sub
    End If

    If Not mState.TimerTickFlag Then
        Exit Sub
    End If

    WriteLog "[UI_THREAD] Poll cycle entered; flag detected."
    mState.ClearTickFlag
    mUIUpdater.UpdateCountdown mState, slideShowWindow

    If mState.RemainingSeconds = 0 Then
        StopCountdown "completed"
    End If

    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] ProcessPendingTick failed: " & Err.Number & " - " & Err.Description
    StopCountdown "poll failure"
End Sub

Public Sub RenderCurrentState(Optional ByVal slideShowWindow As SlideShowWindow)
    On Error GoTo ErrorHandler

    If mState Is Nothing Then
        Exit Sub
    End If

    If Not mState.TimerActive Then
        Exit Sub
    End If

    mUIUpdater.UpdateCountdown mState, slideShowWindow
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] RenderCurrentState failed: " & Err.Number & " - " & Err.Description
End Sub
