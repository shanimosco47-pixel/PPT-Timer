Attribute VB_Name = "basTimerBootstrap"
Option Explicit

Public gLifecycle As clsLifecycle
Public gTimerController As clsTimerController

Public Sub InitializeTimerSystem(Optional ByVal logFilePath As String = "")
    On Error GoTo ErrorHandler

    If gTimerController Is Nothing Then
        Set gTimerController = New clsTimerController
        gTimerController.Initialize logFilePath, "CountdownTimer"
    End If

    If gLifecycle Is Nothing Then
        Set gLifecycle = New clsLifecycle
        gLifecycle.Initialize Application
    End If

    WriteLog "[INIT] Timer system initialized."
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] InitializeTimerSystem failed: " & Err.Number & " - " & Err.Description
    ShutdownTimerSystem
End Sub

Public Sub StartCountdownTimer(ByVal durationSeconds As Long, Optional ByVal logFilePath As String = "")
    On Error GoTo ErrorHandler

    InitializeTimerSystem logFilePath
    gTimerController.StartCountdown durationSeconds
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] StartCountdownTimer failed: " & Err.Number & " - " & Err.Description
    ShutdownTimerSystem
End Sub

Public Sub StopCountdownTimer()
    On Error Resume Next

    If Not gTimerController Is Nothing Then
        gTimerController.StopCountdown "manual stop"
    End If
End Sub

Public Sub PollTimerTick()
    On Error GoTo ErrorHandler

    If Not gTimerController Is Nothing Then
        gTimerController.ProcessPendingTick
    End If

    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] PollTimerTick failed: " & Err.Number & " - " & Err.Description
    ShutdownTimerSystem
End Sub

Public Sub ExperimentalRunPollingLoop(Optional ByVal sleepMilliseconds As Long = 200)
    On Error GoTo ErrorHandler

    If gTimerController Is Nothing Then
        Exit Sub
    End If

    WriteLog "[LOOP] Experimental polling loop started."
    Do While gTimerController.IsPollingActive
        gTimerController.ProcessPendingTick
        DoEvents
        Sleep sleepMilliseconds
    Loop
    WriteLog "[LOOP] Experimental polling loop stopped."
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] ExperimentalRunPollingLoop failed: " & Err.Number & " - " & Err.Description
    ShutdownTimerSystem
End Sub

Public Sub ShutdownTimerSystem()
    On Error Resume Next

    If Not gTimerController Is Nothing Then
        gTimerController.StopCountdown "shutdown"
        Set gTimerController = Nothing
    End If

    If Not gLifecycle Is Nothing Then
        Set gLifecycle = Nothing
    End If

    WriteLog "[CLEANUP] Timer system shutdown complete."
End Sub

Public Sub Auto_Close()
    ShutdownTimerSystem
End Sub
