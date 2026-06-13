VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTimerState"
Option Explicit

Private mTimerTickFlag As Boolean
Private mTargetEndTime As Date
Private mTimerID As LongPtr
Private mTimerActive As Boolean

Public Property Get TimerTickFlag() As Boolean
    TimerTickFlag = mTimerTickFlag
End Property

Public Property Get TargetEndTime() As Date
    TargetEndTime = mTargetEndTime
End Property

Public Property Get TimerID() As LongPtr
    TimerID = mTimerID
End Property

Public Property Get TimerActive() As Boolean
    TimerActive = mTimerActive
End Property

Public Property Get RemainingSeconds() As Long
    Dim secondsRemaining As Long

    secondsRemaining = DateDiff("s", Now, mTargetEndTime)
    If secondsRemaining < 0 Then
        secondsRemaining = 0
    End If

    RemainingSeconds = secondsRemaining
End Property

Public Sub StartTimerState(ByVal newTimerID As LongPtr, ByVal newTargetEndTime As Date)
    mTimerID = newTimerID
    mTargetEndTime = newTargetEndTime
    mTimerActive = True
    mTimerTickFlag = False
End Sub

Public Sub StopTimerState()
    mTimerTickFlag = False
    mTimerActive = False
    mTimerID = 0
    mTargetEndTime = 0
End Sub

Public Sub SetTickFlag()
    mTimerTickFlag = True
End Sub

Public Sub ClearTickFlag()
    mTimerTickFlag = False
End Sub
