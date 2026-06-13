VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsUIUpdater"
Option Explicit

Private mTargetShapeName As String

Private Sub Class_Initialize()
    mTargetShapeName = "CountdownTimer"
End Sub

Public Property Get TargetShapeName() As String
    TargetShapeName = mTargetShapeName
End Property

Public Property Let TargetShapeName(ByVal value As String)
    mTargetShapeName = value
End Property

Public Function FormatRemainingTime(ByVal targetEndTime As Date, ByVal currentTime As Date) As String
    Dim totalSeconds As Long
    Dim hoursPart As Long
    Dim minutesPart As Long
    Dim secondsPart As Long

    totalSeconds = DateDiff("s", currentTime, targetEndTime)
    If totalSeconds < 0 Then
        totalSeconds = 0
    End If

    hoursPart = totalSeconds \ 3600
    minutesPart = (totalSeconds Mod 3600) \ 60
    secondsPart = totalSeconds Mod 60

    If hoursPart > 0 Then
        FormatRemainingTime = Format$(hoursPart, "00") & ":" & Format$(minutesPart, "00") & ":" & Format$(secondsPart, "00")
    Else
        FormatRemainingTime = Format$(minutesPart, "00") & ":" & Format$(secondsPart, "00")
    End If
End Function

Public Sub UpdateCountdown(ByVal state As clsTimerState, Optional ByVal slideShowWindow As SlideShowWindow)
    Dim countdownShape As Shape

    On Error GoTo SafeFail

    Set countdownShape = ResolveCountdownShape(slideShowWindow)
    If countdownShape Is Nothing Then
        WriteLog "[WARN] CountdownTimer shape missing; update skipped."
        Exit Sub
    End If

    If countdownShape.HasTextFrame <> msoTrue Then
        WriteLog "[WARN] CountdownTimer shape has no text frame; update skipped."
        Exit Sub
    End If

    On Error Resume Next
    countdownShape.TextFrame.TextRange.Text = FormatRemainingTime(state.TargetEndTime, Now)
    If Err.Number <> 0 Then
        WriteLog "[WARN] CountdownTimer text update failed: " & Err.Number & " - " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo SafeFail

    WriteLog "[UI_THREAD] CountdownTimer updated."
    Exit Sub

SafeFail:
    WriteLog "[WARN] UI update failed safely: " & Err.Number & " - " & Err.Description
End Sub

Private Function ResolveCountdownShape(Optional ByVal slideShowWindow As SlideShowWindow) As Shape
    Dim activeWindow As SlideShowWindow

    On Error Resume Next

    If slideShowWindow Is Nothing Then
        If SlideShowWindows.Count = 0 Then
            Exit Function
        End If

        Set activeWindow = SlideShowWindows(1)
    Else
        Set activeWindow = slideShowWindow
    End If

    If activeWindow Is Nothing Then
        Exit Function
    End If

    Set ResolveCountdownShape = activeWindow.View.Slide.Shapes(mTargetShapeName)
End Function
