VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsLifecycle"
Option Explicit

Private WithEvents mApp As Application

Public Sub Initialize(ByVal hostApp As Application)
    Set mApp = hostApp
End Sub

Private Sub mApp_SlideShowBegin(ByVal Wn As SlideShowWindow)
    On Error GoTo ErrorHandler

    WriteLog "[EVENT] SlideShowBegin triggered."
    If Not gTimerController Is Nothing Then
        gTimerController.RenderCurrentState Wn
    End If
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] SlideShowBegin handler failed: " & Err.Number & " - " & Err.Description
End Sub

Private Sub mApp_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    On Error GoTo ErrorHandler

    WriteLog "[EVENT] SlideShowNextSlide triggered."
    If Not gTimerController Is Nothing Then
        gTimerController.RenderCurrentState Wn
    End If
    Exit Sub

ErrorHandler:
    WriteLog "[ERROR] SlideShowNextSlide handler failed: " & Err.Number & " - " & Err.Description
End Sub

Private Sub mApp_SlideShowEnd(ByVal Pres As Presentation)
    On Error Resume Next

    WriteLog "[CLEANUP] SlideShowEnd triggered."
    If Not gTimerController Is Nothing Then
        gTimerController.StopCountdown "SlideShowEnd"
    End If
End Sub

Private Sub mApp_PresentationClose(ByVal Pres As Presentation)
    On Error Resume Next

    WriteLog "[CLEANUP] PresentationClose triggered."
    If Not gTimerController Is Nothing Then
        gTimerController.StopCountdown "PresentationClose"
    End If
End Sub

Private Sub mApp_PresentationCloseFinal(ByVal Pres As Presentation)
    On Error Resume Next

    WriteLog "[CLEANUP] PresentationCloseFinal triggered."
    If Not gTimerController Is Nothing Then
        gTimerController.StopCountdown "PresentationCloseFinal"
    End If
End Sub
