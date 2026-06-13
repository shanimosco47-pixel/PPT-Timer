Attribute VB_Name = "basWin32API"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal uIDEvent As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function StartWin32Timer(ByVal intervalMilliseconds As Long) As LongPtr
    StartWin32Timer = SetTimer(0, 0, intervalMilliseconds, AddressOf TimerProc)
End Function

Public Sub StopWin32Timer(ByVal timerID As LongPtr)
    If timerID <> 0 Then
        KillTimer 0, timerID
    End If
End Sub

Public Sub TimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As LongPtr)
    On Error Resume Next

    If Not gTimerController Is Nothing Then
        gTimerController.MarkCallbackTick
    End If
End Sub
