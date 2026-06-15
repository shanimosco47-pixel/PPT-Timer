Attribute VB_Name = "StartCountdown"
Option Explicit

Sub StartCountdown()
    Dim endTime As Date
    Dim slideIndex As Long
    Dim shp As Shape
    Dim ssw As SlideShowWindow
    Dim lastSecond As Long
    Dim currentSecond As Long

    On Error GoTo CleanExit

    If SlideShowWindows.Count = 0 Then
        MsgBox "יש להפעיל את המצגת לפני הפעלת הטיימר.", vbExclamation
        Exit Sub
    End If

    Set ssw = SlideShowWindows(1)

    On Error Resume Next
    Set shp = ssw.View.Slide.Shapes("countdown")
    On Error GoTo CleanExit

    If shp Is Nothing Then
        MsgBox "לא נמצאה צורה בשם 'countdown' על השקף הנוכחי." & vbNewLine & _
               "צור צורת טקסט ושנה את שמה ל-countdown.", vbExclamation
        Exit Sub
    End If

    slideIndex = ssw.View.CurrentShowPosition
    endTime = DateAdd("s", 300, Now)
    lastSecond = -1

    Do While Now < endTime
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Do
        If ssw.View.CurrentShowPosition <> slideIndex Then Exit Do

        currentSecond = DateDiff("s", Now, endTime)
        If currentSecond <> lastSecond Then
            lastSecond = currentSecond
            shp.TextFrame.TextRange.Text = Format(TimeSerial(0, currentSecond \ 60, currentSecond Mod 60), "mm:ss")
        End If
    Loop

    If SlideShowWindows.Count > 0 Then
        If ssw.View.CurrentShowPosition = slideIndex Then
            shp.TextFrame.TextRange.Text = "00:00"
        End If
    End If

    Exit Sub

CleanExit:
    Debug.Print Err.Number, Err.Description
End Sub
