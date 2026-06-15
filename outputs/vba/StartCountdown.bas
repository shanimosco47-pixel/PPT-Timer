Attribute VB_Name = "StartCountdown"
Option Explicit

Sub StartCountdown()
    Dim endTime As Date
    Dim shp As Shape
    Dim ssw As SlideShowWindow
    Dim lastSecond As Long
    Dim currentSecond As Long

    On Error GoTo CleanExit

    If SlideShowWindows.Count = 0 Then
        MsgBox "Start the slideshow before running the timer.", vbExclamation
        Exit Sub
    End If

    Set ssw = SlideShowWindows(1)
    endTime = DateAdd("s", 300, Now)
    lastSecond = -1

    Do While Now < endTime
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Do

        currentSecond = DateDiff("s", Now, endTime)
        If currentSecond <> lastSecond Then
            lastSecond = currentSecond

            ' Find the countdown shape on the current slide
            Set shp = Nothing
            On Error Resume Next
            Set shp = ssw.View.Slide.Shapes("countdown")
            On Error GoTo CleanExit

            If Not shp Is Nothing Then
                shp.TextFrame.TextRange.Text = Format(TimeSerial(0, currentSecond \ 60, currentSecond Mod 60), "mm:ss")
            End If
        End If
    Loop

    ' Show 00:00 on the current slide if countdown shape exists
    If SlideShowWindows.Count > 0 Then
        Set shp = Nothing
        On Error Resume Next
        Set shp = ssw.View.Slide.Shapes("countdown")
        On Error GoTo CleanExit
        If Not shp Is Nothing Then
            shp.TextFrame.TextRange.Text = "00:00"
        End If
    End If

    Exit Sub

CleanExit:
    Debug.Print Err.Number, Err.Description
End Sub
