Attribute VB_Name = "StartCountdown"
Option Explicit

Sub StartCountdown()
    Dim endTime As Date
    Dim shp As Shape
    Dim ssw As SlideShowWindow
    Dim lastSecond As Long
    Dim lastSlide As Long
    Dim currentSecond As Long
    Dim currentSlide As Long

    On Error GoTo CleanExit

    If SlideShowWindows.Count = 0 Then
        MsgBox "Start the slideshow before running the timer.", vbExclamation
        Exit Sub
    End If

    Set ssw = SlideShowWindows(1)
    endTime = DateAdd("s", 300, Now)
    lastSecond = -1
    lastSlide = -1

    On Error Resume Next

    Do While Now < endTime
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Do

        currentSecond = DateDiff("s", Now, endTime)
        currentSlide = ssw.View.CurrentShowPosition

        ' Update when second changes OR when slide changes
        If currentSecond <> lastSecond Or currentSlide <> lastSlide Then
            lastSecond = currentSecond
            lastSlide = currentSlide

            Set shp = Nothing
            Set shp = ssw.View.Slide.Shapes("countdown")

            If Not shp Is Nothing Then
                shp.TextFrame.TextRange.Text = Format(TimeSerial(0, currentSecond \ 60, currentSecond Mod 60), "mm:ss")
            End If
        End If
    Loop

    Set shp = Nothing
    Set shp = ssw.View.Slide.Shapes("countdown")
    If Not shp Is Nothing Then
        shp.TextFrame.TextRange.Text = "00:00"
    End If

    Exit Sub

CleanExit:
    Debug.Print Err.Number, Err.Description
End Sub
