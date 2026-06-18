Attribute VB_Name = "StartCountdown"
Option Explicit

' ============================================================
'  SET YOUR COUNTDOWN DURATION HERE (in seconds)
'  Examples: 300 = 5 min | 600 = 10 min | 900 = 15 min
' ============================================================
Private Const DURATION_SECONDS As Long = 300


' ------------------------------------------------------------
'  START COUNTDOWN — run this when the slideshow is active
' ------------------------------------------------------------
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
    endTime = DateAdd("s", DURATION_SECONDS, Now)
    lastSecond = -1
    lastSlide = -1

    On Error Resume Next

    Do While Now < endTime
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Do

        currentSecond = DateDiff("s", Now, endTime)
        currentSlide = ssw.View.CurrentShowPosition

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


' ------------------------------------------------------------
'  COPY COUNTDOWN TO ALL SLIDES
'  Run this ONCE in edit mode (not during slideshow)
'  It copies the "countdown" shape from the current slide
'  to every other slide, keeping position, size, and name.
' ------------------------------------------------------------
Sub CopyCountdownToAllSlides()
    Dim pres As Presentation
    Dim sourceSlide As Slide
    Dim targetSlide As Slide
    Dim sourceShape As Shape
    Dim newShape As Shape
    Dim i As Integer

    Set pres = ActivePresentation
    Set sourceSlide = ActiveWindow.View.Slide

    ' Find the countdown shape on the current slide
    On Error Resume Next
    Set sourceShape = sourceSlide.Shapes("countdown")
    On Error GoTo 0

    If sourceShape Is Nothing Then
        MsgBox "No shape named 'countdown' found on the current slide." & vbNewLine & _
               "Please select the slide that has the countdown shape first.", vbExclamation
        Exit Sub
    End If

    Dim copied As Integer
    copied = 0

    For i = 1 To pres.Slides.Count
        Set targetSlide = pres.Slides(i)

        ' Skip the source slide
        If targetSlide.SlideIndex = sourceSlide.SlideIndex Then GoTo NextSlide

        ' Remove existing countdown shape on target slide if any
        Dim existing As Shape
        On Error Resume Next
        Set existing = targetSlide.Shapes("countdown")
        On Error GoTo 0
        If Not existing Is Nothing Then existing.Delete

        ' Copy and paste
        sourceShape.Copy
        targetSlide.Shapes.Paste

        ' Rename and reposition to match source
        Set newShape = targetSlide.Shapes(targetSlide.Shapes.Count)
        newShape.Name = "countdown"
        newShape.Left = sourceShape.Left
        newShape.Top = sourceShape.Top
        newShape.Width = sourceShape.Width
        newShape.Height = sourceShape.Height

        copied = copied + 1

NextSlide:
    Next i

    MsgBox "Done! Countdown shape copied to " & copied & " slides.", vbInformation
End Sub
