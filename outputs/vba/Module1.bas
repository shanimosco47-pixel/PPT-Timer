Attribute VB_Name = "Module1"
Option Explicit

Dim g_totalSecs As Long

Sub StartCountdown()
    Dim mins As String
    mins = InputBox("Enter countdown duration in minutes:", "Countdown Timer", "5")
    If mins = "" Or Not IsNumeric(mins) Then Exit Sub
    g_totalSecs = CLng(CDbl(mins) * 60)
    RunTimer
End Sub

Sub SetupShapes()
    Dim pres As Presentation
    Dim sl As Slide
    Dim shp As Shape
    Dim foundTimer As Boolean
    Dim foundBtn As Boolean

    Set pres = ActivePresentation

    Dim tW As Single: tW = 130
    Dim tH As Single: tH = 55
    Dim tL As Single: tL = pres.PageSetup.SlideWidth - tW - 20
    Dim tT As Single: tT = pres.PageSetup.SlideHeight - tH - 20

    Dim bW As Single: bW = 90
    Dim bH As Single: bH = 38
    Dim bL As Single: bL = tL - bW - 10
    Dim bT As Single: bT = pres.PageSetup.SlideHeight - bH - 28

    For Each sl In pres.Slides
        foundTimer = False
        foundBtn = False
        For Each shp In sl.Shapes
            If shp.Name = "countdown" Then foundTimer = True
            If shp.Name = "timerBtn" Then foundBtn = True
        Next shp

        If Not foundTimer Then
            Set shp = sl.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, tL, tT, tW, tH)
            shp.Name = "countdown"
            With shp.TextFrame
                .TextRange.Text = "00:00"
                .TextRange.Font.Size = 40
                .TextRange.Font.Bold = True
                .TextRange.Font.Color.RGB = RGB(0, 0, 0)
                .TextRange.ParagraphFormat.Alignment = ppAlignRight
                .AutoSize = ppAutoSizeNone
                .WordWrap = msoFalse
            End With
            shp.Fill.Transparency = 1
            shp.Line.Visible = msoFalse
        End If

        If Not foundBtn Then
            Set shp = sl.Shapes.AddShape( _
                msoShapeRoundedRectangle, bL, bT, bW, bH)
            shp.Name = "timerBtn"
            With shp
                .TextFrame.TextRange.Text = Chr(9654) & " Start"
                .TextFrame.TextRange.Font.Size = 16
                .TextFrame.TextRange.Font.Bold = True
                .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                .Fill.ForeColor.RGB = RGB(0, 112, 192)
                .Line.Visible = msoFalse
                .ActionSettings(ppMouseClick).Action = ppActionRunMacro
                .ActionSettings(ppMouseClick).Run = "StartCountdown"
            End With
        End If
    Next sl
End Sub

Sub RunTimer()
    Dim endTime As Date
    Dim lastSecond As Long
    Dim currentSecond As Long
    Dim timeStr As String
    Dim sl As Slide
    Dim shp As Shape

    If SlideShowWindows.Count = 0 Then Exit Sub

    endTime = DateAdd("s", g_totalSecs, Now)
    lastSecond = -1

    Do While Now < endTime
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Do

        currentSecond = DateDiff("s", Now, endTime)
        If currentSecond <> lastSecond Then
            lastSecond = currentSecond
            timeStr = Format(currentSecond \ 60, "00") & ":" & _
                      Format(currentSecond Mod 60, "00")
            For Each sl In ActivePresentation.Slides
                Set shp = Nothing
                On Error Resume Next
                Set shp = sl.Shapes("countdown")
                On Error GoTo 0
                If Not shp Is Nothing Then
                    shp.TextFrame.TextRange.Text = timeStr
                End If
            Next sl
        End If

        PauseSeconds 0.1
    Loop

    Dim i As Integer
    For i = 1 To 6
        For Each sl In ActivePresentation.Slides
            Set shp = Nothing
            On Error Resume Next
            Set shp = sl.Shapes("countdown")
            On Error GoTo 0
            If Not shp Is Nothing Then
                shp.TextFrame.TextRange.Font.Color.RGB = _
                    IIf(i Mod 2 = 0, RGB(0, 0, 0), RGB(204, 0, 0))
                shp.TextFrame.TextRange.Text = "00:00"
            End If
        Next sl
        Dim t As Date
        t = Now + TimeValue("00:00:01") / 2
        Do While Now < t: DoEvents: Loop
    Next i

    For Each sl In ActivePresentation.Slides
        Set shp = Nothing
        On Error Resume Next
        Set shp = sl.Shapes("countdown")
        On Error GoTo 0
        If Not shp Is Nothing Then
            shp.TextFrame.TextRange.Font.Color.RGB = RGB(204, 0, 0)
            shp.TextFrame.TextRange.Text = "00:00"
        End If
    Next sl
End Sub

Sub PauseSeconds(ByVal seconds As Single)
    Dim startTime As Single
    startTime = Timer
    Do While Timer - startTime < seconds
        DoEvents
        If SlideShowWindows.Count = 0 Then Exit Sub
    Loop
End Sub
