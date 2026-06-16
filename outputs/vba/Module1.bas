Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Private g_timerId As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

Private g_timerId As Long
#End If

Private g_totalSecs As Long
Private g_endTime As Date
Private g_lastSecond As Long
Private g_isRunning As Boolean
Private g_promptShown As Boolean

Sub StartCountdown()
    If SlideShowWindows.Count = 0 Then
        ActivePresentation.SlideShowSettings.Run
        Exit Sub
    End If

    g_promptShown = False
    StartCountdownForShow SlideShowWindows(1)
End Sub

Sub CountdownShowStarted(ByVal Wn As SlideShowWindow)
    g_promptShown = False
    SetupShapes
End Sub

Sub CountdownShowEnded()
    StopCountdown
    g_promptShown = False
End Sub

Sub OnSlideShowPageChange(ByVal Wn As SlideShowWindow)
    If Not g_promptShown Then
        StartCountdownForShow Wn
    Else
        RefreshCountdownOnSlide Wn
    End If
End Sub

Sub StartCountdownForShow(ByVal Wn As SlideShowWindow)
    Dim mins As String
    Dim minsValue As Double
    Dim displayText As String

    If g_promptShown Then
        RefreshCountdownOnSlide Wn
        Exit Sub
    End If

    g_promptShown = True
    StopCountdown

    If g_totalSecs <= 0 Then
        g_totalSecs = GetInitialTimerSeconds(Wn)
    End If

    SetupShapes
    displayText = FormatSeconds(g_totalSecs)

    mins = InputBox("Enter countdown duration in minutes:", "Countdown Timer", _
                    FormatMinutesForInput(g_totalSecs))

    If Trim$(mins) = "" Then
        UpdateCurrentSlide Wn, displayText, RGB(0, 0, 0)
        Exit Sub
    End If

    If Not IsNumeric(mins) Then
        MsgBox "Please enter a valid number of minutes.", vbExclamation, "Countdown Timer"
        UpdateCurrentSlide Wn, displayText, RGB(0, 0, 0)
        Exit Sub
    End If

    minsValue = CDbl(mins)

    If minsValue <= 0 Then
        MsgBox "Please enter a duration greater than zero.", vbExclamation, "Countdown Timer"
        UpdateCurrentSlide Wn, displayText, RGB(0, 0, 0)
        Exit Sub
    End If

    g_totalSecs = Fix(minsValue * 60)

    If g_totalSecs <= 0 Then
        MsgBox "Duration is too short. Please enter a larger value.", vbExclamation, "Countdown Timer"
        UpdateCurrentSlide Wn, displayText, RGB(0, 0, 0)
        Exit Sub
    End If

    g_endTime = DateAdd("s", g_totalSecs, Now)
    g_lastSecond = -1
    g_isRunning = True

    UpdateAllTimerShapes FormatSeconds(g_totalSecs), RGB(0, 0, 0)

    g_timerId = SetTimer(0, 0, 250, AddressOf CountdownTimerTick)
    If g_timerId = 0 Then
        MsgBox "Could not start the countdown timer.", vbExclamation, "Countdown Timer"
        g_isRunning = False
    End If
End Sub

Sub SetupShapes()
    Dim sl As Slide
    Dim shp As Shape
    Dim timerShp As Shape
    Dim foundTimer As Boolean
    Dim foundBtn As Boolean
    Dim pres As Presentation
    Dim displayText As String

    Set pres = ActivePresentation

    If g_totalSecs <= 0 Then
        g_totalSecs = GetInitialTimerSeconds(Nothing)
    End If

    If g_totalSecs > 0 Then
        displayText = FormatSeconds(g_totalSecs)
    Else
        displayText = "05:00"
    End If

    For Each sl In pres.Slides
        foundTimer = False
        foundBtn = False

        For Each shp In sl.Shapes
            If shp.Name = "countdown" Then
                foundTimer = True
                With shp.TextFrame.TextRange
                    .Text = displayText
                    With .Font
                        .Size = 40
                        .Bold = True
                        .Color.RGB = RGB(0, 0, 0)
                    End With
                End With
                shp.TextFrame.AutoSize = ppAutoSizeNone
                shp.TextFrame.WordWrap = msoFalse
                shp.Fill.Visible = msoFalse
                shp.Line.Visible = msoFalse
            ElseIf shp.Name = "timerBtn" Then
                foundBtn = True
            End If
        Next shp

        If Not foundTimer Then
            Set timerShp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                pres.PageSetup.SlideWidth - 160, pres.PageSetup.SlideHeight - 65, 150, 55)

            With timerShp
                .Name = "countdown"
                .TextFrame.TextRange.Text = displayText

                With .TextFrame.TextRange.Font
                    .Size = 40
                    .Bold = True
                    .Color.RGB = RGB(0, 0, 0)
                End With

                .TextFrame.AutoSize = ppAutoSizeNone
                .TextFrame.WordWrap = msoFalse
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
            End With
        End If

        If foundBtn Then
            sl.Shapes("timerBtn").Delete
        End If
    Next sl
End Sub

#If VBA7 Then
Sub CountdownTimerTick(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Sub CountdownTimerTick(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    Dim currentSecond As Long

    If Not g_isRunning Then Exit Sub

    If SlideShowWindows.Count = 0 Then
        StopCountdown
        g_promptShown = False
        Exit Sub
    End If

    currentSecond = DateDiff("s", Now, g_endTime)
    If currentSecond < 0 Then currentSecond = 0

    If currentSecond <> g_lastSecond Then
        g_lastSecond = currentSecond
        UpdateAllTimerShapes FormatSeconds(currentSecond), RGB(0, 0, 0)
    End If

    If currentSecond = 0 Then
        StopCountdown
        UpdateAllTimerShapes "00:00", RGB(204, 0, 0)
    End If
End Sub

Sub StopCountdown()
    If g_timerId <> 0 Then
        KillTimer 0, g_timerId
        g_timerId = 0
    End If

    g_isRunning = False
End Sub

Sub RefreshCountdownOnSlide(ByVal Wn As SlideShowWindow)
    Dim currentSecond As Long

    If Wn Is Nothing Then Exit Sub

    If Not g_isRunning Then
        UpdateCurrentSlide Wn, FormatSeconds(g_totalSecs), RGB(0, 0, 0)
        Exit Sub
    End If

    currentSecond = DateDiff("s", Now, g_endTime)
    If currentSecond < 0 Then currentSecond = 0

    UpdateCurrentSlide Wn, FormatSeconds(currentSecond), RGB(0, 0, 0)
End Sub

Sub UpdateCurrentSlide(ByVal Wn As SlideShowWindow, ByVal timeText As String, ByVal textColor As Long)
    Dim shp As Shape

    If Wn Is Nothing Then Exit Sub

    Set shp = GetTimerShape(Wn.View.Slide)
    If Not shp Is Nothing Then
        shp.TextFrame.TextRange.Text = timeText
        shp.TextFrame.TextRange.Font.Color.RGB = textColor
    End If
End Sub

Sub UpdateAllTimerShapes(ByVal timeText As String, ByVal textColor As Long)
    Dim sl As Slide
    Dim shp As Shape

    For Each sl In ActivePresentation.Slides
        Set shp = GetTimerShape(sl)
        If Not shp Is Nothing Then
            shp.TextFrame.TextRange.Text = timeText
            shp.TextFrame.TextRange.Font.Color.RGB = textColor
        End If
    Next sl
End Sub

Private Function GetTimerShape(ByVal sl As Slide) As Shape
    On Error Resume Next
    Set GetTimerShape = sl.Shapes("countdown")
    On Error GoTo 0
End Function

Private Function GetInitialTimerSeconds(ByVal Wn As SlideShowWindow) As Long
    Dim shp As Shape
    Dim sl As Slide
    Dim parsedSeconds As Long

    If Not Wn Is Nothing Then
        Set shp = GetTimerShape(Wn.View.Slide)
        parsedSeconds = ParseTimerText(shp)
        If parsedSeconds > 0 Then
            GetInitialTimerSeconds = parsedSeconds
            Exit Function
        End If
    End If

    For Each sl In ActivePresentation.Slides
        Set shp = GetTimerShape(sl)
        parsedSeconds = ParseTimerText(shp)
        If parsedSeconds > 0 Then
            GetInitialTimerSeconds = parsedSeconds
            Exit Function
        End If
    Next sl

    GetInitialTimerSeconds = 300
End Function

Private Function ParseTimerText(ByVal shp As Shape) As Long
    Dim parts() As String
    Dim rawText As String
    Dim mins As Long
    Dim secs As Long

    If shp Is Nothing Then Exit Function

    rawText = Trim$(shp.TextFrame.TextRange.Text)
    If InStr(rawText, ":") = 0 Then Exit Function

    parts = Split(rawText, ":")
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function

    mins = CLng(parts(0))
    secs = CLng(parts(1))

    If mins < 0 Or secs < 0 Or secs > 59 Then Exit Function

    ParseTimerText = mins * 60 + secs
End Function

Private Function FormatSeconds(ByVal totalSeconds As Long) As String
    If totalSeconds < 0 Then totalSeconds = 0

    FormatSeconds = Format$(totalSeconds \ 60, "00") & ":" & _
                    Format$(totalSeconds Mod 60, "00")
End Function

Private Function FormatMinutesForInput(ByVal totalSeconds As Long) As String
    If totalSeconds <= 0 Then
        FormatMinutesForInput = "5"
    ElseIf totalSeconds Mod 60 = 0 Then
        FormatMinutesForInput = CStr(totalSeconds \ 60)
    Else
        FormatMinutesForInput = Format$(totalSeconds / 60, "0.##")
    End If
End Function
