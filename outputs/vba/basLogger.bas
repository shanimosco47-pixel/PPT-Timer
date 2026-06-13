Attribute VB_Name = "basLogger"
Option Explicit

Private mLogFilePath As String

Public Sub ConfigureLogFilePath(Optional ByVal logFilePath As String = "")
    If LenB(logFilePath) = 0 Then
        mLogFilePath = Environ$("TEMP") & "\PPT_Timer_Debug.log"
    Else
        mLogFilePath = logFilePath
    End If

    EnsureLogFolderExists mLogFilePath
End Sub

Public Function GetLogFilePath() As String
    If LenB(mLogFilePath) = 0 Then
        ConfigureLogFilePath
    End If

    GetLogFilePath = mLogFilePath
End Function

Public Sub WriteLog(ByVal message As String)
    Dim fileNumber As Integer

    On Error Resume Next
    fileNumber = FreeFile
    Open GetLogFilePath() For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss"); " "; message
    Close #fileNumber
End Sub

Private Sub EnsureLogFolderExists(ByVal logFilePath As String)
    Dim fileSystem As Object
    Dim folderPath As String

    folderPath = Left$(logFilePath, InStrRev(logFilePath, "\") - 1)
    If LenB(folderPath) = 0 Then
        Exit Sub
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub
