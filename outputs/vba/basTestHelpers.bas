Attribute VB_Name = "basTestHelpers"
Option Explicit

Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then
        Err.Raise vbObjectError + 2000, "AssertTrue", message
    End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Public Sub AssertEqualsString(ByVal actualValue As String, ByVal expectedValue As String, ByVal message As String)
    If StrComp(actualValue, expectedValue, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 2001, "AssertEqualsString", message & " | expected: " & expectedValue & " | actual: " & actualValue
    End If
End Sub

Public Sub AssertEqualsLong(ByVal actualValue As Variant, ByVal expectedValue As Variant, ByVal message As String)
    If CLng(actualValue) <> CLng(expectedValue) Then
        Err.Raise vbObjectError + 2002, "AssertEqualsLong", message & " | expected: " & CStr(expectedValue) & " | actual: " & CStr(actualValue)
    End If
End Sub

Public Sub AssertDateSet(ByVal actualValue As Date, ByVal message As String)
    If actualValue = 0 Then
        Err.Raise vbObjectError + 2003, "AssertDateSet", message
    End If
End Sub
