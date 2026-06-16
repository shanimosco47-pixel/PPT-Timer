VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisPresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Presentation_Open()
    SetupShapes
End Sub

Private Sub Presentation_SlideShowBegin(ByVal Wn As SlideShowWindow)
    SetupShapes
End Sub
