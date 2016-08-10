VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShibbyAppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub app_PresentationSave(ByVal Pres As Presentation)
    If ShibbySettings.ExportOnSave Then
        CodeUtils.ExportAllMsgBox
    End If
End Sub
