VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitConsoleForm 
   Caption         =   "VB Git Console"
   ClientHeight    =   6375
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6210
   OleObjectBlob   =   "GitConsoleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitConsoleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' execute command when enter is pressed
Private Sub CommandBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
     Debug.Print "Command: " & CommandBox.Text
      Dim message As String
      message = Git.GitOther(CommandBox.Text)
      OutputBox.value = message
  End If
End Sub


Private Sub OutputBox_AfterUpdate()
    CommandBox.SetFocus
    CommandBox.SelStart = 0
    CommandBox.SelLength = Len(CommandBox.value)
End Sub


Private Sub OutputBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CommandBox.SetFocus
    CommandBox.SelStart = 0
    CommandBox.SelLength = Len(CommandBox.value)
End Sub