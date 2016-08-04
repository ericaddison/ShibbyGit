VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitConsoleForm 
   Caption         =   "User Git Command"
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


Private Sub OKButton_Click()
    Dim options As String
    options = OptionsBox.Text
    
    If options = "" Then
        MsgBox "Please enter git options"
        Exit Sub
    End If
    
    Git.GitOther (options)
        
End Sub


Private Sub OutputBox_Enter()
    CommandBox.SetFocus
End Sub
