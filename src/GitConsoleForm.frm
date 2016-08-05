VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitConsoleForm 
   Caption         =   "VB Git Console"
   ClientHeight    =   6375
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6204
   OleObjectBlob   =   "GitConsoleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitConsoleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private CommandHistory As New Collection
Private CommandIndex As Integer




' execute command when enter is pressed
Private Sub CommandBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' commandIndex checking
    If CommandIndex <= 0 Then
        CommandIndex = 1
    End If
    
    If CommandIndex > CommandHistory.Count Then
        CommandIndex = CommandHistory.Count
    End If
    
    ' Add a blank item if empty commandHistory
    If CommandHistory.Count = 0 Then
        CommandHistory.Add ""
    End If

    ' return key: process command
    If KeyCode = vbKeyReturn Then
        Debug.Print "Command: " & CommandBox.Text
     
             
        ' allow "git " to preceed options, for muscle memory!
        ' process "export" and "import" differently
        Dim output As String
        If CommandBox.Text Like "git *" Then
            Dim command As String
            command = Right(CommandBox.Text, Len(CommandBox.Text) - 4)
            output = Git.GitOther(command)
        ElseIf CommandBox.Text = "export" Then
            output = CodeUtils.ExportAll
        ElseIf CommandBox.Text = "import" Then
            output = CodeUtils.ImportAll
        Else
            output = Git.GitOther(CommandBox.Text)
        End If
        
        ' push the command on the history
        If CommandBox.Text <> CommandHistory.Item(CommandIndex) Then
            CommandHistory.Add CommandBox.Text, After:=CommandIndex
            CommandIndex = CommandIndex + 1
        End If
        
        ' display the output
        OutputBox.value = output
        KeyCode.value = 0
        
    ' up key: show previous command
    ElseIf KeyCode = vbKeyUp Then
        If CommandIndex > 1 Then
            CommandIndex = CommandIndex - 1
        End If
        CommandBox.Text = CommandHistory(CommandIndex)
        KeyCode.value = 0
        
    ' down key: show next command
    ElseIf KeyCode = vbKeyDown Then
        If CommandIndex < CommandHistory.Count Then
            CommandIndex = CommandIndex + 1
        End If
        CommandBox.Text = CommandHistory(CommandIndex)
        KeyCode.value = 0
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
