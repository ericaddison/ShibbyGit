VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitConsoleForm 
   Caption         =   "ShibbyGit Console"
   ClientHeight    =   6084
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





Option Explicit
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
     
        Dim useShell As Boolean
        useShell = (Shift = 1)
             
        ' allow "git " to preceed options, for muscle memory!
        ' process "export" and "import" differently
        Dim output As String
        If CommandBox.Text Like "git *" Then
            CommandBox.Text = Right(CommandBox.Text, Len(CommandBox.Text) - 4)
        End If
        
        ' parse for available options
        If CommandBox.Text = "export" Then
            output = GitIO.GitExport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
        ElseIf CommandBox.Text = "import" Then
            output = GitIO.GitImport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
        Else
            If useShell Then
                output = "Shell exectution"
                GitCommands.RunGitInShell (CommandBox.Text)
            Else
                output = GitCommands.RunGitAsProcess(CommandBox.Text, 1500)
            End If
        End If
        
        ' push the command on the history
        If CommandBox.Text <> CommandHistory.Item(CommandIndex) Then
            CommandHistory.Add CommandBox.Text, After:=CommandIndex
            CommandIndex = CommandIndex + 1
        End If
        
        ' display the output
        OutputBox.value = output
        OutputBox.SelLength = 0
        OutputBox.SelStart = 0
        OutputBox.SetFocus
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


Private Sub CommandBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GiveCommandBoxFocusAndSelect
    End If
End Sub

Private Sub OutputBox_AfterUpdate()
    GiveCommandBoxFocusAndSelect
End Sub


Private Sub OutputBox_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    GiveCommandBoxFocusAndSelect
End Sub

Private Sub OutputBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    GiveCommandBoxFocusAndSelect
End Sub


Private Sub GiveCommandBoxFocusAndSelect()
    CommandBox.SetFocus
    CommandBox.SelStart = 0
    CommandBox.SelLength = Len(CommandBox.value)
End Sub

Private Sub OutputBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = vbKeyMButton Then
        With CommandBox
            .SelText = OutputBox.SelText
            .SetFocus
        End With
    End If
End Sub

Private Sub CommandBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = vbKeyMButton Then
        CommandBox.SelText = OutputBox.SelText
    End If
End Sub

