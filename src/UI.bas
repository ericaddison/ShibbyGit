Attribute VB_Name = "UI"
Public Sub ShowGitSettingsForm()
    Load GitSettingsForm
    GitSettingsForm.initialize
    GitSettingsForm.Show
End Sub

Public Sub ShowGitRemoteForm()

    Load GitRemoteForm
    
    Set GitRemoteForm.branches = GitParser.ParseBranches
    GitRemoteForm.AddBranchesToList
    
    Set GitRemoteForm.remotes = GitParser.ParseRemotes
    GitRemoteForm.AddPushRemotesToList
    
    GitRemoteForm.Show False

End Sub


Public Sub ShowSetExportDirectoryForm()
    Load SetWorkingDirectoryForm
    
    On Error Resume Next
        Dim dir As String
        dir = DocPropIO.GetItemFromDocProperties(CodeUtils.EXPORT_DIRECTORY_PROPERTY)
        SetWorkingDirectoryForm.DirTextBox.Text = dir
    On Error GoTo 0
        
    SetWorkingDirectoryForm.Show
End Sub



Public Sub ShowGitCommitForm()
    
    Load GitCommitMessageForm
    GitCommitMessageForm.Show
    
End Sub

Public Sub ShowGitOtherForm()
    
    Load GitConsoleForm
    GitConsoleForm.OutputBox.ScrollBars = fmScrollBarsVertical
    GitConsoleForm.Show False
    
End Sub

Public Function FolderDialog() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show = -1 Then
            FolderDialog = .SelectedItems(1)
            Exit Function
        End If
    End With
    FolderDialog = ""
End Function


