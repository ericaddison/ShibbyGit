Attribute VB_Name = "UI"
Public Sub ShowSetExportDirectoryForm()
    Load SetExportDirectoryForm
    
    On Error Resume Next
        Dim dir As String
        dir = DocPropIO.GetItemFromDocProperties(Constants.EXPORT_DIRECTORY_PROPERTY)
        SetExportDirectoryForm.DirTextBox.Text = dir
    On Error GoTo 0
        
    SetExportDirectoryForm.Show
End Sub



Public Sub ShowSetGitPathForm()
    Load GitPathForm
    
    On Error Resume Next
        Dim gitExe As String
        gitExe = GetSetting("CVX_CodeUtils", "FileInfo", Git.GIT_PATH_PROPERTY, "")
        GitPathForm.DirTextBox.Text = gitExe
    On Error GoTo 0
        
    GitPathForm.Show
End Sub

Public Sub ShowGitCommitForm()
    
    Load GitCommitMessageForm
    GitCommitMessageForm.Show
    
End Sub

Public Sub ShowGitOtherForm()
    
    Load GitOtherForm
    GitOtherForm.Show
    
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


