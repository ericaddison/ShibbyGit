Attribute VB_Name = "UI"
Option Explicit

Public Sub ShowGitSettingsForm()
    Load GitSettingsForm
    GitSettingsForm.resetForm
    GitSettingsForm.Show
End Sub

Public Sub ShowGitRemoteForm()
    Load GitRemoteForm
    GitRemoteForm.resetForm
    GitRemoteForm.Show False
End Sub


Public Sub ShowSetExportDirectoryForm()
    Load SetExportDirectoryForm
    SetExportDirectoryForm.resetForm
    SetExportDirectoryForm.Show
End Sub

Public Sub ShowGitCommitForm()
    GitCommitMessageForm.Show
End Sub

Public Sub ShowGitOtherForm()
    
    Load GitConsoleForm
    GitConsoleForm.OutputBox.ScrollBars = fmScrollBarsVertical
    GitConsoleForm.Show False
    
End Sub

Public Sub NonModalMsgBox(ByVal message As String)
    Load NonModalMsgBoxForm
    NonModalMsgBoxForm.Show False
    NonModalMsgBoxForm.Label1.Caption = message
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


