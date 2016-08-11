Attribute VB_Name = "UI"
Option Explicit

Public Sub ShowGitSettingsForm()
    Load GitSettingsForm
    GitSettingsForm.resetForm
    MoveFormOnApplication GitSettingsForm
    GitSettingsForm.Show
    Unload GitSettingsForm
End Sub

Public Sub ShowGitRemoteForm()
    Load GitRemoteForm
    GitRemoteForm.resetForm
    MoveFormOnApplication GitRemoteForm
    GitRemoteForm.Show False
End Sub


Public Sub ShowSetExportDirectoryForm()
    Load SetExportDirectoryForm
    SetExportDirectoryForm.resetForm
    MoveFormOnApplication SetExportDirectoryForm
    SetExportDirectoryForm.Show
    Unload SetExportDirectoryForm
End Sub

Public Sub ShowGitCommitForm()
    Load GitCommitMessageForm
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show
End Sub

Public Sub ShowGitOtherForm()
    If ShibbySettings.ExportOnGit Then
        CodeUtils.ExportAll
    End If
    Load GitConsoleForm
    MoveFormOnApplication GitConsoleForm
    GitConsoleForm.OutputBox.ScrollBars = fmScrollBarsVertical
    GitConsoleForm.Show False
End Sub

Public Sub NonModalMsgBox(ByVal message As String)
    Load NonModalMsgBoxForm
    MoveFormOnApplication NonModalMsgBoxForm
    NonModalMsgBoxForm.Show False
    NonModalMsgBoxForm.Label1.Caption = message
End Sub


Public Function FolderDialog(Optional ByVal sTitle As String = "Browse") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = sTitle
    With fd
        If .Show = -1 Then
            FolderDialog = .SelectedItems(1)
            Exit Function
        End If
    End With
    FolderDialog = ""
End Function

Public Function FileDialog(Optional ByVal sTitle As String = "Browse") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = sTitle
    With fd
        If .Show = -1 Then
            FileDialog = .SelectedItems(1)
            Exit Function
        End If
    End With
    FileDialog = ""
End Function


Private Sub MoveFormOnApplication(ByVal form As Variant)
    form.Left = Application.ActiveWindow.Left
    form.Top = Application.ActiveWindow.Top
End Sub
