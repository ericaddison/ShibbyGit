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
        CodeUtils.ExportAllString ShibbySettings.ImportExportPath
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

Public Sub HideNonModalMsgBox()
    NonModalMsgBoxForm.hide
End Sub


Private Sub MoveFormOnApplication(ByVal form As Variant)
    form.Left = Application.ActiveWindow.Left
    form.Top = Application.ActiveWindow.Top
End Sub
