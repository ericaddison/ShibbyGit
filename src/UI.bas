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


' public interface for export all
' input: folder - the folder to export code modules to
Public Sub ExportAllMsgBox()
    Dim folder As String
    folder = FileUtils.FolderBrowser("Browse for folder for export")
    If folder = "" Then
        Exit Sub
    End If
    NonModalMsgBox "Exporting files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = CodeUtils.ExportAllString(folder)
    
    HideNonModalMsgBox
    MsgBox output
End Sub


' public interface for import all
' input: folder - the folder to import code modules from
Public Sub ImportAllMsgBox()
    Dim folder As String
    folder = FileUtils.FolderBrowser("Browse for code folder to import")
    If folder = "" Then
        Exit Sub
    End If
    NonModalMsgBox "Importing files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = CodeUtils.ImportAllString(folder)
    
    HideNonModalMsgBox
    MsgBox output
End Sub
