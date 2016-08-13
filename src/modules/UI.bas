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


Public Sub ShowGitCommitForm()
    Load GitCommitMessageForm
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show
End Sub

Public Sub ShowGitOtherForm()
    If ShibbySettings.ExportOnGit Then
        GitIO.GitExport ShibbySettings.GitProjectPath, ShibbySettings.fileStructure
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


' public interface for import from
Public Sub ImportSelectedMsgBox()
    Dim files As FileDialogSelectedItems
    Set files = FileUtils.FileBrowserMultiSelect("Browse for code files to import", _
            "VBA Code Module", "*.bas; *.frm; *.cls")
    
    If files Is Nothing Then
        Exit Sub
    End If
    
    NonModalMsgBox "Importing files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = CodeUtils.ImportSelectedString(files)
    
    HideNonModalMsgBox
    MsgBox output
End Sub


' public interface for GitExport
Public Sub GitExportMsgBox()
    NonModalMsgBox "Exporting files to Git Folder" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = GitIO.GitExport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
    
    HideNonModalMsgBox
    MsgBox output
End Sub

' public interface for GitImport
Public Sub GitImportMsgBox()
    NonModalMsgBox "Importing files from Git Folder" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = GitIO.GitImport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
    
    HideNonModalMsgBox
    MsgBox output
End Sub
