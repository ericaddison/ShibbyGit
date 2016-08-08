VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitSettingsForm 
   Caption         =   "ShibbyGit Settings"
   ClientHeight    =   5568
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   8580
   OleObjectBlob   =   "GitSettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub initialize()
    ' set the gitExe path text
    On Error Resume Next
        Dim gitExe As String
        gitExe = GetSetting(CodeUtils.APPNAME, "FileInfo", GitCommands.EXE_PATH_PROPERTY, "")
        GitExeTextBox.Text = gitExe
    On Error GoTo 0
    
    ' set the project path text
    On Error Resume Next
        Dim gitPath As String
        gitPath = DocPropIO.GetItemFromDocProperties(GitCommands.PROJECT_PATH_PROPERTY)
        ProjectPathTextBox.Text = gitPath
    On Error GoTo 0
End Sub


Private Sub CancelButton_Click()
    GitSettingsForm.Hide
End Sub

Private Sub OKButton_Click()
    SaveGitExe
    SaveProjectPath
    GitSettingsForm.Hide
End Sub


Private Sub GitExeBrowseButton_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        If .Show = -1 Then
            GitExeTextBox.Text = .SelectedItems(1)
        End If
    End With
End Sub


Private Sub ProjectPathBrowseButton_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show = -1 Then
            ProjectPathTextBox.Text = .SelectedItems(1)
        End If
    End With
End Sub


' Save the project path as a document property
Private Sub SaveProjectPath()
    Dim newPath As String
    newPath = ProjectPathTextBox.Text
    
    If FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    DocPropIO.AddStringToDocProperties GitCommands.PROJECT_PATH_PROPERTY, newPath
End Sub


' save the gitExe path as a registry property
Private Sub SaveGitExe()
    Dim newPath As String
    newPath = GitExeTextBox.Text
    
    If FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    Call SaveSetting(CodeUtils.APPNAME, "FileInfo", GitCommands.EXE_PATH_PROPERTY, newPath)
End Sub
