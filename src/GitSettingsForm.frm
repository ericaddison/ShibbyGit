VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitSettingsForm 
   Caption         =   "ShibbyGit Settings"
   ClientHeight    =   7320
   ClientLeft      =   30
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





Private needGitUserNameUpdate As Boolean
Private needGitUserEmailUpdate As Boolean


'****************************************************************
' initialize

Public Sub resetForm()
    ' set the gitExe path text
    GitExeTextBox.Text = ShibbySettings.GitExePath
    
    ' set the project path text
    ProjectPathTextBox.Text = ShibbySettings.GitProjectPath
    
    If GitExeTextBox.Text <> "" Then
        ' set the username field
        Dim userName As String
        If ProjectPathTextBox.Text = "" Then
            userName = GitCommands.RunGitAsProcess("config user.name", UseProjectPath:=False)
        Else
            userName = GitCommands.RunGitAsProcess("config user.name")
        End If
        If Len(userName) > 0 Then
            userName = Left(userName, Len(userName) - 1)
        End If
        UserNameBox.value = userName
        
        ' set the email field
        Dim userEmail As String
        If ProjectPathTextBox.Text = "" Then
            userEmail = GitCommands.RunGitAsProcess("config user.email", UseProjectPath:=False)
        Else
            userEmail = GitCommands.RunGitAsProcess("config user.email")
        End If
        If Len(userEmail) > 0 Then
            userEmail = Left(userEmail, Len(userEmail) - 1)
        End If
        UserEmailBox.value = userEmail
    End If
    
    ' set the frx box value
    FrxCleanupBox.value = ShibbySettings.FrxCleanup
    
    ' set the frx box value
    ExportOnGitBox.value = ShibbySettings.ExportOnGit
    
    ' Add items to the file structure box
    FileStructureBox.AddItem "Flat File Stucture"
    FileStructureBox.AddItem "Simple Src Structure"
    FileStructureBox.AddItem "Separated Src Structure"
    Dim fsIndex As ShibbyFileStructure
    fsIndex = ShibbySettings.FileStructure
    FileStructureBox.ListIndex = fsIndex
    
    needGitUserNameUpdate = False
    needGitUserEmailUpdate = False
    
End Sub


'****************************************************************
' component callbacks

Private Sub CancelButton_Click()
    GitSettingsForm.Hide
End Sub

Private Sub OKButton_Click()
    SaveGitExe
    SaveProjectPath
    SaveUserName
    SaveUserEmail
    SaveFrxCleanup
    SaveExportOnGit
    SaveFileStructure
    GitSettingsForm.Hide
End Sub

Private Sub UserEmailBox_Change()
    needGitUserEmailUpdate = True
End Sub

Private Sub UserNameBox_Change()
    needGitUserNameUpdate = True
End Sub


Private Sub GitExeBrowseButton_Click()
    GitExeTextBox.Text = UI.FileDialog("Browser for git.exe")
End Sub


Private Sub ProjectPathBrowseButton_Click()
    ProjectPathTextBox.Text = UI.FolderDialog("Browse for Git project folder")
End Sub


'****************************************************************
' save methods

' Save the project path as a document property
Private Sub SaveProjectPath()
    Dim newPath As String
    newPath = ProjectPathTextBox.Text
    
    If newPath <> "" And FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitProjectPath = newPath
End Sub


' save the gitExe path as a registry property
Private Sub SaveGitExe()
    Dim newPath As String
    newPath = GitExeTextBox.Text
    
    If newPath <> "" And FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitExePath = newPath
End Sub

' save the user email to the git repo
Private Sub SaveUserEmail()
    If needGitUserEmailUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.email """ & UserEmailBox.value & """")
    End If
    needGitUserEmailUpdate = False
End Sub


' save the user name to the git repo
Private Sub SaveUserName()
    If needGitUserNameUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.name """ & UserNameBox.value & """")
    End If
    needGitUserNameUpdate = False
End Sub

' save the frx setting
Private Sub SaveFrxCleanup()
    ShibbySettings.FrxCleanup = FrxCleanupBox.value
End Sub

' save the export on git setting
Private Sub SaveExportOnGit()
    ShibbySettings.ExportOnGit = ExportOnGitBox.value
End Sub

' save the File structure
Private Sub SaveFileStructure()
    ShibbySettings.FileStructure = FileStructureBox.ListIndex
End Sub
