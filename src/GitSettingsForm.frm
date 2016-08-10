VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitSettingsForm 
   Caption         =   "ShibbyGit Settings"
   ClientHeight    =   7812
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




Private needGitUserNameUpdate As Boolean
Private needGitUserEmailUpdate As Boolean


'****************************************************************
' initialize

Public Sub resetForm()
    ' set the gitExe path text
    Dim gitExe As String
    gitExe = GetSetting(CodeUtils.APPNAME, "FileInfo", GitCommands.EXE_PATH_PROPERTY, "")
    GitExeTextBox.Text = gitExe
    
    ' set the project path text
    Dim gitPath As String
    gitPath = DocPropIO.GetItemFromDocProperties(GitCommands.PROJECT_PATH_PROPERTY)
    ProjectPathTextBox.Text = gitPath
    
    ' set the username and email fields
    Dim userName As String
    userName = GitCommands.RunGitAsProcess("config user.name")
    If Len(userName) > 0 Then
        userName = Left(userName, Len(userName) - 1)
    End If
    UserNameBox.value = userName
    
    Dim userEmail As String
    userEmail = GitCommands.RunGitAsProcess("config user.email")
    If Len(userEmail) > 0 Then
        userEmail = Left(userEmail, Len(userEmail) - 1)
    End If
    UserEmailBox.value = userEmail
    
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
    DocPropIO.
        
End Sub
