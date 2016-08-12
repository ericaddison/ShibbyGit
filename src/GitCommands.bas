Attribute VB_Name = "GitCommands"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems


Public Sub GitCommit(ByVal message As String)
    Dim out As String
    out = RunGitAsProcess("commit -am """ & message & """")
    If out = "" Then
        out = "No output from commit"
    End If
    MsgBox out
End Sub

Public Sub GitStatus()
    If ShibbySettings.ExportOnGit Then
        CodeUtils.ExportAllString ShibbySettings.ImportExportPath
    End If
    Dim out As String
    out = RunGitAsProcess("status")
    If out = "" Then
        out = "No output from status"
    End If
    MsgBox out
End Sub


Public Sub GitLog()
    Dim out As String
    out = RunGitAsProcess("log --max-count=5")
    If out = "" Then
        out = "No output from log"
    End If
    MsgBox out
End Sub

Public Sub GitAddAll()
    Dim out As String
    out = RunGitAsProcess("add -A")
    If out = "" Then
        MsgBox "Staged all files for commit"
    Else
        MsgBox "Git response: " & vbCrLf & out
    End If
End Sub

' Rin git in a command shell with incoming options.
' Output is not returned, but full cmd shell interactivity
' is possible. Ends with a call to "pause" to keep the window open
Public Sub RunGitInShell(ByVal options As String, Optional ByVal UseProjectPath As Boolean = True)
    Dim gitExe As String
    
    If UseProjectPath Then
        gitExe = GitExeWithPath
    Else
        gitExe = GetGitExe
    End If

    Dim command As String
    command = "cmd /c echo Running 'git " & options & "'" & _
        " & " & gitExe & options & " & pause"
        
    Debug.Print command
    shell command, 1
End Sub


' Run git with specified options ... calls "git -C <path> options",
' Launches a new process and returns the output
' path to git executable and project directory come from settings
' VBA will wait "waitTime" milliseconds for the process to complete. Default value is 20000 for 20s
Public Function RunGitAsProcess(ByVal options As String, Optional ByVal waitTime As Long = 20000, _
        Optional ByVal UseProjectPath As Boolean = True) As String

    Dim gitExe As String
    gitExe = GetGitExe(False)
    
    Dim parms As String
    If UseProjectPath Then
        options = GitPathOption(GetWorkingDir) & " " & options
    End If
    
    If gitExe = "" Then
        RunGitAsProcess = "Failed to run Git as Process"
        Exit Function
    End If

    ' call git
    Dim output As String
    output = ShellRedirect.Redirect(gitExe, options, waitTime)
    
    RunGitAsProcess = output
End Function


' get the gitExe path with the -C working directory parameter added
' optional boolean quoteGitExe, if true, adds quotes around GitExe path if spaces in path
Public Function GitExeWithPath() As String
    GitExeWithPath = ""

    Dim git As String
    git = GetGitExe
    If git = "" Then
        Exit Function
    End If
    
    Dim workingDir As String
    workingDir = GetWorkingDir
    If workingDir = "" Then
        Exit Function
    End If

    GitExeWithPath = git & " " & GitPathOption(workingDir) & " "
    
End Function

Private Function GetGitExe(Optional quoteGitExe As Boolean = True) As String
    GetGitExe = ""

    ' get the git executable path
    Dim gitExe As String
    gitExe = ShibbySettings.GitExePath
    
    If gitExe = "" Or IsNull(gitExe) Then
        MsgBox "Please set the git executable path"
        gitExe = FileUtils.FileBrowser
        If (gitExe = "") Then
            Exit Function
        Else
            ShibbySettings.GitExePath = gitExe
        End If
    End If
    
    ' bad directory check
    If FileOrDirExists(gitExe) = False Then
        MsgBox "Cannot find git executable: " & gitExe
        Exit Function
    End If
    
    ' add quotes if spaces in the path
    If quoteGitExe And InStr(1, gitExe, " ") Then
        gitExe = """" & gitExe & """"
    End If
    
    GetGitExe = gitExe
End Function


Private Function GetWorkingDir() As String
    GetWorkingDir = ""
    
    ' get the working directory path
    Dim workingDir As String
    workingDir = ShibbySettings.GitProjectPath
    
    ' not found in doc props, browse for one
    If workingDir = "" Then
        MsgBox "Please set the git Project Path"
        workingDir = FileUtils.FolderBrowser
        If (workingDir = "") Then
            Exit Function
        Else
            ShibbySettings.GitProjectPath = workingDir
        End If
    End If
    
    ' bad directory check
    If FileOrDirExists(workingDir) = False Then
        MsgBox "Cannot find project folder: " & workingDir
        Exit Function
    End If
    
    GetWorkingDir = workingDir
End Function

Private Function GitPathOption(ByVal path As String) As String
    If InStr(1, path, " ") Then
        GitPathOption = " -C """ & path & """ "
    Else
        GitPathOption = " -C " & path
    End If
    
End Function
