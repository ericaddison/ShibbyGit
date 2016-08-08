Attribute VB_Name = "GitCommands"
Option Explicit
Public Const EXE_PATH_PROPERTY As String = "code_GitExecutablePath"
Public Const PROJECT_PATH_PROPERTY As String = "code_GitProjectPath"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems


Public Sub GitCommit(ByVal message As String)
    Dim out As String
    out = RunGitAsProcess("commit -am """ & message & """")
    MsgBox out
End Sub

Public Sub GitStatus()
    Dim out As String
    out = RunGitAsProcess("status", 1500)
    Debug.Print "status out = " & out
    MsgBox out
End Sub


Public Sub GitLog()
    Dim out As String
    out = RunGitAsProcess("log --max-count=5")
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
Public Sub RunGitInShell(ByVal options As String)
    Dim command As String
    command = "cmd /c echo Running 'git " & options & "'" & _
    " & " & GitCommands.GitExeWithPath & " " & options & " & pause"
    Shell command, 1
End Sub


' Run git with specified options ... calls "git -C <path> options",
' Launches a new process and returns the output
' path to git executable and project directory come from settings
' VBA will wait "waitTime" milliseconds for the process to complete. Default value is 10000 for 10s
Public Function RunGitAsProcess(ByVal options As String, Optional ByVal waitTime As Long = 10000) As String

    Dim gitExe As String
    gitExe = GetGitExe
    
    Dim workingDir As String
    workingDir = GetWorkingDir

    ' crate the parameter string
    Dim parms As String
    parms = " -C """ & workingDir & """ " & options
    
    ' call git
    Dim output As String
    output = ShellRedirect.Redirect(gitExe, parms, 1500)
    
    RunGitAsProcess = output
End Function


' get the gitExe path with the -C working directory parameter added
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
    
    ' crate the parameter string
    Dim command As String
    If InStr(1, git, " ") Then
        command = """" & git & """"
    Else
        command = git
    End If
    
    If InStr(1, workingDir, " ") Then
        command = command & " -C """ & workingDir & """ "
    Else
        command = command & " -C " & workingDir & " "
    End If

    GitExeWithPath = command
    
End Function

Private Function GetGitExe() As String
    GetGitExe = ""

    ' get the git executable path
    Dim gitExe As String
    gitExe = GetSetting(CodeUtils.APPNAME, "FileInfo", GitCommands.EXE_PATH_PROPERTY, "")
    
    If gitExe = "" Or IsNull(gitExe) Then
        MsgBox "Please set the git executable path"
        Exit Function
    End If
    
    ' bad directory check
    If FileOrDirExists(gitExe) = False Then
        MsgBox "Cannot find git executable: " & gitExe
        Exit Function
    End If
    
    GetGitExe = gitExe
End Function


Private Function GetWorkingDir() As String
    GetWorkingDir = ""
    
    ' get the working directory path
    Dim workingDir As String
    workingDir = DocPropIO.GetItemFromDocProperties(PROJECT_PATH_PROPERTY)
    
    ' not found in doc props, browse for one
    If workingDir = "" Then
        workingDir = UI.FolderDialog
    End If
    
    ' browse cancelled, exit
    If (workingDir = "") Then
        Exit Function
    End If
        
    ' bad directory check
    If FileOrDirExists(workingDir) = False Then
        MsgBox "Cannot find folder: " & workingDir
        Exit Function
    End If
    
    GetWorkingDir = workingDir
End Function
