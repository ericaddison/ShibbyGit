Attribute VB_Name = "GitCommands"
Public Const EXE_PATH_PROPERTY As String = "code_GitExecutablePath"
Public Const PROJECT_PATH_PROPERTY As String = "code_GitProjectPath"
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems

' testing push origin master for web remote
Public Sub GitRemotes()

    Debug.Print "Git Remotes Dev"

    ' get the git executable path
    Dim gitExe As String
    gitExe = GetSetting(CodeUtils.APPNAME, "FileInfo", Git.EXE_PATH_PROPERTY, "")
 
     ' get the working directory path
    Dim workingDir As String
    workingDir = DocPropIO.GetItemFromDocProperties(PROJECT_PATH_PROPERTY)
    
 
    ' crate the parameter string
    Dim command As String
    command = "cmd /c "
    If InStr(1, gitExe, " ") Then
        command = command & """" & gitExe & """ "
    Else
        command = command & gitExe & " "
    End If
    
    If InStr(1, workingDir, " ") Then
        command = command & "-C """ & workingDir & """ "
    Else
        command = command & "-C " & workingDir & " "
    End If
    
    Dim file As String
    file = workingDir & "\tmp.txt"
    
    command = command & " push origin master & pause"
    Shell command, 1

    
End Sub


Public Sub GitCommit(ByVal message As String)
    Dim out As String
    out = GitOther("commit -am """ & message & """")
    MsgBox out
End Sub

Public Sub GitStatus()
    Dim out As String
    out = GitOther("status")
    MsgBox out
End Sub


Public Sub GitLog()
    Dim out As String
    out = GitOther("log")
    MsgBox out
End Sub

Public Sub GitAddAll()
    Dim out As String
    out = GitOther("add -A")
    If out = "" Then
        MsgBox "Staged all files for commit"
    Else
        MsgBox "Git response: " & vbCrLf & out
    End If
End Sub


' Main function to call git ... calls "git -C <path> options",
' where the options come from the incoming string, and the
' path is from the export directory
Public Function GitOther(ByVal options As String) As String
    
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
        
    ' crate the parameter string
    Dim parms As String
    parms = "-C """ & workingDir & """ " & options
    
    ' call git
    Dim output As String
    output = ShellRedirect.Redirect(gitExe, parms, 1000)
    
    GitOther = output
End Function
