Attribute VB_Name = "Git"
Public Const GIT_PATH_PROPERTY As String = "code_GitExecutablePath"


Public Sub GitCommit(ByVal message As String)
    GitOther ("commit -am """ & message & """")
End Sub

Public Sub GitStatus()
    GitOther ("status")
End Sub


Public Sub GitLog()
    GitOther ("log")
End Sub

Public Sub GitAddAll()
    GitOther ("add -A")
End Sub


' Main function to call git ... calls "git -C <path> options",
' where the options come from the incoming string, and the
' path is from the export directory
Public Sub GitOther(ByVal options As String)
    
    ' get the git executable path
    Dim gitExe As String
    gitExe = GetSetting("CVX_CodeUtils", "FileInfo", Git.GIT_PATH_PROPERTY, "")
    
    If gitExe = "" Or IsNull(gitExe) Then
        MsgBox "Please set the git executable path"
        Exit Sub
    End If
    
    ' bad directory check
    If FileOrDirExists(gitExe) = False Then
        MsgBox "Cannot find git executable: " & gitExe
        Exit Sub
    End If
    
    ' get the working directory path
    Dim workingDir As String
    workingDir = DocPropIO.GetItemFromDocProperties(CodeUtils.EXPORT_DIRECTORY_PROPERTY)
    
    ' not found in doc props, browse for one
    If workingDir = "" Then
        workingDir = UI.FolderDialog
    End If
    
    ' browse cancelled, exit
    If (workingDir = "") Then
        Exit Sub
    End If
        
    ' bad directory check
    If FileOrDirExists(workingDir) = False Then
        MsgBox "Cannot find folder: " & workingDir
        Exit Sub
    End If
        
    ' crate the parameter string
    Dim parms As String
    parms = "-C """ & workingDir & """ " & options
    
    ' call git
    Dim output As String
    output = ShellRedirect.Redirect(gitExe, parms)
    
    MsgBox output
End Sub

