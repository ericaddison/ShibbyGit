Attribute VB_Name = "BootStrapLoader"
' self-contained function to load ShibbyGit
' Run this routine once to load the ShibbyGit source
' then forget about it


Public Sub LoadShibbyGitCode()

    MsgBox "Please Browse to and select the ShibbyGit src folder in the upcoming file browser"

    ' folder dialog to find source folder
    Dim fd As FileDialog
    Dim srcFolder As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Browse to ShibbyGit src Folder"
    If fd.Show = -1 Then
        srcFolder = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' import files
    Dim file As String
    file = dir(srcFolder & "\modules\")
    On Error Resume Next
        While file <> ""
            Debug.Print file
            If file Like "*.bas" Then
                Application.VBE.ActiveVBProject.VBComponents.Import (srcFolder & "\modules\" & file)
            End If
            
            If err.Number = 1004 Then
                MsgBox "You must ""trust access to the VBA project object model"" in " & vbCrLf & _
                    "File->Options->Trust Center->Trust Center Settings->Macro Settings"
                Exit Sub
            End If
            
            file = dir
        Wend
    On Error GoTo 0
    
    file = dir(srcFolder & "\forms\")
    On Error Resume Next
        While file <> ""
            If file Like "*.frm" Then
                Application.VBE.ActiveVBProject.VBComponents.Import (srcFolder & "\forms\" & file)
            End If
            file = dir
        Wend
    On Error GoTo 0
    
    file = dir(srcFolder & "\classModules\")
    On Error Resume Next
        While file <> ""
            If file Like "*.cls" Then
                Application.VBE.ActiveVBProject.VBComponents.Import (srcFolder & "\classModules\" & file)
            End If
            file = dir
        Wend
    On Error GoTo 0

End Sub







