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
    With Application.VBE.ActiveVBProject.VBComponents
        Dim file As String
        file = dir(srcFolder & "\modules\")
        On Error GoTo LoadError
            While file <> ""
                If file Like "*.bas" Then
                    .Import (srcFolder & "\modules\" & file)
                End If
                file = dir
            Wend
        
        file = dir(srcFolder & "\forms\")
        On Error Resume Next
            While file <> ""
                If file Like "*.frm" Then
                    .Import (srcFolder & "\forms\" & file)
                End If
                file = dir
            Wend

        file = dir(srcFolder & "\classModules\")
        On Error Resume Next
            While file <> ""
                If file Like "*.cls" Then
                    .Import (srcFolder & "\classModules\" & file)
                End If
                file = dir
            Wend
        On Error GoTo 0
    
        .Remove .Item("BootStrapLoader1")
    End With
    
    Exit Sub
LoadError:
    MsgBox "Error bootstrapping ShibbyGit: " & Err.Number & vbCrLf & _
        "You may need to ""trust access to the VBA project object model"" in " & vbCrLf & _
        "File->Options->Trust Center->Trust Center Settings->Macro Settings"
    Err.Clear
    Exit Sub
End Sub



