Attribute VB_Name = "BootStrapLoader"
' self-contained function to load ShibbyGit
' Run this routine once to load the ShibbyGit source
' then forget about it


Public Sub LoadShibbyGitCode()

    ' folder dialog to find source folder
    Dim fd As FileDialog
    Dim srcFolder As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Browse to ShibbyGit Source Folder"
    If fd.Show = -1 Then
        srcFolder = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' import files
    Dim file As String
    file = dir(srcFolder & "\")
    On Error Resume Next
        While file <> ""
            If file Like "*.bas" Or file Like "*.cls" Or file Like "*.frm" Then
                Application.VBE.ActiveVBProject.VBComponents.Import (srcFolder & "\" & file)
            End If
            file = dir
        Wend
    On Error GoTo 0

End Sub




