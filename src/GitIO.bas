Attribute VB_Name = "GitIO"
Option Explicit


Public Function GitExportAll() As String
    
    ' get the export directory
    Dim gitDir As String
    gitDir = ShibbySettings.GitProjectPath
    
    ' not found in doc props, browse for one
    If gitDir = "" Then
        gitDir = UI.FolderDialog
        ShibbySettings.GitProjectPath = gitDir
        If (gitDir = "") Then
            Exit Function
        End If
    End If
    
    ' bad directory
    If FileOrDirExists(gitDir) = False Then
        MsgBox "Cannot find folder: " & gitDir
        Exit Function
    End If
    
    ' create folders if needed
    CheckCodeFolders gitDir
    
    ' write files
    Dim projectInd As Integer
    projectInd = CodeUtils.FindActiveFileVBProject
    If projectInd = -1 Then
        GitExportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.name
        Exit Function
    End If
    
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
        Dim ind As Integer
        Dim filesWritten As String
        Dim extension As String
        For ind = 1 To .Count
            extension = ""
            Select Case .Item(ind).Type
               Case ClassModule
                   extension = ".cls"
               Case form
                   extension = ".frm"
               Case Module
                   extension = ".bas"
            End Select

            If (extension <> "") Then
                If ShibbySettings.FileStructure = Flat Then
                    .Item(ind).Export (gitDir & "\" & .Item(ind).name & extension)
                    filesWritten = filesWritten & vbCrLf & .Item(ind).name & extension
                ElseIf ShibbySettings.FileStructure = SimpleSrc Then
                    .Item(ind).Export (gitDir & "\src\" & .Item(ind).name & extension)
                    filesWritten = filesWritten & vbCrLf & "src\" & .Item(ind).name & extension
                Else
                    Dim subFolder As String
                    Select Case .Item(ind).Type
                    Case ClassModule
                        subFolder = "classModules"
                    Case form
                        subFolder = "forms"
                    Case Module
                        subFolder = "modules"
                    End Select
                    .Item(ind).Export (gitDir & "\src\" & subFolder & "\" & .Item(ind).name & extension)
                    filesWritten = filesWritten & vbCrLf & "src\" & subFolder & "\" & .Item(ind).name & extension
                End If
                
            End If
        Next ind
    
    End With
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    GitExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & gitDir & vbCrLf & filesWritten

End Function



Private Sub CheckCodeFolders(ByVal gitDir As String)
    ' create folders if needed
    If ShibbySettings.FileStructure <> Flat Then
        If dir(gitDir & "\src\") = "" Then
            MkDir gitDir & "\src\"
        End If
        If ShibbySettings.FileStructure = SeparatedSrc Then
            If dir(gitDir & "\src\modules\") = "" Then
                MkDir gitDir & "\src\modules\"
            End If
            If dir(gitDir & "\src\forms\") = "" Then
                MkDir gitDir & "\src\forms\"
            End If
            If dir(gitDir & "\src\classModules\") = "" Then
                MkDir gitDir & "\src\classModules\"
            End If
        End If
    End If
End Sub
