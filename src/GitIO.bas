Attribute VB_Name = "GitIO"
Option Explicit
Private Const MODULEFOLDER As String = "modules"
Private Const CLASSFOLDER As String = "classModules"
Private Const FORMFOLDER As String = "forms"
Private Const SOURCEFOLDER As String = "src"

' Export all code modules to git directory
' based on the selected file structure
Private Function GitExportAll() As String
    
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
    
    Dim compInd As Integer
    Dim filesWritten As String
    Dim nextFile As String
    Dim nComps As Integer
    nComps = Application.VBE.VBProjects.Item(projectInd).VBComponents.Count
    
    For compInd = 1 To nComps
        nextFile = ExportToProperFolder(projectInd, compInd, gitDir)
        filesWritten = filesWritten & vbCrLf & nextFile
    Next compInd
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    GitExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & gitDir & vbCrLf & filesWritten

End Function


' return the correct file extension based on the type of module
' module type constants defined in CodeUtils
Private Function GetExtensionFromModuleType(ByVal codeType As Integer)
    Dim extension As String
    Select Case codeType
       Case CodeUtils.ClassModule
           extension = ".cls"
       Case CodeUtils.form
           extension = ".frm"
       Case CodeUtils.Module
           extension = ".bas"
    End Select
    GetExtensionFromModuleType = extension
End Function


' export one module to the proper directory
' input: projectInd - the index of the desired VBProject in Application.VBE.VBProjects
' input: compInd - the index of the desired component in project.VBComponents
' input: gitDir - the root directory of the export
' output: the path of the output file, relative to gitDir
Private Function ExportToProperFolder(ByVal projectInd As Integer, ByVal compInd As Integer, ByVal gitDir As String)
    With Application.VBE.VBProjects.Item(projectInd).VBComponents.Item(compInd)
        
        Dim extension As String
        extension = GetExtensionFromModuleType(.Type)

        If (extension <> "") Then
            Dim file As String
            file = SOURCEFOLDER & "\"
            
            ' flat file structure
            If ShibbySettings.FileStructure = Flat Then
                file = .name & extension
                .Export (gitDir & "\" & file)
                
            ' simple source folder structure
            ElseIf ShibbySettings.FileStructure = SimpleSrc Then
                file = file & .name & extension
                .Export (gitDir & "\" & file)
                
            ' separated source folder structure
            Else
                Select Case .Type
                Case ClassModule
                    file = file & CLASSFOLDER
                Case form
                    file = file & FORMFOLDER
                Case Module
                    file = file & MODULEFOLDER
                End Select
                file = file & "\" & .name & extension
                .Export (gitDir & "\" & file)
            End If
         End If
    End With
    ExportToProperFolder = file
End Function


' Check for existence of required code folders based
' on the file structure type. Create if necessary
Private Sub CheckCodeFolders(ByVal gitDir As String)
    ' create folders if needed
    If ShibbySettings.FileStructure <> Flat Then
        If dir(gitDir & "\" & SOURCEFOLDER & "\") = "" Then
            MkDir gitDir & "\" & SOURCEFOLDER & "\"
        End If
        If ShibbySettings.FileStructure = SeparatedSrc Then
            If dir(gitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\") = "" Then
                MkDir gitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\"
            End If
            If dir(gitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\") = "" Then
                MkDir gitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\"
            End If
            If dir(gitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\") = "" Then
                MkDir gitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\"
            End If
        End If
    End If
End Sub
