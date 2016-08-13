Attribute VB_Name = "GitIO"
Option Explicit
Private Const MODULEFOLDER As String = "modules"
Private Const CLASSFOLDER As String = "classModules"
Private Const FORMFOLDER As String = "forms"
Private Const SOURCEFOLDER As String = "src"
Private pFileStructure As ShibbyFileStructure
Private pGitDir As String
Private pProjectInd As Integer


'****************************************************
' Public functions
'****************************************************

Public Sub test()
    Debug.Print GitExport(ShibbySettings.GitProjectPath, SimpleSrc)
End Sub


' Public entry point for Git Import
Public Function GitImport(ByVal gitDir As String, ByVal fileStructure As ShibbyFileStructure) As String
    pFileStructure = fileStructure
    pGitDir = gitDir
    GitImport = GitImportAll
End Function


' Public entry point for Git Export
Public Function GitExport(ByVal gitDir As String, ByVal fileStructure As ShibbyFileStructure) As String
    pFileStructure = fileStructure
    pGitDir = gitDir
    GitExport = GitExportAll
End Function


'****************************************************
' Private functions
'****************************************************

' check that the incoming folder is valid
Private Function CheckGitFolder() As Boolean
    CheckGitFolder = True
    ' check the incoming folder
    Dim folderCheck As String
    folderCheck = FileUtils.VerifyFolder(pGitDir)
    If folderCheck = FileUtils.BAD_FOLDER Then
        CheckGitFolder = False
        Exit Function
    ElseIf folderCheck <> FileUtils.GOOD_FOLDER Then
        pGitDir = folderCheck
        ShibbySettings.GitProjectPath = pGitDir
    End If
End Function


' Import all code modules from git directory
' based on the selected file structure
Private Function GitImportAll() As String

    If Not CheckGitFolder Then
        Exit Function
    End If
    
    ' import files
    Dim file As String
    Dim ModuleName As String
    Dim filesRead As String
    file = dir(pGitDir & "\")
    While file <> ""
        ModuleName = RemoveAndImportModule(pProjectInd, pGitDir & "\" & file)
        If ModuleName <> "" Then
            filesRead = filesRead & vbCrLf & ModuleName
        End If
        file = dir
    Wend


    GitImportAll = "ShibbyGit Modules Loaded: " & filesRead

End Function



' Export all code modules to git directory
' based on the selected file structure
Private Function GitExportAll() As String
    
    If Not CheckGitFolder Then
        Exit Function
    End If
    
    
    ' create folders if needed
    CheckCodeFolders
    
    ' write files
    pProjectInd = CodeUtils.FindFileVBProject
    If pProjectInd = -1 Then
        GitExportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.name
        Exit Function
    End If
    
    Dim compInd As Integer
    Dim filesWritten As String
    Dim nextFile As String
    Dim nComps As Integer
    nComps = Application.VBE.VBProjects.Item(pProjectInd).VBComponents.Count
    
    For compInd = 1 To nComps
        nextFile = ExportToProperFolder(compInd)
        filesWritten = filesWritten & vbCrLf & nextFile
    Next compInd
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    GitExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & pGitDir & vbCrLf & filesWritten

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
' input: compInd - the index of the desired component in project.VBComponents.Item(pProjectInd)
' output: the path of the output file, relative to pGitDir
Private Function ExportToProperFolder(ByVal compInd As Integer)
    With Application.VBE.VBProjects.Item(pProjectInd).VBComponents.Item(compInd)
        
        Dim extension As String
        extension = GetExtensionFromModuleType(.Type)

        If (extension <> "") Then
            Dim file As String
            file = SOURCEFOLDER & "\"
            
            ' flat file structure
            If pFileStructure = flat Then
                file = .name & extension
                .Export (pGitDir & "\" & file)
                
            ' simple source folder structure
            ElseIf pFileStructure = SimpleSrc Then
                file = file & .name & extension
                .Export (pGitDir & "\" & file)
                
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
                .Export (pGitDir & "\" & file)
            End If
         End If
    End With
    ExportToProperFolder = file
End Function


' Check for existence of required code folders based
' on the file structure type. Create if necessary
Private Sub CheckCodeFolders()
    ' create folders if needed
    If pFileStructure <> flat Then
        If dir(pGitDir & "\" & SOURCEFOLDER & "\") = "" Then
            MkDir pGitDir & "\" & SOURCEFOLDER & "\"
        End If
        If pFileStructure = SeparatedSrc Then
            If dir(pGitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\") = "" Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\"
            End If
            If dir(pGitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\") = "" Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\"
            End If
            If dir(pGitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\") = "" Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\"
            End If
        End If
    End If
End Sub
