Attribute VB_Name = "CodeUtils"
' Any functions to help with the actual coding process
Option Explicit

Public Const EXPORT_DIRECTORY_PROPERTY As String = "code_ExportDirectory"

' component type constants
Public Const Module As Integer = 1
Public Const ClassModule As Integer = 2
Public Const form As Integer = 3
Public Const Document As Integer = 100
Public Const Padding As Integer = 24

Private pFolder As String

'****************************************************
' Public functions
'****************************************************

' public interface for export all, no msg box
' input: folder - the folder to export code modules to
' output: String with list of modules exported
Public Function ExportAllString(ByVal folder As String) As String
    pFolder = folder
    ExportAllString = ExportAll
End Function


' public interface for import all, no msg box
' input: folder - the folder to import code modules from
' output: String with list of modules imported
Public Function ImportAllString(ByVal folder As String) As String
    pFolder = folder
    ImportAllString = ImportAll
End Function


' public interface for export all
' input: folder - the folder to export code modules to
Public Sub ExportAllMsgBox()
    pFolder = FileUtils.FolderBrowser("Browse for folder for export")
    If pFolder = "" Then
        Exit Sub
    End If
    UI.NonModalMsgBox "Exporting files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = ExportAll
    
    UI.HideNonModalMsgBox
    MsgBox output
End Sub

' public interface for import all
' input: folder - the folder to import code modules from
Public Sub ImportAllMsgBox()
    pFolder = FileUtils.FolderBrowser("Browse for code folder to import")
    If pFolder = "" Then
        Exit Sub
    End If
    UI.NonModalMsgBox "Importing files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = ImportAll
    
    UI.HideNonModalMsgBox
    MsgBox output
End Sub


' Find the index of the VBProject corresponding to
' the active presentation
' output: the index of the VBProject corresponding to the active filename
'           -1 if not found
Public Function FindFileVBProject(Optional ByVal fileName As String = "") As Integer
    If fileName = "" Then
        fileName = ActivePresentation.FullName
    End If

    With Application
        Dim ind As Integer
        For ind = 1 To .VBE.VBProjects.Count
            Dim VBFileName As String
            On Error Resume Next
                VBFileName = .VBE.VBProjects.Item(ind).fileName
            On Error GoTo 0
            If VBFileName = fileName Then
                FindFileVBProject = ind
                Exit Function
            End If
        Next ind
    End With
    
    MsgBox "Could not find VB Project associated with open file: " & fileName _
            & vbCrLf & "Is this a new, unsaved presentation?"
    FindFileVBProject = -1
End Function


' remove a code module and import
' input: projectInd - the index of the desired VBProject in Application.VBE.VBProjects
' input: file - the full path to the file for import
' output: the name of the imported module, or "" if none
Public Function RemoveAndImportModule(ByVal projectInd As Integer, ByVal file As String) As String
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
        If CheckCodeType(file) <> -1 Then
            Dim ModuleName As String
            ModuleName = FileBaseName(file)
            
            ' don't import modules with running code!
            If ModuleName = "NonModalMsgBoxForm" Or ModuleName = "CodeUtils" Then
                Exit Function
            End If
            
            ' check if module already exists in project
            Dim moduleExists As Boolean
            On Error Resume Next
                .Item (ModuleName)
                moduleExists = (err = 0)
                err.Clear
            On Error GoTo 0
            
            ' rename and remove
            If moduleExists Then
                If CheckCodeType(file) = 3 Then
                    .Remove .Item(ModuleName)
                Else
                    .Item(ModuleName).name = ModuleName & "R"
                    .Remove .Item(ModuleName & "R")
                End If
                DoEventsAndWait 10, 2
            End If
            
            ' import new module
            Dim newModule As Variant
            Set newModule = .Import(file)
            RemoveAndImportModule = newModule.name
        End If
    End With
End Function


' return a Module type based on the file extension
' input: file - filename of a code module
' output: integer corresponding to module type
Public Function CheckCodeType(ByVal file As String) As Integer

    If file Like "*.bas" Then
        CheckCodeType = Module
    ElseIf file Like "*.frm" Then
        CheckCodeType = form
    ElseIf file Like "*.cls" Then
        CheckCodeType = ClassModule
    Else
        CheckCodeType = -1
    End If

End Function


'****************************************************
' Private functions
'****************************************************

Private Function ExportAll() As String

    ' write files
    Dim projectInd As Integer
    projectInd = FindFileVBProject
    If projectInd = -1 Then
        ExportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.name
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
                .Item(ind).Export (pFolder & "\" & .Item(ind).name & extension)
                filesWritten = filesWritten & vbCrLf & .Item(ind).name & extension
            End If
        Next ind
    
    End With
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    ExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & pFolder & vbCrLf & filesWritten

End Function


Private Function ImportAll() As String

    ' get project index from active file name
    Dim projectInd As Integer
    projectInd = FindFileVBProject
    If projectInd = -1 Then
        ImportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.name
        Exit Function
    End If

    ' import files
    Dim file As String
    Dim ModuleName As String
    Dim filesRead As String
    file = dir(pFolder & "\")
    While file <> ""
        ModuleName = RemoveAndImportModule(projectInd, pFolder & "\" & file)
        If ModuleName <> "" Then
            filesRead = filesRead & vbCrLf & ModuleName
        End If
        file = dir
    Wend


    ImportAll = "ShibbyGit Modules Loaded: " & filesRead

End Function
