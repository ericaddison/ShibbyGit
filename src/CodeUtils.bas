Attribute VB_Name = "CodeUtils"
' Any functions to help with the actual coding process
Option Explicit

Public Const EXPORT_DIRECTORY_PROPERTY As String = "code_ExportDirectory"
Public Const APPNAME As String = "ShibbyGit"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' component type constants
Private Const Module As Integer = 1
Private Const ClassModule As Integer = 2
Private Const form As Integer = 3
Private Const Document As Integer = 100
Private Const Padding As Integer = 24


Public Sub ExportAllMsgBox()
    Dim output As String
    UI.NonModalMsgBox "Exporting files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    DoEventsAndWait 10, 2
    output = ExportAll
    NonModalMsgBoxForm.Hide
    MsgBox output
End Sub

Public Sub ImportAllMsgBox()
    Dim output As String
    UI.NonModalMsgBox "Importing files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    DoEventsAndWait 10, 2
    output = ImportAll
    NonModalMsgBoxForm.Hide
    MsgBox output
End Sub



Public Function ExportAll() As String
    
    ' get the export directory
    Dim exportDir As String
    exportDir = DocPropIO.GetItemFromDocProperties(EXPORT_DIRECTORY_PROPERTY)
    
    ' not found in doc props, browse for one
    If exportDir = "" Then
        exportDir = UI.FolderDialog
    End If
    
    ' browse cancelled, exit
    If (exportDir = "") Then
        Exit Function
    End If
    
    ' bad directory
    If FileOrDirExists(exportDir) = False Then
        MsgBox "Cannot find folder: " & exportDir
        Exit Function
    End If
    
    ' write files
    Dim projectInd As Integer
    projectInd = FindActiveFileVBProject
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
                .Item(ind).Export (exportDir & "\" & .Item(ind).name & extension)
                filesWritten = filesWritten & vbCrLf & .Item(ind).name & extension
            End If
        Next ind
    
    End With
     
    ExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & exportDir & vbCrLf & filesWritten

End Function


Public Function ImportAll() As String

    ' get the export directory
    Dim importDir As String
    importDir = DocPropIO.GetItemFromDocProperties(EXPORT_DIRECTORY_PROPERTY)
    
    ' not found in doc props, browse for one
    If importDir = "" Then
        importDir = UI.FolderDialog
    End If
    
    ' browse cancelled, exit
    If (importDir = "") Then
        Exit Function
    End If
    
    ' bad directory check
    If FileOrDirExists(importDir) = False Then
        MsgBox "Cannot find folder: " & importDir
        Exit Function
    End If
    
    ' get project index from active file name
    Dim projectInd As Integer
    projectInd = FindActiveFileVBProject
    If projectInd = -1 Then
        ImportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.name
        Exit Function
    End If

    ' import files
    Dim file As String
    Dim ModuleName As String
    Dim filesRead As String
    file = dir(importDir & "\")
    While file <> ""
        ModuleName = RemoveAndImportModule(projectInd, importDir & "\" & file)
        If ModuleName <> "" Then
            filesRead = filesRead & vbCrLf & ModuleName
        End If
        file = dir
    Wend


    ImportAll = "ShibbyGit Modules Loaded: " & filesRead

End Function

Private Function CheckCodeType(ByVal file As String) As Integer

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


Private Function FileBaseName(ByVal file As String) As String
    FileBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(file)
End Function

'used to test filepaths of commmand button   links to see if they work - change their color if not working
' from http://superuser.com/questions/649745/check-if-path-to-file-is-correct-on-excel-column
Public Function FileOrDirExists(PathName As String) As Boolean
    'Macro Purpose: Function returns TRUE if the specified file
    Dim iTemp As Integer
    
    'Ignore errors to allow for error evaluation
    On Error Resume Next
        iTemp = GetAttr(PathName)
        
        'Check if error exists and set response appropriately
        Select Case err.Number
            Case Is = 0
                FileOrDirExists = True
            Case Else
                FileOrDirExists = False
        End Select
    
    'Resume error checking
    On Error GoTo 0
End Function


' Find the index of the VBProject corresponding to
' the active presentation
Private Function FindActiveFileVBProject() As Integer

    With Application
        Dim ind As Integer
        For ind = 1 To .VBE.VBProjects.Count
            Dim VBFileName As String
            On Error Resume Next
                VBFileName = .VBE.VBProjects.Item(ind).FileName
            On Error GoTo 0
            If VBFileName = .ActivePresentation.FullName Then
                FindActiveFileVBProject = ind
                Exit Function
            End If
        Next ind
    End With
    MsgBox "Could not find VB Project associated with open file: " & ActivePresentation.FullName _
            & vbCrLf & "Is this a new, unsaved presentation?"
    FindActiveFileVBProject = -1
End Function


Private Function RemoveAndImportModule(ByVal projectInd As Integer, ByVal file As String) As String
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


Private Sub DoEventsAndWait(ByVal nLoops As Integer, ByVal sleepTimeMs As Integer)
    Dim ind As Integer
    For ind = 1 To nLoops
        DoEvents
        Sleep sleepTimeMs
    Next ind
End Sub
