Attribute VB_Name = "CodeUtils"
' Any functions to help with the actual coding process
Option Explicit

Public Const EXPORT_DIRECTORY_PROPERTY As String = "code_ExportDirectory"
Public Const APPNAME As String = "ShibbyGit"
Private Const OldTag As String = "O"


Public Sub ExportAllMsgBox()
    MsgBox ExportAll
End Sub

Public Sub ImportAllMsgBox()
    MsgBox ImportAll
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
        ExportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.Name
        Exit Function
    End If
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
    
        Dim ind As Integer
        Dim filesWritten As String
        Dim extension As String
        For ind = 1 To .Count
            extension = ""
            Select Case .Item(ind).Type
               Case .Item("dummyClass").Type
                   extension = ".cls"
               Case .Item("dummyForm").Type
                   extension = ".frm"
               Case .Item("dummyModule").Type
                   extension = ".bas"
            End Select
            
            If (extension <> "") Then
                .Item(ind).Export (exportDir & "\" & .Item(ind).Name & extension)
                filesWritten = filesWritten & vbCrLf & .Item(ind).Name & extension
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
    
    ' import files
    Dim projectInd As Integer
    projectInd = FindActiveFileVBProject
    If projectInd = -1 Then
        ExportAll = "Uh oh! Could not find VBProject associated with " & ActivePresentation.Name
        Exit Function
    End If
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
    

        ' first loop through files and delete modules to be imported
        Dim file As String
        Dim ModuleName As String
        Dim filesRead As String
        file = dir(importDir & "\")
        While file <> ""
            If CheckCodeType(file) <> -1 And file <> "CodeUtils.bas" Then
                On Error Resume Next
                    ModuleName = FileBaseName(file)
                    .Item(ModuleName).Name = ModuleName & OldTag
                    .Remove .Item(ModuleName & OldTag)
                    .Import importDir & "\" & file
                On Error GoTo 0
                filesRead = filesRead & vbCrLf & ModuleName
            End If
            file = dir
        Wend
        
    End With

    ImportAll = "ShibbyGit Modules Loaded: " & filesRead

End Function

Private Function CheckCodeType(ByVal file As String) As Integer

    With Application.VBE.ActiveVBProject.VBComponents
        
        If file Like "*.bas" Then
            CheckCodeType = .Item("dummyModule").Type
        ElseIf file Like "*.frm" Then
            CheckCodeType = .Item("dummyForm").Type
        ElseIf file Like "*.cls" Then
            CheckCodeType = .Item("dummyClass").Type
        Else
            CheckCodeType = -1
        End If
        
    End With
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
Select Case Err.Number
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
            If .VBE.VBProjects.Item(ind).FileName = .ActivePresentation.FullName Then
                FindActiveFileVBProject = ind
                Exit Function
            End If
        Next ind
    End With
    FindActiveFileVBProject = -1
End Function

