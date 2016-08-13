Attribute VB_Name = "ShibbySettings"
Option Explicit

Private Const APPNAME As String = "ShibbyGit"
Private Const EXE_PATH_PROPERTY As String = "shibby_GitExecutablePath"
Private Const PROJECT_PATH_PROPERTY As String = "shibby_GitProjectPath"
Private Const FRX_CLEANUP_PROPERTY As String = "shibby_FrxCleanup"
Private Const EXPORT_ON_GIT_PROPERTY As String = "shibby_ExportOnGit"
Private Const FILESTRUCTURE_PROPERTY As String = "shibby_FileStructure"
Public Enum ShibbyFileStructure
    flat = 0
    SimpleSrc = 1
    SeparatedSrc = 2
End Enum

'***************************************************************
' App dependent info
'***************************************************************

Public Function GetProjectFileName() As String
    Dim name As String
    name = Application.name
    
    Dim app As Object
    Set app = Application
    
    Select Case name
        Case "Microsoft PowerPoint"
            GetProjectFileName = app.ActivePresentation.FullName
        Case "Microsoft Excel"
            GetProjectFileName = app.ActiveSheet.FullName
        Case "Microsoft Word"
            GetProjectFileName = app.ActiveDocument.FullName
      End Select
    
End Function

Public Function GetProjectName() As String
    Dim name As String
    name = Application.name
    
    Dim app As Object
    Set app = Application
    
    Select Case name
        Case "Microsoft PowerPoint"
            GetProjectName = app.ActivePresentation.name
        Case "Microsoft Excel"
            GetProjectName = app.ActiveSheet.name
        Case "Microsoft Word"
            GetProjectName = app.ActiveDocument.name
      End Select
End Function


'***************************************************************
' Property accessors
'***************************************************************


' get the git exe path
Public Property Get GitExePath() As String
    GitExePath = GetSetting(APPNAME, "FileInfo", EXE_PATH_PROPERTY, "")
End Property

' set the git exe path
Public Property Let GitExePath(ByVal newPath As String)
    Call SaveSetting(APPNAME, "FileInfo", EXE_PATH_PROPERTY, newPath)
End Property

' get the Git Project path
Public Property Get GitProjectPath() As String
    GitProjectPath = DocPropIO.GetItemFromDocProperties(PROJECT_PATH_PROPERTY)
End Property

' set the git project path
Public Property Let GitProjectPath(ByVal newPath As String)
    DocPropIO.AddStringToDocProperties PROJECT_PATH_PROPERTY, newPath
End Property

' get the FrxCleanup setting
Public Property Get FrxCleanup() As Boolean
    FrxCleanup = DocPropIO.GetBooleanFromDocProperties(FRX_CLEANUP_PROPERTY)
End Property

' set the FrxCleanup setting
Public Property Let FrxCleanup(ByVal newVal As Boolean)
    DocPropIO.AddBooleanToDocProperties FRX_CLEANUP_PROPERTY, newVal
End Property

' get the export on save setting
Public Property Get ExportOnGit() As Boolean
    ExportOnGit = DocPropIO.GetBooleanFromDocProperties(EXPORT_ON_GIT_PROPERTY)
End Property

' set the git project path
Public Property Let ExportOnGit(ByVal newVal As Boolean)
    DocPropIO.AddBooleanToDocProperties EXPORT_ON_GIT_PROPERTY, newVal
End Property

' get the export on save setting
Public Property Get fileStructure() As ShibbyFileStructure
    Dim fs As Variant
    fs = DocPropIO.GetItemFromDocProperties(FILESTRUCTURE_PROPERTY)
    If fs = "" Then
        fileStructure = flat
    Else
        fileStructure = fs
    End If
End Property

' set the git project path
Public Property Let fileStructure(ByRef newVal As ShibbyFileStructure)
    DocPropIO.AddNumberToDocProperties FILESTRUCTURE_PROPERTY, newVal
End Property
