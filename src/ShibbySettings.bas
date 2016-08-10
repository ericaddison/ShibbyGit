Attribute VB_Name = "ShibbySettings"
Option Explicit

Private Const APPNAME As String = "ShibbyGit"
Private Const EXE_PATH_PROPERTY As String = "code_GitExecutablePath"
Private Const PROJECT_PATH_PROPERTY As String = "code_GitProjectPath"
Private Const FRX_CLEANUP_PROPERTY As String = "code_FrxCleanup"
Private Const EXPORT_ON_SAVE_PROPERTY As String = "code_ExportOnSave"

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
Public Property Get ExportOnSave() As Boolean
    ExportOnSave = DocPropIO.GetBooleanFromDocProperties(EXPORT_ON_SAVE_PROPERTY)
End Property

' set the git project path
Public Property Let ExportOnSave(ByVal newVal As Boolean)
    DocPropIO.AddBooleanToDocProperties EXPORT_ON_SAVE_PROPERTY, newVal
End Property



