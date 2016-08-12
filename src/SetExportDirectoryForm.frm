VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetExportDirectoryForm 
   Caption         =   "Set Import/Export Directory"
   ClientHeight    =   2472
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6195
   OleObjectBlob   =   "SetExportDirectoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetExportDirectoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub resetForm()
    Dim dir As String
    dir = ShibbySettings.ImportExportPath
    DirTextBox.Text = dir
End Sub


Private Sub BrowseButton_Click()
   DirTextBox.Text = FileUtils.FolderBrowser("Browse for Import/Export Folder")
End Sub

Private Sub CancelButton_Click()
    Me.hide
End Sub

Private Sub OKButton_Click()

    Dim newPath As String
    newPath = DirTextBox.Text
    
    If FileUtils.FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find folder: " & newPath
        Exit Sub
    End If

    ShibbySettings.ImportExportPath = newPath
    Me.hide
End Sub



