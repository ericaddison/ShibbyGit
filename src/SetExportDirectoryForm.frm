VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetExportDirectoryForm 
   Caption         =   "Set Import/Export Directory"
   ClientHeight    =   2472
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6192
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
    dir = DocPropIO.GetItemFromDocProperties(CodeUtils.EXPORT_DIRECTORY_PROPERTY)
    DirTextBox.Text = dir
End Sub


Private Sub BrowseButton_Click()
   DirTextBox.Text = UI.FolderDialog("Browse for Import/Export Folder")
End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub OKButton_Click()

    Dim newPath As String
    newPath = DirTextBox.Text
    
    If FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find folder: " & newPath
        Exit Sub
    End If

    DocPropIO.AddStringToDocProperties CodeUtils.EXPORT_DIRECTORY_PROPERTY, newPath
    Me.Hide
End Sub



