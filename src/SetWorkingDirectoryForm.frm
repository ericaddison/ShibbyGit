VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetWorkingDirectoryForm 
   Caption         =   "Set Import/Export Directory"
   ClientHeight    =   2475
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6192
   OleObjectBlob   =   "SetWorkingDirectoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetWorkingDirectoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub BrowseButton_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show = -1 Then
            DirTextBox.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub CancelButton_Click()
    SetWorkingDirectoryForm.Hide
End Sub

Private Sub OKButton_Click()

    Dim newPath As String
    newPath = DirTextBox.Text
    
    If FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find folder: " & newPath
        Exit Sub
    End If

    DocPropIO.AddStringToDocProperties CodeUtils.EXPORT_DIRECTORY_PROPERTY, newPath
    SetWorkingDirectoryForm.Hide
End Sub

