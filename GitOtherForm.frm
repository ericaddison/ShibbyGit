VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitOtherForm 
   Caption         =   "User Git Command"
   ClientHeight    =   2640
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6204
   OleObjectBlob   =   "GitOtherForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitOtherForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CancelButton_Click()
    GitOtherForm.Hide
End Sub

Private Sub OKButton_Click()
    Dim options As String
    options = OptionsBox.Text
    
    If options = "" Then
        MsgBox "Please enter git options"
        Exit Sub
    End If
    
    Git.GitOther (options)
        
End Sub
