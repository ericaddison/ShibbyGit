VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitCommitMessageForm 
   Caption         =   "Git Commit Message"
   ClientHeight    =   2256
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   7512
   OleObjectBlob   =   "GitCommitMessageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitCommitMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()
    GitCommitMessageForm.Hide
End Sub

Private Sub OKButton_Click()
    Dim commitMessage As String
    commitMessage = MessageTextBox.Text
    
    If commitMessage = "" Then
        MsgBox "Please enter a commit message"
        Exit Sub
    End If
    
    GitCommands.GitCommit (commitMessage)
    GitCommitMessageForm.Hide
End Sub

