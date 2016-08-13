Attribute VB_Name = "GitProject"


' Look for .frx files in the git status
' if present, check for associated .frm
' with the same status. If that fails,
' checkout the previous .frx
Public Sub RemoveUnusedFrx()
    
    Dim status As String
    status = GitCommands.RunGitAsProcess("status -s")
  
    ' put all the status lines in a collection
    Dim strArray() As String
    strArray = Split(status, vbLf)
    Dim statusLines As New Collection
    Dim i As Integer
    For i = LBound(strArray) To UBound(strArray)
        If Not InCollection(statusLines, strArray(i)) Then
            statusLines.Add strArray(i), strArray(i)
        End If
    Next i
  
    ' loop through to see if frx has accomapnying frm
    ' if not, checkout the frx
    Dim checkoutFiles As String
    Dim line As Variant
    For Each line In statusLines
        If line Like "*.frx" Then
            Dim form As String
            form = Left(line, Len(line) - 3)
            form = form & "frm"
            If Not InCollection(statusLines, form) Then
                checkoutFiles = checkoutFiles & " " & GetFileNameFromStatusLine(line)
            End If
        End If
    Next line
        
    GitCommands.RunGitAsProcess ("checkout -- " & checkoutFiles)
End Sub


Public Function InCollection(col As Collection, key As String) As Boolean
    On Error GoTo incol
    col.Item key
incol:
    InCollection = (err.Number = 0)
    On Error GoTo 0
End Function


Private Function GetFileNameFromStatusLine(ByVal line As String) As String
    If Len(line) > 3 Then
        GetFileNameFromStatusLine = Right(line, Len(line) - 3)
    End If
End Function
