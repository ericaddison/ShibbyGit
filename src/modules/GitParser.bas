Attribute VB_Name = "GitParser"
Public Function ParseBranches() As Collection

    Dim branches As New Collection
    Dim output As String
    output = GitCommands.RunGitAsProcess("branch")
    
    Dim branchNames() As String
    branchNames = Split(output, vbLf)
    Dim ind As Integer
    Dim newBranch As GitBranch
    For ind = LBound(branchNames) To UBound(branchNames)
        If Not branchNames(ind) = "" Then
            Set newBranch = New GitBranch
            newBranch.name = branchNames(ind)
            branches.Add newBranch
        End If
    Next ind

    Set ParseBranches = branches

End Function


Public Function ParseRemotes() As Collection

    Dim remotes As New Collection
    Dim output As String
    output = GitCommands.RunGitAsProcess("remote -v")
    
    Dim remoteInfo() As String
    remoteInfo = Split(output, vbLf)
    Dim ind As Integer
    Dim newRemote As GitRemote
    For ind = LBound(remoteInfo) To UBound(remoteInfo)
        If Not remoteInfo(ind) = "" Then
            Set newRemote = New GitRemote
            Dim remoteLine() As String
            remoteInfo(ind) = Replace(remoteInfo(ind), vbTab, " ")
            remoteLine = Split(remoteInfo(ind), " ")
            newRemote.name = remoteLine(0)
            newRemote.Url = remoteLine(1)
            newRemote.RemoteType = remoteLine(2)
            remotes.Add newRemote
        End If
    Next ind

    Set ParseRemotes = remotes
End Function
