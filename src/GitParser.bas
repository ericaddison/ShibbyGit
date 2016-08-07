Attribute VB_Name = "GitParser"
Public Function ParseBranches() As Collection

    Dim branches As New Collection
    Dim output As String
    output = GitCommands.GitOther("branch")
    
    Dim branchNames() As String
    branchNames = Split(output, vbLf)
    Dim ind As Integer
    Dim newBranch As GitBranch
    For ind = LBound(branchNames) To UBound(branchNames)
        If Not branchNames(ind) = "" Then
            Set newBranch = New GitBranch
            newBranch.Name = branchNames(ind)
            branches.Add newBranch
            Debug.Print newBranch.Name
        End If
    Next ind

    Set ParseBranches = branches

End Function


Public Function ParseRemotes() As Collection

    Dim remotes As New Collection
    Dim output As String
    output = GitCommands.GitOther("remote -v")
    
    Dim branchNames() As String
    branchNames = Split(output, vbLf)
    Dim ind As Integer
    Dim newBranch As GitBranch
    For ind = LBound(branchNames) To UBound(branchNames)
        If Not branchNames(ind) = "" Then
            Set newBranch = New GitBranch
            newBranch.Name = branchNames(ind)
            branches.Add newBranch
            Debug.Print newBranch.Name
        End If
    Next ind

    Set ParseBranches = branches

End Function
