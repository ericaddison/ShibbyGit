Attribute VB_Name = "DocPropIO"


Public Function GetItemFromDocProperties(ByVal name As String) As Variant
  Dim docProps As Office.DocumentProperties
  Set docProps = GetDocProps

  On Error Resume Next
    
    Dim val As String
    val = docProps.Item(name).value
    If err.Number <> 0 Then
        err.Clear
        val = ""
    End If

  On Error GoTo 0
  
  GetItemFromDocProperties = val
End Function


Public Function GetBooleanFromDocProperties(ByVal name As String) As Boolean
  Dim docProps As Office.DocumentProperties
  Set docProps = GetDocProps
    
  On Error Resume Next
    
    Dim val As Boolean
    val = docProps.Item(name).value
    If err.Number <> 0 Then
        err.Clear
        val = False
    End If

  On Error GoTo 0
  
  GetBooleanFromDocProperties = val
End Function


Public Sub AddStringToDocProperties(ByVal name As String, ByVal value As Variant)
  Dim docProps As Office.DocumentProperties
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeString
  
End Sub

Public Sub AddBooleanToDocProperties(ByVal name As String, ByVal value As Boolean)
  Dim docProps As Office.DocumentProperties
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeBoolean
  
End Sub


Public Sub AddNumberToDocProperties(ByVal name As String, ByVal value As Variant)
  Dim docProps As Office.DocumentProperties
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeNumber
  
End Sub


Private Function GetDocProps() As DocumentProperties
    #If APPNAME = "PowerPoint" Then
        Set GetDocProps = ActivePresentation.CustomDocumentProperties
    #ElseIf APPNAME = "Excel" Then
        Set GetDocProps = ActiveWorkbook.CustomDocumentProperties
    #ElseIf APPNAME = "Word" Then
        Set GetDocProps = ActiveDocument.CustomDocumentProperties
    #End If
End Function

