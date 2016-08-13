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


Public Function GetDocProps() As DocumentProperties
    Dim name As String
    name = Application.name
    
    Dim app As Object
    Set app = Application
    
    Select Case name
        Case "Microsoft PowerPoint"
            Set GetDocProps = ActivePresentation.CustomDocumentProperties
        Case "Microsoft Excel"
            Set GetDocProps = ActiveWorkbook.CustomDocumentProperties
        Case "Microsoft Word"
            Set GetDocProps = ActiveDocument.CustomDocumentProperties
    End Select
End Function

