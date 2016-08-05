Attribute VB_Name = "DocPropIO"
Public Function GetItemFromDocProperties(ByVal Name As String) As Variant
  Dim docProps As Office.DocumentProperties
  Set docProps = ActivePresentation.CustomDocumentProperties
    
  On Error Resume Next
    
    Dim dir As String
    dir = docProps.Item(Name).value
    If Err.Number <> 0 Then
        Err.Clear
        dir = ""
    End If

  On Error GoTo 0
  
  GetItemFromDocProperties = dir
End Function

Public Sub AddStringToDocProperties(ByVal Name As String, ByVal value As Variant)
  Dim docProps As Office.DocumentProperties
  Set docProps = ActivePresentation.CustomDocumentProperties
    
  On Error Resume Next
    docProps.Item(Name).Delete
  On Error GoTo 0
  docProps.Add Name:=Name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeString
  
End Sub



Public Sub AddNumberToDocProperties(ByVal Name As String, ByVal value As Variant)
  Dim docProps As Office.DocumentProperties
  Set docProps = ActivePresentation.CustomDocumentProperties
    
  On Error Resume Next
    docProps.Item(Name).Delete
  On Error GoTo 0
  docProps.Add Name:=Name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeNumber
  
End Sub

