VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' class gitBranch
' represents one git project branch

'*******************************
' Attributes

Private pName As String
Private pActive As Boolean


'*******************************
' Accessors

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(ByVal value As String)
    If value Like "[*]*" Then
        pActive = True
        value = Replace(value, "*", "")
    Else
        pActive = False
    End If
    value = Replace(value, " ", "")
    pName = value
End Property

Public Property Get Active() As Boolean
    Active = pActive
End Property

Public Property Let Active(ByVal value As Boolean)
    pActive = value
End Property

