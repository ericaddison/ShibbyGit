VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GitRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' class gitBranch
' represents one git project branch

'*******************************
' Attributes

Private pName As String
Private pUrl As String
Private pType As String

'*******************************
' Accessors

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

Public Property Get Url() As String
    Url = pUrl
End Property

Public Property Let Url(ByVal value As String)
    pUrl = value
End Property

Public Property Get RemoteType() As String
    RemoteType = pType
End Property

Public Property Let RemoteType(ByVal value As String)
    If InStr(1, LCase(value), "fetch") Then
        pType = "fetch"
    Else
        pType = "push"
    End If
End Property
