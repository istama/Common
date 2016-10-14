'
' 日付: 2016/04/28
'
Namespace Account

Public Class UserInfo
  Private name As String
  Private id As String
  Private password As String
  
  Public Sub New(name As String, id As String, password As String)
    Me.name = name
    Me.id = id
    Me.password = password
  End Sub
  
  Public Function GetName As String
    Return name
  End Function
  
  Public Function GetId As String
    Return id
  End Function
  
  Public Function GetSimpleId As String
    Return id.Substring(id.Length - 3, 3)
  End Function
  
  Public Function GetPassword As String
    Return password
  End Function
End Class

End Namespace