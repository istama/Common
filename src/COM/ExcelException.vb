'
' 日付: 2016/04/29
'
Imports System
Imports System.Runtime.Serialization

Namespace COM
  
  ''' <summary>
''' ExcelのCOMコンポーネントを生成する際に異常が発生したときに投げられる例外。
''' </summary>
Public Class ExcelException
  Inherits Exception
  Implements ISerializable
  
  Public Sub New()
  End Sub
  
  Public Sub New(message As String)
    MyBase.New(message)
  End Sub
  
  Public Sub New(message As String, innerException As Exception)
    MyBase.New(message, innerException)
  End Sub
  
  ' This constructor is needed for serialization.
  Protected Sub New(info As SerializationInfo, context As StreamingContext)
    MyBase.New(info, context)
  End Sub
End Class

End Namespace