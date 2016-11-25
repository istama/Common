'
' 日付: 2016/11/23
'
Imports System.Data

Namespace Extensions

Public Module DataRowCollectionExtensions
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function ToEnumerable(collection As DataRowCollection) As IEnumerable(Of DataRow)
    For Each row As DataRow In collection
      Yield row
    Next
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub ForEach(collection As DataRowCollection, f As Action(Of DataRow))
    For Each row As DataRow In collection
      f(row)
    Next
  End Sub
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function Convert(OF T)(collection As DataRowCollection, f As Func(Of DataRow, T)) As IEnumerable(Of T)
    For Each row As DataRow In collection
      Yield f(row)
    Next
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function Where(collection As DataRowCollection, f As Func(Of DataRow, Boolean)) As IEnumerable(Of DataRow)
    For Each row As DataRow In collection
      If f(row) Then
        Yield row
      End If
    Next
  End Function
End Module

End Namespace