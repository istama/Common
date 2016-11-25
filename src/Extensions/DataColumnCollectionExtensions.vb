'
' 日付: 2016/11/25
'
Imports System.Data

Namespace Extensions
  
Public Module DataColumnCollectionExtensions
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function ToEnumerable(collection As DataColumnCollection) As IEnumerable(Of DataColumn)
    For Each col As DataColumn In collection
      Yield col
    Next
  End Function
End Module

End Namespace