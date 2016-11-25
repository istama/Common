'
' 日付: 2016/09/22
' 
Namespace Extensions

''' <summary>
''' IEnumerableの拡張メソッド。
''' </summary>
Public Module IEnumerableExtensions
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub ForEach(Of T)(c As IEnumerable(Of T), f As Action(Of T))
    For Each i As T In c
      f(i)
    Next
  End Sub
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub ForEach(Of T)(c As IEnumerable(Of T), f As Action(Of T, Integer))
    Dim idx As Integer = 0
    For Each i As T In c
      f(i, idx)
      idx += 1
    Next
  End Sub
    
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function All(Of T)(c As IEnumerable(Of T), f As Func(Of T, Boolean)) As Boolean
    Dim fulfill As Boolean = True
    For Each i As T In c
      If Not f(i) Then
        fulfill = False
        Exit For
      End If
    Next
    Return fulfill
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Any(Of T)(c As IEnumerable(Of T), f As Func(Of T, Boolean)) As Boolean
    Dim fulfill As Boolean = False
    For Each i As T In c
      If f(i) Then
        fulfill = True
        Exit For
      End If
    Next
    Return fulfill
  End Function  
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Find(Of T)(c As IEnumerable(Of T), f As Func(Of T, Boolean)) As T
    Dim found As T = Nothing
    For Each i As T In c
      If f(i) Then
        found = i
        Exit For
      End If
    Next
    Return found
  End Function  
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Convert(Of T, T2)(c As IEnumerable(Of T), f As Func(Of T, T2)) As IEnumerable(Of T2)
    Dim list As New List(Of T2)
    For Each i As T In c
      list.Add(f(i))
    Next
    Return list
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function _Where(Of T)(c As IEnumerable(Of T), f As Func(Of T, Boolean)) As IEnumerable(Of T)
    For Each i As T In c
      If f(i) Then
        Yield i
      End If
    Next
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Hold(Of T, T2)(c As IEnumerable(Of T), init As T2, f As Func(Of T, T2, T2)) As T2
    For Each i As T In c
      init = f(i, init)
    Next
    Return init
  End Function
End Module

End Namespace