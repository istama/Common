'
' 日付: 2016/09/04
'
Imports System

Namespace Util

''' <summary>
''' 値のあるOption型
''' </summary>
Public Structure Some(Of T)
  Implements IOption(Of T)
  
  Public Shared Function Create(v As T) As IOption(Of T)
    Return New Some(Of T)(v)
  End Function
  
  Private v As T
  
  Private Sub New(v As T)
    Me.v = v
  End Sub
  
  Public Function GetOrDefault(def As T) As T Implements IOption(Of T).GetOrDefault
    Return v
  End Function
End Structure

End Namespace