'
' 日付: 2016/09/04
'
Imports System

Namespace Util

''' <summary>
''' 値のないOption型
''' </summary>
Public Structure None(Of T)
  Implements IOption(Of T)
  
  Public Shared Function Create() As IOption(Of T)
    Return New None(Of T)(0)
  End Function
  
  Private v As Integer
  
  Private Sub New(v As Integer)
    Me.v = v
  End Sub
  
  Public Function GetOrDefault(def As T) As T Implements IOption(Of T).GetOrDefault
    Return def
  End Function
End Structure

End Namespace