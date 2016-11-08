'
' 日付: 2016/11/08
'
Imports System.Text

Namespace Extensions

Public Module StringExtensions
  
  ''' <summary>
  ''' 文字列の表示幅を返す
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function LengthByHalf(myStr As String) As Integer
    Return Encoding.GetEncoding("Shift_JIS").GetByteCount(myStr)
  End Function
  
  ''' <summary>
  ''' 文字列の右側を指定された長さになるよう、指定された文字で埋める。
  ''' 幅は半角を１文字を１として指定する。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function PadRightByHalf(myStr As String, size As Integer, padChar As Char) As String
    Dim len As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(myStr)
    Dim lack As Integer = size - len
    If lack <= 0 Then
      Return myStr
    Else
      For i = 1 To lack
        myStr = myStr & " "
      Next
      Return myStr
    End If
  End Function
  
  ''' <summary>
  ''' 文字列の左側を指定された長さになるよう、指定された文字で埋める。
  ''' 幅は半角を１文字を１として指定する。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function PadLeftByHalf(myStr As String, size As Integer, padChar As Char) As String
    Dim len As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(myStr)
    Dim lack As Integer = size - len
    If lack <= 0 Then
      Return myStr
    Else
      For i = 1 To lack
        myStr = " " & myStr
      Next
      Return myStr
    End If
  End Function
End Module

End Namespace