'
' 日付: 2016/10/10
'
Namespace Util
  
Public Class TextUtils
  
  Public Shared Function ToCharCode(text As String, figure As Integer) As Integer
    If text Is Nothing Then Throw New NullReferenceException("text is null")
    
    Dim rest As Integer = figure ' 文字コードの残り桁数
    Dim codeStr As String = String.Empty
    
    For Each c As Char In text.ToCharArray
      ' Unicode文字を数値に変換する
      Dim code As String = AscW(c).ToString
      ' 戻り値として返す文字コードの残りの長さを取得
      Dim len As Integer = Math.Min(rest, code.Length)
      
      codeStr += code.Substring(0, len)
      rest -= len
      
      If rest <= 0 Then Exit For
    Next
    
    Return Integer.Parse(codeStr)
  End Function
 
End Class

End Namespace