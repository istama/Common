'
' 日付: 2016/10/09
'
Imports System.Text.RegularExpressions

Imports Common.Extensions

Namespace Text

''' <summary>
''' 文字列のマッチング用クラス。
''' 文字列とマッチング方法がセットになっている。
''' </summary>
Public Class MatchingText
  ''' マッチングする文字列。  
  Private _words As String()
  ''' マッチング方法。
  Private _mode As MatchingMode
  
  Public Sub New(word As String, mode As MatchingMode)
    If word Is Nothing Then Throw New NullReferenceException("word is null")
    
    Me._words = word.Split(" "c)
    Me._mode   = mode
  End Sub
  
  Public Function Word As String
    Return String.Join(" ", _words)
  End Function
  
  Public Function Words As List(Of String)
    Return New List(Of String)(_words)
  End Function
  
  Public Function Mode As MatchingMode
    Return Me._mode
  End Function
  
  ''' <summary>
  ''' ほかのマッチング文字列とマッチングするか判定する。
  ''' マッチング方法は、マッチング文字列を受け取ったインスタンスにセットされた方法で行う。
  ''' 文字列を受け取った方のマッチング文字列が空の場合は判定は常にTrueとする。
  ''' 引数のマッチング文字列が空のときの判定結果はisMatchEmptyTextの値で決める。
  ''' </summary>
  Public Function Matching(t As MatchingText, isMatchEmptyText As Boolean) As Boolean
    If t Is Nothing Then Return False
    
    ' マッチング文字列がセットされていない場合は常にTrueを返す
    If Me.Word = String.Empty Then Return True
    ' 引数のマッチング文字列が空の場合、isMatchEmptyTextがtrueならtrueを返す
    If t.Word = String.Empty AndAlso isMatchEmptyText Then Return True
    
    ' 分割されたマッチング文字列の中に空文字がある場合は、マッチング対象としない
    Return _words.Any(Function(w) w <> String.Empty AndAlso Regex.IsMatch(t.Word, ToPettern(w))) 
  End Function
  
  Private Function ToPettern(word As String) As String
    If Me._mode = MatchingMode.Forward Then
      Return "^" & word
    ElseIf Me._mode = MatchingMode.Backward
      Return word & "$"
    ElseIf Me._mode = MatchingMode.Perfection
      Return "^" & word & "$"
    Else
      Return word
    End If
  End Function    
End Class

End Namespace