'
' 日付: 2016/04/28
'
Imports System.Text.RegularExpressions

Namespace Format

Public Class CSV
  
  ''' <summary>
  ''' CSVテキストを各フィールドに分解してリストに格納して返す
  ''' </summary>
  ''' <param name="csv">CSVテキスト</param>
  ''' <returns>フィールドリスト</returns>
  Public Shared Function Decode(csv As String) As List(Of String)
    Dim fields As New List(Of String)
    Dim text As String = csv
    
    Do While text IsNot Nothing
      Dim t As Tuple(Of String, String) = Extract(text)
      fields.Add(t.Item1)
      If t.Item2.IndexOf(",") = 0 Then
        text = t.Item2.Substring(1)
      Else
        text = Nothing
      End If
    Loop
    
    Return fields
  End Function
  
  ''' <summary>
  ''' 先頭のフィールドと残りのテキストに分割して返す
  ''' </summary>
  ''' <param name="csv">CSVテキスト</param>
  ''' <returns>フィールドと残りのテキストを持つタプル</returns>
  Private Shared Function Extract(csv As String) As Tuple(Of String, String)
    Dim field As String
    Dim rest As String
    Dim trimed As String = csv.TrimStart(" "c)
    
    If trimed.IndexOf("""") = 0 Then
      ' フィールドが "" で囲まれている場合
      Dim qIdx As Integer = GetQuoteIndex(trimed, 1)
      If qIdx > 0 Then
        field = trimed.Substring(1, qIdx - 1).Replace("""""", """")
        rest  = trimed.Substring(qIdx + 1).TrimStart(" "c)
      Else	
        Throw New ArgumentException("CSVの書式が不正です")
      End If
    Else
      ' フィールドが "" で囲まれていない場合
      Dim cIdx As Integer = csv.IndexOf(",")
      If cIdx >= 0 Then
        field = csv.Substring(0, cIdx)
        rest  = csv.Substring(cIdx)
      Else
        field = csv
        rest  = ""
      End If
    End If
    
    Return Tuple.Create(OF String, String)(field, rest)
  End Function
  
  ''' <summary>
  ''' 引数のテキストのうち、最初に見つかったダブルクォートの位置を返す
  ''' ただしエスケープされたダブルクォートは省く
  ''' 見つからなかった場合は-1を返す
  ''' </summary>
  ''' <param name="csv">テキスト</param>
  ''' <param name="start">検索の開始位置</param>
  ''' <returns>ダブルクォートのあるインデックス</returns>
  Private Shared Function GetQuoteIndex(csv As String, start As Integer) As Integer
    Dim ca As Char() = csv.ToCharArray
    Dim idx As Integer = start
    Dim found As Boolean = False
    
    Do While ca.Length > idx
      ' ダブルクォートが連続していた場合はエスケープ文字なのでスキップする
      If ca(idx) = """" Then
        found = Not found
        If idx + 1 = ca.Length Then
          Exit Do
        Else
          idx += 1
        End If
      Else
        If found Then
          ' idxが1つ余分に加算されているので引く
          idx -= 1
          Exit Do
        Else
          idx += 1
        End If
      End If
    Loop
    
    If idx < ca.Length Then
      Return idx
    Else
      Return -1
    End If	
  End Function
  
End Class

End Namespace