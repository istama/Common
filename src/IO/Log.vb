'
' SharpDevelopによって生成
' ユーザ: Blue
' 日付: 2016/09/26
' 時刻: 21:59
' 
' このテンプレートを変更する場合「ツール→オプション→コーディング→標準ヘッダの編集」
'
Namespace IO
  
Public Class Log
  Private Shared file As TextFile
  
  Public Shared Sub SetFilePath(path As String)
    file = New TextFile(path, System.Text.Encoding.UTF8)
    file.Create
  End Sub
  
  Public Shared Sub out(text As String)
    If file IsNot Nothing Then
      file.Append(text)
    End If
  End Sub
End Class

End Namespace