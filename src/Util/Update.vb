'
' 日付: 2016/05/04
'
Imports Common.IO
Imports System.IO

Namespace Util

Public MustInherit Class Update
  Private latestVersionFile As TextFile	
  Private batchFile As TextFile
  
  Public Sub New(latestVersionFilePath As String, batchFilePath As String)
    If latestVersionFilePath = String.Empty Then
      Throw New ArgumentException("バージョンファイルへのパスが空文字です。")
    ElseIf batchFilePath = String.Empty Then
      Throw New ArgumentException("バッチファイルへのパスが空文字です。")
    End If
    
    latestVersionFile = New TextFile(latestVersionFilePath, System.Text.Encoding.Default)
    batchFile = New TextFile(batchFilePath, System.Text.Encoding.Default)
  End Sub
  
  ''' <summary>
  ''' バッチファイルを作成する。
  ''' </summary>
  Public Function CreateUpdateBatch() As Boolean
    Dim s As String = Script()
    If s <> String.Empty Then
      ' バッチファイルが存在しない場合、新規作成する
      batchFile.Create()
      ' 新規作成したファイルにスクリプトを上書きする
      batchFile.Write(s)
      Return True
    Else
      Return False
    End If
  End Function
  
  ''' <summary>
  ''' バッチファイルを削除する。
  ''' </summary>
  Public Sub DeleteUpdateBatch()
    batchFile.Delete
  End Sub
  
  ''' <summary>
  ''' 現在バージョンより新しいバージョンがある場合、バッチファイルを実行する。
  ''' </summary>
  ''' <param name="currentVersion"></param>
  Public Function RunUpdateBatch(currentVersion As Version) As System.Diagnostics.Process
    If CreateUpdateBatch() AndAlso existsUpdateVersion(currentVersion) Then
      Dim fullpath As String = Path.GetFullPath(batchFile.Path())
      Return System.Diagnostics.Process.Start(fullpath)
    Else
      Return Nothing
    End If
  End Function
  
  ''' <summary>
  ''' 更新されたバージョンのファイルがある場合はTrueを返す。
  ''' </summary>
  ''' <param name="currentVersion"></param>
  ''' <returns></returns>
  Public Function existsUpdateVersion(currentVersion As Version) As Boolean
    Dim exists As Boolean = False
    Try
      Dim filetext As List(Of String) = latestVersionFile.Read()
      Dim latest As Version = Nothing
      
      exists =
        filetext.Count > 0 AndAlso _
        Version.TryParse(filetext(0), latest) AndAlso _
        currentVersion < latest
    Catch ex As FileNotFoundException
    End Try
    
    Return exists
  End Function
  
  ''' <summary>
  ''' バッチファイルに記述するスクリプトを返す。
  ''' </summary>
  ''' <returns></returns>
  Protected MustOverride Function Script() As String
  
End Class

End Namespace