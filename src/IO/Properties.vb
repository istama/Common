'
' 日付: 2016/04/27
'
Imports System.Collections
Imports System.IO
Imports Common.Util

Namespace IO

Public Class Properties
  ' IOストリーム
  Private file As TextFile
  ' プロパティテーブル
  Private table As IDictionary(Of String, String)
  ' 読み込んだプロパティテーブルが最新かどうか
  Private latest As Boolean
  
  ''' <summary>
  ''' コンストラクタ
  ''' </summary>
  ''' <param name="filepath">プロパティを読み書きするファイルのパス</param>
  Public Sub New(filepath As String)
    file = New TextFile(filepath, System.Text.Encoding.Default)
    table = New Dictionary(Of String, String)
    latest = False
  End Sub
  
  ''' <summary>
  ''' プロパティの値を取得
  ''' </summary>
  ''' <param name="key">プロパティのキー</param>
  ''' <returns>プロパティの値</returns>
  Public Function GetValue(key As String) As IOption(Of String)
    Dim dict As IDictionary(Of String, String) = Load()
    
    If dict.ContainsKey(Key) Then
      Return Some(Of String).Create(dict(key))
    Else
      Return None(Of String).Create()
    End If
  End Function
  
  ''' <summary>
  ''' プロパティをファイルに書き込む
  ''' </summary>
  ''' <param name="key">プロパティのキー</param>
  ''' <param name="value">プロパティの値</param>
  Public Sub Add(key As String, value As String)
    Dim prop As String = String.Format("{0}={1}", key, value)
    file.Append(prop)
    latest = False
  End Sub
  
  ''' <summary>
  ''' プロパティファイルを構築する
  ''' すでにあるプロパティはリセットされる
  ''' </summary>
  ''' <param name="table">プロパティテーブル</param>
  Public Sub Build(table As IDictionary(Of String, String))
    file.Reset()
    For Each key In table.Keys
      Call Add(key, table(key))
    Next
  End Sub
  
  ''' <summary>
  ''' プロパティを読み込みテーブルにして返す
  ''' キーが重複して書き込まれていた場合、後に読み込んだ方の値を返す
  ''' </summary>
  ''' <returns>プロパティのテーブル</returns>
  Public Function Load() As IDictionary(Of String, String)
    If latest Then
      Return table
    Else
      Try
        Dim newTable As New Dictionary(Of String, String)
        Dim list As List(Of String) = file.Read()
        list.ForEach(
          Sub(prop)
            Dim idx As Integer = prop.IndexOf("=")
            If idx > 0 Then
              Dim key As String = prop.Substring(0, idx)
              Dim value As String
              If prop.Length > idx + 1 Then
                value = prop.Substring(idx+1)
              Else
                value = String.Empty
              End If
              
              If newTable.ContainsKey(key) Then
                newTable(key) = value
              Else
                newTable.Add(key, value)
              End If
            End If
          End Sub)
      
        table = newTable
        latest = True
      Catch ex As FileNotFoundException
        file.Create()
      End Try
      
      Return table
    End If
  End Function
End Class

End Namespace