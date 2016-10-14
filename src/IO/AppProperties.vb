'
' 日付: 2016/04/28
'
Imports System.Linq

Imports System.Text
Imports Common.Extensions
Imports Common.Util


Namespace IO

''' <summary>
''' プロパティファイルの構築とアクセスを管理する
''' </summary>
Public MustInherit Class AppProperties
  ''' プロパティ	
  Private prop As Properties
  ''' プロパティを再構築したかどうか	
  Private rebuilded As Boolean
  
  ''' <summary>
  '''  コンストラクタ
  ''' </summary>
  ''' <param name="filepath">プロパティファイルのパス</param>
  Public Sub New(filepath As String)
    prop = New Properties(filepath)
    rebuilded = False
  End Sub
  
  ''' <summary>
  ''' プロパティの値を取得
  ''' </summary>
  ''' <param name="key">プロパティのキー</param>
  ''' <returns>プロパティの値</returns>
  Public Function GetValue(key As String) As IOption(Of String)
    ReBuild()
    Return prop.GetValue(key)
  End Function
  
  ''' <summary>
  ''' プロパティを追加
  ''' </summary>
  ''' <param name="key">プロパティのキー</param>
  ''' <param name="value">プロパティの値</param>
  Public Sub AddValue(key As String, value As String)
    prop.Add(key, value)
  End Sub
  
  ''' <summary>
  ''' デフォルトにあってファイルにないプロパティを追加し、
  ''' そしてデフォルト以外のプロパティが許可されてない場合は、ファイルにあってデフォルトにないプロパティは削除するように
  ''' ファイルプロパティを構築しなおす
  ''' デフォルトとファイルでプロパティの値が異なる場合は、ファイルが優先される
  ''' </summary>
  Private Sub ReBuild()
    If Not rebuilded Then
      Dim def As IDictionary(Of String, String) = DefaultProperties()
      Dim current As IDictionary(Of String, String) = prop.Load()
      Dim newTable As New Dictionary(Of String, String)
      Dim changed As Boolean = False
      
      ' デフォルトにないプロパティの使用が許可されている場合
      If AllowNonDefaultProperty() Then
        ' 既存のすべてのプロパティを新しいプロパティに残す
        current.Keys.ForEach(Sub(k) newTable.Add(k, current(k)))
      End If
      
      ' デフォルトのプロパティがまだセットされていない場合セットする
      def.Keys.ForEach(
        Sub(k)
          If Not newTable.ContainsKey(k) Then
            If current.ContainsKey(k) Then
              newTable.Add(k, current(k))
            Else
              newTable.Add(k, def(k))
            End If
            changed = True
          End If
        End Sub)
    
     If changed OrElse newTable.Count <> current.Count Then
       prop.Build(newTable)
      End If
    
     rebuilded = True
    End if
  End Sub
  
  ''' <summary>
  ''' デフォルトのプロパティとその値を返す。
  ''' ここで指定されたプロパティは必須であり、プロパティファイルに存在しない場合は、
  ''' アプリケーション起動時に自動でセットされる。
  ''' </summary>
  ''' <returns>プロパティのテーブル</returns>
  Protected MustOverride Function DefaultProperties() As IDictionary(Of String, String) 
  
  ''' <summary>
  ''' デフォルトにないプロパティの存在を認めるかどうか。
  ''' 認めるならTrueを返す。
  ''' 認めると、プロパティファイルにデフォルト値以外のプロパティが記述されていても、起動時に削除されない。
  ''' </summary>
  ''' <returns></returns>
  Protected MustOverride Function AllowNonDefaultProperty() As Boolean

End Class

End Namespace