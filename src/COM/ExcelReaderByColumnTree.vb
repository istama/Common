'
' 日付: 2016/10/18
'
Namespace COM
  
''' <summary>
''' Excelの指定した行の列を、列の木構造のコレクションにもとづいて読み込むクラス。
''' </summary>
Public Class ExcelReaderByColumnTree
  Private ReadOnly excel As Excel3
  
  Public Sub New(excel As Excel3)
    If excel Is Nothing Then Throw New ArgumentNullException("excel is null")
    
    Me.excel = excel
  End Sub
  
  Public Sub Open(filepath As String, readMode As Boolean)
    'Me.excel.Open(filepath, readMode)
  End Sub
  
  Public Sub Close(filepath As String)
    'Me.excel.Close(filepath)
  End Sub
  
  ''' <summary>
  ''' Excelの指定した行の列を、列の木構造のコレクションにもとづいて読み込む。
  ''' 木構造の親の列にデータがセットされていない場合、子の列は読み込まない。
  ''' </summary>
  Public Function Read(row As Integer, filepath As String, sheetName As String, columnNodes As ExcelColumnNode) As IDictionary(Of String, String)
    If filepath Is Nothing Then Throw New ArgumentNullException("filepath is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    
    Dim result As New Dictionary(Of String, String)
    Read(row, filepath, sheetName, columnNodes, result)
    
    Return result
  End Function
  
  ''' <summary>
  ''' 列を再帰で読み込む。
  ''' </summary>
  Private Sub Read(row As Integer, filepath As String, sheetName As String, columnNodes As ExcelColumnNode, result As IDictionary(Of String, String))
    Dim cell As Cell = Cell.Create(row, columnNodes.GetCol)
    'Dim value As String = excel.Read(New ExcelData("", filepath, sheetName, cell))
    Dim value As String = debugRead(row, filepath, sheetName, cell).ToString
    
    ' データテーブルに含めるノードかどうか判定
    If columnNodes.ContainedToDataTable Then
      ' 列が重複していないか確かめる
      If Not result.ContainsKey(columnNodes.GetName) Then
        result.Add(columnNodes.GetName, value)
      Else
        ' 重複していた場合は後に読み込んだデータで上書きする
        result(columnNodes.GetName) = value      
      End If
    End If
    
    If value IsNot String.Empty Then
      columnNodes.GetChilds.ForEach(
        Sub(node) Read(row, filepath, sheetName, node, result))
    End If
  End Sub
  
  Private Function debugRead(row As Integer, filepath As String, sheetName As String, cell As Cell) As String
    Dim value As Integer = (cell.Row + Asc(cell.Col))
    If value Mod 6 = 0 Then
      Return String.Empty
    End If
    
    Dim m As Integer = 1
    For i = 1 To 12
      If sheetName = i.ToString & "月分" Then
        m = i
        Exit For
      End If
    Next
    
    Return (value + m).ToString
  End Function
End Class

End Namespace