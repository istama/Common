'
' 日付: 2016/10/18
'
Namespace COM
  
'''' <summary>
'''' 読み込むExcelの列とその列名を持つクラス。
'''' 複数のこのクラスのインスタンスを木構造に結合することができる。
'''' 列に対して処理を行うときに、子ノードの処理を親ノードの状態に依存させることなどができる。
'''' </summary>
Public Structure ExcelColumnNode
  Private ReadOnly name As String
  Private ReadOnly col As String
  
  Private ReadOnly notContainedToDataTable As Boolean
  
  Private ReadOnly childs As List(Of ExcelColumnNode)
  
  Public Sub New(col As String, Optional name As String="", Optional notContainedToDataTable As Boolean=False)
    If col Is Nothing Then Throw New ArgumentNullException("col is null")
    If Not Cell.ValidColumn(col) Then Throw New ArgumentException("col is invalid value / " & col)
    
    Me.col = col
    Me.name = name
    Me.notContainedToDataTable = notContainedToDataTable
    Me.childs = New List(Of ExcelColumnNode)
  End Sub
  
  Public Function GetName() As String
    Return Me.name
  End Function
  
  Public Function GetCol() As String
    Return Me.col
  End Function
  
  Public Function ContainedToDataTable() As Boolean
    Return Me.notContainedToDataTable = False
  End Function
  
  Public Function GetChilds() As List(Of ExcelColumnNode)
    Return New List(Of ExcelColumnNode)(childs.ToArray)
  End Function
  
  Public Sub AddChild(node As ExcelColumnNode)
    Me.childs.Add(node)
  End Sub
  
  ''' <summary>
  ''' このExcelColumnNodeをDataTableに変換したときの列のコレクションを取得する。
  ''' </summary>
  Public Function ToDataColumnCollection As DataColumnCollection
    Return ToDataTable.Columns
  End Function
  
  ''' <summary>
  ''' このExcelColumnNodeのツリーをDataTableに変換する。
  ''' </summary>
  Public Function ToDataTable() As DataTable
    Dim table As New DataTable
    Me.AddColumns(table)
    
    Return table
  End Function
  
  Private Sub AddColumns(table As DataTable)
    If Me.ContainedToDataTable Then
      table.Columns.Add(Me.CreateColumn(Me.name))
    End If
    
    Me.GetChilds.ForEach(Sub(n) n.AddColumns(table))
  End Sub
    
  Private Function CreateColumn(name As String) As DataColumn
    Dim col As New DataColumn
    col.ColumnName = name
    col.AutoIncrement = False
		
		Return col
	End FUnction
End Structure

End Namespace