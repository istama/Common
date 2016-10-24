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
  
  Private ReadOnly notContaindeToDataTable As Boolean
  
  Private ReadOnly childs As List(Of ExcelColumnNode)
  
  Public Sub New(col As String, Optional name As String="", Optional notContainedToDataTable As Boolean=False)
    If col Is Nothing Then Throw New ArgumentNullException("col is null")
    If Not Cell.ValidColumn(col) Then Throw New ArgumentException("col is invalid value / " & col)
    
    Me.col = col
    Me.name = name
    Me.notContaindeToDataTable = notContaindeToDataTable
    Me.childs = New List(Of ExcelColumnNode)
  End Sub
  
  Public Function GetName() As String
    Return Me.name
  End Function
  
  Public Function GetCol() As String
    Return Me.col
  End Function
  
  Public Function ContainedToDataTable() As Boolean
    Return Not Me.notContaindeToDataTable
  End Function
  
  Public Function GetChilds() As List(Of ExcelColumnNode)
    Return New List(Of ExcelColumnNode)(childs.ToArray)
  End Function
  
  Public Sub AddChild(node As ExcelColumnNode)
    Me.childs.Add(node)
  End Sub
  
End Structure

End Namespace