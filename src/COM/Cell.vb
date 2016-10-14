'
' 日付: 2016/04/29
'
Imports System.Text.RegularExpressions

Namespace COM

''' <summary>
''' Excelのセル位置を表すクラス。
''' </summary>
Public Class Cell
  Implements System.IEquatable(Of Cell)	
  ''' 行	
  Private _row As Integer
  Public ReadOnly Property Row() As Integer
    Get
      Return _row
    End Get
  End Property
  
  ''' 列
  Private _col As String
  Public ReadOnly Property Col() As String
    Get
      Return _col
    End Get
  End Property
  
  Private Sub New(row As Integer, col As String)
    Me._row = row
    Me._col = col
  End Sub
  
  Public Function Point() As String
    Return _col & _row
  End Function
  
  Public Overrides Function ToString() As String
    Return String.Format("row: {0} col: {1}", _row, _col)
  End Function
  
  Public Overloads Function Equals(ByVal other As Cell) As Boolean _
    Implements System.IEquatable(Of Cell).Equals
    'objがNothingのときは、等価でない
    If other Is Nothing Then
      Return False
    End If
    
    Return Me = other
  End Function
  
  Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
    'objがNothingか、型が違うときは、等価でない
    If (obj Is Nothing) OrElse Not (Me.GetType() Is obj.GetType()) Then
      Return False
    End If
    
    Return Me.Equals(CType(obj, Cell))
  End Function	
  
  Public Overrides Function GetHashCode() As Integer
    Return _row + _col.GetHashCode
  End Function
  
  Public Shared Function Create(row As Integer, col As Integer) As Cell
    If row < 1 Then
      Throw New ArgumentException("値が範囲外です。 row: " & row)
    ElseIf col < 1 Then
      Throw New ArgumentException("値が範囲外です。 col: " & col)
    End If
    
    Return Create(row, ToColWord(col))
  End Function
  
  Public Shared Function Create(row As Integer, col As String) As Cell
    If row < 1 Then
      Throw New ArgumentException("値が範囲外です。 row: " & row)
    ElseIf Not ValidColumn(col) Then
      Throw New ArgumentException("値が不正です。 col: " & col)
    End If
    
    Return New Cell(row, col)
  End Function
  
  ''' <summary>
  ''' Excelの列を表す文字列が正しい書式か判定する。
  ''' </summary>
  ''' <param name="col">Excelの列</param>
  ''' <returns></returns>
  Public Shared Function ValidColumn(col As String) As Boolean
    Return col IsNot Nothing AndAlso Regex.IsMatch(col, "^[A-Z]+$")
  End Function
  
  ''' <summary>
  ''' 数値をExcelの列を表す文字に変換する。
  ''' 1がA、26がZ、27がAA に変換される。
  ''' </summary>
  ''' <param name="value">変換する値</param>
  ''' <returns>列を表す文字</returns>
  Public Shared Function ToColWord(value As Integer) As String
    If value < 1 Then
      Throw New ArgumentException("1以上の数値を渡してください。 value: " & value)
    End If
    
    Const BASE_NUM As Integer = 26
    
    If value <= BASE_NUM Then
      Return ToAlph(value)
    Else
      Dim left As Integer = (value - 1) \ BASE_NUM
      Console.WriteLine("value: " & value)
      Console.WriteLine("left: " & left)
      Return ToColWord(left) & ToAlph(value - (BASE_NUM * left))
    End If		
  End Function
  
  ''' <summary>
  ''' 数値をアルファベット１文字に変換する。
  ''' 1がA, 2がB ... 26がZ に変換される。
  ''' 1～26以外の数値を渡された場合は例外を投げる。
  ''' </summary>
  ''' <param name="offset">アルファベットに変換する数値</param>
  ''' <returns>アルファベット１文字</returns>
  Private Shared Function ToAlph(offset As Integer) As Char
    ' アルファベットの範囲外のため
    If offset < 1 OrElse offset > 26 Then
      Throw New ArgumentException("数値が範囲の外です offset / " & offset)
    End If
    
    ' "A"の文字コードにoffsetを加算した値のコードのアルファベット１文字を返す
    Return Convert.ToChar(Asc("A"c) + offset - 1)
  End Function
  
  Public Shared Operator =(ByVal cell1 As Cell, ByVal cell2 As Cell) As Boolean
    Return cell1.Row = cell2.Row AndAlso cell1.Col = cell2.Col
  End Operator
  
  Public Shared Operator <>(ByVal cell1 As Cell, ByVal cell2 As Cell) As Boolean
    Return cell1.Row <> cell2.Row OrElse cell1.Col <> cell2.Col
  End Operator
End Class

End Namespace