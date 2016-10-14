'
' 日付: 2016/07/29
'
Imports System.Collections.Generic
Imports System.Data

Namespace Data

Public Class EnumerableDataRowCollection
	Implements IEnumerable(Of DataRow), IEnumerable
	Private rows As DataRowCollection
	
	Public Sub New(rows As DataRowCollection)
		Me.rows = rows
	End Sub
	
	Public Function GetEnumerator() As IEnumerator(Of DataRow) Implements IEnumerable(Of DataRow).GetEnumerator
		Dim list As New List(Of DataRow)
		For Each row As DataRow In rows
			list.Add(row)
		Next
		Return list.GetEnumerator
	End Function
	
	Private Function GetEnumerator1() As IEnumerator Implements IEnumerable(Of DataRow).GetEnumerator
		Dim ie As IEnumerable(Of DataRow) = Me
		Return ie.GetEnumerator
	End Function
	
End Class

Public Class EnumerableList(Of T)
	Implements IEnumerable(Of T), IEnumerable
	Private items As IEnumerable
	
	Public Sub New(rows As IEnumerable)
		Me.items = rows
	End Sub
	
	Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
		Dim list As New List(Of T)
		For Each row As T In items
			list.Add(row)
		Next
		Return list.GetEnumerator
	End Function
	
	Private Function GetEnumerator1() As IEnumerator Implements IEnumerable(Of T).GetEnumerator
		Dim ie As IEnumerable(Of T) = Me
		Return ie.GetEnumerator
	End Function
	
End Class
End Namespace