'
' 日付: 2016/10/25
'
Imports System.Data
Imports Common.IO

Namespace Extensions
  
Public Module DataRowExtensions
  
  
  ''' <summary>
  ''' すべての列をDouble型で加算する。
  ''' Double型に変換できない列は何もしない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub PlusByDouble(dataRow As DataRow, addedRow As DataRow)
    Dim table As DataTable = addedRow.Table
    For Each col As DataColumn In table.Columns
      Dim name As String = col.ColumnName
      Dim value As Double
      If Not System.Convert.IsDBNull(addedRow(name)) AndAlso Double.TryParse(addedRow(name).ToString, value) Then
        Dim sum As Double
        If Not System.Convert.IsDBNull(dataRow(name)) AndAlso Double.TryParse(dataRow(name).ToString, sum) Then
          dataRow(name) = (value + sum).ToString
        Else
          dataRow(name) = value.ToString
        End If
      End If
    Next
  End Sub
  
  ''' <summary>
  ''' すべての列をDouble型で減算する。
  ''' Double型に変換できない列は何もしない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub MinusByDouble(dataRow As DataRow, addedRow As DataRow)
    Dim table As DataTable = addedRow.Table
    For Each col As DataColumn In table.Columns
      Dim name As String = col.ColumnName
      Dim value As Double
      If Not System.Convert.IsDBNull(addedRow(name)) AndAlso Double.TryParse(addedRow(name).ToString, value) Then
        Dim sum As Double
        If Not System.Convert.IsDBNull(dataRow(name)) AndAlso Double.TryParse(dataRow(name).ToString, sum) Then
          dataRow(name) = (value - sum).ToString
        Else
          dataRow(name) = value.ToString
        End If
      End If
    Next
  End Sub
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function IsNull(dataRow As DataRow, columnName As String) As Boolean
    If Not dataRow.HasColumn(columnName) Then
      Throw New ArgumentException("指定された列名は存在しません。 / " & columnName)
    End If
    
    Return System.Convert.IsDBNull(dataRow(columnName))
  End Function
  
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function HasColumn(dataRow As DataRow, columnName As String) As Boolean
    Return dataRow.Table.Columns.Contains(columnName)
  End Function
End Module

End Namespace