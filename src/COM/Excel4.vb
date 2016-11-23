'
' 日付: 2016/11/21
'
Option Strict Off

Imports System.Threading
Imports System.Collections.Concurrent
Imports System.Linq
Imports System.IO

Imports Common.IO
Imports Common.Extensions

Namespace COM
  
''' <summary>
''' Excelにアクセスするクラス。
''' このクラスはスレッドセーフでないので単一のスレッドから実行されるべき。
''' </summary>
Public Class Excel4
  Private ReadOnly KEY_BOOK As String = "__book__*"
  
  Private initialized As Boolean = False
  
  Private excel As Object
  Private workbooks As Object
  Private bookTable As Dictionary(Of String, Dictionary(Of String, Object))
  
  Public Sub New()
  End Sub
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' </summary>
  Public Sub Init()
    If Not initialized Then
      Me.excel = CreateObject("Excel.Application")
      Me.workbooks = excel.WorkBooks
      Me.bookTable = New Dictionary(Of String, Dictionary(Of String, Object))
      
      initialized = True
      Log.out("initialized excel")
    End If
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' </summary>
  Public Sub Quit()
    If initialized Then
      Me.bookTable.Keys.ForEach(Sub(path) Close(path))
      Resource.Release(Me.workBooks)
      Me.excel.Quit()
      
      initialized = False
      Log.out("quit excel")
      End If
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルを開く。
  ''' </summary>
  Public Sub Open(filepath As String, readMode As Boolean)
    If filepath Is Nothing Then Throw New ArgumentNullException("filepath is null")
    
    If Not initialized Then Throw New ExcelException("初期処理が実行されていません。")
    
    Dim fullpath As String = Path.GetFullPath(filepath)
  
    If Not File.Exists(fullpath) Then
      Throw New FileNotFoundException("指定したファイルは存在しません。 / filepath " & filepath)
    End If
    
    If Not Me.bookTable.ContainsKey(fullpath) Then
      Dim sheetTable As New Dictionary(Of String, Object)
      
      If Me.initialized Then
        sheetTable.Add(KEY_BOOK, workbooks.Open(fullPath, Nothing, readMode))
        Me.bookTable.Add(fullpath, sheetTable)
        Log.out("open book / filepath: " & filepath & " readMode: " & readMode.ToString)
      End If
    End If  
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルを閉じる。
  ''' </summary>
  Public Sub Close(filepath As String)
    If filepath Is Nothing Then Throw New ArgumentNullException("filepath is null")
    
    Dim fullpath As String = Path.GetFullPath(filePath)
    
    If Me.bookTable.ContainsKey(fullpath) Then
      Dim sheetTable As Dictionary(Of String, Object) = Me.bookTable(fullpath)
      Me.bookTable.Remove(fullpath)
      
      If sheetTable.ContainsKey(KEY_BOOK) Then
        Dim book As Object = sheetTable(KEY_BOOK)
        sheetTable.Remove(KEY_BOOK)
        book.Close(False)
        Resource.Release(book)
        Log.out("book close / filepath: " & filepath)
      End If
    End If
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのセルを読み込む。
  ''' </summary>
  Public Function Read(filepath As String, sheetName As String, cell As Cell) As String
    If filepath  Is Nothing Then Throw New ArgumentNullException("filepath is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    If cell      Is Nothing Then Throw New ArgumentNullException("cell is null")
    
    'Log.out("read excel / filepath: " & filepath & " sheetName: " & sheetName & " cell: " & cell.ToString)
    Dim fullpath As String = Path.GetFullPath(filepath)
    
    Return AccessCell(fullpath, sheetName, cell, Function(rng) rng.Value)
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルのセルに書き込む。
  ''' </summary>
  Public Sub Write(text As String, filepath As String, sheetName As String, cell As Cell)
    If filepath  Is Nothing Then Throw New ArgumentNullException("filepath is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    If cell      Is Nothing Then Throw New ArgumentNullException("cell is null")
    
    'Log.out("write to excel / filepath: " & filepath & " sheetName: " & sheetName & " cell: " & cell.ToString & " text: " & text)
    Dim fullpath As String = Path.GetFullPath(filepath)
    
    AccessCell(fullpath, sheetName, cell,
      Function(rng)
        rng.Value = text
        Return Nothing
      End Function)
    
    Dim book As Object = GetBook(fullpath)
    book.Save()
  End Sub
  
  ''' <summary>
  ''' Excelのセルにアクセスする。
  ''' </summary>
  Private Function AccessCell(fullpath As String, sheetName As String, cell As Cell, f As Func(Of Object, String)) As String
    If Not Me.initialized Then Throw New InvalidOperationException("初期処理が行われていません。")
    
    Dim sheet As Object = GetSheet(fullpath, sheetName)
    
    Dim result As String = String.Empty
    Dim rng As Object = sheet.Range(cell.Point)
    If rng IsNot Nothing Then
      result = f(rng)
      Resource.Release(rng)
    End If
    
    Return result
  End Function  
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを収めたテーブルを取得する。
  ''' 指定したExcelファイルが開かれていない場合は例外を投げる。
  ''' </summary>
  Private Function GetSheetTable(fullpath As String) As Dictionary(Of String, Object)
    If Not Me.bookTable.ContainsKey(fullpath) Then
      Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
    End If
    
    Return Me.bookTable(fullpath)
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルのブックを取得する。
  ''' 指定したExcelファイルが開かれていない場合は例外を投げる。
  ''' </summary>
  Private Function GetBook(fullpath As String) As Object
    Dim sheetTable As Dictionary(Of String, Object) = GetSheetTable(fullpath)
    If Not sheetTable.ContainsKey(KEY_BOOK) Then
      Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
    End If
    
    Return sheetTable(KEY_BOOK)
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルの指定したシートを取得する。
  ''' 指定したExcelファイルが開かれていない場合、
  ''' 指定した名前のシートが存在しない場合は例外を投げる。
  ''' </summary>
  Private Function GetSheet(fullpath As String, sheetName As String) As Object
    Dim sheetTable As Dictionary(Of String, Object) = GetSheetTable(fullpath)
    Dim sheet As Object = Nothing
    If Not sheetTable.ContainsKey(sheetName) Then
      If Not sheetTable.ContainsKey(KEY_BOOK) Then
        Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
      End If
      Dim book As Object = sheetTable(KEY_BOOK)
      
      ' ブックから指定した名前のシートを探す
      For Each sh As Object In book.worksheets
        If sheetName = sh.Name Then
          sheet = sh
          sheetTable.Add(sheetName, sheet)
          Exit For
        End If
      Next
      
      If sheet Is Nothing Then
        Throw New ArgumentException("指定した名前のExcelシートが見つかりません。 / " & sheetName)
      End If
    Else
      sheet = sheetTable(sheetName)
    End If
    
    Return sheet
  End Function
End Class

End Namespace