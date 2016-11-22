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
''' このクラスは同一のExcelファイルの開く/閉じるの操作を読み込み/書き込みの操作と同じスレッドで行えばスレッドセーフである。
''' </summary>
Public Class Excel4
  Private ReadOnly KEY_BOOK As String = "__book__*"
  
  Private initialized As Boolean = False
  
  Private excel As Object
  Private workbooks As Object
  Private bookTable As ConcurrentDictionary(Of String, ConcurrentDictionary(Of String, Object))
  
  ''' <summary>
  ''' 初期処理と終了処理を行うときはWriterLock、
  ''' 開く、閉じる、読み込み、書き込みを行うときはReaderLockをかける。
  ''' </summary>
  Private ReadOnly rwLock As New ReaderWriterLock
  
  Public Sub New()
  End Sub
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' </summary>
  Public Sub Init()
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)
      If Not initialized Then
        Me.excel = CreateObject("Excel.Application")
        Me.workbooks = excel.WorkBooks
        Me.bookTable = New ConcurrentDictionary(Of String, ConcurrentDictionary(Of String, Object))
        
        initialized = True
        Log.out("initialized excel")
      End If
    Finally
      Me.rwLock.ReleaseWriterLock()
    End Try
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' </summary>
  Public Sub Quit()
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)
      If initialized Then
        Me.bookTable.Keys.ForEach(Sub(path) Close(path))
        Resource.Release(Me.workBooks)
        Me.excel.Quit()
        
        initialized = False
        Log.out("quit excel")
      End If
    Finally
      Me.rwLock.ReleaseWriterLock
    End Try
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
      Dim sheetTable As New ConcurrentDictionary(Of String, Object)
      Try
        Me.rwLock.AcquireReaderLock(Timeout.Infinite)       
        SyncLock Me
          If Not Me.initialized Then
            Log.out("open book / filepath: " & filepath & " readMode: " & readMode.ToString)
            sheetTable.TryAdd(KEY_BOOK, workbooks.Open(fullPath, Nothing, readMode))
            Me.bookTable.TryAdd(fullpath, sheetTable)
          End If
        End SyncLock
      Finally
        Me.rwLock.ReleaseReaderLock
      End Try
    End If  
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルを閉じる。
  ''' </summary>
  Public Sub Close(filepath As String)
    If filepath Is Nothing Then Throw New ArgumentNullException("filepath is null")
    
    Dim sheetTable As ConcurrentDictionary(Of String, Object) = Nothing
    If Me.bookTable.TryRemove(Path.GetFullPath(filePath), sheetTable) Then
      Try
        Me.rwLock.AcquireReaderLock(Timeout.Infinite)  
        SyncLock Me
          ' シートを解放する
  '        sheetTable.Keys.
  '          Where(Function(k) k <> KEY_BOOK).
  '          ForEach(
  '            Sub(k)
  '              Dim sheet As Object = Nothing
  '              If sheetTable.TryRemove(k, sheet) Then
  '                Resource.Release(sheet)
  '              End If
  '            End Sub)
          
          Dim book As Object = Nothing
          If sheetTable.TryRemove(KEY_BOOK, book) Then
            Resource.Release(book.worksheets)
            book.Close(False)
            Resource.Release(book)
            Log.out("book close / filepath: " & filepath)
          End If
        End SyncLock
      Finally
        Me.rwLock.ReleaseReaderLock        
      End Try
    End If
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのセルを読み込む。
  ''' </summary>
  Public Function Read(filepath As String, sheetName As String, cell As Cell) As String
    If filepath  Is Nothing Then Throw New ArgumentNullException("filepath is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    If cell      Is Nothing Then Throw New ArgumentNullException("cell is null")
    
    Dim fullpath As String = Path.GetFullPath(filepath)
    
    Try
      Me.rwLock.AcquireReaderLock(Timeout.Infinite)
      Return AccessCell(fullpath, sheetName, cell, Function(rng) rng.Value)
    Finally
      Me.rwLock.ReleaseReaderLock
    End Try
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルのセルに書き込む。
  ''' </summary>
  Public Sub Write(text As String, filepath As String, sheetName As String, cell As Cell)
    If filepath  Is Nothing Then Throw New ArgumentNullException("filepath is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    If cell      Is Nothing Then Throw New ArgumentNullException("cell is null")
    
    Dim fullpath As String = Path.GetFullPath(filepath)
    
    Try
      Me.rwLock.AcquireReaderLock(Timeout.Infinite)
      AccessCell(fullpath, sheetName, cell,
        Function(rng)
          rng.Value = text
          Return Nothing
        End Function)
      
      Dim book As Object = GetBook(fullpath)
      SyncLock Me
        book.Save()
      End SyncLock
    Finally
      Me.rwLock.ReleaseReaderLock
    End Try
  End Sub
  
  ''' <summary>
  ''' Excelのセルにアクセスする。
  ''' </summary>
  Private Function AccessCell(fullpath As String, sheetName As String, cell As Cell, f As Func(Of Object, String)) As String
    If Not Me.initialized Then Throw New InvalidOperationException("初期処理が行われていません。")
    
    Dim sheet As Object = GetSheet(fullpath, sheetName)
    
    Dim result As String = String.Empty
    SyncLock Me
      Dim rng As Object = sheet.Range(cell.Point)
      If rng IsNot Nothing Then
        result = f(rng)
        Resource.Release(rng)
      End If
    End SyncLock
    
    Return result
  End Function  
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを収めたテーブルを取得する。
  ''' 指定したExcelファイルが開かれていない場合は例外を投げる。
  ''' </summary>
  Private Function GetSheetTable(fullpath As String) As ConcurrentDictionary(Of String, Object)
    Dim sheetTable As ConcurrentDictionary(Of String, Object) = Nothing
    If Not Me.bookTable.TryGetValue(fullpath, sheetTable) Then
      Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
    End If
    
    Return sheetTable
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルのブックを取得する。
  ''' 指定したExcelファイルが開かれていない場合は例外を投げる。
  ''' </summary>
  Private Function GetBook(fullpath As String) As Object
    Dim sheetTable As ConcurrentDictionary(Of String, Object) = GetSheetTable(fullpath)
    Dim book As Object = Nothing
    If Not sheetTable.TryGetValue(KEY_BOOK, book) Then
      Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
    End If
    
    Return book
  End Function
  
  ''' <summary>
  ''' 指定したExcelファイルの指定したシートを取得する。
  ''' 指定したExcelファイルが開かれていない場合、
  ''' 指定した名前のシートが存在しない場合は例外を投げる。
  ''' </summary>
  Private Function GetSheet(fullpath As String, sheetName As String) As Object
    Dim sheetTable As ConcurrentDictionary(Of String, Object) = GetSheetTable(fullpath)
    
    Dim sheet As Object = Nothing
    If Not sheetTable.TryGetValue(sheet, sheetName) Then
      Dim book As Object = Nothing
      If Not sheetTable.TryGetValue(KEY_BOOK, book) Then
        Throw New ArgumentException("指定されたExcelファイルは開かれていません。 / " & fullpath)
      End If
      
      SyncLock Me
        ' ブックから指定した名前のシートを探す
        For Each sh As Object In book.worksheets
          If sheetName = sh.Name Then
            sheet = sh
            sheetTable.TryAdd(sheetName, sheet)
            Exit For
          End If
        Next
      End SyncLock
      
      If sheet Is Nothing Then
        Throw New ArgumentException("指定した名前のExcelシートが見つかりません。 / " & sheetName)
      End If
    End If
    
    Return sheet
  End Function
End Class

End Namespace