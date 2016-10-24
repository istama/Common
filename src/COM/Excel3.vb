'
' 日付: 2016/10/24
'
Option Strict Off

Imports Common.Threading
Imports System.Threading
Imports System.Collections.Concurrent
Imports System.IO

Imports Common.IO

Namespace COM

''' <summary>
''' Excelの起動、アクセス、終了を行う。
''' スレッドセーフ。
''' 既存のExcelクラスから操作方法を変更している。
''' </summary>
Public Class Excel3
  Implements IExcel	
  ''' Excelを起動しアクセスする	
  Private app As App3
  ''' 初期処理を行ったかどうか
  Private initialized As Boolean
  
  ''' <summary>
  ''' マルチスレッド時にExcelのCOMコンポーネントの生成・削除の操作と読み書きの操作を同時に行えないようにするためにロックの役割を果たす。
  ''' つまり、読み書きが行われている間は生成・削除を行うことはできない。
  ''' SyncLockによるロックを使わない理由は、読み書きの操作をSyncLockで囲むとマルチスレッド時に１件ずつしか読み書きを行えないから。
  ''' </summary>
  Private ReadOnly rwLock As New ReaderWriterLock
  
  Sub New()
    initialized = False
    Log.out("create excel")
  End Sub
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' これを行わないとExcelにアクセスできない。
  ''' </summary>
  Public Sub init() Implements IExcel.init
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)
      If Not initialized Then
        app = New App3()
        initialized = True
        Log.out("initialized excel complete")
      End If
    Finally
      Me.rwLock.ReleaseWriterLock
    End Try
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' これを行わないとExcelのCOMコンポーネントが解放されない。
  ''' </summary>
  Public Sub Quit() Implements IExcel.Quit
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)      
      If initialized Then
        app.Quit()
        initialized = False
        Log.out("quit excel complete")
      End If
    Finally
      Me.rwLock.ReleaseWriterLock      
    End Try
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルを開く。
  ''' 開く際には読み込みモードか読み書きモードかを指定する。
  ''' これを行わないとExcelにはアクセスできない。
  ''' </summary>
  Public Sub Open(filepath As String, readMode As Boolean)
    If Not initialized Then Throw New ExcelException("初期処理が実行されていません。")
    
    Me.app.OpenBook(filepath, readMode)
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを解放する。
  ''' </summary>
  ''' <param name="filepath"></param>
  Public Sub Close(filepath As String) Implements IExcel.Close
    If Not initialized Then Throw New ExcelException("初期処理が実行されていません。")
    
    Me.app.CloseBook(filepath)
  End Sub
  
  ''' <summary>
  ''' Excelファイルを読み込む
  ''' </summary>
  ''' <param name="data"></param>
  ''' <returns></returns>
  Public Function Read(data As ExcelData) As String Implements IExcel.Read
    Return Read(data.Filepath, data.SheetName, data.Cell)
  End Function
  
  ''' <summary>
  ''' Excelファイルを読み込む。
  ''' </summary>
  ''' <param name="filepath">読み込むExcelファイルへのパス</param>
  ''' <param name="sheetName">読み込むExcelファイルのシート名</param>
  ''' <param name="cell">読み込むExcelファイルのセルの位置</param>
  ''' <returns>読み込んだ文字列</returns>
  Public Function Read(filepath As String, sheetName As String, cell As Cell) As String Implements IExcel.Read
    Return Access(Function(app) app.Read(filepath, sheetName, cell))
  End Function
  
  ''' <summary>
  ''' Excelファイルに書き込む。
  ''' </summary>
  ''' <param name="data"></param>
  Public Sub Write(data As ExcelData) Implements IExcel.Write
    Write(data.WrittenText, data.Filepath, data.SheetName, data.Cell)
  End Sub
  
  ''' <summary>
  ''' Excelファイルに書き込む
  ''' </summary>
  ''' <param name="writtenText">書き込む文字列</param>
  ''' <param name="filepath">書き込むExcelファイルへのパス</param>
  ''' <param name="sheetName">書き込むExcelファイルのシート名</param>
  ''' <param name="cell">書き込むExcelファイルのセルの位置</param>
  Public Sub Write(writtenText As String, filepath As String, sheetName As String, cell As Cell) Implements IExcel.Write
    Access(
      Function(app)
        app.Write(writtenText, filepath, sheetName, cell)	
        Return Nothing
      End Function)
  End Sub
  
  Private Function Access(f As Func(Of App3, String)) As String
    If Not initialized Then Throw New ExcelException("初期処理が実行されていません。")
    
    Dim result As String = String.Empty
    Try
      rwLock.AcquireReaderLock(System.Threading.Timeout.Infinite)
      If Me.initialized Then
        result = f(app)
      End If
    Catch ex As Exception
      Throw New ExcelException(ex.Message)
    Finally
      rwLock.ReleaseReaderLock()
    End Try
    
    Return result
  End Function
  
End Class

''' <summary>
''' ExcelのCOMコンポーネントを管理する実体となるクラス。
''' このクラスはExcelクラスからのみ呼び出されることを前提として作られている。
''' </summary>
Class App3
  Private ReadOnly excel As Object
  Private ReadOnly workbooks As Object
  Private ReadOnly bookTable As ConcurrentDictionary(Of String, Book3)
  
  Private ReadOnly rwLock As New ReaderWriterLock
  
  Private closed As Boolean = False
  
  Sub New()
    Me.excel = CreateObject("Excel.Application")
    Me.workbooks = excel.WorkBooks
    Me.bookTable = New ConcurrentDictionary(Of String, Book3)
    
    Log.out("execute excel")
  End Sub
  
  ''' <summary>
  ''' 生成したExcelのCOMコンポーネントを全て解放する。
  ''' このメソッドはExcelクラスのQuit()から呼び出されることを前提としており、
  ''' ReadWriteLockに同期の仕組みを依存しているため、他のメソッドからは呼び出すときは注意する。
  ''' </summary>
  Sub Quit()
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)      
      If Not closed Then
        ' Excelファイルを閉じる
        For Each k In Me.bookTable.Keys
          CloseBook(k)
        Next
        
        ' Excelアプリケーションのリソースを解放する
        Resource.Release(Me.workBooks)
        Me.excel.Quit()
        Resource.Release(Me.excel)
        closed = True
        
        Log.out("quit excel")
      End If
    Finally
      Me.rwLock.ReleaseWriterLock     
    End Try
  End Sub
  
  ''' <summary>
  ''' ExcelファイルのCOMコンポーネントを制御するクラスを開く。
  ''' このメソッドはReadWriteLockに同期の仕組みを依存しているため、扱いに注意する。
  ''' </summary>
  Sub OpenBook(filepath As String, readMode As Boolean)
    If closed Then
      Throw New InvalidOperationException("Excelコンポーネントは既に終了しています。")
    End If
  
    Dim fullpath As String = Path.GetFullPath(filepath)
  
    If Not File.Exists(fullpath) Then
      Throw New FileNotFoundException("指定したファイルは存在しません。 / filepath " & filepath)
    End If
    
    If Not Me.bookTable.ContainsKey(fullpath) Then
      Try
        Me.rwLock.AcquireReaderLock(Timeout.Infinite)
        If Not Me.closed Then
          Me.bookTable.TryAdd(fullpath, Book3.GetInstance(Me.workbooks, fullpath, readMode))
        End If
      Finally
        Me.rwLock.ReleaseReaderLock
      End Try
    End If
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを解放する。
  ''' このメソッドはExcelクラスのClose()から呼び出されることを前提としており、
  ''' ReadWriteLockに同期の仕組みを依存しているため、他のメソッドからは呼び出すときは注意する。
  ''' </summary>
  ''' <param name="filepath"></param>
  Sub CloseBook(filepath As String)
    Dim book As Book3 = Nothing
    If Me.bookTable.TryRemove(Path.GetFullPath(filePath), book) Then
      book.Close()
    End If
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルが開かれた状態か判定する。
  ''' </summary>
  Function Opened(filepath As String) As Boolean
    Dim fullpath As String = Path.GetFullPath(filePath)
    
    If Not File.Exists(fullpath) Then
      Throw New FileNotFoundException("指定したファイルは存在しません。 / filepath " & filepath)
    End If
    
    Return bookTable.ContainsKey(fullpath)
  End Function
  
  ''' <summary>
  ''' Excelファイルを読み込む。
  ''' </summary>
  Function Read(filepath As String, sheetName As String, cell As Cell) As String
    If Not Opened(filepath) Then Throw New ExcelException("指定したファイルは開かれていません。 / " & filepath)
    
    Return Open(filepath, Function(book) book.Read(sheetName, cell))
  End Function
  
  ''' <summary>
  ''' Excelファイルに書き込む。
  ''' </summary>
  Sub Write(writtenText As String, filepath As String, sheetName As String, cell As Cell) 
    If Not Opened(filepath) Then Throw New ExcelException("指定したファイルは開かれていません。 / " & filepath)
    
    Open(
      filepath,
      Function(book)
        book.Write(writtenText, sheetName, cell)
        Return Nothing
      End Function)
  End Sub
  
  Private Function Open(filepath As String, f As Func(Of Book3, String)) As String
    If Not Me.closed Then Throw New InvalidOperationException("Excelアプリケーションはすでに閉じられています。")
    
    Dim result As String = String.Empty
    
    Dim book As Book3 = Nothing
    If Me.bookTable.TryGetValue(Path.GetFullPath(filepath), book) Then
      result = f(book)
    End If
    
    Return result
  End Function
  
End Class

Class Book3
  Private ReadOnly book As Object
  Private ReadOnly worksheets As Object
  Private ReadOnly sheetTable As ConcurrentDictionary(Of String, Sheet3)
  
  Private ReadOnly filepath As String
  
  Private ReadOnly rwLock As New ReaderWriterLock
  Private closed As Boolean = False
  
  Public Shared Function GetInstance(workbooks As Object, filepath As String, readMode As String) As Book3
    Dim fullpath As String = Path.GetFullPath(filePath)
    
    If Not File.Exists(fullpath) Then
      Throw New FileNotFoundException("指定したExcelファイルは存在しません。 / filepath " & filepath)
    End If
      
    Return New Book3(workbooks.Open(fullPath, Nothing, readMode), filepath)
  End Function
  
  Private Sub New(book As Object, filepath As String)
    Me.book = book
    Me.worksheets = book.worksheets
    Me.filepath = filepath
    Me.sheetTable = New ConcurrentDictionary(Of String, Sheet3)
    
    Log.out("create book / file name: " & filepath)
  End Sub
  
  Sub Close()
    Try
      Me.rwLock.AcquireWriterLock(Timeout.Infinite)
      If Not closed Then 
        For Each k In sheetTable.Keys
          CloseSheet(k)				
        Next
        
        Resource.Release(worksheets)
        book.Close(False)
        Resource.Release(book)
        closed = True
        
        Log.out("closed book / file name: " & filepath)
      End If
    Finally 
      Me.rwLock.ReleaseWriterLock
    End Try
  End Sub
  
  Sub CloseSheet(sheetName As String)
    Dim sheet As Sheet3 = Nothing
    If Me.sheetTable.TryRemove(sheetName, sheet) Then
      sheet.Close()
    End If
  End Sub
  
  Function Read(sheetName As String, cell As Cell) As String
    Return OpenSheet(sheetName, Function(sheet) sheet.Read(cell))
  End Function
  
  Sub Write(writtenText As String, sheetName As String, cell As Cell)
    OpenSheet(
      sheetName,
      Function(sheet)
        sheet.Write(writtenText, cell)
        Return Nothing
      End Function)
  End Sub
  
  Sub Save()
    Me.book.Save()
  End Sub
  
  ''' <summary>
  ''' 指定した名前のシートを開いて処理を行う。
  ''' </summary>
  Private Function OpenSheet(sheetName As String, f As Func(Of Sheet3, String)) As String
    If closed Then Throw New InvalidOperationException("このブックは既に閉じられています。")
    
    Dim sheet As Sheet3 = Nothing
    
    If Not Me.sheetTable.TryGetValue(sheetName, sheet) Then
      Try
        Me.rwLock.AcquireReaderLock(Timeout.Infinite)
        If Not Me.closed Then
          sheet = Sheet3.GetInstance(Me.worksheets, sheetName)
          Me.sheetTable.TryAdd(sheetName, sheet)
        End IF
      Finally
        Me.rwLock.ReleaseReaderLock
      End Try
    End If
    
    If sheet IsNot Nothing Then
      Return f(sheet)
    Else
      Return String.Empty
    End If
  End Function
  
End Class

''' <summary>
''' Excelファイルのシートを表すクラス。
''' セルを指定して読み書きできる。
''' スレッドセーフ。
''' </summary>
Class Sheet3
  Private ReadOnly sheet As Object
  
  Private ReadOnly rwLock As New ReaderWriterLock
  Private closed As Boolean = False
  
  ''' <summary>
  ''' 指定した名前のシートのインスタンスをワークシートから取得する。
  ''' </summary>
  Public Shared Function GetInstance(worksheets As Object, sheetName As String) As Sheet3
    If worksheets Is Nothing Then Throw New ArgumentNullException("worksheet is null")
    If sheetName Is Nothing Then Throw New ArgumentNullException("sheetName is null")
    
    For Each sh As Object In worksheets
      If sheetName = sh.Name Then
        Return New Sheet3(sh)
      End If
    Next
    
    Throw New ArgumentException("指定した名前のExcelシートが見つかりません。 / " & sheetName)
  End Function
  
  Private Sub New(sheet As Object)
    Me.sheet = sheet
    
    Log.out("create sheet / sheet name: " & Me.sheet.Name)
  End Sub
  
  Sub Close()
    Me.rwLock.AcquireWriterLock(Timeout.Infinite)
    Try
      If Not Me.closed Then
        Resource.Release(Me.sheet)
        closed = True
        Log.out("closed sheet / sheet name: " & Me.sheet.Name)
      End If
    Finally
      Me.rwLock.ReleaseWriterLock()			
    End Try
  End Sub
  
  Function Read(cell As Cell) As String
    Return OpenRange(cell, Function(rng) rng.Value)
  End Function
  
  Public Sub Write(text As String, cell As Cell)
    OpenRange(
      cell,
      Function(rng)
        rng.Value = text
        Return Nothing
      End Function)
  End Sub
  
  Private Function OpenRange(cell As Cell, f As Func(Of Object, String)) As String
    If closed Then Throw New InvalidOperationException("このシートは既に閉じられています。 / sheet name: " & Me.sheet.Name)
    
    Dim result As String = String.Empty
    
    Me.rwLock.AcquireWriterLock(Timeout.Infinite)
    Try
      If Not Me.closed Then
        Dim rng As Object = Me.sheet.Range(cell.Point)
        If rng IsNot Nothing Then
          result = f(rng)
          Resource.Release(rng)
        End If
      End If
    Finally
      Me.rwLock.ReleaseWriterLock()
    End Try
    
    Return result
  End Function
  
End Class

End Namespace
