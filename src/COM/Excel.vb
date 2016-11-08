'
' 日付: 2016/05/01
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
''' </summary>
Public Class Excel
  Implements IExcel	
  ''' Excelを起動しアクセスする	
  Private app As App
  ''' 初期処理を行ったかどうか
  Private initialized As Boolean
  ''' 安全モードで実行するかどうか
  Private safetyMode As Boolean
  
  ''' <summary>
  ''' マルチスレッド時にExcelのCOMコンポーネントの生成・削除の操作と読み書きの操作を同時に行えないようにするためにロックの役割を果たす。
  ''' つまり、読み書きが行われている間は生成・削除を行うことはできない。
  ''' SyncLockによるロックを使わない理由は、読み書きの操作をSyncLockで囲むとマルチスレッド時に１件ずつしか読み書きを行えないから。
  ''' </summary>
  'Private semaphore As LRSemaphore = New Threading.LRSemaphore()
  Private rwLock As ReaderWriterLock
  
  Sub New(Optional safetyMode As Boolean=True)
    Me.safetyMode = safetyMode
    initialized = False
    'semaphore = New Threading.LRSemaphore()
    rwLock = New ReaderWriterLock()
    Log.out("create excel")
  End Sub
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' これを行わないとExcelにアクセスできない。
  ''' </summary>
  Public Sub init() Implements IExcel.init
    Control(
      Sub()
      If Not initialized Then
        app = New App()
        initialized = True
        Log.out("initialized excel complete")
      End If
    End Sub,
    "他のスレッドからアクセスが行われているため初期処理を実行できません。")
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' これを行わないとExcelのCOMコンポーネントが解放されない。
  ''' </summary>
  Public Sub Quit() Implements IExcel.Quit
    Control(
      Sub()
      If initialized Then
        app.Quit()
        initialized = False
        Log.out("quit excel complete")
      End If
    End Sub,
    "他のスレッドからアクセスが行われているため終了処理を実行できません。")
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを解放する。
  ''' </summary>
  ''' <param name="filepath"></param>
  Public Sub Close(filepath As String) Implements IExcel.Close
    Control(
      Sub()	app.CloseBook(filepath),
      "他のスレッドからアクセスが行われているため閉じられません。")
  End Sub
  
  Private Sub Control(f As Action, errMsg As String)
    ' Read() Write() など、Excelへのアクセス操作が行われている間にこのメソッドを実行するとFalseが返る。
    ' そうでない場合はTrueが返り、Decrement()を実行するまで、Excelへのアクセス操作をブロックする。
    'If semaphore.IncrementLeftIfPass() Then
    rwLock.AcquireWriterLock(System.Threading.Timeout.Infinite)
    Try
      SyncLock Me
        f()
      End SyncLock
    Finally
      'semaphore.Decrement()
      rwLock.ReleaseWriterLock()
    End Try
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
    Return Access(Function() app.Read(filepath, False, sheetName, cell))
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
      Function()
      app.Write(writtenText, filepath, False, sheetName, cell)	
      Return Nothing
    End Function)
    ' 安全モードで起動している場合は、書き込み後にブックを閉じる
    If Me.safetyMode Then
      Me.Close(filepath)
    End If
  End Sub
  
  Private Function Access(f As Func(Of String)) As String
    ' Quit() Close() など、ExcelのCOMコンポーネントへの操作が行われている間にこのメソッドを実行するとFalseが返る。
    ' そうでない場合はTrueが返り、Decrement()を実行するまで、COMコンポーネントの操作をブロックする。
    'If semaphore.IncrementRightIfPass() Then
    rwLock.AcquireReaderLock(System.Threading.Timeout.Infinite)
    Try
      SyncLock Me
        If Not initialized Then
          Throw New ExcelException("初期処理が実行されていません。")
        End If
      End SyncLock
      
      Return f()
    Finally
      'semaphore.Decrement()
      rwLock.ReleaseReaderLock()
    End Try
    'Else
    '	Return ""
    'End If		
  End Function
  
End Class

''' <summary>
''' Excwlにアクセスするための情報をまとめたクラス。
''' </summary>
Public Class ExcelData
  Private _filepath As String
  Public ReadOnly Property Filepath() As String
    Get
      Return _filepath
    End Get
  End Property
  
  Private _sheetName As String
  Public ReadOnly Property SheetName() As String
    Get
      Return _sheetName
    End Get
  End Property
  
  Private _cell As Cell
  Public ReadOnly Property Cell() As Cell
    Get
      Return _cell
    End Get
  End Property
  
  Private _writtenText As String
  Public ReadOnly Property WrittenText() As String
    Get
      Return _writtenText
    End Get
  End Property
  
  Public Sub New(writtenText As String, filepath As String, sheetName As String, cell As Cell)
    _writtenText = writtenText
    _filepath = filepath
    _sheetName = sheetName
    _cell = cell
  End Sub
  
  Public Overrides Function ToString() As String
    Return String.Format("filepath: {0} sheetName: {1} {2} text: {3}", _filepath, _sheetName, _cell, _writtenText)
  End Function
  
  Public Shared Operator =(ByVal d1 As ExcelData, ByVal d2 As ExcelData) As Boolean
    Return _
      d1.WrittenText = d2.WrittenText AndAlso _
      d1.Filepath    = d2.Filepath    AndAlso _
      d1.SheetName   = d2.SheetName   AndAlso _
      d1.Cell        = d2.Cell
  End Operator
  
  Public Shared Operator <>(ByVal d1 As ExcelData, ByVal d2 As ExcelData) As Boolean
    Return _
      d1.WrittenText <> d2.WrittenText OrElse _
      d1.Filepath    <> d2.Filepath    OrElse _
      d1.SheetName   <> d2.SheetName   OrElse _
      d1.Cell        <> d2.Cell
  End Operator
End Class

''' <summary>
''' ExcelのCOMコンポーネントを管理する実体となるクラス。
''' このクラスはExcelクラスからのみ呼び出されることを前提として作られている。
''' 同期の仕組みはExcelクラスに依存しているため、このクラス単体ではスレッドセーフにならない。
''' </summary>
Class App
  Private excel As Object
  Private workbooks As Object
  Private bookTable As ConcurrentDictionary(Of String, Book)
  
  Private closed As Boolean
  
  Sub New()
    excel = CreateObject("Excel.Application")
    workbooks = excel.WorkBooks
    bookTable = New ConcurrentDictionary(Of String, Book)
    closed = False
    
    Log.out("execute excel")
  End Sub
  
  ''' <summary>
  ''' 生成したExcelのCOMコンポーネントを全て解放する。
  ''' このメソッドはExcelクラスのQuit()から呼び出されることを前提としており、
  ''' ReadWriteLockに同期の仕組みを依存しているため、他のメソッドからは呼び出すときは注意する。
  ''' </summary>
  Sub Quit()
    SyncLock Me
      If Not closed Then
        ' Excelファイルを閉じる
        For Each k In bookTable.Keys
          CloseBook(k)
        Next
        
        ' Excelアプリケーションのリソースを解放する
        Resource.Release(workBooks)
        excel.Quit()
        Resource.Release(excel)
        closed = True
        
        Log.out("released excel")
      End If
    End SyncLock
  End Sub
  
  ''' <summary>
  ''' 指定したExcelファイルのCOMコンポーネントを解放する。
  ''' このメソッドはExcelクラスのClose()から呼び出されることを前提としており、
  ''' ReadWriteLockに同期の仕組みを依存しているため、他のメソッドからは呼び出すときは注意する。
  ''' </summary>
  ''' <param name="filepath"></param>
  Sub CloseBook(filepath As String)
    Dim fullpath As String = Path.GetFullPath(filePath)
    
    If bookTable.ContainsKey(fullpath) Then
      SyncLock Me
        If bookTable.ContainsKey(fullpath) Then
          bookTable(fullpath).Close()
          bookTable.TryRemove(fullpath, Nothing)
        End If
      End SyncLock
    End If
  End Sub
  
  ''' <summary>
  ''' Excelファイルを読み込む。
  ''' このメソッドはReadWriteLockに同期の仕組みを依存しているため、扱いに注意する。
  ''' </summary>
  ''' <param name="filepath"></param>
  ''' <param name="readMode"></param>
  ''' <param name="sheetName"></param>
  ''' <param name="cell"></param>
  ''' <returns></returns>
  Function Read(filepath As String, readMode As Boolean, sheetName As String, cell As Cell) As String
    Return OpenBook(filepath, readMode).Read(sheetName, cell)
  End Function
  
  ''' <summary>
  ''' Excelファイルに書き込む。
  ''' このメソッドはReadWriteLockに同期の仕組みを依存しているため、扱いに注意する。
  ''' </summary>
  ''' <param name="writtenText"></param>
  ''' <param name="filepath"></param>
  ''' <param name="readMode"></param>
  ''' <param name="sheetName"></param>
  ''' <param name="cell"></param>
  Sub Write(writtenText As String, filepath As String, readMode As Boolean, sheetName As String, cell As Cell) 
    Dim book As Book = OpenBook(filepath, readMode)
    book.Write(writtenText, sheetName, cell)
    book.Save
  End Sub
  
  ''' <summary>
  ''' ExcelファイルのCOMコンポーネントを制御するクラスを返す。
  ''' このメソッドはReadWriteLockに同期の仕組みを依存しているため、扱いに注意する。
  ''' </summary>
  ''' <param name="filepath"></param>
  ''' <param name="readMode"></param>
  ''' <returns></returns>
  Private Function OpenBook(filepath As String, readMode As Boolean) As Book
    If closed Then
      Throw New InvalidOperationException("Excelコンポーネントは既に終了しています。")
    End If
    
    Dim fullpath As String = Path.GetFullPath(filePath)
    Dim book As Book = Nothing
    
    If bookTable.TryGetValue(fullpath, book) Then
      Return book
    Else
      If Not File.Exists(fullpath) Then
        Throw New FileNotFoundException("指定したファイルは存在しません。 / filepath " & filepath)
      End If
      
      book = New Book(workbooks.Open(fullPath, Nothing, readMode), fullpath)
      bookTable.TryAdd(fullpath, book)
      
      Return book
    End If
  End Function
End Class

Class Book
  Private book As Object
  Private worksheets As Object
  Private sheetTable As ConcurrentDictionary(Of String, Sheet)
  
  Private filepath As String
  Private closed As Boolean
  
  Sub New(book As Object, filepath As String)
    Me.book = book
    Me.worksheets = book.worksheets
    Me.filepath = filepath
    Me.sheetTable = New ConcurrentDictionary(Of String, Sheet)
    Me.closed = False
    
    Log.out("create book /" & filepath)
  End Sub
  
  Sub Close()
    SyncLock Me
      If Not closed Then 
        For Each k In sheetTable.Keys
          CloseSheet(k)				
        Next
        
        Resource.Release(worksheets)
        book.Close(False)
        Resource.Release(book)
        closed = True
        
        Log.out("closed book / " & filepath)
      End if
    End SyncLock
  End Sub
  
  Sub CloseSheet(sheetName As String)
    If sheetTable.ContainsKey(sheetName) Then
      SyncLock Me
        If sheetTable.ContainsKey(sheetName) Then
          sheetTable(sheetName).Close()
          sheetTable.TryRemove(sheetName, Nothing)
        End If
      End SyncLock
    End If
  End Sub
  
  Function Read(sheetName As String, cell As Cell) As String
    Return OpenSheet(sheetName).Read(cell)
  End Function
  
  Sub Write(writtenText As String, sheetName As String, cell As Cell)
    OpenSheet(sheetName).Write(writtenText, cell)
  End Sub
  
  Sub Save()
    Me.book.Save()
  End Sub
  
  Private Function OpenSheet(sheetName As String) As Sheet
    If closed Then 
      Throw New InvalidOperationException("このブックは既に閉じられています。")
    End If
    
    Dim sh As Sheet = Nothing
    
    If sheetTable.TryGetValue(sheetName, sh) Then
      Return sh
    Else
      sh = GetSheet(sheetName)
      If sh IsNot Nothing Then
        sheetTable.TryAdd(sheetName, sh)
      Else
        Throw New ArgumentException("存在しないワークシートです: " & sheetName)
      End If
      
      Return sh
    End If
  End Function
  
  Private Function GetSheet(sheetName As String) As Sheet
    Dim i As Integer = 1
    For Each sh As Object In worksheets
      If sheetName = sh.Name Then
        Return New Sheet(sh, sheetName)
      End If
      i += 1
    Next
    
    Return Nothing
  End Function
End Class

Class Sheet
  Private sheet As Object
  Private sheetName As String
  
  Private closed As Boolean
  Private rwLock As ReaderWriterLock
  
  Sub New(sheet As Object, sheetName As String)
    Me.sheet = sheet
    Me.sheetName = sheetName
    Me.closed = False
    Me.rwLock = New ReaderWriterLock()
    
    Log.out("create sheet / " & sheetName)
  End Sub
  
  Sub Close()
    rwLock.AcquireWriterLock(Timeout.Infinite)
    Try
      SyncLock Me
        If Not closed Then
          Resource.Release(sheet)
          closed = True
          Log.out("closed sheet " & sheetName)
        End If
      End SyncLock
    Finally
      rwLock.ReleaseWriterLock()			
    End Try
  End Sub
  
  Function Read(cell As Cell) As String
    If closed Then
      Throw New InvalidOperationException("このシートは既に閉じられています。")
    End If
    
    rwLock.AcquireReaderLock(Timeout.Infinite)
    Try
      If closed = False Then
        Dim rng As Object = GetRange(cell)
        Dim value As String = rng.Value
        Resource.Release(rng)
        
        If value IsNot Nothing Then
          Return value
        Else
          Return String.Empty
        End If
      Else
        Return String.Empty
      End If
    Finally
      rwLock.ReleaseReaderLock()
    End Try
  End Function
  
  Public Sub Write(text As String, cell As Cell)
    If closed Then
      Throw New InvalidOperationException("このシートは既に閉じられています。")
    End If
    
    rwLock.AcquireWriterLock(Timeout.Infinite)
    Try
      If closed = False Then
        Dim rng As Object = GetRange(cell)
        If rng IsNot Nothing Then
          Log.out("write in Sheet / " & text & " " & cell.ToString)
          rng.Value = text
          Resource.Release(rng)
        End If
      End If
    Finally
      rwLock.ReleaseWriterLock()
    End Try
  End Sub
  
  Private Function GetRange(cell As Cell) As Object
    Return sheet.Range(cell.Point)
  End Function
End Class

End Namespace