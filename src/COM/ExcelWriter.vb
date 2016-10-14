'
' 日付: 2016/05/11
'
'
Imports System.Threading
Imports System.Collections.Concurrent

Imports Common.IO

Namespace COM

''' <summary>
''' Excelに非同期に書き込むためのラッパークラス。
''' Excelへの書き込みは別スレッドで行われるため、書き込みメソッドは呼出し後に即座に処理を戻す。
''' 書き込み中にエラーが発生した場合は、ThreadExceptionEventHandlerにセットされた処理が実行される。
''' </summary>
Public Class ExcelWriter
  '''	Excelにアクセスするクラス
  Private excel As IExcel
  ''' 書き込みデータを一時保存するキュー
  Private queue As BlockingCollection(Of ExcelData)
  ''' 書き込み操作を行うスレッド	
  Private thread As Thread
  
  ''' Excelにアクセス中に例外が発生した場合に実行されるハンドラ
  Private _ThreadExceptionEventHandler As Action(Of Exception)
  Public WriteOnly Property ThreadExceptionEventHandler As Action(Of Exception)
    Set (handler As Action(Of Exception))
      _ThreadExceptionEventHandler = handler
    End Set
  End Property
  
  Public Sub New(excel As IExcel)
    Me.excel = excel
    _ThreadExceptionEventHandler = Sub(ex) Throw ex
    queue = Nothing
  End Sub
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' これを実行しないと書き込めない。
  ''' </summary>
  Public Sub Init()
    SyncLock Me
      If queue Is Nothing OrElse queue.IsAddingCompleted Then
        queue = New BlockingCollection(Of ExcelData)
        thread = New Thread(AddressOf WriteTask)
        thread.IsBackground = True ' メインスレッドが終了したときにこのスレッドをabortさせる
        thread.Start               ' 書き込みスレッドを実行
      End If
    End SyncLock
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' キューにたまっているデータを全て書き込んでから書き込みスレッドを終了させる。
  ''' </summary>
  Public Sub Quit()
    SyncLock Me
      If queue IsNot Nothing AndAlso queue.IsAddingCompleted = False Then
        queue.CompleteAdding ' キューにこれ以上要素を追加できないようにする
        thread.Join()
      End If
    End SyncLock
  End Sub
  
  ''' <summary>
  ''' キューにたまっている処理が終了するまで待つ。
  ''' </summary>
  Public Sub Wait()
    While queue.Count > 0 
      ' 0になるまでループ
    End While
  End Sub
  
  ''' <summary>
  ''' Excelに非同期で書き込む。
  ''' 別スレッドで書き込まれるため、このメソッドは即座に処理を戻す。
  ''' Init()が呼び出されていない場合は例外を投げる。
  ''' </summary>
  ''' <param name="writtenText"></param>
  ''' <param name="filepath"></param>
  ''' <param name="sheetName"></param>
  ''' <param name="cell"></param>
  Public Sub AsyncWrite(writtenText As String, filepath As String, sheetName As String, cell As Cell)
    SyncLock Me
      If queue Is Nothing Then
        Throw New InvalidOperationException("開始処理が実行されていません。")
      ElseIf queue.IsAddingCompleted
        Throw New InvalidOperationException("キューは閉じられました。")
      End If
    End SyncLock
    
    ' 書き込みデータをキューに追加する
    queue.Add(New ExcelData(writtenText, filepath, sheetName, cell))
  End Sub
  
  ''' <summary>
  ''' 書き込み用タスク。
  ''' </summary>
  Private Sub WriteTask()
    Dim data As ExcelData = Nothing
    Dim nextData As ExcelData = Nothing
    
    ' キューが閉じられるまでループする
    While queue.IsCompleted = False
      Try
        ' 先頭のデータを取得。データがない場合は追加されるまでロックされる
        If data Is Nothing Then
          data = queue.Take()
          Log.out("take data in ExcelWritter / " & data.ToString)
        End If
        
        ' 書き込み先が同じデータが連続する場合は、最新のデータが取得されるまでループする
        Do While queue.TryTake(nextData) AndAlso data.Cell = nextData.Cell
          data = nextData
          nextData = Nothing
        Loop
      Catch ex As InvalidOperationException
        ' queueがCompleted状態のときにTake()を呼び出したりするとこの例外が発生する
      End Try
      
      Try
        ' データを書き込む
        If data IsNot Nothing Then
          Log.out("write data in ExcelWritter / " & data.ToString)
          excel.Write(data)
        End If
      Catch ex As Exception
        ' 例外が発生した場合、データをキューの先頭に戻し、ハンドラを実行する
        AddToHeadInQueue(data)
        _ThreadExceptionEventHandler(ex)
      End Try
      
      ' 余分に取得した次のデータをセットする
      data = nextData
    End While
  End Sub
  
  ''' <summary>
  ''' キューの先頭にデータをセットする。
  ''' </summary>
  ''' <param name="data"></param>
  Private Sub AddToHeadInQueue(data As ExcelData)
    Dim list As New List(Of ExcelData)
    list.Add(data)
    
    SyncLock Me
      If Not queue.IsAddingCompleted Then
        Do While queue.Count > 0
          list.Add(queue.Take())
        Loop
        
        For Each d As ExcelData In list
          queue.Add(d)
        Next
      End If
    End SyncLock
  End Sub
  
End Class

End Namespace