'
' 日付: 2016/05/02
'
Imports NUnit.Framework
Imports Common.Threading
Imports System.Threading

<TestFixture> _
Public Class TestLRSemaphore
	Private s As LRSemaphore = New LRSemaphore()
	
	Private Class Messaging
		Private msgList As New List(Of String)
		Private s As LRSemaphore
		
		Sub New(semaphore As LRSemaphore)
			s = semaphore
		End Sub
		
		Sub LeftThread()
			Do While Not s.IncrementLeftIfPass()
			Loop
			send("L")
			s.Decrement()
		End Sub
		
		Sub RightThread()
			Do While Not s.IncrementRightIfPass()
			Loop
			send("R")
			s.Decrement()			
		End Sub
		
		Sub send(msg As String)
			SyncLock Me
				msgList.Add(msg)
			End SyncLock
		End Sub
		
		Function recv() As String
			SyncLock Me
				Dim str As String = String.Empty
				msgList.ForEach(Sub(m) str += m)
				Return str
			End SyncLock
		End Function
	End Class
	
	<Test> _
	Public Sub TestLeft
		' Leftセマフォの処理が終わるまでRightセマフォは待機する
		Dim msg As Messaging =
			StartThread(
				s,
				Sub() s.IncrementLeftIfPass()
			)
		
		Assert.AreEqual("LLLRRR", msg.recv)
	End Sub
	
	<Test> _
	Public Sub TestRight
		' Rightセマフォの処理が終わるまでLeftセマフォは待機する
		Dim msg As Messaging =
			StartThread(
				s,
				Sub() s.IncrementRightIfPass()
			)
		
		Assert.AreEqual("RRRLLL", msg.recv)
	End Sub
	
	Private Function StartThread(s As LRSemaphore, increment As Action) As Messaging
		Dim msg As New Messaging(s) 
			
		increment()
		
		Dim threadList As New List(Of Thread)
		For i As Integer = 0 To 2
			threadList.Add(New Thread(AddressOf msg.LeftThread))
			threadList.Add(New Thread(AddressOf msg.RightThread))
		Next
		
		' LRSemaphoreが正しく動作すれば、LeftかRightのどちらかのスレッドが先に実行される
		For Each th As Thread In threadList
			th.Start()
		Next
		
		' すべてのスレッドが開始するまで待機
		Do Until threadList.TrueForAll(Function(t) t.ThreadState = ThreadState.Running OrElse t.ThreadState = ThreadState.Stopped)
		Loop
		
		s.decrement()

		threadList.ForEach(Sub(t) t.Join())
		
		Return msg
	End Function
	

	

	
End Class
