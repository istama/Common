'
' 日付: 2016/05/02
'
Imports NUnit.Framework
Imports Common.Threading
Imports System.Threading
Imports System.Collections.Concurrent

<TestFixture> _
Public Class TestVolatileBoolean
	Private recv As New List(Of Integer)
	Private boolList As New List(Of Boolean)
	Private b As New VolatileBoolean(False)
	
	<Test> _
	Public Sub TestClass
		Dim t1 As New Thread(AddressOf Thread1)
		Dim t2 As New Thread(AddressOf Thread2)
		Dim t3 As New Thread(AddressOf Thread3)
		Dim t4 As New Thread(AddressOf Thread4)
		Dim t5 As New Thread(AddressOf Thread5)
		
		recv.Add(5)
		t1.Start()
		t2.Start()
		t3.Start()
		t4.Start()
		t5.Start()
		
		t1.Join()
		
		Assert.AreEqual(True, boolList(0))
		Assert.AreEqual(False, boolList(1))
		Assert.AreEqual(True, boolList(2))
		Assert.AreEqual(False, boolList(3))
		Assert.AreEqual(True, boolList(4))
		
	End Sub
	
	Private Function NewThread(id As Integer) As Action
		Return Sub()
		 		Wait(id)
		 		b.Write(Not b.Read)
		 		boolList.Add(b.Read)
		 		recv.Add(id - 1)
			End Sub
	End Function
	
	Private Sub Thread1()
		Run(1)
	End Sub
	
	Private Sub Thread2()
		Run(2)
	End Sub
	
	Private Sub Thread3()
		Run(3)
	End Sub	
	
	Private Sub Thread4()
		Run(4)
	End Sub
	
	Private Sub Thread5()
		Run(5)
	End Sub
	
	Private Sub Run(id As Integer)
		Wait(id)
		b.Write(Not b.Read)
		boolList.Add(b.Read)
		recv.Add(id - 1)		
	End Sub
	
	Private Sub Wait(id As Integer)
		Do While Not recv(recv.Count - 1) = id
		Loop
	End Sub
	
	Private Sub test()
		
	End Sub
End Class
