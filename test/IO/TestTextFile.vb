'
' 日付: 2016/04/27
'
Imports NUnit.Framework
Imports Common.IO
Imports System.IO

<TestFixture> _
Public Class TestTextFile
	Dim filepath As String = "./test.txt"	
	Dim file As New TextFile(filepath, System.Text.Encoding.Default)
	
	<Test> _
	Public Sub Create
		' ファイルを作成する
		Assert.True(file.Create())
		Assert.IsTrue(System.IO.File.Exists(filepath))
		
		' ファイルが存在する場合はなにもしない
		file.Write("abcde")
		Assert.False(file.Create())
		Assert.AreEqual("abcde", Read(filepath)(0))
		
		Delete(filepath)
	End Sub
	
	<Test> _
	Public Sub Reset
		' ファイルを作成する
		file.Reset()
		Assert.IsTrue(System.IO.File.Exists(filepath))
		
		file.Write("abcde")
		file.Write("defghi")
		
		file.Reset()
		Assert.AreEqual(0, Read(filepath).Count)
		
		Delete(filepath)
	End Sub
	
	<Test> _
	Public Sub Write
		' ファイルに文字列を書き込む
		Dim text As String = "あいうえお"
		file.Write(text)
		Assert.AreEqual(text, Read(filepath)(0))
		
		' ファイルに文字列を上書き
		Dim text2 As String = "かきくけこ"
		file.Write(text2)
		Assert.AreEqual(text2, Read(filepath)(0))
		
		Delete(filepath)
	End Sub
	
	<Test> _
	Public Sub Append
		' ファイルに文字列を書き込む
		Dim text As String = "あいうえお"
		file.Append(text)
		Assert.AreEqual(text, Read(filepath)(0))
		
		' ファイルに文字列を追記
		Dim text2 As String = "かきくけこ"
		file.Append(text2)
		Dim result As List(Of String) = Read(filepath)
		Assert.AreEqual(text, result(0))
		Assert.AreEqual(text2, result(1))
		
		Delete(filepath)
	End Sub
	
	<Test> _
	Public Sub Read
		Dim text As String = "あいうえお"
		Dim text2 As String = "かきくこけ"
		file.Append(text)
		file.Append(text2)
		
		' ファイルを読み込む
		Dim results As List(Of String) = file.Read()
		Assert.AreEqual(text, results(0))
		Assert.AreEqual(text2, results(1))
		Assert.AreEqual(2, results.Count)
		
		Delete(filepath)
	End Sub
	
	<Test> _
	Public Sub Read2
		Dim text As String = "あいうえお"
		Dim text2 As String = "かきくこけ"
		file.Append(text)
		file.Append(text2)
		
		' ファイルを読み込む
		Dim idx As Integer = 0
		Dim results As New List(Of String)
		file.Read(
		  Sub(line)
		    results.Add(line & idx.ToString)  
		    idx += 1
		  End Sub
		  )
		Assert.AreEqual(text & "0", results(0))
		Assert.AreEqual(text2 & "1", results(1))
		Assert.AreEqual(2, results.Count)
		
		Delete(filepath)
		End Sub
		
	Private Function Read(path As String) As List(Of String)
		Dim result As New List(Of String)

		Using sr As New StreamReader(path, System.Text.Encoding.Default)
			Dim line As String = sr.ReadLine()
			Do While line IsNot Nothing
				result.Add(line)
				line = sr.ReadLine()
			Loop
			
			Return result
		End Using
	End Function
	
	Private Sub Delete(path As String)
		System.IO.File.Delete(path)
	End Sub
	
		<TestFixtureTearDown> _
		Public Sub Dispose
		Delete(filepath)
	End Sub
End Class
