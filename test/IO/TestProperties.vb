'
' 日付: 2016/04/27
'
Imports NUnit.Framework
Imports Common.IO

<TestFixture> _
Public Class TestProperties
	Private filepath As String = "./test.txt"	
	Private prop As New Properties(filepath)
	
	<Test> _
	Public Sub GetValue
		Dim table As New Dictionary(Of String, String)
		table.Add("width", "140")
		table.Add("height", "150")
		table.Add("depth", "100")
		table.Add("weight", "")
		prop.Build(table)
		
		Dim p As New Properties(filepath)
		Assert.AreEqual("140", p.GetValue("width").GetOrDefault(""))
		Assert.AreEqual("100", p.GetValue("depth").GetOrDefault(""))
		Assert.AreEqual("", p.GetValue("weight").GetOrDefault("abc"))
		Assert.AreEqual("", p.GetValue("length").GetOrDefault(""))
		
		Delete()
	End Sub
	
	<Test> _
	Public Sub Add
		Dim f As New TextFile(filepath, System.Text.Encoding.Default)
		
		' プロパティファイルに書き込む
		Dim key As String = "width"
		Dim value As String = "130"
		prop.Add(key, value)
		Assert.AreEqual("width=130", f.Read()(0))
		
		' プロパティファイルに追記
		Dim key2 As String = "height"
		Dim value2 As String = "80"
		prop.Add(key2, value2)
		Dim result = f.Read()
		Assert.AreEqual("width=130", result(0))
		Assert.AreEqual("height=80", result(1))
		
		Delete()
	End Sub
	
	<Test> _
	Public Sub Build
		' ファイルにプロパティを構築
		Dim table As New Dictionary(Of String, String)
		table.Add("abc", "10")
		table.Add("あ", "い")
		table.Add("width", "100")
		prop.Build(table)
		'Dim table2 As IDictionary(Of String, String) = prop.Load()
		Assert.AreEqual("10", prop.GetValue("abc").GetOrDefault(""))
		Assert.AreEqual("い", prop.GetValue("あ").GetOrDefault(""))
		Assert.AreEqual("100", prop.GetValue("width").GetOrDefault(""))
		
		Delete()
	End Sub
	
	<Test> _
	Public Sub Load
		Dim f As New TextFile(filepath, System.Text.Encoding.Default)
		f.Append("a=1")
		f.Append("b=2")
		f.Append("=555")
		f.Append("err=")
		f.Append("err")
		f.Append("あ=い")
		f.Append("b=3")
		
		' プロパティファイルを読み込む
		Dim table As IDictionary(Of String, String) = prop.Load()
		
		Assert.AreEqual(table("a"), "1")
		Assert.AreEqual(table("あ"), "い")
		Assert.AreEqual(table("b"), "3")
		Assert.AreEqual(table("err"), "")
		Assert.AreEqual(table.Count, 4)
		
		Delete()
	End Sub
		
	<TestFixtureSetUp> _
	Public Sub Init
	End Sub
	
	<TestFixtureTearDown> _
	Public Sub Dispose
		Delete()
	End Sub
	
	Private Sub Delete()
		System.IO.File.Delete(filepath)
	End Sub
End Class
