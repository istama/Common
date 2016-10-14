'
' 日付: 2016/04/28
'
Imports NUnit.Framework
Imports Common.IO

<TestFixture> _
Public Class TestAppProperties
	
	Class SampleProperties
		Inherits AppProperties
		
		Public Sub New(filepath As String)
			MyBase.New(filepath)
		End Sub
		
		Protected Overrides Function DefaultProperties() As IDictionary(Of String, String)
			Dim p As New Dictionary(Of String, String)
			p.Add("width", "200")
			p.Add("height", "150")
			
			Return p
		End Function
		
		Protected Overrides Function AllowNonDefaultProperty As Boolean
		  Return False
		End Function
	End Class
	
	Private filepath As String = "./test.txt"	
	
	<Test> _
	Public Sub Constructor
		Dim table As New Dictionary(Of String, String)
		table.Add("width", "120")
		table.Add("depth", "100")
		Dim p As New Properties(filepath)
		p.Build(table)
		
		' プロパティファイルを再構築
		Dim ap As New SampleProperties(filepath)
		Assert.AreEqual("120", ap.GetValue("width").GetOrDefault(""))
		Assert.AreEqual("150", ap.GetValue("height").GetOrDefault(""))
		Assert.AreEqual("", ap.GetValue("depth").GetOrDefault(""))
		
		Delete()
	End Sub
	
	<TestFixtureTearDown> _
	Public Sub Dispose
		Delete()
	End Sub
	
	Private Sub Delete
		System.IO.File.Delete(filepath)
	End Sub
End Class
