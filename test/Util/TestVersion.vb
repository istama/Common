'
' 日付: 2016/05/04
'
'Imports NUnit.Framework
'Imports Common.Util
'
'<TestFixture> _
'Public Class TestVersion
'	Private v As Version = New Version("1.2.3.4")
'	
'	<Test> _
'	Public Sub TestCompereOperators
'		' バージョン番号の比較
'		Assert.AreEqual(True, v < New Version("1.2.3.5"))
'		Assert.AreEqual(True, v < New Version("2.0"))
'		Assert.AreEqual(True, v <= New Version("2.0"))
'		Assert.AreEqual(True, v <= New Version("1.2.3.4"))
'		Assert.AreEqual(True, v > New Version("1.2.3"))
'		Assert.AreEqual(True, v > New Version("0.2.3.4"))
'		Assert.AreEqual(True, v >= New Version("0.2.3.4"))
'		Assert.AreEqual(True, v >= New Version("1.2.3.4"))
'		Assert.AreEqual(True, v = New Version("1.2.3.4"))
'		Assert.AreEqual(True, v <> New Version("1.0.0.4"))
'	End Sub
'	
'	<Test> _
'	Public Sub TestTryParse
'		Dim ver1 As Version
'		Dim res1 As Boolean = Version.TryParse("1.2.3.4", ver1)
'		Assert.True(res1)
'		Assert.AreEqual("1.2.3.4", ver1.ToString)
'		
'		Dim ver2 As Version
'		Dim res2 As Boolean = Version.TryParse("1.2.3.x", ver1)
'		Assert.False(res2)
'	End Sub
'	
'	<Test> _
'	Public Sub TestConstruction
'		' バージョンに数字以外の文字が含まれている
'		Dim ex As Exception =
'			Assert.Throws(Of ArgumentException)(
'				Function() New Version("1.2.3.x")
'				)
'		' . で終了している
'		Dim ex2 As Exception =
'			Assert.Throws(Of ArgumentException)(
'				Function() New Version("1.2.3.")
'				)
'		' . が連続している
'		Dim ex3 As Exception =
'			Assert.Throws(Of ArgumentException)(
'				Function() New Version("1.2..3")
'				)
'		' マイナーバージョンがない
'		Dim ex4 As Exception =
'			Assert.Throws(Of ArgumentException)(
'				Function() New Version("1")
'			)
'	End Sub
'End Class
