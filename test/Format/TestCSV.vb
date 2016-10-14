'
' 日付: 2016/04/28
'
Imports NUnit.Framework
Imports Common.Format

<TestFixture> _
Public Class TestCSV
	<Test> _
	Public Sub Decode
		' 通常のフィールド
		Dim f1 As List(Of String) = CSV.Decode("abc,defg,hijkl")
		Assert.AreEqual(3, f1.Count)
		Assert.AreEqual("abc", f1(0))
		Assert.AreEqual("defg", f1(1))
		Assert.AreEqual("hijkl", f1(2))
		
		' ""で囲まれたフィールド
		Dim f2 As List(Of String) = CSV.Decode("""012"",xyz,""3456"",""ttt""")
		Assert.AreEqual(4, f2.Count)
		Assert.AreEqual("012", f2(0))
		Assert.AreEqual("xyz", f2(1))
		Assert.AreEqual("ttt", f2(3))
		
		' フィールドの中に "" と ,
		Dim f3 As List(Of String) = CSV.Decode("""This is mine, """"pen""""""")
		Assert.AreEqual("This is mine, ""pen""", f3(0))
		
		' 空フィールド
		Dim f4 As List(Of String) = CSV.Decode(",,,")
		Assert.AreEqual(4, f4.Count)
		Assert.AreEqual("", f4(0))
		Assert.AreEqual("", f4(3))
		
		' ""が閉じていない
		Dim ex As Exception =
			Assert.Throws(Of ArgumentException)(
				Function() CSV.Decode("abc,""def,ghi")
			)

	End Sub
End Class
