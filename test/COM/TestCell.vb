'
' 日付: 2016/05/05
'
Imports NUnit.Framework
Imports Common.COM

<TestFixture> _
Public Class TestCell
	<Test> _
	Public Sub TestCreate
		' Cellを生成する
		Dim cell1 As Cell = Cell.Create(1, 2)
		Assert.AreEqual(1, cell1.Row)
		Assert.AreEqual("B", cell1.Col)
		
		Dim cell2 As Cell = Cell.Create(5, "D")
		Assert.AreEqual(5, cell2.Row)
		Assert.AreEqual("D", cell2.Col)
		
		' 行が0以下
		Dim ex As Exception =
			Assert.Throws(Of ArgumentException)(
				Function() Cell.Create(0, 1)
				)	
		' 列が0以下
		Dim ex2 As Exception =
			Assert.Throws(Of ArgumentException)(
				Function() Cell.Create(1, 0)
				)	
		' 列文字が不正
		Dim ex3 As Exception =
			Assert.Throws(Of ArgumentException)(
				Function() Cell.Create(0, "5")
				)		
	End Sub
	
	<Test> _
	Public Sub TestToColWord
		' 数値を列文字に変換する
		Assert.AreEqual("Z", Cell.ToColWord(26))
		Assert.AreEqual("AA", Cell.ToColWord(27))
		Assert.AreEqual("AZ", Cell.ToColWord(52))
		Assert.AreEqual("BA", Cell.ToColWord(53))
		Assert.AreEqual("CV", Cell.ToColWord(100))
		Assert.AreEqual("ZA", Cell.ToColWord(677))
		Assert.AreEqual("ZZ", Cell.ToColWord(702))
		Assert.AreEqual("AAA", Cell.ToColWord(703))
	End Sub
	
	<Test> _
	Public Sub TestEqual
		Dim c1 As Cell = Cell.Create(10, 5)
		Dim c2 As Cell = Cell.Create(10, 5)
		Dim c3 As Cell = Cell.Create(9, 5)
		Dim c4 As Cell = Cell.Create(10, 4)
		
		Assert.IsTrue(c1 = c2)
		Assert.IsFalse(c1 = c3)
		Assert.IsFalse(c1 = c4)
	End Sub
	
	<Test> _
	Public Sub TestNotEqual
		Dim c1 As Cell = Cell.Create(10, 5)
		Dim c2 As Cell = Cell.Create(10, 5)
		Dim c3 As Cell = Cell.Create(9, 5)
		Dim c4 As Cell = Cell.Create(10, 4)
		
		Assert.IsFalse(c1 <> c2)
		Assert.IsTrue(c1 <> c3)
		Assert.IsTrue(c1 <> c4)
	End Sub
End Class
