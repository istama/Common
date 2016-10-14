'
' 日付: 2016/05/20
'
Imports NUnit.Framework
'Imports NUnit.Mocks
Imports NSubstitute
Imports Moq
Imports Common.COM

<TestFixture> _
Public Class TestExcelWriter
	<Test> _
	Public Sub TestAsyncWrite
		Dim m As New Mock(Of IExcel)()
		m.
		Setup(Sub(x) x.Write(It.IsAny(Of ExcelData))).
		Callback(Of ExcelData)(
			Sub(e)
			Assert.AreEqual("1", e.WrittenText)
			Assert.AreEqual("test.xls", e.filepath)
			Assert.AreEqual("sheet1", e.sheetName)
			Assert.AreEqual(1, e.cell.Row)
			Assert.AreEqual("A", e.Cell.Col)
			End Sub
		)
		
		Dim w As New ExcelWriter(m.Object)
		w.Init
		w.AsyncWrite("1", "test.xls", "sheet1", Cell.Create(1,1))
		m.Verify(Sub(x) x.Write(It.IsAny(Of ExcelData)()), Times.Once)
		
		w.Quit

' NUnit.Mocksによるテスト。
' NUnit自身がこのモックを推奨していない。		
'		Dim mockExcel As New NUnit.Mocks.DynamicMock("Excel", GetType(Common.COM.Excel))
'		mockExcel.Expect("Write", New ExcelData("2", "test.xls", "sheet1", Cell.Create(1,1)))
'		mockExcel.Expect("Write", New ExcelData("3", "test.xls", "sheet1", Cell.Create(2,2)))
'		mockExcel.Expect("Write", New ExcelData("5", "test.xls", "sheet1", Cell.Create(2,2)))
'		mockExcel.Expect("Write", New ExcelData("6", "test.xls", "sheet2", Cell.Create(2,2)))
'		mockExcel.Verify

' NSubsutituteによるテスト。
' Subsutituteから生成されるモックはオリジナルの実装がそのまま実行されてしまう。
' interfaceの実装でないクラスの場合、オリジナルの実装が足かせになる場合も。
'		Dim mock As IExcel = Substitute.For(Of IExcel)()
'		Dim w As New ExcelWriter(mock)
'		w.Init
'		w.AsyncWrite("1", "test.xls", "sheet1", Cell.Create(1,1))
'		mock.Received().Write(New ExcelData("1", "test.xls", "sheet1", Cell.Create(1,1)))
		
'		w.AsyncWrite("2", "test.xls", "sheet1", Cell.Create(1,1))
'		w.AsyncWrite("3", "test.xls", "sheet1", Cell.Create(2,2))
'		w.AsyncWrite("4", "test.xls", "sheet1", Cell.Create(2,2))
'		w.AsyncWrite("5", "test.xls", "sheet1", Cell.Create(2,2))
'		w.AsyncWrite("6", "test.xls", "sheet2", Cell.Create(2,2))

		
'		mock.ReceivedWithAnyArgs().Write("2", "test.xls", "sheet1", Cell.Create(1,1))
'		mock.ReceivedWithAnyArgs().Write(New ExcelData("3", "test.xls", "sheet1", Cell.Create(2,2)))
'		mock.Received().Write(New ExcelData("1", "test.xls", "sheet1", Cell.Create(1,1)))
'		w.AsyncWrite("2", "test.xls", "sheet1", Cell.Create(1,1))
'		mock.Received().Write(New ExcelData("2", "test.xls", "sheet1", Cell.Create(1,1)))
'		
'		w.AsyncWrite("3", "test.xls", "sheet1", Cell.Create(2,2))
'		mock.Received().Write(New ExcelData("3", "test.xls", "sheet1", Cell.Create(2,2)))
'		w.AsyncWrite("4", "test.xls", "sheet1", Cell.Create(2,2))
'		w.AsyncWrite("5", "test.xls", "sheet1", Cell.Create(2,2))
'		mock.Received().Write(New ExcelData("5", "test.xls", "sheet1", Cell.Create(2,2)))
'		w.AsyncWrite("6", "test.xls", "sheet2", Cell.Create(2,2))
'		mock.Received().Write(New ExcelData("6", "test.xls", "sheet1", Cell.Create(2,2)))
'		w.Quit
	End Sub
End Class
