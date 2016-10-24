'
' SharpDevelopによって生成
' ユーザ: Blue
' 日付: 2016/10/18
' 時刻: 12:41
' 
' このテンプレートを変更する場合「ツール→オプション→コーディング→標準ヘッダの編集」
'
Imports NUnit.Framework
Imports Common.COM

<TestFixture> _
Public Class TestExcelColumnNode
  <Test> _
  Public Sub TestConstructor
    Dim n As New ExcelColumnNode("A", "col1")
    
    Assert.AreEqual("A", n.GetCol)
    Assert.AreEqual("col1", n.GetName)
    Assert.AreEqual(0, n.GetChilds.Count)
  End Sub
  
  <Test>
  Public Sub TestAddChild
    Dim n1 As New ExcelColumnNode("A", "col1")
    Dim n2 As New ExcelColumnNode("B", "col2")
    Dim n3 As New ExcelColumnNode("C", "col3")
    Dim n4 As New ExcelColumnNode("D", "col4")
    
    n1.AddChild(n2)
    n2.AddChild(n3)
    n2.AddChild(n4)
    
    Assert.AreEqual("A", n1.GetCol)
    Assert.AreEqual("B", n1.GetChilds(0).GetCol)
    Assert.AreEqual("C", n1.GetChilds(0).GetChilds(0).GetCol)
    Assert.AreEqual("D", n1.GetChilds(0).GetChilds(1).GetCol)
  End Sub
  
  <Test>
  Public Sub TestToDateTable
    Dim n1 As New ExcelColumnNode("A", "root", True)
    Dim n2 As New ExcelColumnNode("B", "c1")
    Dim n3 As New ExcelColumnNode("C", "c2")
    Dim n4 As New ExcelColumnNode("D", "c1-1")
    Dim n5 As New ExcelColumnNode("E", "c1-2")
    Dim n6 As New ExcelColumnNode("F", "c2-1", True)
    Dim n7 As New ExcelColumnNode("G", "c2-1-1")
    
    n1.AddChild(n2)
    n1.AddChild(n3)
    n2.AddChild(n4)
    n2.AddChild(n5)
    n3.AddChild(n6)
    n3.AddChild(n7)
    
    Dim res As DataTable = n1.ToDataTable
    Assert.AreEqual(5, res.Columns.Count)
    Assert.AreEqual("c1", res.Columns(0).ColumnName)
    Assert.AreEqual("c1-1", res.Columns(1).ColumnName)
    Assert.AreEqual("c1-2", res.Columns(2).ColumnName)
    Assert.AreEqual("c2", res.Columns(3).ColumnName)
    Assert.AreEqual("c2-1-1", res.Columns(4).ColumnName)
  End Sub
  
End Class
