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
  
End Class
