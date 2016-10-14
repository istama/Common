'
' 日付: 2016/05/05
'
Imports NUnit.Framework
Imports Common.Util
Imports System.IO
Imports Common.IO

<TestFixture> _
Public Class TestUpdate
	Private versionFilePath As String = Path.GetFullPath("./testversion.txt")
	Private batchFilePath As String = Path.GetFullPath("./testbatch.bat")
	Private resultFilePath As String = Path.GetFullPath("./testUpdate.txt")
	Private u As New SampleUpdate(versionFilePath, batchFilePath)	
	Private versionIO As New TextFile(versionFilePath, System.Text.Encoding.Default)
	Private batchIO As New TextFile(batchFilePath, System.Text.Encoding.Default)
	Private resultIO As New TextFile(resultFilePath, System.Text.Encoding.Default)
	
	Private version As New Version("1.0.0.0")
	
	Class SampleUpdate
		Inherits Update
		
		Public Sub New(vPath As String, bPath As String)
			MyBase.New(vPath, bPath)
		End Sub
		
		Protected Overrides Function Script() As String
			Return "echo abc> testUpdate.txt"
		End Function
	End Class
	
	<Test> _
	Public Sub TestUpdateClass
		' バッチファイルを生成
		u.CreateUpdateBatch()
		Assert.True(File.Exists(batchFilePath))
		Assert.AreEqual("echo abc> testUpdate.txt", batchIO.Read(0))
		
		' 最新のバージョンファイルがあるか確認
		Assert.False(u.existsUpdateVersion(version))
		versionIO.Write("1.0.0.1")
		Assert.True(u.existsUpdateVersion(version))
		
		' バッチファイル実行
		Dim p As System.Diagnostics.Process = u.RunUpdateBatch(version)
		p.WaitForExit()
		Assert.AreEqual("abc", resultIO.Read(0))
		
	End Sub
	
	<TestFixtureTearDown> _
	Public Sub Dispose
		File.Delete(resultFilePath)
		File.Delete(versionFilePath)
		File.Delete(batchFilePath)
	End Sub	
End Class
