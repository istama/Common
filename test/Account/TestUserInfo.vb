'
' 日付: 2016/05/02
'
Imports NUnit.Framework
Imports Common.Account

<TestFixture> _
Public Class TestUserInfo
	Private ui As New UserInfo("John", "d501", "pass")
	
	<Test> _
	Public Sub TestGetName
		Assert.AreEqual("John", ui.GetName)
	End Sub
	
	<Test> _
	Public Sub TestGetId
		Assert.AreEqual("d501", ui.GetId)
	End Sub
	
	<Test> _
	Public Sub TestGetSimpleId
		Assert.AreEqual("501", ui.GetSimpleId)
	End Sub
	
	<Test> _
	Public Sub TestGetPassword
		Assert.AreEqual("pass", ui.GetPassword)
	End Sub
End Class
