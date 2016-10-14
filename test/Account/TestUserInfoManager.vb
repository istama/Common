'
' 日付: 2016/05/02
'
Imports NUnit.Framework
Imports Common.Account
Imports Common.IO

<TestFixture> _
Public Class TestUserInfoManager
	Private filepath As String = "./csvtest.txt"
	Private m As UserInfoManager
	
	<TestFixtureSetUp> _
	Public Sub Init
		Dim f As New TextFile(filepath, System.Text.Encoding.Default)
		f.Append("101,guitar,john")
		f.Append("102,base,paul")
		f.Append("""103"",""guitar, sitar"",""george""")
		f.Append("104,""drum"",""ringo""")
		
		m = UserInfoManager.Create(filepath)
	End Sub
	
	<Test> _
	Public Sub TestGetUserInfo
		' ID、パスワードが一致するユーザ情報を取得する
		Dim ui1 As UserInfo = m.GetUserInfo("101", "guitar")
		Assert.AreEqual("john", ui1.GetName)
		Dim ui2 As UserInfo = m.GetUserInfo("102", "base")
		Assert.AreEqual("paul", ui2.GetName)
		Dim ui3 As UserInfo = m.GetUserInfo("103", "guitar, sitar")
		Assert.AreEqual("george", ui3.GetName)
		Dim ui4 As UserInfo = m.GetUserInfo("104", "drum")
		Assert.AreEqual("ringo", ui4.GetName)
	End Sub	
	
	<Test> _
	Public Sub TestCertify
		Assert.True(m.Certify("101", "guitar"))
		Assert.False(m.Certify("102", "guitar"))
		Assert.False(m.Certify("101", "guiter"))
	End Sub
	
	<TestFixtureTearDown> _
	Public Sub Dispose
		System.IO.File.Delete(filepath)
	End Sub
End Class
