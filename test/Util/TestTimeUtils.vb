'
' SharpDevelopによって生成
' ユーザ: Blue
' 日付: 2016/09/24
' 時刻: 22:24
' 
' このテンプレートを変更する場合「ツール→オプション→コーディング→標準ヘッダの編集」
'
Imports NUnit.Framework
Imports Common.Util

<TestFixture> _
Public Class TestTimeUtils
	<Test> _
	Public Sub TestToValue
		' 60を1.0とした数値に変換する
		Assert.AreEqual(1.0, TimeUtils.ToHour(60, 1))
		Assert.AreEqual(0.92, TimeUtils.ToHour(55, 2))
		Assert.AreEqual(0.83, TimeUtils.ToHour(50, 2))
		Assert.AreEqual(0.75, TimeUtils.ToHour(45, 2))
		Assert.AreEqual(0.67, TimeUtils.ToHour(40, 2))
		Assert.AreEqual(0.58, TimeUtils.ToHour(35, 2))
		Assert.AreEqual(0.5, TimeUtils.ToHour(30, 2))
		Assert.AreEqual(0.42, TimeUtils.ToHour(25, 2))
		Assert.AreEqual(0.33, TimeUtils.ToHour(20, 2))
		Assert.AreEqual(0.25, TimeUtils.ToHour(15, 2))
		Assert.AreEqual(0.17, TimeUtils.ToHour(10, 2))
		Assert.AreEqual(0.08, TimeUtils.ToHour(5, 2))
		Assert.AreEqual(0.0, TimeUtils.ToHour(0, 1))
		Assert.AreEqual(2.58, TimeUtils.ToHour(155, 2))
	End Sub
	
	<Test> _
	Public Sub TestToMinute
		' 1.0を60とした数値に変換する
		Assert.AreEqual(60, TimeUtils.ToMinute(1.0, 5))
		Assert.AreEqual(55, TimeUtils.ToMinute(0.92, 5))
		Assert.AreEqual(50, TimeUtils.ToMinute(0.83, 5))
		Assert.AreEqual(45, TimeUtils.ToMinute(0.75, 5))
		Assert.AreEqual(40, TimeUtils.ToMinute(0.67, 5))
		Assert.AreEqual(35, TimeUtils.ToMinute(0.58, 5))
		Assert.AreEqual(30, TimeUtils.ToMinute(0.50, 5))
		Assert.AreEqual(25, TimeUtils.ToMinute(0.42, 5))
		Assert.AreEqual(20, TimeUtils.ToMinute(0.33, 5))
		Assert.AreEqual(15, TimeUtils.ToMinute(0.25, 5))
		Assert.AreEqual(10, TimeUtils.ToMinute(0.17, 5))
		Assert.AreEqual(5, TimeUtils.ToMinute(0.08, 5))
		Assert.AreEqual(0, TimeUtils.ToMinute(0.00, 5))
		
		Assert.AreEqual(50, TimeUtils.ToMinute(0.92, 10))
		Assert.AreEqual(40, TimeUtils.ToMinute(0.75, 10))
	End Sub
End Class
