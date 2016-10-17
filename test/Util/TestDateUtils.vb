'
' 日付: 2016/06/22
'
Imports NUnit.Framework
Imports Common.Util

<TestFixture> _
Public Class TestDateUtils
	
'	<Test> _
'	Public Sub TestToDateTime
'		Dim d1 As DateTime
'		Assert.True(DateUtils.ToDateTime("20000101", d1))
'		Assert.AreEqual(2000, d1.year)
'		Assert.AreEqual(1, d1.month)
'		Assert.AreEqual(1, d1.day)
'		
'		Dim d2 As DateTime
'		Assert.True(DateUtils.ToDateTime("20120229", d2))
'		Assert.AreEqual(2012, d2.year)
'		Assert.AreEqual(2, d2.month)
'		Assert.AreEqual(29, d2.day)
'		
'		Dim d3 As DateTime
'		Assert.True(DateUtils.ToDateTime("20191231", d3))
'		Assert.AreEqual(2019, d3.year)
'		Assert.AreEqual(12, d3.month)
'		Assert.AreEqual(31, d3.day)
'		
'		Dim ed As DateTime
'		Assert.False(DateUtils.ToDateTime("160511", ed))
'		Assert.False(DateUtils.ToDateTime("20110229", ed))
'		Assert.False(DateUtils.ToDateTime("200012311", ed))
'		Assert.False(DateUtils.ToDateTime("201x0101", ed))
'	End Sub
	
	<Test> _
	Public Sub TestGetDateListOfEveryMonth
		' ２つの日付間の日付オブジェクトを、月間隔で生成する
		
		' 同じ年の日付間
		Dim d1 As New DateTime(2000, 5, 31)
		Dim d2 As New DateTime(2000, 11, 1)
		Dim r1 As List(Of DateTime) = DateUtils.GetDateListOfEveryMonth(d1, d2)
		Assert.AreEqual(7, r1.Count)
		For i = 0 To r1.Count - 1
			Assert.AreEqual(2000, r1(i).Year)
			Assert.AreEqual(i + 5, r1(i).Month)
		Next
		
		' １年違いの日付間
		Dim d3 As New DateTime(2000, 11, 30)
		Dim d4 As New DateTime(2001, 2, 1)
		Dim r2 As List(Of DateTime) = DateUtils.GetDateListOfEveryMonth(d3, d4)
		Assert.AreEqual(4, r2.Count)
		For i = 0 To r2.Count - 1
			If i < 2
				Assert.AreEqual(2000, r2(i).Year)
				Assert.AreEqual(i + 11, r2(i).Month)
			Else
				Assert.AreEqual(2001, r2(i).Year)
				Assert.AreEqual(i - 1, r2(i).Month)
			End If
		Next
		
		' ２年以上違いの日付間
		Dim d5 As New DateTime(2000, 12, 30)
		Dim d6 As New DateTime(2003, 1, 1)
		Dim r3 As List(Of DateTime) = DateUtils.GetDateListOfEveryMonth(d5, d6)
		Assert.AreEqual(26, r3.Count)
		For i = 0 To r3.Count - 1
			If i < 1
				Assert.AreEqual(2000, r3(i).Year)
				Assert.AreEqual(i + 12, r3(i).Month)
			Else If i < 13
				Assert.AreEqual(2001, r3(i).Year)
				Assert.AreEqual(i, r3(i).Month)	
			Else If i < 25
				Assert.AreEqual(2002, r3(i).Year)
				Assert.AreEqual(i - 12, r3(i).Month)
			Else
				Assert.AreEqual(2003, r3(i).Year)
				Assert.AreEqual(i - 24, r3(i).Month)
			End If
		Next
		
		' fromDateの日付の方がtoDateよりも新しい場合、例外を投げる
		Dim ex As Exception =
			Assert.Throws(Of ArgumentException)(
				Function() DateUtils.GetDateListOfEveryMonth(d2, d1)
			)
	End Sub
	
	<Test> _
	Public Sub TestGetDateOfNextWeekDay
		' 指定した日付の直後の指定した曜日の日付を取得する
		
		Dim d1 As DateTime = #07/01/2016#
		Dim res1 As DateTime = DateUtils.GetDateOfNextWeekDay(d1, DayOfWeek.Saturday)
		Assert.AreEqual(2, res1.Day)
		Assert.AreEqual(7, res1.Month)
		Assert.AreEqual(2016, res1.Year)
		
		' 同じ曜日の場合のチェック
		Dim res2 As DateTime = DateUtils.GetDateOfNextWeekDay(d1, DayOfWeek.Friday)
		Assert.AreEqual(1, res2.Day)
		
		' 翌月にまたぐ場合のチェック
		Dim d2 As DateTime = #07/26/2016#
		Dim res3 As DateTime = DateUtils.GetDateOfNextWeekDay(d2, DayOfWeek.Monday)
		Assert.AreEqual(1, res3.Day)
		Assert.AreEqual(8, res3.Month)
		Assert.AreEqual(2016, res3.Year)
		
		Dim d3 As DateTime = #07/31/2016#
		Dim res4 As DateTime = DateUtils.GetDateOfNextWeekDay(d3, DayOfWeek.Monday)
		Assert.AreEqual(1, res4.Day)
		Assert.AreEqual(8, res4.Month)
		Assert.AreEqual(2016, res4.Year)
		
		Dim res5 As DateTime = DateUtils.GetDateOfNextWeekDay(d3, DayOfWeek.Saturday)
		Assert.AreEqual(6, res5.Day)
		Assert.AreEqual(8, res5.Month)
		
		' 翌年にまたぐ場合のチェック
		Dim d4 As DateTime = #12/31/2016#
		Dim res6 As DateTime = DateUtils.GetDateOfNextWeekDay(d4, DayOfWeek.Sunday)
		Assert.AreEqual(1, res6.Day)
		Assert.AreEqual(1, res6.Month)
		Assert.AreEqual(2017, res6.Year)
	End Sub
	
	<Test> _
	Public Sub TestGetDays()
		Dim res1 As List(Of Integer) = DateUtils.GetDaysOf(2016, 7, DayOfWeek.Saturday, DayOfWeek.Sunday)
		Assert.AreEqual(10, res1.Count)
		Dim ex1() As Integer = {2, 3, 9, 10, 16, 17, 23, 24, 30, 31}
		For Each e In ex1
			Assert.Contains(e, res1)
		Next
	End Sub
	
	<Test>
	Public Sub TestGetWeekCountInMonth
	  Dim d1 = New DateTime(2016, 10, 1)
	  Assert.AreEqual(1, DateUtils.GetWeekCountInMonth(d1, DayOfWeek.Saturday))
	  Assert.AreEqual(1, DateUtils.GetWeekCountInMonth(d1, DayOfWeek.Sunday))
	  
	  Dim d2 = New DateTime(2016, 10, 2)
	  Assert.AreEqual(2, DateUtils.GetWeekCountInMonth(d2, DayOfWeek.Saturday))
	  Assert.AreEqual(1, DateUtils.GetWeekCountInMonth(d2, DayOfWeek.Sunday))	  
	  
	  Dim d3 = New DateTime(2016, 10, 30)
	  Assert.AreEqual(6, DateUtils.GetWeekCountInMonth(d3, DayOfWeek.Saturday))
	  Assert.AreEqual(5, DateUtils.GetWeekCountInMonth(d3, DayOfWeek.Sunday))	
	End Sub
	
'	<Test> _
'	Public Sub Test
'	  Dim d As DateTime
'	  Assert.AreEqual(True, DateTime.TryParse("2016/09/24", d))
'	  Assert.AreEqual(d.Year, 2016)
'	  Assert.AreEqual(d.Month, 9)
'	  Assert.AreEqual(d.Day, 24)
'	End Sub
End Class
