'
' 日付: 2016/10/17
'
Imports NUnit.Framework

Imports Common.Util

<TestFixture> _
Public Class TestDateTerm
  
  <Test> _
  Public Sub TestConstructor
    Dim min As New DateTime(2010, 1, 1)
    Dim max As New DateTime(2015, 12, 31)
    Dim str As String = "abc"
    Dim t As New DateTerm(min, max, str)
    
    Assert.AreEqual(min, t.BeginDate)
    Assert.AreEqual(max, t.EndDate)
    Assert.AreEqual(str, t.Label)
  End Sub
  
  <Test>
  Public Sub TestDailyTerms
    Dim t As New DateTerm(New DateTime(2016, 09, 30), New DateTime(2016, 11, 1))
    Dim l As List(Of DateTerm) = t.DailyTerms(Function(d) String.Format("{0}月{1}日", d.Month, d.Day))
    
    Assert.AreEqual(33, l.Count)
    AssertDailyTerm(2016,  9, 30, l(0))
    AssertDailyTerm(2016, 10,  1, l(1))
    AssertDailyTerm(2016, 10, 31, l(31))
    AssertDailyTerm(2016, 11,  1, l(32))
  End Sub
  
  Private Sub AssertDailyTerm(y As Integer, m As Integer, d As Integer, term As DateTerm)
    AssertDate(y, m, d, term.BeginDate)
    AssertDate(y, m, d, term.EndDate)
    Assert.AreEqual(String.Format("{0}月{1}日", m, d), term.Label)
  End Sub
  
  <Test>
  Public Sub TestWeeklyTerm
    Dim t As New DateTerm(New DateTime(2016, 12, 3), New DateTime(2016, 12, 31))
    Dim l As List(Of DateTerm) = t.WeeklyTerms(DayOfWeek.Saturday, Function(b, e) String.Format("{0}日-{1}日", b.Day, e.Day))
    
    Assert.AreEqual(5, l.Count)
    AssertDate(2016, 12,  3, l(0).BeginDate)
    AssertDate(2016, 12,  3, l(0).EndDate)
    Assert.AreEqual("3日-3日", l(0).Label)
    
    AssertDate(2016, 12,  4, l(1).BeginDate)
    AssertDate(2016, 12, 10, l(1).EndDate)
    
    AssertDate(2016, 12, 25, l(4).BeginDate)
    AssertDate(2016, 12, 31, l(4).EndDate)
    
    Dim t2 As New DateTerm(New DateTime(2016, 12, 31), New DateTime(2017, 1, 1))
    Dim l2 As List(Of DateTerm) = t2.WeeklyTerms(DayOfWeek.Saturday, Function(b, e) String.Empty)
    
    Assert.AreEqual(2, l2.Count)
    AssertDate(2016, 12, 31, l2(0).BeginDate)
    AssertDate(2016, 12, 31, l2(0).EndDate)
    
    AssertDate(2017,  1,  1, l2(1).BeginDate)
    AssertDate(2017,  1,  1, l2(1).EndDate)
  End Sub
  
  <Test>
  Public Sub TestMonthlyTerms()
    Dim t As New DateTerm(New DateTime(2016, 11, 30), New DateTime(2017, 1, 1))
    Dim l As List(Of DateTerm) = t.MonthlyTerms(Function(b, e) String.Format("{0}月{1}日-{2}日", b.Month, b.Day, e.Day))
    
    Assert.AreEqual(3, l.Count)
    AssertDate(2016, 11, 30, l(0).BeginDate)
    AssertDate(2016, 11, 30, l(0).EndDate)
    Assert.AreEqual("11月30日-30日", l(0).Label)
    
    AssertDate(2016, 12,  1, l(1).BeginDate)
    AssertDate(2016, 12, 31, l(1).EndDate)
        
    AssertDate(2017,  1,  1, l(2).BeginDate)
    AssertDate(2017,  1,  1, l(2).EndDate)
  End Sub
  
  Private Sub AssertDate(y As Integer, m As Integer, d As Integer, _date As DateTime)
    Assert.AreEqual(New DateTime(y, m, d), _date)
  End Sub
End Class
