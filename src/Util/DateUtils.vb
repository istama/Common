'
' 日付: 2016/06/22
'
Namespace Util

''' <summary>
''' 日付に関するユーティリティクラス。
''' </summary>
Public Class DateUtils
  
  ''' <summary>
  ''' 開始月から終了月までの月をリストに格納して返す。
  ''' 日にちはすべて一日にセットされる。
  ''' toDateよりfromDateの方が新しい日付の場合は例外を投げる。
  ''' </summary>
  Public Shared Function GetDateListOfEveryMonth(fromDate As DateTime, toDate As DateTime) As List(Of DateTime)
    If fromDate > toDate Then
      Throw New ArgumentException("fromDate is recent than toDate")
    End If
    
    Dim dates As New List(Of DateTime)
    
    ' fromとtoが同じ年の場合
    If fromDate.Year = toDate.Year Then
      For m = fromDate.Month To toDate.Month
        dates.Add(New DateTime(fromDate.Year, m, 1))
      Next
    Else
      ' fromの年のDateを12月まで格納
      For m = fromDate.Month To 12
        dates.Add(New DateTime(fromDate.Year, m, 1))
      Next
      
      ' fromとtoの年が2年以上開いている場合
      If fromDate.Year + 1 < toDate.Year Then
        ' fromとtoの間の年のDateを1月から12月まで格納
        For y = fromDate.Year + 1 To toDate.Year - 1
          For m = 1 To 12
            dates.Add(New DateTime(y, m, 1))
          Next
        Next
      End If
      
      For m = 1 To toDate.Month
        dates.Add(New DateTime(toDate.Year, m, 1))
      Next
    End If
    
    Return dates
  End Function	
  
  ''' <summary>
  ''' 指定した日付の直後の指定した曜日の日付を取得する。
  ''' 指定した日付と同じ曜日を指定した場合、指定した日付をそのまま返す。
  ''' </summary>
  Public Shared Function GetDateOfNextWeekDay(d As DateTime, weekDay As DayOfWeek) As DateTime
    Dim interval As Integer
    If weekDay = d.DayOfWeek Then
      Return d
    ElseIf weekDay > d.DayOfWeek Then
      interval = weekDay - d.DayOfWeek
    Else
      interval = (weekDay + 7) - d.DayOfWeek
    End If
    
    Return d.AddDays(interval)
  End Function
  
  ''' <summary>
  ''' 指定した年月のうち、指定した曜日の日にちをリストにして返す。
  ''' </summary>
  Public Shared Function GetDaysOf(year As Integer, month As Integer, ParamArray daysOfWeek() As DayOfWeek) As List(Of Integer)
    If month < 1 OrElse month > 12 Then Throw New ArgumentException("month is invalid / " & month)
    
    Dim days As New List(Of Integer)
    Dim d As New DateTime(year, month, 1)
    Do
      If Array.Exists(daysOfWeek, Function(w) w = d.DayOfWeek) Then
        days.Add(d.Day)
      End If
      d = d.AddDays(1)
    Loop While d.Month = month
    
    Return days
  End Function
End Class

End Namespace