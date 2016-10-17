
Option Strict On
'
' 日付: 2016/10/17
'
Namespace Util
  
''' <summary>
''' 期間を表す構造体。
''' </summary>
Public Structure DateTerm
  ''' 開始日
  Private ReadOnly min As DateTime
  ''' 終了日
  Private ReadOnly max As DateTime
  ''' このオブジェクトを文字列表現するためのラベル
  Private ReadOnly _label As String
  
  ''' このオブジェクトの期間を１日おきに分割したリスト
  Private dailyTermList As List(Of DateTerm)
  ''' このオブジェクトの期間を１週間おきに分割したリスト
  Private weeklyTermList As List(Of DateTerm)
  ''' このオブジェクトの期間を１ヶ月おきに分割したリスト
  Private monthlyTermList As List(Of DateTerm)
  
  Public Sub New(min As DateTime, max As DateTime, Optional label As String="")
    Me.min = min
    Me.max = max
    Me._label = label
  End Sub
  
  ''' <summary>
  ''' 期間の開始日を返す。
  ''' </summary>
  Public Function BeginDate() As DateTime
    Return Me.min
  End Function
  
  ''' <summary>
  ''' 期間の終了日を返す。
  ''' </summary>
  Public Function EndDate() As DateTime
    Return Me.max
  End Function
  
  ''' <summary>
  ''' ラベルを返す。
  ''' </summary>
  Public Function Label() As String
    Return Me._label
  End Function
  
  ''' <summary>
  ''' 期間を１日おきに分割したリストを返す。
  ''' </summary>
  Public Function DailyTerms() As List(Of DateTerm)
    If Me.dailyTermList Is Nothing Then
      Return DailyTerms(Function(d) String.Empty)
    Else
      Return Me.dailyTermList
    End If
  End Function
  
  ''' <summary>
  ''' 期間を１日おきに分割したリストを返す。
  ''' 引数にはこの各日付のラベルを返す関数。
  ''' </summary>
  Public Function DailyTerms(f As Func(Of DateTime, String)) As List(Of DateTerm)
    If f Is Nothing Then Throw New ArgumentNullException("f is null")
    
    Dim l As New List(Of DateTerm)
      
    Dim d As DateTime = Me.BeginDate
    While d <= Me.EndDate
      l.Add(New DateTerm(d, d, f(d)))
      d = d.AddDays(1)
    End While
    
    Me.dailyTermList = l
  
    Return Me.dailyTermList
  End Function
  
  ''' <summary>
  ''' 期間を１週間おきに分割したリストを返す。
  ''' </summary>
  Public Function WeeklyTerms() As List(Of DateTerm)
    If Me.WeeklyTerms Is Nothing Then
      Return WeeklyTerms(DayOfWeek.Saturday, Function(b, e) String.Empty)
    Else
      Return Me.weeklyTermList
    End If
  End Function
  
  ''' <summary>
  ''' 期間を１週間おきに分割したリストを返す。
  ''' 引数にはこの各日付のラベルを返す関数。
  ''' </summary>
  Public Function WeeklyTerms(weekend As DayOfWeek, f As Func(Of DateTime, DateTime, String)) As List(Of DateTerm)
    If f Is Nothing Then Throw New ArgumentNullException("f is null")
  
    Dim l As New List(Of DateTerm)
    
    Dim beginDate As DateTime = Me.BeginDate   ' 週の開始日
    Dim weekCntInMonth As Integer = 1  		' １ヶ月の中での週をカウント
    
    While beginDate <= Me.EndDate
      ' 週の開始日と終了日のセットを生成
      Dim endDate As DateTime = DateUtils.GetDateOfNextWeekDay(beginDate, weekend)
      If endDate > Me.EndDate Then
        endDate = Me.EndDate
      End If
      
      l.Add(New DateTerm(beginDate, endDate, f(beginDate, endDate)))
      
      ' 週の開始日を更新
      beginDate = endDate.AddDays(1)
    End While
    
    Me.weeklyTermList = l
'			' この週を表す文字列を生成
'			Dim label As String
'			' 週の終了日が翌月にまたがない場合
'			If endDate.Month = beginDate.Month Then
'				label = String.Format("{0:00}月第{1}週", beginDate.Month, weekCntInMonth.ToString)
'				' 終了日が月末日の場合
'				If endDate.Day = DateTime.DaysInMonth(endDate.Year, endDate.Month) Then
'					weekCntInMonth = 1 ' 翌月の第１週からカウントを開始
'				Else
'					weekCntInMonth += 1 ' 週のカウントを加算
'				End If
'			Else
'				label = String.Format("{0:00}月第{1}週/{2:00}月第1週", beginDate.Month, weekCntInMonth.ToString, endDate.Month)
'				weekCntInMonth = 2	' 翌月の第２週からカウントを開始
'			End If
    
    Return Me.WeeklyTermList
  End Function
  
  ''' <summary>
  ''' 期間を１ヶ月おきに分割したリストを返す。
  ''' </summary>
  Public Function MonthlyTerms() As List(Of DateTerm)
    If Me.monthlyTermList Is Nothing Then
      Return monthlyTerms(Function(b, e) String.Empty)
    Else
      Return Me.monthlyTermList
    End If
  End Function
  
  ''' <summary>
  ''' 期間を１ヶ月おきに分割したリストを返す。
  ''' 引数にはこの各日付のラベルを返す関数。
  ''' </summary>
  Public Function MonthlyTerms(f As Func(Of DateTime, DateTime, String)) As List(Of DateTerm)
    If f Is Nothing Then Throw New ArgumentNullException("f is null")
  
    Dim l As New List(Of DateTerm)
    
    Dim min = Me.BeginDate()
    Dim max = Me.EndDate()
    
    DateUtils.GetDateListOfEveryMonth(min, max).ForEach(
      Sub(d)
        ' 月の開始日と月末日のセットを作成
        Dim begin As DateTime = d
        If min > begin Then
          begin = min
        End If
        
        Dim _end As New DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month))
        If max < _end Then
          _end = max
        End If
        
        l.Add(New DateTerm(begin, _end, f(begin, _end)))
      End Sub)
    
    Me.monthlyTermList = l
  
    Return Me.monthlyTermList
  End Function
  
  Public Overrides Function ToString As String
    Return String.Format("{0} - {1}", Me.BeginDate().ToString, Me.EndDate().ToString)
  End Function
  
End Structure

End Namespace