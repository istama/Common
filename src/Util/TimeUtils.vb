'
' SharpDevelopによって生成
' ユーザ: Blue
' 日付: 2016/09/22
' 時刻: 22:46
' 
' このテンプレートを変更する場合「ツール→オプション→コーディング→標準ヘッダの編集」
'
Namespace Util

Public Class TimeUtils
  
  ''' <summary>
  ''' 60を1.0とした実数値に変換する。
  ''' </summary>
  ''' <param name="minute">変換する数値</param>
  ''' <param name="digits">小数点以下の桁数</param>
  ''' <returns></returns>
  Public Shared Function ToHour(minute As Integer, digits As Integer) As Double
    Return Math.Round(minute / 60, digits)
  End Function
  
  ''' <summary>
  ''' 1.0を60とした整数値に変換する。
  ''' 小数点以下は四捨五入される。
  ''' intervalで数値の間隔を指定できる。変換した値がintervalの中間の値だった場合は、端数は切り捨てる。
  ''' たとえばintervalが5なら、戻り値は5おきの値で返される。
  ''' </summary>
  ''' <param name="value">変換する数値</param>
  ''' <param name="granularity">変換後の数値の間隔</param>
  ''' <returns></returns>
  Public Shared Function ToMinute(value As Double, granularity As Integer) As Integer
    Dim min As Integer = CType(Math.Round(value * 60, MidpointRounding.AwayFromZero), Integer)
    Dim fraction As Integer = min Mod granularity
    Return min - fraction
  End Function
  
  ''' <summary>
  ''' 1.0を1時間とする数値をTimeSpanオブジェクトに変換する。
  ''' </summary>
  ''' <param name="hourtime"></param>
  ''' <returns></returns>
  Public Shared Function ToTimeSpanFrom(hourtime As Double) As TimeSpan
    Dim minTime As Integer = TimeUtils.ToMinute(hourtime, 1)
    Dim hour As Integer = minTime \ 60
    Dim min As Integer = minTime Mod 60
    
    Return New TimeSpan(hour, min, 0)
  End Function
End Class

End Namespace