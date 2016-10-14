'
' 日付: 2016/05/01
'
Namespace Threading

''' <summary>
''' LeftかRightかどちらか一方のみの通過を許可するセマフォ。
''' </summary>
Public Class LRSemaphore
  Private Shared LEFT_FLAG As String = "Left"
  Private Shared RIGHT_FLAG As String = "Right"
  
  Private passableFlag As String
  Private count As Integer
  
  Public Sub New()
    passableFlag = Nothing
    count = 0
  End Sub
  
  ''' <summary>
  ''' Leftが通過可能ならばカウンタを１つ増加させ、Trueを返す。
  ''' 誰かがRightを通過している状態の場合は、Flaseを返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function IncrementLeftIfPass() As Boolean
    SyncLock Me
      If passableFlag = LEFT_FLAG OrElse count = 0 Then
        passableFlag = LEFT_FLAG
        count += 1
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' Rightが通過可能ならばカウンタを１つ増加させ、Trueを返す。
  ''' 誰かがLeftを通過している状態の場合は、Flaseを返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function IncrementRightIfPass() As Boolean
    SyncLock Me
      If passableFlag = RIGHT_FLAG OrElse count = 0 Then
        passableFlag = RIGHT_FLAG
        count += 1
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' カウンタを１つ減らす。
  ''' 0未満にはならない。
  ''' カウンタが0のときは、LeftもRightも通過可能。
  ''' </summary>
  Public Sub Decrement()
    SyncLock Me
      If count > 0 Then
        count -= 1
      Else
        count = 0
        passableFlag = Nothing
      End If
    End SyncLock
  End Sub
End Class

End Namespace