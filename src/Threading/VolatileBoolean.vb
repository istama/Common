'
' 日付: 2016/05/02
'
Namespace Threading
	
''' <summary>
''' 揮発性のBooleanクラス。
''' スレッドセーフであることを確かめるテストによると、Booleanはもともと揮発性かもしれない...
''' </summary>
Public Class VolatileBoolean
	Private bool As Boolean
	
	Public Sub New(bool As Boolean)
		Me.bool = bool		
	End Sub
	
	Public Function Read() As Boolean
		SyncLock Me
			Return bool
		End SyncLock
	End Function
	
	Public Sub Write(bool As Boolean)
		SyncLock Me
			Me.bool = bool
		End SyncLock
	End Sub
End Class

End Namespace