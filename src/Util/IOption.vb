'
' 日付: 2016/09/04
'
Namespace Util
	
''' <summary>
''' Option型インタフェース
''' </summary>
Public Interface IOption(Of T)
	Function GetOrDefault(def As T) As T
End Interface

End Namespace