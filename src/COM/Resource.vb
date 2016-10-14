'
' 日付: 2016/04/29
'
Imports System.Runtime.InteropServices.Marshal

Namespace COM

Public Class Resource
  
  ''' <summary>
  ''' COMオブジェクトを解放する。
  ''' </summary>
  ''' <param name="resource">COMオブジェクト</param>
  Public Shared Sub Release(ByRef resource As Object)
    Try
      If resource IsNot Nothing Then
        FinalReleaseComObject(resource)
      End If
    Finally
      resource = Nothing
    End Try
  End Sub
End Class

End Namespace