'
' 日付: 2016/09/26
'
Namespace IO
  
Public Class Log
  Private Shared file As TextFile
  
  Public Shared Sub SetFilePath(path As String)
    file = New TextFile(path, System.Text.Encoding.UTF8)
    file.Reset
  End Sub
  
  Public Shared Sub out(text As String)
    If file IsNot Nothing Then
      SyncLock file
        If file IsNot Nothing Then
          file.Append(text)
        End If
      End SyncLock
    End If
  End Sub
End Class

End Namespace