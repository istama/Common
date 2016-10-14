'
' 日付: 2016/07/29
'
Imports System.Threading.Tasks

Namespace Threading

Public Class MultiTask
  
  Public Shared Sub Run(Of T)(correction As IEnumerable(Of T), f As Func(Of T, Action(Of Object)), args As Object)
    Dim taskArray As New List(Of Task)
    For Each e In correction
      taskArray.Add(Task.Factory.StartNew(f(e), args))
    Next
    Task.WaitAll(taskArray.ToArray)
  End Sub
  
  Public Shared Sub Run(Of T)(correction As IEnumerable(Of T), filter As Func(Of T, Boolean), f As Func(Of T, Action(Of Object)), args As Object)
    Dim taskArray As New List(Of Task)
    For Each e In correction
      If filter(e) Then
        taskArray.Add(Task.Factory.StartNew(f(e), args))
      End If
    Next
    Task.WaitAll(taskArray.ToArray)
  End Sub
  
  Public Shared Function Run(Of T, R)(correction As IEnumerable(Of T), f As Func(Of T, Func(Of Object, R)), args As Object) As List(Of R)
    Dim taskArray As New List(Of Task(Of R))
    For Each e In correction
      taskArray.Add(Task.Factory.StartNew(f(e), args))
    Next
    Task.WaitAll(taskArray.ToArray)
    
    Return taskArray.ConvertAll(Function(task) task.Result)
  End Function
  
End Class

End Namespace