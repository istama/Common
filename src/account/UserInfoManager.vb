'
' 日付: 2016/04/28
'
Imports Common.IO
Imports Common.Format

Namespace Account	

''' <summary>
''' ユーザ情報を管理
''' </summary>
Public Class UserInfoManager
  Private _userInfoList As List(Of UserInfo)
  Public ReadOnly Property UserInfoList() As List(Of UserInfo)
    Get
      Dim newlist As New List(Of UserInfo)
      _userInfoList.ForEach(Sub(ui) newlist.Add(ui))
      
      Return _userInfoList
    End Get
  End Property
  
  ''' <summary>
  ''' コンストラクタ
  ''' </summary>
  ''' <param name="userInfoList">ユーザ情報のリスト</param>
  Private Sub New(userInfoCollection As List(Of UserInfo))
    Me._userInfoList = New List(Of UserInfo)
    userInfoCollection.ForEach(Sub(ui) _userInfoList.Add(ui))
  End Sub
  
  ''' <summary>
  ''' IDとパスワードが一致するユーザがあればTrueを返す
  ''' </summary>
  ''' <param name="id">id</param>
  ''' <param name="password">パスワード</param>
  ''' <returns>一致するユーザがあればTrue, そうでなければFalse</returns>
  Public Function Certify(id As String, password As String) As Boolean
    Return GetUserInfo(id, password) IsNot Nothing
  End Function
  
  ''' <summary>
  ''' IDとパスワードが一致するユーザを取得する
  ''' 一致しなければNothingを返す
  ''' </summary>
  ''' <param name="id">id</param>
  ''' <param name="password">パスワード</param>
  ''' <returns>IDとパスワードが一致するUserInfoクラス</returns>
  Public Function GetUserInfo(id As String, password As String) As UserInfo
    Dim idx As Integer = _userInfoList.FindIndex(Function(ui) ui.GetId = id AndAlso ui.GetPassword = password)
    If idx >= 0 Then
      Return _userInfoList(idx)
    Else
      Return Nothing
    End If
  End Function
  
  ''' <summary>
  ''' ユーザ情報を読み込み、UserInfoManagerクラスにして返す。
  ''' </summary>
  ''' <param name="filepath">ユーザ情報が記述されたファイルのパス</param>
  ''' <returns>ユーザ情報を管理するクラス</returns>
  Public Shared Function Create(filepath As String) As UserInfoManager
    Dim userlist As New List(Of UserInfo)
    
    Dim f As New TextFile(filepath, System.Text.Encoding.Default)
    f.Read().ForEach(
      Sub(csvText)
      Try
        Dim fields As List(Of String) = CSV.Decode(csvText)
        Dim userInfo As New UserInfo(fields(2), fields(0), fields(1))
        userlist.Add(userInfo)
      Catch ex As Exception
      End Try
    End Sub
  )
  
  Return New UserInfoManager(userlist)
End Function
End Class

End Namespace