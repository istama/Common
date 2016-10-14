'
' 日付: 2016/04/27
'
Imports System.IO
Imports System.Text

Namespace IO

''' <summary>
''' ファイルの文字列を読み書きする
''' </summary>
Public Class TextFile
  ' 入出力を行うファイルのパス	
  Private filepath As String	
  ' エンコード
  Private encoding As Encoding
  
  ''' <summary>
  ''' コンストラクタ
  ''' </summary>
  ''' <param name="filepath">入出力を行うファイルのパス</param>
  Sub New(filepath As String, encoding As Encoding)
    If filepath = String.Empty Then
      Throw New ArgumentException("ファイルパスが空文字です。")
    End If
    
    Me.filepath = filepath
    Me.encoding = encoding
  End Sub
  
  ''' <summary>
  ''' アクセスするファイルのパスを返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function Path() As String
    Return filepath
  End Function
  
  ''' <summary>
  ''' ファイルを作成　すでにある場合は何もしない
  ''' </summary>
  Public Function Create() As Boolean
    If Not File.Exists(filepath) Then
      Reset()
      Return True
    Else
      Return False
    End If
  End Function
  
  ''' <summary>
  ''' ファイルを空にする
  ''' ファイルが存在しない場合、作成される
  ''' </summary>
  Public Sub Reset()
    If File.Exists(filepath) Then
      File.Delete(filepath)
    End If
    
    Using stream As FileStream = File.Create(filepath)
    End Using
  End Sub
  
  ''' <summary>
  ''' ファイルが存在するか判定する。
  ''' </summary>
  ''' <returns></returns>
  Public Function Exists() As Boolean
    Return File.Exists(filepath)
  End Function
  
  ''' <summary>
  ''' ファイルを削除する。
  ''' </summary>
  Public Sub Delete() 
    File.Delete(filepath)
  End Sub
  
  ''' <summary>ファイルに文字列を上書きする</summary>
  ''' <param name="text">書き込む文字列</param>
  Public Sub Write(text As String)
    Call Output(text, False)
  End Sub
  
  ''' <summary>
  ''' ファイルに文字列を追記する
  ''' </summary>
  ''' <param name="text">追記する文字列</param>
  Public Sub Append(text As String)
    Call Output(text, True)
  End Sub
  
  ''' <summary>
  ''' ファイルを読み込み、文字列を１行おきに分割してListに格納して返す
  ''' </summary>
  ''' <returns>ファイルの文字列</returns>
  Public Function Read() As List(Of String)
    Using stream As New StreamReader(filepath, encoding)
      Dim list As New List(Of String)
      Dim line As String = stream.ReadLine()
      Do While line IsNot Nothing
        list.Add(line)
        line = stream.ReadLine()
      Loop
      
      Return list
    End Using
  End Function
  
  ''' <summary>
  ''' ファイルを１行ずつ読み込みそのつど引数の関数に文字列を渡す。
  ''' </summary>
  ''' <param name="f">コールバック関数</param>
  Public Sub Read(f As Action(Of String))
    Using stream As New StreamReader(filepath, encoding)
      Dim line As String = stream.ReadLine()
      Do While line IsNot Nothing
        f(line)
        line = stream.ReadLine()
      Loop
    End Using    
  End Sub
  
  ''' <summary>ファイルに文字列を上書きする</summary>
  ''' <param name="text">書き込む文字列</param>
  ''' <param name="append">Trueなら追記する、Falseなら上書きする</param>
  Private Sub Output(text As String, append As Boolean)
    Using stream As New StreamWriter(filepath, append, encoding)
      stream.Write(text + stream.NewLine)
    End Using
  End Sub
  
End Class

End Namespace