'
' 日付: 2016/05/22
'
Namespace COM

Public Interface IExcel
  Sub init()
  Sub Quit()
  Sub Close(filepath As String)
  Function Read(data As ExcelData) As String
  Function Read(filepath As String, sheetName As String, cell As Cell) As String
  Sub Write(data As ExcelData)
  Sub Write(writtenText As String, filepath As String, sheetName As String, cell As Cell)
End Interface

End Namespace