Attribute VB_Name = "utils"
Option Explicit

' ��鹤�����Ƿ����
Public Function Utils_SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    Utils_SheetExists = (Not ThisWorkbook.Sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function
