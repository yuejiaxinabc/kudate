Attribute VB_Name = "Mainprocedures"
' 标准模块 (如 Module1)
Option Explicit

' 初始化库存管理表
Public Sub InitializeInventorySheet()
    ' 这里添加实际的初始化代码
    ' 示例:
    Dim invSheet As Worksheet
    Set invSheet = ThisWorkbook.Sheets("库存管理")
    
    ' 清除旧数据（保留表头）
    If invSheet.Range("A2").Value <> "" Then
        invSheet.Range("A2", invSheet.Cells(invSheet.Rows.Count, "A").End(xlUp)).EntireRow.Delete
    End If
    
    ' 设置初始值或格式
    invSheet.Range("A1").Value = "产品ID"
    invSheet.Range("B1").Value = "产品名称"
    invSheet.Range("C1").Value = "库存数量"
    ' 添加更多初始化代码...
End Sub

' 刷新所有剩余天数
Public Sub RefreshAllRemainingDays()
    ' 这里添加实际的刷新代码
    ' 示例:
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("数据管理")
    
    Dim lastRow As Long
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' 假设第1行是标题
        Dim expiryDate As Date
        On Error Resume Next ' 防止无效日期
        expiryDate = dataSheet.Cells(i, "C").Value ' 假设C列是有效期
        
        If Err.Number = 0 Then
            ' 计算剩余天数
            Dim daysLeft As Long
            daysLeft = expiryDate - Date
            dataSheet.Cells(i, "D").Value = daysLeft ' 假设D列显示剩余天数
        Else
            Err.Clear
            dataSheet.Cells(i, "D").Value = "无效日期"
        End If
    Next i
End Sub
