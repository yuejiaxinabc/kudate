Attribute VB_Name = "inventory"
Option Explicit

' 初始化库存管理表
Public Sub Inventory_Initialize()
    ' 这里添加库存管理表的初始化代码
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
    invSheet.Range("D1").Value = "有效期"
    
    ' 设置标题格式
    With invSheet.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241) ' 浅蓝色背景
    End With
    
    ' 设置列宽
    invSheet.Columns("A:D").AutoFit
End Sub
