Attribute VB_Name = "模块6"
Sub RefreshAllWarehouseStock()
    Dim invSht As Worksheet, whSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    Set whSht = ThisWorkbook.Sheets("仓库")
    
    ' 清空仓库数量
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    For i = 2 To whLastRow
        whSht.Cells(i, "F").Value = 0
    Next i
    
    ' 重新计算所有仓库库存
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    For i = 2 To invLastRow
        Dim whName As String
        whName = invSht.Cells(i, "W").Value
        Dim qty As Long
        qty = 0
        If IsNumeric(invSht.Cells(i, "AB").Value) Then
            qty = invSht.Cells(i, "AB").Value
        End If
        
        If whName <> "" Then
            ' 查找仓库并增加数量
            For j = 2 To whLastRow
                If whSht.Cells(j, "C").Value = whName Then
                    whSht.Cells(j, "F").Value = whSht.Cells(j, "F").Value + qty
                    Exit For
                End If
            Next j
        End If
    Next i
    
    msgBox "仓库库存已刷新!", vbInformation
End Sub
