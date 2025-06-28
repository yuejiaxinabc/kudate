Attribute VB_Name = "模块2"
' 在模块中添加以下代码
Option Explicit

' 更新仓库库存数量的主过程
Sub UpdateWarehouseStock()
    On Error Resume Next
    Dim invSht As Worksheet, whSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    Set whSht = ThisWorkbook.Sheets("仓库")
    
    ' 获取仓库列表
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    
    ' 重置所有仓库数量为0
    Dim i As Long
    For i = 2 To whLastRow
        whSht.Cells(i, "F").Value = 0
    Next i
    
    ' 统计每个仓库的在库数量
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    For i = 2 To invLastRow
        Dim whName As String
        whName = invSht.Cells(i, "W").Value ' W列是所属仓库
        
        If whName <> "" Then
            ' 查找仓库名称在仓库表中的位置
            Dim j As Long
            For j = 2 To whLastRow
                If whSht.Cells(j, "C").Value = whName Then
                    If invSht.Cells(i, "AA").Value = "在库" Then
                        whSht.Cells(j, "F").Value = whSht.Cells(j, "F").Value + 1
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' 添加库存预警
    For i = 2 To whLastRow
        If whSht.Cells(i, "F").Value < 5 Then
            whSht.Cells(i, "F").Interior.Color = RGB(255, 200, 200) ' 红色背景
        Else
            whSht.Cells(i, "F").Interior.ColorIndex = xlNone ' 清除背景色
        End If
    Next i
    
    msgBox "仓库库存已更新!", vbInformation
End Sub

' 更新单个仓库库存（加减操作）
Sub UpdateSingleWarehouse(whName As String, operation As String)
    Dim whSht As Worksheet
    Set whSht = ThisWorkbook.Sheets("仓库")
    
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To whLastRow
        If whSht.Cells(i, "C").Value = whName Then
            Select Case operation
                Case "add"
                    whSht.Cells(i, "F").Value = whSht.Cells(i, "F").Value + 1
                Case "subtract"
                    If whSht.Cells(i, "F").Value > 0 Then
                        whSht.Cells(i, "F").Value = whSht.Cells(i, "F").Value - 1
                    End If
            End Select
            Exit For
        End If
    Next i
End Sub

' 库存管理表状态变化监控
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    ' 如果改变的是在库状态列(AA列)
    If Not Intersect(Target, Me.Columns("AA")) Is Nothing Then
        If Target.Row > 1 Then
            Dim whName As String
            whName = Me.Cells(Target.Row, "W").Value ' W列是所属仓库
            
            If whName <> "" Then
                ' 判断状态变化
                If Target.Value = "在库" Then
                    ' 增加对应仓库库存
                    UpdateSingleWarehouse whName, "add"
                ElseIf Target.Value = "出库" Then
                    ' 减少对应仓库库存
                    UpdateSingleWarehouse whName, "subtract"
                End If
            End If
        End If
    End If
    
    Application.EnableEvents = True
End Sub
