Attribute VB_Name = "模块3"
' 检测表代码
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    Dim rowNum As Long
    rowNum = Target.Row
    
    ' 只处理第2行及以下的数据行
    If rowNum < 2 Then Exit Sub
    
    ' 1. 管理号变化时自动填充其他信息
    If Not Intersect(Target, Me.Columns("C")) Is Nothing Then
        If Target.Value <> "" Then
            AutoFillInspectionInfo rowNum
        End If
    End If
    
    ' 2. 合格证编号变化时更新库存管理表
    If Not Intersect(Target, Me.Columns("G")) Is Nothing Then
        If Target.Value <> "" Then
            UpdateInventoryCert rowNum
        End If
    End If
    
    Application.EnableEvents = True
End Sub

' 自动填充检测表信息
Sub AutoFillInspectionInfo(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ActiveSheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    
    Dim mgmtNo As String
    mgmtNo = inspSht.Cells(rowNum, "C").Value
    
    ' 查找管理号对应的行
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To invLastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            ' 自动填充信息
            inspSht.Cells(rowNum, "A") = inspSht.Cells(inspSht.Rows.Count, "A").End(xlUp).Row ' 序号
            inspSht.Cells(rowNum, "B").Value = invSht.Cells(i, "B").Value ' 所属部门
            inspSht.Cells(rowNum, "H").Value = invSht.Cells(i, "Q").Value ' 使用地点
            inspSht.Cells(rowNum, "I").Value = invSht.Cells(i, "R").Value ' 使用用途
            inspSht.Cells(rowNum, "J").Value = invSht.Cells(i, "S").Value ' 分类
            inspSht.Cells(rowNum, "K").Value = invSht.Cells(i, "V").Value ' 当前位置
            inspSht.Cells(rowNum, "L").Value = invSht.Cells(i, "W").Value ' 所属仓库
            
            ' 自动计算剩余天数
            If IsDate(invSht.Cells(i, "L").Value) Then
                Dim currentDate As Date
                currentDate = ThisWorkbook.Sheets("首页").Range("B15").Value
                inspSht.Cells(rowNum, "F").Value = invSht.Cells(i, "L").Value - currentDate
            End If
            
            ' 自动设置状态
            Dim remainingDays As Integer
            remainingDays = inspSht.Cells(rowNum, "F").Value
            If remainingDays <= 3 Then
                inspSht.Cells(rowNum, "E").Value = "待检"
            ElseIf remainingDays <= 10 Then
                inspSht.Cells(rowNum, "E").Value = "即将到期"
            Else
                inspSht.Cells(rowNum, "E").Value = "正常"
            End If
            
            ' 记录出库时间
            inspSht.Cells(rowNum, "D").Value = Now
            
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        msgBox "未找到管理号: " & mgmtNo, vbExclamation
    End If
End Sub

' 更新库存管理表的合格证编号
Sub UpdateInventoryCert(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ActiveSheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    
    Dim mgmtNo As String
    mgmtNo = inspSht.Cells(rowNum, "C").Value
    Dim certNo As String
    certNo = inspSht.Cells(rowNum, "G").Value
    
    ' 查找管理号对应的行
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To invLastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            ' 更新合格证编号
            invSht.Cells(i, "P").Value = certNo
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        msgBox "合格证编号已更新到库存管理表!", vbInformation
    Else
        msgBox "未找到管理号: " & mgmtNo, vbExclamation
    End If
End Sub
