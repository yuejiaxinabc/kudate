VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 在检测工作表的代码中
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Dim rowNum As Long
    rowNum = Target.Row
    
    ' 只处理第2行及以下的数据行
    If rowNum < 2 Then Exit Sub
    
    ' 1. 管理号变化时自动填充其他信息
    If Not Intersect(Target, Me.Range("C:C")) Is Nothing Then
        If Target.Value <> "" Then
            ' 调用自动填充函数
            Call AutoFillInspectionInfo(rowNum)
        End If
    End If
    
    ' 2. 合格证编号变化时更新库存管理表
    If Not Intersect(Target, Me.Range("G:G")) Is Nothing Then
        If Target.Value <> "" Then
            ' 调用更新函数
            Call UpdateInventoryCert(rowNum)
        End If
    End If
    
ErrorHandler:
    Application.EnableEvents = True
End Sub

' 自动填充检测表信息
Sub AutoFillInspectionInfo(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ThisWorkbook.Sheets("检测")
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
            Dim remainingDays As Long
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
            
            ' 自动生成序号
            If inspSht.Cells(rowNum, "A").Value = "" Then
                inspSht.Cells(rowNum, "A").Value = inspSht.Cells(inspSht.Rows.Count, "A").End(xlUp).Row
            End If
            
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        msgBox "未找到管理号: " & mgmtNo, vbExclamation, "错误"
    End If
End Sub

' 更新库存管理表的合格证编号
Sub UpdateInventoryCert(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ThisWorkbook.Sheets("检测")
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
        msgBox "合格证编号已更新到库存管理表!", vbInformation, "成功"
    Else
        msgBox "未找到管理号: " & mgmtNo, vbExclamation, "错误"
    End If
End Sub
    Application.EnableEvents = True
End Sub

' 同步检测表合格证编号
Sub SyncCertWithInspection(rowNum As Long)
    Dim invSht As Worksheet, inspSht As Worksheet
    Set invSht = ActiveSheet
    Set inspSht = ThisWorkbook.Sheets("检测")
    
    Dim mgmtNo As String
    mgmtNo = invSht.Cells(rowNum, "E").Value
    Dim certNo As String
    certNo = invSht.Cells(rowNum, "P").Value
    
    ' 在检测表查找对应管理号
    Dim inspLastRow As Long
    inspLastRow = inspSht.Cells(inspSht.Rows.Count, "C").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To inspLastRow
        If inspSht.Cells(i, "C").Value = mgmtNo Then
            ' 更新合格证编号
            inspSht.Cells(i, "G").Value = certNo
            found = True
        End If
    Next i
    
    If Not found Then
        ' 如果没有找到，创建新记录
        Dim newRow As Long
        newRow = inspSht.Cells(inspSht.Rows.Count, "C").End(xlUp).Row + 1
        inspSht.Cells(newRow, "C").Value = mgmtNo
        inspSht.Cells(newRow, "G").Value = certNo
        msgBox "在检测表创建了新记录", vbInformation
    End If
End Sub
Private Sub Worksheet_Chan1ge(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' 检查是否是数量列(AB列)发生变化
    If Not Intersect(Target, Me.Columns("AB")) Is Nothing Then
        If Target.Row > 1 Then
            Dim whName As String
            whName = Me.Cells(Target.Row, "W").Value ' W列是所属仓库
            
            If whName <> "" Then
                ' 获取旧值（从备份列获取）
                Dim oldValue As Long
                oldValue = 0
                If Not IsEmpty(Me.Cells(Target.Row, "AC").Value) Then
                    oldValue = Me.Cells(Target.Row, "AC").Value
                End If
                
                ' 获取新值
                Dim newValue As Long
                newValue = 0
                If IsNumeric(Target.Value) Then
                    newValue = Target.Value
                End If
                
                ' 更新仓库数量
                UpdateWarehouseStock whName, oldValue, newValue
                
                ' 保存当前值为备份
                Me.Cells(Target.Row, "AC").Value = newValue
            End If
        End If
    End If

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 更新仓库库存函数
Sub UpdateWarehouseStock(whName As String, oldQty As Long, newQty As Long)
    Dim whSht As Worksheet
    Set whSht = ThisWorkbook.Sheets("仓库")
    
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    ' 查找仓库并更新数量
    Dim i As Long
    For i = 2 To whLastRow
        If whSht.Cells(i, "C").Value = whName Then
            ' 先减去旧值，再加新值
            whSht.Cells(i, "F").Value = whSht.Cells(i, "F").Value - oldQty + newQty
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        ' 如果仓库不存在，添加新行
        whLastRow = whLastRow + 1
        whSht.Cells(whLastRow, "C").Value = whName
        whSht.Cells(whLastRow, "F").Value = newQty
    End If
End Sub
' 在Worksheet_Change事件中添加
If Not Intersect(Target, Me.Columns("W")) Is Nothing Then
    If Target.Row > 1 Then
        Dim oldWh As String, newWh As String
        oldWh = ""
        newWh = Target.Value
        
        ' 获取旧仓库名称（从备份列）
        If Not IsEmpty(Me.Cells(Target.Row, "AD").Value) Then
            oldWh = Me.Cells(Target.Row, "AD").Value
        End If
        
        ' 获取数量
        Dim qty As Long
        qty = 0
        If IsNumeric(Me.Cells(Target.Row, "AB").Value) Then
            qty = Me.Cells(Target.Row, "AB").Value
        End If
        
        ' 更新仓库库存
        If oldWh <> "" Then
            UpdateWarehouseStock oldWh, qty, 0 ' 从旧仓库移除
        End If
        If newWh <> "" Then
            UpdateWarehouseStock newWh, 0, qty ' 添加到新仓库
        End If
        
        ' 保存新仓库为备份
        Me.Cells(Target.Row, "AD").Value = newWh
    End If
End If
