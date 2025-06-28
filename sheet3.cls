VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 出库表管理号输入监控
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    ' 如果改变的是管理号列(C列)
    If Not Intersect(Target, Me.Columns("C")) Is Nothing Then
        If Target.Row > 1 And Target.Value <> "" Then
            Dim invSht As Worksheet
            Set invSht = ThisWorkbook.Sheets("库存管理")
            
            Dim mgmtNo As String
            mgmtNo = Target.Value
            
            ' 查找管理号对应的行
            Dim invLastRow As Long
            invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
            
            Dim foundRow As Long
            foundRow = 0
            
            Dim i As Long
            For i = 2 To invLastRow
                If invSht.Cells(i, "E").Value = mgmtNo Then
                    foundRow = i
                    Exit For
                End If
            Next i
            
            If foundRow > 0 Then
                ' 1. 更新库存管理表状态为"出库"
                invSht.Cells(foundRow, "AA").Value = "出库"
                
                ' 2. 减少对应仓库库存
                Dim whName As String
                whName = invSht.Cells(foundRow, "W").Value
                UpdateSingleWarehouse whName, "subtract"
                
                ' 3. 自动填充其他信息
                Me.Cells(Target.Row, "B").Value = invSht.Cells(foundRow, "B").Value ' 所属部门
                Me.Cells(Target.Row, "F").Value = invSht.Cells(foundRow, "N").Value ' 剩余天数
                Me.Cells(Target.Row, "G").Value = invSht.Cells(foundRow, "Q").Value ' 使用地点
                Me.Cells(Target.Row, "H").Value = invSht.Cells(foundRow, "R").Value ' 使用用途
                Me.Cells(Target.Row, "I").Value = invSht.Cells(foundRow, "S").Value ' 分类
                Me.Cells(Target.Row, "J").Value = invSht.Cells(foundRow, "V").Value ' 当前位置
                Me.Cells(Target.Row, "K").Value = invSht.Cells(foundRow, "W").Value ' 所属仓库
                
                ' 4. 记录出库时间
                Me.Cells(Target.Row, "D").Value = Now
            Else
                msgBox "未找到管理号: " & mgmtNo, vbExclamation
            End If
        End If
    End If
    
    Application.EnableEvents = True
End Sub
Private Sub Worksheet_Chang1e(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    ' 只处理管理号列(C列)和出库数量列(E列)
    If Not Intersect(Target, Me.Range("C:C, E:E")) Is Nothing Then
        If Target.Row > 1 And Target.Value <> "" Then
            ' 获取管理号和出库数量
            Dim mgmtNo As String
            Dim outQty As Long
            mgmtNo = Me.Cells(Target.Row, "C").Value
            outQty = Me.Cells(Target.Row, "E").Value
            
            If mgmtNo <> "" And outQty > 0 Then
                ' 更新库存管理表
                UpdateInventory mgmtNo, outQty
                
                ' 更新仓库库存
                UpdateWarehouseStock mgmtNo, outQty
                
                ' 记录出库时间
                Me.Cells(Target.Row, "D").Value = Now
                
                ' 设置状态
                Me.Cells(Target.Row, "F").Value = "已出库"
            End If
        End If
    End If

ErrorHandler:
    Application.EnableEvents = True
End Sub

' 更新库存管理表
Sub UpdateInventory(mgmtNo As String, outQty As Long)
    Dim invSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    
    Dim lastRow As Long
    lastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            ' 减少库存数量
            Dim currentQty As Long
            currentQty = invSht.Cells(i, "AB").Value
            invSht.Cells(i, "AB").Value = currentQty - outQty
            
            ' 增加出库数量
            Dim outTotal As Long
            outTotal = invSht.Cells(i, "Ad").Value
            invSht.Cells(i, "AC").Value = outTotal + outQty
            
            ' 更新在库状态
            If invSht.Cells(i, "AB").Value <= 0 Then
                invSht.Cells(i, "AA").Value = "出库"
            End If
            
            Exit For
        End If
    Next i
End Sub

' 更新仓库库存
Sub UpdateWarehouseStock(mgmtNo As String, outQty As Long)
    Dim invSht As Worksheet, whSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    Set whSht = ThisWorkbook.Sheets("仓库")
    
    ' 获取所属仓库
    Dim whName As String
    whName = ""
    
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To invLastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            whName = invSht.Cells(i, "W").Value
            Exit For
        End If
    Next i
    
    If whName = "" Then Exit Sub
    
    ' 更新仓库库存
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To whLastRow
        If whSht.Cells(i, "C").Value = whName Then
            whSht.Cells(i, "F").Value = whSht.Cells(i, "F").Value - outQty
            Exit For
        End If
    Next i
End Sub
