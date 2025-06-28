VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 只在修改G列（合格证编号）时运行
    If Target.Column <> 7 Or Target.Row < 2 Then Exit Sub
    
    Dim mgmtNo As String, certNo As String
    mgmtNo = Cells(Target.Row, "C").Value
    certNo = Target.Value
    
    ' 更新库存管理表
    If mgmtNo <> "" And certNo <> "" Then
        UpdateCertInInventory mgmtNo, certNo
    End If
End Sub

Sub UpdateCertInInventory(mgmtNo As String, certNo As String)
    Dim invSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("库存管理")
    
    Dim lastRow As Long
    lastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim updated As Boolean
    updated = False
    
    ' 查找并更新合格证编号
    For i = 2 To lastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            invSht.Cells(i, "P").Value = certNo
            updated = True
            Exit For
        End If
    Next i
    
    If updated Then
        msgBox "合格证编号已更新!", vbInformation
    Else
        msgBox "未找到管理号: " & mgmtNo, vbExclamation
    End If
End Sub
Private Sub Worksheet_Chang1e(ByVal Target As Range)
    ' 只在修改G列（合格证编号）时运行
    If Target.Column <> 7 Or Target.Row < 2 Then Exit Sub
    
    Dim mgmtNo As String, certNo As String
    mgmtNo = Cells(Target.Row, "C").Value
    certNo = Target.Value
    
    ' 更新库存管理表
    If mgmtNo <> "" And certNo <> "" Then
        UpdateCertInInventory mgmtNo, certNo
        
        ' 锁定出库时间（转换为固定值）
        Cells(Target.Row, "D").Value = Cells(Target.Row, "D").Value
    End If
End Sub
