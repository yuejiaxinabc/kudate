VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' 确保首页存在
    Dim homeSheet As Worksheet
    Set homeSheet = Nothing
    On Error Resume Next
    Set homeSheet = ThisWorkbook.Sheets("首页")
    On Error GoTo ErrorHandler
    
    If homeSheet Is Nothing Then
        msgBox "找不到【首页】工作表，请检查工作簿结构", vbCritical
        Exit Sub
    End If
    
    ' 确保首页有当前日期
    With homeSheet.Range("B15")
        If IsEmpty(.Value) Or Not IsDate(.Value) Then
            .Value = Date
            .NumberFormat = "yyyy-mm-dd"
        End If
    End With
    
    ' 初始化库存管理表
    If Utils_SheetExists("库存管理") Then
        Inventory_Initialize
    Else
        msgBox "找不到【库存管理】工作表，跳过初始化", vbExclamation
    End If
    
    ' 刷新所有剩余天数
    If Utils_SheetExists("数据管理") Then
        Data_RefreshRemainingDays
    Else
        msgBox "找不到【数据管理】工作表，跳过剩余天数刷新", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "错误号: " & Err.Number & vbCrLf & _
             "错误描述: " & Err.Description & vbCrLf & _
             "发生在: Workbook_Open"
    
    msgBox "工作簿打开时发生错误:" & vbCrLf & errMsg, vbCritical
    Err.Clear
End Sub

