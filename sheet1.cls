VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 在模块中粘贴以下代码
Sub GoToInventory()
    Application.Goto Sheets("库存管理").Range("A1"), True
End Sub

Sub GoToWarehouse()
    Application.Goto Sheets("仓库").Range("A1"), True
End Sub

Sub GoToInspection()
    Application.Goto Sheets("检测").Range("A1"), True
End Sub

Sub GoToQuery()
    Application.Goto Sheets("查询").Range("A1"), True
End Sub

Sub GoToOutbound()
    Application.Goto Sheets("出库").Range("A1"), True
End Sub

Sub GoToPersonnel()
    Application.Goto Sheets("人员").Range("A1"), True
End Sub

Sub GoToHome()
    Application.Goto Sheets("首页").Range("A1"), True
End Sub

' 在ThisWorkbook对象中粘贴以下代码
Private Sub Workbook_Open()
    ' 首次打开时创建首页按钮
    Call CreateHomeButtons
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' 每次激活工作表时创建首页按钮
    Call CreateHomeButtons
End Sub

Sub CreateHomeButtons()
    On Error Resume Next
    Dim ws As Worksheet
    Dim btn As Button
    
    For Each ws In ThisWorkbook.Worksheets
        ' 删除现有按钮（如果存在）
        ws.Buttons("HomeBtn").Delete
        
        ' 添加新按钮
        Set btn = ws.Buttons.Add(10, 10, 80, 30)
        With btn
            .Caption = "返回首页"
            .OnAction = "GoToHome"
            .Name = "HomeBtn"
            .Font.Bold = True
            .Font.Size = 10
            .Interior.Color = RGB(91, 155, 213)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next ws
End Sub
Private Sub Worksheet_Activate()
    ' 清除可能存在的旧消息
    On Error Resume Next
    Me.Shapes("HomeMessageBox").Delete
End Sub
