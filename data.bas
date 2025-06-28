Attribute VB_Name = "data"
Option Explicit

' 刷新所有剩余天数
Public Sub Data_RefreshRemainingDays()
    ' 这里添加刷新剩余天数的代码
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("数据管理")
    
    Dim lastRow As Long
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        msgBox "数据管理表中没有数据", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To lastRow ' 假设第1行是标题
        Dim expiryDate As Date
        On Error Resume Next ' 防止无效日期
        expiryDate = dataSheet.Cells(i, "C").Value ' 假设C列是有效期
        
        If Err.Number = 0 Then
            ' 计算剩余天数
            Dim daysLeft As Long
            daysLeft = expiryDate - Date
            
            ' 根据剩余天数设置格式
            With dataSheet.Cells(i, "D") ' 假设D列显示剩余天数
                .Value = daysLeft
                Select Case daysLeft
                    Case Is < 0
                        .Interior.Color = RGB(255, 199, 206) ' 红色
                        .Font.Color = RGB(156, 0, 6)
                    Case 0 To 3
                        .Interior.Color = RGB(255, 235, 156) ' 黄色
                        .Font.Color = RGB(156, 101, 0)
                    Case Else
                        .Interior.Color = RGB(198, 239, 206) ' 绿色
                        .Font.Color = RGB(0, 97, 0)
                End Select
            End With
        Else
            Err.Clear
            dataSheet.Cells(i, "D").Value = "无效日期"
            dataSheet.Cells(i, "D").Interior.Color = RGB(242, 242, 242) ' 灰色
        End If
    Next i
End Sub
