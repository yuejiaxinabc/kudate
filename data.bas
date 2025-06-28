Attribute VB_Name = "data"
Option Explicit

' ˢ������ʣ������
Public Sub Data_RefreshRemainingDays()
    ' �������ˢ��ʣ�������Ĵ���
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("���ݹ���")
    
    Dim lastRow As Long
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        msgBox "���ݹ������û������", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To lastRow ' �����1���Ǳ���
        Dim expiryDate As Date
        On Error Resume Next ' ��ֹ��Ч����
        expiryDate = dataSheet.Cells(i, "C").Value ' ����C������Ч��
        
        If Err.Number = 0 Then
            ' ����ʣ������
            Dim daysLeft As Long
            daysLeft = expiryDate - Date
            
            ' ����ʣ���������ø�ʽ
            With dataSheet.Cells(i, "D") ' ����D����ʾʣ������
                .Value = daysLeft
                Select Case daysLeft
                    Case Is < 0
                        .Interior.Color = RGB(255, 199, 206) ' ��ɫ
                        .Font.Color = RGB(156, 0, 6)
                    Case 0 To 3
                        .Interior.Color = RGB(255, 235, 156) ' ��ɫ
                        .Font.Color = RGB(156, 101, 0)
                    Case Else
                        .Interior.Color = RGB(198, 239, 206) ' ��ɫ
                        .Font.Color = RGB(0, 97, 0)
                End Select
            End With
        Else
            Err.Clear
            dataSheet.Cells(i, "D").Value = "��Ч����"
            dataSheet.Cells(i, "D").Interior.Color = RGB(242, 242, 242) ' ��ɫ
        End If
    Next i
End Sub
