Attribute VB_Name = "ģ��3"
' �������
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    Dim rowNum As Long
    rowNum = Target.Row
    
    ' ֻ�����2�м����µ�������
    If rowNum < 2 Then Exit Sub
    
    ' 1. ����ű仯ʱ�Զ����������Ϣ
    If Not Intersect(Target, Me.Columns("C")) Is Nothing Then
        If Target.Value <> "" Then
            AutoFillInspectionInfo rowNum
        End If
    End If
    
    ' 2. �ϸ�֤��ű仯ʱ���¿������
    If Not Intersect(Target, Me.Columns("G")) Is Nothing Then
        If Target.Value <> "" Then
            UpdateInventoryCert rowNum
        End If
    End If
    
    Application.EnableEvents = True
End Sub

' �Զ���������Ϣ
Sub AutoFillInspectionInfo(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ActiveSheet
    Set invSht = ThisWorkbook.Sheets("������")
    
    Dim mgmtNo As String
    mgmtNo = inspSht.Cells(rowNum, "C").Value
    
    ' ���ҹ���Ŷ�Ӧ����
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To invLastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            ' �Զ������Ϣ
            inspSht.Cells(rowNum, "A") = inspSht.Cells(inspSht.Rows.Count, "A").End(xlUp).Row ' ���
            inspSht.Cells(rowNum, "B").Value = invSht.Cells(i, "B").Value ' ��������
            inspSht.Cells(rowNum, "H").Value = invSht.Cells(i, "Q").Value ' ʹ�õص�
            inspSht.Cells(rowNum, "I").Value = invSht.Cells(i, "R").Value ' ʹ����;
            inspSht.Cells(rowNum, "J").Value = invSht.Cells(i, "S").Value ' ����
            inspSht.Cells(rowNum, "K").Value = invSht.Cells(i, "V").Value ' ��ǰλ��
            inspSht.Cells(rowNum, "L").Value = invSht.Cells(i, "W").Value ' �����ֿ�
            
            ' �Զ�����ʣ������
            If IsDate(invSht.Cells(i, "L").Value) Then
                Dim currentDate As Date
                currentDate = ThisWorkbook.Sheets("��ҳ").Range("B15").Value
                inspSht.Cells(rowNum, "F").Value = invSht.Cells(i, "L").Value - currentDate
            End If
            
            ' �Զ�����״̬
            Dim remainingDays As Integer
            remainingDays = inspSht.Cells(rowNum, "F").Value
            If remainingDays <= 3 Then
                inspSht.Cells(rowNum, "E").Value = "����"
            ElseIf remainingDays <= 10 Then
                inspSht.Cells(rowNum, "E").Value = "��������"
            Else
                inspSht.Cells(rowNum, "E").Value = "����"
            End If
            
            ' ��¼����ʱ��
            inspSht.Cells(rowNum, "D").Value = Now
            
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        msgBox "δ�ҵ������: " & mgmtNo, vbExclamation
    End If
End Sub

' ���¿������ĺϸ�֤���
Sub UpdateInventoryCert(rowNum As Long)
    Dim inspSht As Worksheet, invSht As Worksheet
    Set inspSht = ActiveSheet
    Set invSht = ThisWorkbook.Sheets("������")
    
    Dim mgmtNo As String
    mgmtNo = inspSht.Cells(rowNum, "C").Value
    Dim certNo As String
    certNo = inspSht.Cells(rowNum, "G").Value
    
    ' ���ҹ���Ŷ�Ӧ����
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To invLastRow
        If invSht.Cells(i, "E").Value = mgmtNo Then
            ' ���ºϸ�֤���
            invSht.Cells(i, "P").Value = certNo
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        msgBox "�ϸ�֤����Ѹ��µ��������!", vbInformation
    Else
        msgBox "δ�ҵ������: " & mgmtNo, vbExclamation
    End If
End Sub
