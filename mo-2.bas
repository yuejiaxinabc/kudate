Attribute VB_Name = "ģ��2"
' ��ģ����������´���
Option Explicit

' ���²ֿ���������������
Sub UpdateWarehouseStock()
    On Error Resume Next
    Dim invSht As Worksheet, whSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("������")
    Set whSht = ThisWorkbook.Sheets("�ֿ�")
    
    ' ��ȡ�ֿ��б�
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    
    ' �������вֿ�����Ϊ0
    Dim i As Long
    For i = 2 To whLastRow
        whSht.Cells(i, "F").Value = 0
    Next i
    
    ' ͳ��ÿ���ֿ���ڿ�����
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    For i = 2 To invLastRow
        Dim whName As String
        whName = invSht.Cells(i, "W").Value ' W���������ֿ�
        
        If whName <> "" Then
            ' ���Ҳֿ������ڲֿ���е�λ��
            Dim j As Long
            For j = 2 To whLastRow
                If whSht.Cells(j, "C").Value = whName Then
                    If invSht.Cells(i, "AA").Value = "�ڿ�" Then
                        whSht.Cells(j, "F").Value = whSht.Cells(j, "F").Value + 1
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' ��ӿ��Ԥ��
    For i = 2 To whLastRow
        If whSht.Cells(i, "F").Value < 5 Then
            whSht.Cells(i, "F").Interior.Color = RGB(255, 200, 200) ' ��ɫ����
        Else
            whSht.Cells(i, "F").Interior.ColorIndex = xlNone ' �������ɫ
        End If
    Next i
    
    msgBox "�ֿ����Ѹ���!", vbInformation
End Sub

' ���µ����ֿ��棨�Ӽ�������
Sub UpdateSingleWarehouse(whName As String, operation As String)
    Dim whSht As Worksheet
    Set whSht = ThisWorkbook.Sheets("�ֿ�")
    
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

' �������״̬�仯���
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    ' ����ı�����ڿ�״̬��(AA��)
    If Not Intersect(Target, Me.Columns("AA")) Is Nothing Then
        If Target.Row > 1 Then
            Dim whName As String
            whName = Me.Cells(Target.Row, "W").Value ' W���������ֿ�
            
            If whName <> "" Then
                ' �ж�״̬�仯
                If Target.Value = "�ڿ�" Then
                    ' ���Ӷ�Ӧ�ֿ���
                    UpdateSingleWarehouse whName, "add"
                ElseIf Target.Value = "����" Then
                    ' ���ٶ�Ӧ�ֿ���
                    UpdateSingleWarehouse whName, "subtract"
                End If
            End If
        End If
    End If
    
    Application.EnableEvents = True
End Sub
