Attribute VB_Name = "Mainprocedures"
' ��׼ģ�� (�� Module1)
Option Explicit

' ��ʼ���������
Public Sub InitializeInventorySheet()
    ' �������ʵ�ʵĳ�ʼ������
    ' ʾ��:
    Dim invSheet As Worksheet
    Set invSheet = ThisWorkbook.Sheets("������")
    
    ' ��������ݣ�������ͷ��
    If invSheet.Range("A2").Value <> "" Then
        invSheet.Range("A2", invSheet.Cells(invSheet.Rows.Count, "A").End(xlUp)).EntireRow.Delete
    End If
    
    ' ���ó�ʼֵ���ʽ
    invSheet.Range("A1").Value = "��ƷID"
    invSheet.Range("B1").Value = "��Ʒ����"
    invSheet.Range("C1").Value = "�������"
    ' ��Ӹ����ʼ������...
End Sub

' ˢ������ʣ������
Public Sub RefreshAllRemainingDays()
    ' �������ʵ�ʵ�ˢ�´���
    ' ʾ��:
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Sheets("���ݹ���")
    
    Dim lastRow As Long
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' �����1���Ǳ���
        Dim expiryDate As Date
        On Error Resume Next ' ��ֹ��Ч����
        expiryDate = dataSheet.Cells(i, "C").Value ' ����C������Ч��
        
        If Err.Number = 0 Then
            ' ����ʣ������
            Dim daysLeft As Long
            daysLeft = expiryDate - Date
            dataSheet.Cells(i, "D").Value = daysLeft ' ����D����ʾʣ������
        Else
            Err.Clear
            dataSheet.Cells(i, "D").Value = "��Ч����"
        End If
    Next i
End Sub
