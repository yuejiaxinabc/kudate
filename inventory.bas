Attribute VB_Name = "inventory"
Option Explicit

' ��ʼ���������
Public Sub Inventory_Initialize()
    ' ������ӿ������ĳ�ʼ������
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
    invSheet.Range("D1").Value = "��Ч��"
    
    ' ���ñ����ʽ
    With invSheet.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241) ' ǳ��ɫ����
    End With
    
    ' �����п�
    invSheet.Columns("A:D").AutoFit
End Sub
