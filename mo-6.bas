Attribute VB_Name = "ģ��6"
Sub RefreshAllWarehouseStock()
    Dim invSht As Worksheet, whSht As Worksheet
    Set invSht = ThisWorkbook.Sheets("������")
    Set whSht = ThisWorkbook.Sheets("�ֿ�")
    
    ' ��ղֿ�����
    Dim whLastRow As Long
    whLastRow = whSht.Cells(whSht.Rows.Count, "C").End(xlUp).Row
    For i = 2 To whLastRow
        whSht.Cells(i, "F").Value = 0
    Next i
    
    ' ���¼������вֿ���
    Dim invLastRow As Long
    invLastRow = invSht.Cells(invSht.Rows.Count, "E").End(xlUp).Row
    
    For i = 2 To invLastRow
        Dim whName As String
        whName = invSht.Cells(i, "W").Value
        Dim qty As Long
        qty = 0
        If IsNumeric(invSht.Cells(i, "AB").Value) Then
            qty = invSht.Cells(i, "AB").Value
        End If
        
        If whName <> "" Then
            ' ���ҲֿⲢ��������
            For j = 2 To whLastRow
                If whSht.Cells(j, "C").Value = whName Then
                    whSht.Cells(j, "F").Value = whSht.Cells(j, "F").Value + qty
                    Exit For
                End If
            Next j
        End If
    Next i
    
    msgBox "�ֿ�����ˢ��!", vbInformation
End Sub
