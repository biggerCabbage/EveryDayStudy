Attribute VB_Name = "testsub20181119"
Sub test01()
    '
    '�����Ƿ�capping����У��capping���÷�ʽ
    '��һ�����������˵�,�ڶ�����Ԫ���е���������һ�������ݾ�������ĳЩ����»�Υ����һԭ����һ��У�鹦�ܿ��Ʋ�����Υ����һԭ��
    '1����������ƹ���������ʱ�򣬷�Χ��Ҫѡ������
    '�ձ�:1234  ����:5678
    
    With ThisWorkbook.Sheets("Sheet3")
        Dim maxRow
        Dim i
        maxRow = .Range("I65536").End(xlUp).row
        For i = 1 To maxRow
            .Range("I" & i).Interior.Color = RGB(255, 255, 255) '����ĳ����Ԫ�����ɫ
            If .Range("H" & i) = "�ձ�" Then
                If .Range("I" & i) <> "1" And .Range("I" & i) <> "2" And .Range("I" & i) <> "3" And .Range("I" & i) <> "4" And .Range("I" & i) <> "" Then
                    MsgBox "��������"
                    .Range("I" & i).Interior.Color = RGB(255, 0, 0) '����ĳ����Ԫ�����ɫ
                End If
            End If
            If .Range("H" & i) = "����" Then
                If .Range("I" & i) <> "5" And .Range("I" & i) <> "6" And .Range("I" & i) <> "7" And .Range("I" & i) <> "8" And .Range("I" & i) <> "" Then
                    MsgBox "��������"
                    .Range("I" & i).Interior.Color = RGB(255, 0, 0)
                End If
            End If
'            .Range("a10:g26").Borders.LineStyle = xlContinuous
            .Range("I" & i & ":" & "I" & i).Borders.LineStyle = xlSlantDashDot '����Ŀǰû�ҵ����ʵ���
        Next
    End With
    
End Sub
