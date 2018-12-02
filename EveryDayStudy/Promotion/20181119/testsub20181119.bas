Attribute VB_Name = "testsub20181119"
Sub test01()
    '
    '根据是否capping内容校验capping设置方式
    '做一个二级下拉菜单,第二级单元格中的内容由上一级的内容决定，有某些情况下会违背这一原则，做一个校验功能控制不允许违背这一原则
    '1设置添加名称管理器内容时候，范围需要选择工作簿
    '日本:1234  韩国:5678
    
    With ThisWorkbook.Sheets("Sheet3")
        Dim maxRow
        Dim i
        maxRow = .Range("I65536").End(xlUp).row
        For i = 1 To maxRow
            .Range("I" & i).Interior.Color = RGB(255, 255, 255) '设置某个单元格的颜色
            If .Range("H" & i) = "日本" Then
                If .Range("I" & i) <> "1" And .Range("I" & i) <> "2" And .Range("I" & i) <> "3" And .Range("I" & i) <> "4" And .Range("I" & i) <> "" Then
                    MsgBox "输入有误"
                    .Range("I" & i).Interior.Color = RGB(255, 0, 0) '设置某个单元格的颜色
                End If
            End If
            If .Range("H" & i) = "韩国" Then
                If .Range("I" & i) <> "5" And .Range("I" & i) <> "6" And .Range("I" & i) <> "7" And .Range("I" & i) <> "8" And .Range("I" & i) <> "" Then
                    MsgBox "输入有误"
                    .Range("I" & i).Interior.Color = RGB(255, 0, 0)
                End If
            End If
'            .Range("a10:g26").Borders.LineStyle = xlContinuous
            .Range("I" & i & ":" & "I" & i).Borders.LineStyle = xlSlantDashDot '划线目前没找到合适的线
        Next
    End With
    
End Sub
