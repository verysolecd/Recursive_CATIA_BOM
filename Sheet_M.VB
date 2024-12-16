Sub Arry2sht(oArry, xlsht, RowNb) '读取到excel
        If oArry(0) = 0 Then ' 读取属性到excel进行修改
        '表格格式化
         xlsht.Cells(RowNb, 1).Value = oArry(9)  'iThickness
         xlsht.Cells(RowNb, 2).Value = oArry(6)  'iMass
            For i = 1 To 5
                 colNb = 2 * i + 1
                 xlsht.Cells(RowNb, colNb).Value = oArry(i)
            Next
                 xlsht.Cells(RowNb, 2).Value = oArry(6)
                 xlsht.Cells(RowNb, 13).Value = oArry(8)
        Else ' 读取属性到BOM
         xlsht.Cells(RowNb, 1).Value = RowNb - 2
            For i = 0 To UBound(oArry) - 2
                If i <> 4 Then
                    colNb = i + 2
                    xlsht.Cells(RowNb, colNb).Value = oArry(i)
                End If
            Next
            xlsht.Cells(RowNb, 11).Value = oArry(9)
'            Dim rg: Set rg = xlsht.Cells(RowNb, 10)
'            rg.FormulaR1C1 = "=RC[-2]*RC[-1]"
        End If
End Sub
Function sht2Arry(xlsht, RowNb) '========将Excel属性输入到Arry=======
        Dim oArry(0 To 8), i, colNb
        For i = 1 To 5
            colNb = 2 * i + 2
            oArry(i) = xlsht.Cells(RowNb, colNb).Value
        Next
        oArry(6) = ""
        oArry(7) = ""
        colNb = 14
        oArry(8) = xlsht.Cells(RowNb, colNb).Value
        sht2Arry = oArry
End Function
Function getsht(xlApp)
    Dim xlsht, xlWorkbook
    Dim oArry(0 To 11), i, indrow, indcol, RowNb
        Set xlWorkbook = xlApp.Workbooks.Add ' 创建一个新的工作簿
        Set xlsht = xlWorkbook.Sheets.Add  ' 在新工作簿中添加一个新的工作表
        oArry(1) = "No." & vbCrLf & "编号"
        oArry(2) = "Layout" & vbCrLf & "层级"
        oArry(3) = "PN" & vbCrLf & "零件号"
        oArry(4) = "Nomenclature" & vbCrLf & "英文名称"
        oArry(5) = "Definition" & vbCrLf & "中文名称"
        oArry(6) = "Picture" & vbCrLf & "图像"
        oArry(7) = "Material" & vbCrLf & "材料"
        oArry(8) = "Weight" & vbCrLf & "单质量(kg/PCS)"
        oArry(9) = "Quantity" & vbCrLf & "数量(PCS)"
        oArry(10) = "TotalWeight" & vbCrLf & "总质量(kg)"
        oArry(11) = "Thickness" & vbCrLf & "厚度(mm)"
        For i = 1 To UBound(oArry)
            xlsht.Cells(1, i).Value = i  '写入表头信息
            If oArry(i) <> "" Then
                xlsht.Cells(2, i).Value = oArry(i)  '写入表头信息
            End If
        Next
            xlsht.Columns(3).NumberFormatLocal = "@"  '设置格式
            xlsht.Columns(7).NumberFormatLocal = "@"
            xlsht.Columns(8).NumberFormatLocal = "0.000"
            xlsht.Columns(9).NumberFormatLocal = "0"
            xlsht.Columns(10).NumberFormatLocal = "0.000"
            xlsht.Columns(11).NumberFormatLocal = "0.00"
            xlsht.Rows(1).NumberFormatLocal = "0"
        Set getsht = xlsht
End Function
Sub LvMg(xlsht)
    Dim lastrow, i, LVcol
    With xlsht
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
            Set LVcol = .Columns(2)
            LVcol.HorizontalAlignment = xlLeft
            lastrow = .Cells(.Rows.Count, 2).End(xlUp).Row
        For i = 4 To lastrow
            If LVcol.Cells(i).Value <> "" Then
                .Rows(i).OutlineLevel = LVcol.Cells(i).Value ' 设置当前行的大纲级别
                LVcol.Cells(i).IndentLevel = LVcol.Cells(i).Value - 1 ' 设置当前单元格的缩进级别
            End If
        Next
    End With
    lastrow = xlsht.Cells(Rows.Count, 3).End(xlUp).Row
    xlsht.Rows(lastrow + 3 & ":" & Rows.Count).EntireRow.Hidden = True
    Columns("M:M").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
End Sub