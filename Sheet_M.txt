'============Sheet操作=============
Sub Arry2sht(oArry, xlsht, RowNb) '读取到excel
        xlsht.Cells(RowNb, 1).Value = RowNb - 2
        xlsht.Cells(RowNb, 2).Value = oArry(0)
        If oArry(0) = 0 Then ' 读取属性到excel进行修改
            For i = 1 To UBound(oArry) - 2
                 colNb = 2 * i + 1
                 xlsht.Cells(RowNb, colNb).Value = oArry(i)
            Next
                 xlsht.Cells(RowNb, 14).Value = oArry(8)
        Else ' 读取属性到BOM
            For i = 0 To UBound(oArry) - 1
                If i <> 4 Then
                    colNb = i + 2
                    xlsht.Cells(RowNb, colNb).Value = oArry(i)
                End If
            Next
            Dim rg: Set rg = xlsht.Cells(RowNb, 10)
                   rg.FormulaR1C1 = "=RC[-2]*RC[-1]"
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
        colNb = 15
        oArry(8) = xlsht.Cells(RowNb, colNb).Value
        sht2Arry = oArry
End Function
Function getsht(xlApp)
    Dim xlsht, xlWorkbook
    Dim oArry(0 To 10), i, indrow, indcol, RowNb
        Set xlWorkbook = xlApp.Workbooks.Add ' 创建一个新的工作簿
        Set xlsht = xlWorkbook.Sheets.Add  ' 在新工作簿中添加一个新的工作表
            oArry(0) = "No." & vbCrLf & "编号"
            oArry(1) = "Layout" & vbCrLf & "层级"
            oArry(2) = "PN" & vbCrLf & "零件号"
            oArry(3) = "Nomenclature" & vbCrLf & "英文名称"
            oArry(4) = "Definition" & vbCrLf & "中文名称"
            oArry(5) = "Picture" & vbCrLf & "图像"
            oArry(6) = "Material" & vbCrLf & "材料"
            oArry(7) = "Weight" & vbCrLf & "单质量(kg/PCS)"
            oArry(8) = "Quantity" & vbCrLf & "数量(PCS)"
            oArry(9) = "TotalWeight" & vbCrLf & "总质量(kg)"
            oArry(10) = "Thickness" & vbCrLf & "厚度(mm)"
        For i = LBound(oArry) To UBound(oArry)
            xlsht.Cells(1, i + 1).Value = i  '写入表头信息
            If oArry(i) <> "" Then
                xlsht.Cells(2, i + 1).Value = oArry(i)  '写入表头信息
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
Sub LvMg(xlsht, rowcol())
    Dim lastrow, i, colB
    With xlsht
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
            Set colB = .Columns(rowcol(1) - 1)
            colB.HorizontalAlignment = xlLeft
            lastrow = .Cells(.Rows.Count, rowcol(1) - 1).End(xlUp).Row
        For i = 4 To lastrow
            If colB.Cells(i).Value <> "" Then
                .Rows(i).OutlineLevel = colB.Cells(i).Value ' 设置当前行的大纲级别
                colB.Cells(i).IndentLevel = colB.Cells(i).Value ' 设置当前单元格的缩进级别
            End If
        Next
    End With
    lastrow = xlsht.Cells(Rows.Count, 3).End(xlUp).Row
    xlsht.Rows(lastrow + 3 & ":" & Rows.Count).EntireRow.Hidden = True
    Columns("M:M").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
End Sub