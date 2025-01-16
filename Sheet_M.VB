Sub Arry2sht(oArry, xlsht, rownb) '===写入到excel====
    Dim i, ind, k, colNb
    Dim rng
    With xlsht
        If oArry(0) = 0 Then ' 读取属性到excel进行修改
        k = 3: Set rng = Union(.Cells(rownb, k), .Cells(rownb, k + 2), .Cells(rownb, k + 4), .Cells(rownb, k + 6))
        ind = 1
        For i = 1 To 4
           rng.Areas(i).Value = oArry(ind)(i)
        Next i
        ind = 3
                .Cells(rownb, 1).Value = oArry(ind)(4) '质量
                .Cells(rownb, 2).Value = oArry(ind)(5) '厚度
                .Cells(rownb, 11).Value = oArry(ind)(2) '材料
                .Cells(rownb, 13).Value = oArry(ind)(3) '密度
        Else ' 读取属性到BOM
            .Cells(rownb, 1).Value = rownb - 2
            .Cells(rownb, 2).Value = oArry(0)
            .Cells(rownb, 7).Value = oArry(2)  '数量
       ind = 1: colNb = 3
             Set rng = .Range(.Cells(rownb, colNb), .Cells(rownb, colNb - 1 + UBound(oArry(ind))))
             oArry(ind)(4) = ""
             rng.Value = oArry(ind)
       ind = 3 ' =======usp参数
         .Cells(rownb, 8).Value = oArry(ind)(4)  '质量
         .Cells(rownb, 11).Value = oArry(ind)(2) '材料
         .Cells(rownb, 12).Value = oArry(ind)(5) '厚度
         Dim rg: Set rg = .Cells(rownb, 9): rg.FormulaR1C1 = "=RC[-2]*RC[-1]"
        End If
    End With
End Sub
Function sht2Arry(xlsht, rownb) '========将Excel属性输入到Arry=======
        Dim i, ind, k, colNb
        Dim rng, LV: LV = 0
        Dim att_default(1 To 4)
        Dim att_usp(1 To 4)
        With xlsht
        colNb = 4: Set rng = Union(.Cells(rownb, colNb), .Cells(rownb, colNb + 2), .Cells(rownb, colNb + 4), .Cells(rownb, colNb + 6))
        ind = 1
        For i = 1 To rng.Areas.Count
               att_default(i) = rng.Areas(i).Value  '默认属性提取
        Next
        '========usp属性提取
        ind = 3
        att_usp(2) = .Cells(rownb, 12).Value
          att_usp(3) = .Cells(rownb, 14).Value
        sht2Arry = Array(0, att_default(), 0, att_usp())
        End With
        End Function
Function getsht(xlApp)
    Dim xlsht, xlWorkbook
    Dim oArry(1 To 16), i, indrow, indcol, rownb
        Set xlWorkbook = xlApp.Workbooks.Add ' 创建一个新的工作簿
        Set xlsht = xlWorkbook.Sheets.Add  ' 在新工作簿中添加一个新的工作表
        oArry(1) = "No." & vbCrLf & "编号"
        oArry(2) = "Layout" & vbCrLf & "层级"
        oArry(3) = "PN" & vbCrLf & "零件号"
        oArry(4) = "Nomenclature" & vbCrLf & "英文名称"
        oArry(5) = "Definition" & vbCrLf & "中文名称"
        oArry(6) = "Picture" & vbCrLf & "图像"
        oArry(7) = "Quantity" & vbCrLf & "数量(PCS)"
        oArry(8) = "Weight" & vbCrLf & "单质量(kg/PCS)"
        oArry(9) = "TotalWeight" & vbCrLf & "总质量(kg)"
        oArry(10) = ""
        oArry(11) = "Material" & vbCrLf & "材料"
        oArry(12) = "Thickness" & vbCrLf & "厚度(mm)"
        oArry(13) = "TS" & vbCrLf & "抗拉Mpa"
        oArry(14) = "YS" & vbCrLf & "屈服Mpa"
        oArry(15) = "EL" & vbCrLf & "延伸率%"
        Dim rng
       Set rng = xlsht.Range(xlsht.Cells(2, 1), xlsht.Cells(2, UBound(oArry())))
       rng.Value = oArry
            xlsht.Columns(3).NumberFormatLocal = "@"  '设置格式
            xlsht.Columns(7).NumberFormatLocal = "0"
            xlsht.Columns(8).NumberFormatLocal = "0.000"
            xlsht.Columns(9).NumberFormatLocal = "0.000"
            'xlsht.Columns(10).NumberFormatLocal = "@"
            xlsht.Columns(11).NumberFormatLocal = "@"
            xlsht.Columns(12).NumberFormatLocal = "0.00"
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
    Columns("Q:Q").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
End Sub

