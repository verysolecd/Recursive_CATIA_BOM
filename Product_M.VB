Sub iniPrd(oPrd, oDict)
        Dim refPrd, oPrt, sPara, oArry(0 To 4), bDict
        Set refPrd = oPrd.ReferenceProduct
        Set sPara = refPrd.UserRefProperties
        oArry(0) = "cm"
        oArry(1) = "iBodys"
        oArry(2) = "iDensity"
        oArry(3) = "iMass"
        oArry(4) = "iMaterial"
        If oDict.exists(oPrd.PartNumber) = False Then
            oDict(oPrd.PartNumber) = 1
            Call createUsrProp(oPrd, oArry())
        End If
        If oPrd.Products.Count > 0 Then
            For Each Product In oPrd.Products
                Call iniPrd(Product, oDict)
            Next Product
        End If
End Sub
Sub createUsrProp(oPrd, Arry())
    Dim refPrd, sPara, i, oArry(), xPara(0 To 10)
    Set refPrd = oPrd.ReferenceProduct
    Set xPara(0) = refPrd.UserRefProperties
    oArry() = Arry()
        On Error Resume Next  '========对于prd 创建mass和material
                Set xPara(3) = xPara(0).Item(oArry(3))
                    If Err.Number <> 0 Then
                        Set xPara(3) = xPara(0).CreateDimension(oArry(3), "MASS", 0#)
                        Err.Clear
                    End If
                Set xPara(4) = xPara(0).Item(oArry(4))
                If Err.Number <> 0 Then
                    Set xPara(4) = xPara(0).CreateString(oArry(4), "未赋值")
                    Err.Clear
                End If
'        On Error GoTo 0
'        On Error Resume Next
                Set oPrt = refPrd.Parent.Part
            If Err.Number = 0 Then
                Call iniPrt(oPrd, oArry())
            Else
                i = 3
                Set xPara(0) = refPrd.Publications
                Set oRef = refPrd.CreateReferenceFromName(xPara(i).Name)
                Set pubid = xPara(0).Item(oArry(i))
                    If Error <> 0 Then
                        Set oPub = xPara(0).Add(oArry(i)) ' 添加发布
                        xPara(0).SetDirect oArry(i), oRef ' 设置发布元素
                       Err.Clear
                    End If
            End If
        On Error GoTo 0
End Sub
Sub iniPrt(oPrd, Arry())
    Dim refPrd, oPrt, oArry(), xPara(0 To 10), i, pubid, MBD, oRef
    Set refPrd = oPrd.ReferenceProduct
    Set oPrt = refPrd.Parent.Part
    Set MBD = oPrt.MainBody
         oArry() = Arry()
    On Error Resume Next
			Set xPara(0) = oPrt.Parameters.RootParameterSet.ParameterSets.Item(oArry(0)).DirectParameters
		If Err.Number <> 0 Then
			Set xPara(0) = oPrt.Parameters.RootParameterSet.ParameterSets.CreateSet(oArry(0))
			Set xPara(0) = oPrt.Parameters.RootParameterSet.ParameterSets.Item(oArry(0)).DirectParameters
			Set xPara(1) = xPara(0).CreateList(oArry(1))  ' 创建 iBodys 清单
			Set xPara(2) = xPara(0).CreateDimension(oArry(2), "DENSITY", 7850#)  ' 创建密度参数
			Set xPara(3) = xPara(0).CreateDimension(oArry(3), "MASS", 0#)  ' 创建质量参数
			Err.Clear
		Else
			Set xPara(1) = xPara(0).Item(oArry(1))
					If Err.Number <> 0 Then
						Set xPara(1) = xPara(0).CreateList(oArry(1))  ' 创建 iBodys 清单
						   Err.Clear
					 End If
						Dim lst, lstobj
			Set lst = xPara(1)
			Set lstobj = lst.ValueList.Item(MBD.Name)
					If Err.Number <> 0 Then
						lst.ValueList.Add MBD
						 Err.Clear
					End If
			Set xPara(2) = xPara(0).Item(oArry(2))
					If Err.Number <> 0 Then
						  Set xPara(2) = xPara(0).CreateDimension(oArry(2), "DENSITY", 7850#)  ' 创建密度参数
								Err.Clear
					  End If
			 Set xPara(3) = xPara(0).Item(oArry(3))
					If Err.Number <> 0 Then
							Set xPara(3) = xPara(0).CreateDimension(oArry(3), "MASS", 0#)  ' 创建质量参数
					   Err.Clear
					End If
			Err.Clear
		End If
    On Error GoTo 0
      oArry(0) = "cm\"
    ' === 定义 EKL 规则 ===
     xVol = "=Vol*"
     sp = " "
        Dim oEKL(0 To 10)
        oEKL(2) = "let lst(list) set lst=" & oArry(0) & oArry(1)              '2   'let lst(list) 'set lst= cm\iBodys
        oEKL(3) = " let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1}"
        oEKL(4) = oArry(0) & oArry(3) & xVol & oArry(0) & oArry(2)  '4   cm\iMass=Vol*cm\iDensity  '& xVol & oArry(0)& oArry(3)
        oEKL(0) = "CalM"
        oEKL(1) = oEKL(2) & sp & oEKL(3) & sp & oEKL(4)
    ' === 创建 EKL 规则 ===
    On Error Resume Next
        Set oRule = oPrt.Relations.Item(oEKL(0))
        If Err.Number <> 0 Then
            Set oRule = oPrt.Relations.CreateProgram(oEKL(0), "cal of mass", oEKL(1))
        Else
            If oRule.Value <> oEKL(1) Then
                oRule.Modify oEKL(1)
            End If
            Err.Clear
        End If
      On Error GoTo 0
        ' === 创建 关系 ===
    Set refPrd = oPrd.ReferenceProduct
    Set xPara(0) = refPrd.UserRefProperties
    Dim Fc, Fcname, Fcnote, Fctarget, myFc
    On Error Resume Next
            Fcname = "CMAS"
            Fcnote = "产品iMass=cm/iMass"
    Set Fctarget = xPara(0).Item("iMass")
            myFc = "cm\iMass"
            Set Fc = oPrt.Relations.Item(Fcname)
            If Err.Number <> 0 Then
                 Set Fc = oPrt.Relations.CreateFormula(Fcname, Fcnote, Fctarget, myFc)
                 Err.Clear
            Else
                Fc.Modify myFc
            End If
On Error GoTo 0
    oPrt.Update:    oPrd.Update
        Set xPara(0) = refPrd.Publications
                Set xPara(4) = refPrd.UserRefProperties.Item(oArry(4))
            For i = 2 To 3
                Set oRef = refPrd.CreateReferenceFromName(xPara(i).Name)
                On Error Resume Next
                    Set pubid = xPara(0).Item(oArry(i))
                If Error <> 0 Then
                    Set oPub = xPara(0).Add(oArry(i)) ' 添加发布
                    xPara(0).SetDirect oArry(i), oRef ' 设置发布元素
                End If
                On Error GoTo 0
            Next
    oPrt.Update: oPrd.Update
End Sub
Function getPropArry(oPrd) '获取产品信息
    Dim refPrd, oPrt, sPara, oArry(0 To 8), xPara(0 To 10), i, pubid, MBD, oRef:
    Set refPrd = oPrd.ReferenceProduct
    Set sPara = refPrd.UserRefProperties
          With refPrd
            oArry(0) = 0  'LV
            oArry(1) = .PartNumber
            oArry(2) = .Nomenclature
            oArry(3) = .Definition
            oArry(4) = oPrd.Name
            oArry(5) = sPara.Item("iMaterial").Value
            oArry(6) = sPara.Item("iMass").Value
            oArry(7) = 0  '数量
           On Error Resume Next
               Set oPrt = refPrd.Parent.Part
                 If Err.Number = 0 Then
                   Dim sPara1: Set sPara1 = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                   oArry(8) = sPara1.Item("iDensity").Value
                 Else
                    oArry(8) = ""
                    Err.Clear
                 End If
          End With
       getPropArry = oArry()
End Function
Function changeProp(oPrd, oArry()) '修改产品属性 数组到 to catia
    Dim refPrd: Set refPrd = oPrd.ReferenceProduct
    Dim sPara: Set sPara = refPrd.UserRefProperties
    Dim pArry():   pArry() = getPropArry(oPrd)
    With refPrd
        For i = 1 To UBound(oArry())
            If oArry(i) <> "" Or oArry(i) <> 0 Then
               pArry(i) = oArry(i)
            End If
        Next
            .PartNumber = pArry(1)
            .Nomenclature = pArry(2)
            .Definition = pArry(3)
            oPrd.Name = pArry(4)
            sPara.Item("iMaterial").Value = pArry(5)
        End With
            On Error Resume Next
               Set oPrt = refPrd.Parent.Part
                 If Err.Number = 0 Then
                   Dim sPara1: Set sPara1 = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                    sPara1.Item("iDensity").Value = pArry(8)
                 Else
                    Err.Clear
                 End If
End Function
Function Assmass(oPrd)
    Dim total, children: total = 0
        Set children = oPrd.Products
        If oPrd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass (children.Item(i))
                total = children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value + total
            Next
                oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oPrd.Parameters.Item("iMass").Value
        End If
   Assmass = total
End Function
Sub recurPrd3(oPrd, xlsht, oRowNb, LV)
    Dim bDict, oArry, i
            oArry = infoPrd(oPrd, LV)
            Call Arry2sht(oArry, xlsht, oRowNb)
            If oPrd.Products.Count > 0 Then '若产品1子级>0
                Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
                For i = 1 To oPrd.Products.Count
                  If bDict.exists(oPrd.Products.Item(i).PartNumber) = False Then
                        bDict(oPrd.Products.Item(i).PartNumber) = 1
                        oRowNb = oRowNb + 1
                        Call recurPrd3(oPrd.Products.Item(i), xlsht, oRowNb, LV + 1)
                    End If
               Next
            End If
End Sub
Function infoPrd(oPrd, LV)
    Dim indqty, indLV, oArry, oDict
    indLV = 0: indqty = 7
        oArry = getPropArry(oPrd) '获取产品参数，数量暂设为0
        oArry(0) = LV ''赋予oprd LV值
    If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
        Set oDict = Dictbros(oPrd.Parent.Parent)
        oArry(indqty) = oDict(oPrd.PartNumber)       '获取oprd数量
    Else
        oArry(indqty) = 1  '若无父级产品，证明自己是根产品，数量为1
    End If
    infoPrd = oArry
End Function
Function Dictbros(oPrd)  '获取兄弟字典
    Dim oDict, children, i, pn
    Set oDict = CreateObject("Scripting.Dictionary")
    Set children = oPrd.Products
    For i = 1 To children.Count
        pn = children.Item(i).PartNumber
        If oDict.exists(pn) = True Then
            oDict(pn) = oDict(pn) + 1
        Else
            oDict(pn) = 1
        End If
        Next
    Set Dictbros = oDict
End Function
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