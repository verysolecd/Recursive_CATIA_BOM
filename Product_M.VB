Sub recurPrd(oprd, xlsht, currRow, LV)
    Dim bDict, i
            Arry2sht bominfoPrd(oprd, LV), xlsht, currRow
            If oprd.Products.Count > 0 Then '若产品1子级>0
                Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
                For i = 1 To oprd.Products.Count
                  If bDict.exists(oprd.Products.Item(i).PartNumber) = False Then
                        bDict(oprd.Products.Item(i).PartNumber) = 1
                        recurPrd oprd.Products.Item(i), xlsht, oRowNb + 1, LV + 1
                    End If
               Next
            End If
End Sub
Function bominfoPrd(oprd, LV)
    Dim indqty, indLV, oArry(), oDict
    indLV = 0
    indqty = 5
        oArry() = getPropArry(oprd) '获取产品参数，数量暂设为0
        oArry(0) = LV ''赋予oprd LV值
    If TypeOf oprd.Parent Is Products Then    '若有父级产品'获取兄弟字典
        Set oDict = Dictbros(oprd.Parent.Parent)
        oArry(2)(indqty) = oDict(oprd.PartNumber)       '获取oprd数量
    Else
        oArry(2)(indqty) = 1  '若无父级产品，证明自己是根产品，数量为1
    End If
    oArry(1)(4) = ""
    bominfoPrd = oArry
End Function
Function Assmass(oprd)
    Dim i, total, children: total = 0
        Set children = oprd.Products
        If children.Count > 0 Then
            For i = 1 To children.Count
                Assmass (children.Item(i))
                total = children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value + total
            Next
                oprd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oprd.ReferenceProduct.UserRefProperties.Item("iMass").Value 'oprd.Parameters.Item("iMass").Value
        End If
   Assmass = total
End Function
Function getPropArry(oprd) '获取产品信息
    Dim oprt, refprd: Set refprd = oprd.ReferenceProduct
    Dim colls, i, LV: LV = 0
    Dim oArry(0 To 2)
'====arr_def 我是默认属性数组===
With refprd
    Dim att_default(1 To 4)
    att_default(1) = .PartNumber
    att_default(2) = .Nomenclature
    att_default(3) = .definition
    att_default(4) = oprd.Name
End With
    '====arr_usp() 我是自定义属性数组===
    Dim att_usp(1 To 5)
    Dim att_usp_Names(1 To 5)
    
    att_usp_Names(1) = "iMaterial"
    att_usp_Names(2) = "iThickness"
    att_usp_Names(3) = "iMass"
    att_usp_Names(4) = "iDensity"
    att_usp_Names(5) = "qty" '数量    
    For i = 1 To 5
        Select Case i
        Case 1, 3
            Set colls = refprd.UserRefProperties
            att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
            Set colls = Nothing
        Case 2, 4
            On Error Resume Next
                Set oprt = refprd.Parent.Part
                Set colls = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                '非part情况，无cm情况，，有cm但无参，不是零件或零件未创建则密度为空,thickness 为空
            If Err.Number <> 0 Then
                att_usp(i) = "N\A"
                Err.Clear
            Else
                att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
            End If
        Case 5: att_usp(i) = ""
        End Select
    Next    
    '====
        oArry(0) = LV
        oArry(1) = att_default
        oArry(2) = att_usp
        getPropArry = oArry
End Function
Function Dictbros(oprd)  '获取兄弟字典
    Dim oDict, i, pn: Set oDict = CreateObject("Scripting.Dictionary")
        For i = 1 To oprd.Products.Count
            pn = oprd.Products.Item(i).PartNumber
            If oDict.exists(pn) = True Then
                oDict(pn) = oDict(pn) + 1
            Else
                oDict(pn) = 1
            End If
        Next
    Set Dictbros = oDict
End Function
Function hasProperty(colls, parameterName)
    Dim propertyItem
    On Error Resume Next
        Set propertyItem = colls.Item(parameterName)
        If Err.Number <> 0 Then
            Err.Clear
            hasProperty = False
        Else
            hasProperty = True
        End If
    On Error GoTo 0
End Function
Function thisParameterValue(colls, parameterName)
        On Error Resume Next
            thisParameterValue = colls.Item(parameterName).Value
          If Err.Number <> 0 Then
                Err.Clear
                thisParameterValue = "N\A"
               End If
        On Error GoTo 0
End Function
Function whois2rv()
   '获取要读取或修改的  全局变量 Prd2wt 定义
    Dim infobox: infobox = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
        Select Case infobox
        Case 6  '===选择“是”,进行产品选择====
            xlApp.Windows(1).WindowState = xlMinimized
            Dim oSel, status, iType(0), selPrd
            Set oSel = CATIA.Activedocument.Selection
            oSel.Clear
            iType(0) = "Product"
            If oSel.Count2 = 0 Then
                status = oSel.SelectElement3(iType, "请在结构树选择被连接产品", True, 2, False)
            End If
            If status = "Cancel" Then
                MsgBox "用户取消操作，程序将退出"
                Exit Function
            End If
            If status = "Normal" And oSel.Count2 = 1 Then
                    Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
                    oSel.Clear
            Else
                aMsgBox "请只选择一个产品"
                Exit Function
            End If
        Case 7 '===选择“否”====
            Set selPrd = CATIA.Activedocument.product
        Case 2 '===选择“取消”====
            aMsgBox "程序运行取消，将退出"
            Exit Function
        End Select
        
          Set whois2rv = selPrd
            xlApp.Windows(1).WindowState = xlMaximized
            infobox = "你选择的产品是" & selPrd.PartNumber & "是否继续"
          MsgBox infobox
End Function


