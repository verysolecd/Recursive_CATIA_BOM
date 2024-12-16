Sub recurPrd3(oprd, xlsht, oRowNb, LV)
    Dim bDict, oArry, i
            oArry = infoPrd(oprd, LV)
            Call Arry2sht(oArry, xlsht, oRowNb)
            If oprd.Products.Count > 0 Then '若产品1子级>0
                Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
                For i = 1 To oprd.Products.Count
                  If bDict.exists(oprd.Products.Item(i).PartNumber) = False Then
                        bDict(oprd.Products.Item(i).PartNumber) = 1
                        oRowNb = oRowNb + 1
                        Call recurPrd3(oprd.Products.Item(i), xlsht, oRowNb, LV + 1)
                    End If
               Next
            End If
End Sub
Function infoPrd(oprd, LV)
    Dim indqty, indLV, oArry, oDict
    indLV = 0
    indqty = 7
        oArry = getPropArry(oprd) '获取产品参数，数量暂设为0
        oArry(0) = LV ''赋予oprd LV值
    If TypeOf oprd.Parent Is Products Then    '若有父级产品'获取兄弟字典
        Set oDict = Dictbros(oprd.Parent.Parent)
        oArry(indqty) = oDict(oprd.PartNumber)       '获取oprd数量
    Else
        oArry(indqty) = 1  '若无父级产品，证明自己是根产品，数量为1
    End If
    infoPrd = oArry
End Function
Function Assmass(oprd)
    Dim total, children: total = 0
        Set children = oprd.Products
        If oprd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass (children.Item(i))
                total = children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value + total
            Next
                oprd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oprd.Parameters.Item("iMass").Value
        End If
   Assmass = total
End Function
Function getPropArry(oprd) '获取产品信息
    Dim refprd: Set refprd = oprd.ReferenceProduct
    Dim oprt, oArry(0 To 9), oUSPs, i
    With refprd
            oArry(0) = 0  'LV 请勿更改，因为写入excel时按oArry(0)的值做了识别
            oArry(1) = .PartNumber
            oArry(2) = .Nomenclature
            oArry(3) = .Definition
            oArry(4) = oprd.Name
            oArry(7) = 0  '数量
    End With
   With oUSPs
    For i = 5 To 9
        Dim prp(10)
            prp(5) = "iMaterial"
            prp(6) = "iMass"
            prp(8) = "iDensity"
            prp(9) = "iThickness"
            Select Case i
                Case 5, 6
                    Set oUSPs = refprd.UserRefProperties
                    If hasProperty(oUSPs, prp(i)) = False Then
                        oArry(i) = "N/A"
                    Else
                        oArry(i) = .Item(prp(i)).Value
                    End If
                Case 8, 9
                    On Error Resume Next
                            Set oUSPs = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                        If Err.Number = 0 Then '有cm情况
                            Set oUSPs = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                            If hasProperty(oUSPs, prp(i)) = False Then
                                oArry(i) = "N/A"
                            Else  '无cm情况
                                oArry(i) = .Item(prp(i)).Value
                            End If
                        Else '不是零件或零件未创建则密度为空,thickness 为空
                            oArry(8) = "N/A"
                            oArry(9) = "N/A"
                            Err.Clear
                        End If
                    On Error GoTo 0
            End Select
    Next
 End With
'    oArry(5) = .Item("iMaterial").Value
'    oArry(6) = .Item("iMass").Value
'    oArry(8) = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters.Item("iDensity").Value  '如果是part则存在密度，
'    oArry(9) = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters.Item("iThickness").Value
       getPropArry = oArry()
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
Function hasProperty(colls, propertyName)
    Dim propertyItem
    On Error Resume Next
        Set propertyItem = colls.Item(propertyName)
        If Err.Number <> 0 Then
            Err.Clear
            hasProperty = False
        Else
            hasProperty = True
        End If
    On Error GoTo 0
End Function