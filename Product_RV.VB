Function changeProp(oprd, oArry()) '修改产品属性 数组到 to catia
    Dim refprd: Set refprd = oprd.ReferenceProduct
    Dim pArry():   pArry() = getPropArry(oprd)
    Dim ParameterName(10)
        For i = 1 To UBound(oArry())
        If oArry(i) <> "" And pArry(i) <> oArry(i) Then
            Select Case i
                Case 1
            refprd.PartNumber = oArry(i)
                Case 2
            refprd.Nomenclature = oArry(i)
                Case 3
            refprd.Definition = oArry(i)
                Case 4
            oprd.Name = oArry(i)
                Case 5
            refprd.UserRefProperties.Item("iMaterial").Value = oArry(i)
                Case 8, 9
                    ParameterName(8) = "iDensity"
                    ParameterName(9) = "iThickness"
                    On Error Resume Next
                        Set oprt = refprd.Parent.Part
                        If Err.Number = 0 Then
                            oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters.Item(ParameterName(i)).Value = oArry(i)
                        Else
                            Err.Clear
                        End If
                    On Error GoTo 0
            End Select
        End If
        Next
End Function
Sub iniPrd(oprd, oDict)
    Dim refprd, oprt, sPara
    Set refprd = oprd.ReferenceProduct
    If oDict.exists(oprd.PartNumber) = False Then
        oDict(oprd.PartNumber) = 1
        Call iniUsp(oprd)
    End If
    Set refprd = oprd.ReferenceProduct
    If oprd.Products.Count > 0 Then
        For Each Product In oprd.Products
            Call iniPrd(Product, oDict)
        Next Product
    End If
End Sub
Sub iniUsp(oprd)
    Set refprd = oprd.ReferenceProduct
    Dim colls
    Dim parameterNames(5), parameterObj(5), propertyType, i
    parameterNames(0) = "cm"
    parameterNames(1) = "iBodys"
    parameterNames(2) = "iMaterial"
    parameterNames(3) = "iDensity"
    parameterNames(4) = "iMass"
    parameterNames(5) = "iThickness"
    '对于PRD创建参数，"iMaterial" "iMass" "iThickness"
    Set propertyType = CreateObject("Scripting.Dictionary")
    propertyType(parameterNames(0)) = "String"
    propertyType(parameterNames(1)) = "list"
    propertyType(parameterNames(2)) = "String"
    propertyType(parameterNames(3)) = "Density"
    propertyType(parameterNames(4)) = "Mass"
    propertyType(parameterNames(5)) = "Length"
    Set colls = refprd.UserRefProperties
    For i = 2 To 5
        If hasProperty(colls, parameterNames(i)) = False Then
            Select Case i
                Case 2
                   Set parameterObj(i) = colls.CreateString(parameterNames(i), "TBD")
                Case 4 To 5
                   Set parameterObj(i) = colls.CreateDimension(parameterNames(i), propertyType(parameterNames(i)), 0#)
            End Select
         Else
             Set parameterObj(i) = colls.Item(parameterNames(i))
        End If
    Next i
    Set colls = Nothing
    Dim oprt
    On Error Resume Next
    Set oprt = refprd.Parent.Part
        If Err.Number = 0 Then
           Call iniPrt(oprd, parameterNames())
        Else
           Err.Clear
           Dim oref, pubid, oPub
            Set colls = refprd.Publications
            i = 4
            If hasProperty(colls, parameterNames(i)) = False Then
                Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                Set oPub = colls.Add(parameterNames(i)) ' 添加发布
                colls.SetDirect parameterNames(i), oref ' 设置发布元素
                Set colls = Nothing
            End If
         End If
    On Error GoTo 0
    oprd.Update
    '释放对象
    Set colls = Nothing
End Sub
Sub iniPrt(oprd, parameterNames())
    Dim refprd, oprt
    Set refprd = oprd.ReferenceProduct
    Set oprt = refprd.Parent.Part
    Dim MBD: Set MBD = oprt.MainBody
    '==============创建cm参数集合
    On Error Resume Next
                Set colls = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
        If Err.Number <> 0 Then
                Err.Clear
                Set colls = oprt.Parameters.RootParameterSet.ParameterSets.CreateSet("cm")
                Set colls = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
        End If
    On Error GoTo 0
'    parameterNames(0) = "cm"
'    parameterNames(1) = "iBodys"
'    parameterNames(2) = "iMaterial"
'    parameterNames(3) = "iDensity"
'    parameterNames(4) = "iMass"
'    parameterNames(5) = "iThickness"
    Dim propertyType As Object
    Set propertyType = CreateObject("Scripting.Dictionary")
    propertyType(parameterNames(0)) = "String"
    propertyType(parameterNames(1)) = "list"
    propertyType(parameterNames(2)) = "String"
    propertyType(parameterNames(3)) = "Density"
    propertyType(parameterNames(4)) = "Mass"
    propertyType(parameterNames(5)) = "Length"
'==============创建PRT参数
    Dim parameterObj(5)
    For i = 1 To 5
        If hasProperty(colls, parameterNames(i)) = False Then
            Select Case i
                Case 1
                    Set parameterObj(i) = colls.CreateList(parameterNames(i)) 'list参数创建  iBodys
                Case 2
                   'Set parameterObj(i) = colls.CreateString(parameterNames(i), "TBD")  '材料参数创建 iMaterial
                Case 3 To 5
                   Set parameterObj(i) = colls.CreateDimension(parameterNames(i), propertyType(parameterNames(i)), 0#) 'dimension参数创建
            End Select
        Else
            Set parameterObj(i) = colls.Item(parameterNames(i))
        End If
    Next i
 '=============对于list增加mainbody
        Dim lst, lstobj
        Set lst = parameterObj(1)
        If hasProperty(lst.ValueList, MBD.Name) = False Then
            lst.ValueList.Add MBD
        End If
    Set colls = Nothing
    oprt.Update:    oprd.Update
'===============创建发布
'    parameterNames(0) = "cm"
'    parameterNames(1) = "iBodys"
'    parameterNames(2) = "iMaterial"
'    parameterNames(3) = "iDensity"
'    parameterNames(4) = "iMass"
'    parameterNames(5) = "iThickness"
    i = 1
    For i = 3 To 5
        Set pubs = refprd.Publications
        If hasProperty(pubs, parameterNames(i)) = False Then
            Select Case i
                Case 3, 4, 5
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(parameterNames(i)) ' 添加发布
                    pubs.SetDirect parameterNames(i), oref ' 设置发布元素
                Case 2
                    Set parameterObj(i) = refprd.UserRefProperties.Item(parameterNames(i))
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(parameterNames(i)) ' 添加发布
                    pubs.SetDirect parameterNames(i), oref ' 设置发布元素
            End Select
        End If
    Next
    oprt.Update: oprd.Update
' === 创建 EKL 规则 ===
    Dim strEKL(2)
    strEKL(0) = "CalM"
    strEKL(1) = "let lst(list) set lst=cm\iBodys let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1} cm\iMass=Vol*cm\iDensity"
    Set colls = oprt.Relations
    If hasProperty(colls, strEKL(0)) = False Then
        Set oRule = colls.CreateProgram(strEKL(0), "cal of mass", strEKL(1))
    Else
        Set oRule = colls.Item(strEKL(0))
        If oRule.Value <> strEKL(1) Then
            oRule.Modify strEKL(1)
        End If
    End If
'==============创建关系
    Dim oFormula, RLname, RLnote, RLtarget, RLsource
    Set colls = oprt.Relations
'----重量计算
    RLname = "CMAS"
    Set RLtarget = refprd.UserRefProperties.Item(parameterNames(4))
    RLsource = "cm\iMass"
        If hasProperty(colls, RLname) = False Then
            Set oFormula = colls.CreateFormula(RLname, " ", RLtarget, RLsource)
        Else
            Set oFormula = colls.Item(RLname)
                If oFormula.Value <> RLsource Then
                    oFormula.Modify RLsource
                End If
        End If
        Debug.Print oFormula.Value
    RLname = "CTK"
    Set RLtarget = refprd.UserRefProperties.Item(parameterNames(5))
    RLsource = "cm\iThickness"
        If hasProperty(colls, RLname) = False Then
            Set oFormula = colls.CreateFormula(RLname, " ", RLtarget, RLsource)
        Else
            Set oFormula = colls.Item(RLname)
                If oFormula.Value <> RLsource Then
                    oFormula.Modify RLsource
                End If
        End If
End Sub


