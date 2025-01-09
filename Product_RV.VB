Function changeProp(oprd, oArry()) '修改产品属性 数组到 to catia
    Dim refprd: Set refprd = oprd.ReferenceProduct
    Dim oprt
    Dim pArry():   pArry() = getPropArry(oprd)
    Dim colls
    Dim uspNames(1 To 5)
    uspNames(1) = "iMaterial"
    uspNames(2) = "iThickness"
    uspNames(3) = "iMass"
    uspNames(4) = "iDensity"
    uspNames(5) = ""   '数量
    Dim parameterName(10), i, ind
    ind = 1
        For i = 1 To 4
        If oArry(ind)(i) <> "" And oArry(ind)(i) <> pArry(ind)(i) Then
            Select Case i
                Case 1: refprd.PartNumber = oArry(1)(i)
                Case 2: refprd.Nomenclature = oArry(1)(i)
                Case 3: refprd.definition = oArry(1)(i)
                Case 4: oprd.Name = oArry(1)(i)
            End Select
        End If
        Next
    ind = 2
     i = 1
     If oArry(ind)(i) <> "" And pArry(ind)(i) <> oArry(ind)(i) Then
     Set colls = refprd.UserRefProperties
     colls.Item(uspNames(i)).Value = oArry(ind)(i)
     End If
     i = 4
   On Error Resume Next
    Set oprt = refprd.Parent.Part
    Set colls = oprt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
    If Err.Number = 0 Then
        If oArry(ind)(i) <> "" And oArry(ind)(i) <> pArry(ind)(i) Then
            colls.Item(uspNames(i)).Value = oArry(ind)(i)
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function
Sub iniPrd(oprd, oDict)
'Dim product
    If oDict.exists(oprd.PartNumber) = False Then
        oDict(oprd.PartNumber) = 1
        Call initMyRefPrd(oprd)
    End If
    If oprd.Products.Count > 0 Then
        For Each product In oprd.Products
            Call iniPrd(product, oDict)
        Next product
    End If
End Sub
Sub initMyRefPrd(oprd)
    Set refprd = oprd.ReferenceProduct
    Dim colls
    Dim attNames(5), parameterObj(5), attType, i
    attNames(0) = "cm"
    attNames(1) = "iBodys"
    attNames(2) = "iMaterial"
    attNames(3) = "iDensity"
    attNames(4) = "iMass"
    attNames(5) = "iThickness"
    '对于PRD创建参数，"iMaterial" "iMass" "iThickness"
    Set attType = CreateObject("Scripting.Dictionary")
    attType(attNames(0)) = "String"
    attType(attNames(1)) = "list"
    attType(attNames(2)) = "String"
    attType(attNames(3)) = "Density"
    attType(attNames(4)) = "Mass"
    attType(attNames(5)) = "Length"
    Set colls = refprd.UserRefProperties
    For i = 2 To 5
        If hasProperty(colls, attNames(i)) = False Then
            Select Case i
                Case 2
                   Set parameterObj(i) = colls.CreateString(attNames(i), "TBD")
                Case 4 To 5
                   Set parameterObj(i) = colls.CreateDimension(attNames(i), attType(attNames(i)), 0#)
            End Select
         Else
             Set parameterObj(i) = colls.Item(attNames(i))
        End If
    Next i
    Set colls = Nothing
    Dim oprt
    On Error Resume Next
    Set oprt = refprd.Parent.Part
        If Err.Number = 0 Then
           Call iniPrt(oprd, attNames())
        Else
           Err.Clear
           Dim oref, pubid, oPub
            Set colls = refprd.Publications
            i = 4
            If hasProperty(colls, attNames(i)) = False Then
                Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                Set oPub = colls.Add(attNames(i)) ' 添加发布
                colls.SetDirect attNames(i), oref ' 设置发布元素
                Set colls = Nothing
            End If
         End If
    On Error GoTo 0
    oprd.Update
    '释放对象
    Set colls = Nothing
End Sub
Sub iniPrt(oprd, attNames())
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
'    attNames(0) = "cm"
'    attNames(1) = "iBodys"
'    attNames(2) = "iMaterial"
'    attNames(3) = "iDensity"
'    attNames(4) = "iMass"
'    attNames(5) = "iThickness"
    Dim attType As Object
    Set attType = CreateObject("Scripting.Dictionary")
    attType(attNames(0)) = "String"
    attType(attNames(1)) = "list"
    attType(attNames(2)) = "String"
    attType(attNames(3)) = "Density"
    attType(attNames(4)) = "Mass"
    attType(attNames(5)) = "Length"
'==============创建PRT参数
    Dim parameterObj(5)
    For i = 1 To 5
        If hasProperty(colls, attNames(i)) = False Then
            Select Case i
                Case 1
                    Set parameterObj(i) = colls.CreateList(attNames(i)) 'list参数创建  iBodys
                Case 2
                   'Set parameterObj(i) = colls.CreateString(attNames(i), "TBD")  '材料参数创建 iMaterial
                Case 3 To 5
                   Set parameterObj(i) = colls.CreateDimension(attNames(i), attType(attNames(i)), 0#) 'dimension参数创建
            End Select
        Else
            Set parameterObj(i) = colls.Item(attNames(i))
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
'    attNames(0) = "cm"
'    attNames(1) = "iBodys"
'    attNames(2) = "iMaterial"
'    attNames(3) = "iDensity"
'    attNames(4) = "iMass"
'    attNames(5) = "iThickness"
    i = 1
    For i = 3 To 5
        Set pubs = refprd.Publications
        If hasProperty(pubs, attNames(i)) = False Then
            Select Case i
                Case 3, 4, 5
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(attNames(i)) ' 添加发布
                    pubs.SetDirect attNames(i), oref ' 设置发布元素
                Case 2
                    Set parameterObj(i) = refprd.UserRefProperties.Item(attNames(i))
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(attNames(i)) ' 添加发布
                    pubs.SetDirect attNames(i), oref ' 设置发布元素
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
    Set RLtarget = refprd.UserRefProperties.Item(attNames(4))
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
    Set RLtarget = refprd.UserRefProperties.Item(attNames(5))
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
