Function AttModify(oPrd, oArry) '修改产品属性 数组到 to catia
    Dim refprd: Set refprd = oPrd.ReferenceProduct
    Dim oPrt, colls, i, ind
    Dim pArry():   pArry() = infoPrd(oPrd)
    ind = 1
        For i = 1 To 4
        If oArry(ind)(i) <> "" And oArry(ind)(i) <> pArry(ind)(i) Then
            Select Case i
                Case 1: refprd.PartNumber = oArry(1)(i)
                Case 2: refprd.Nomenclature = oArry(1)(i)
                Case 3: refprd.definition = oArry(1)(i)
                Case 4: oPrd.Name = oArry(1)(i)
            End Select
        End If
        Next
    ind = 3
     i = 2
     If oArry(ind)(i) <> "" And pArry(ind)(i) <> oArry(ind)(i) Then
     Set colls = refprd.UserRefProperties
     colls.Item(att_usp_Names(i)).Value = oArry(ind)(i)
     Set colls = Nothing
     End If
     i = 3
     On Error Resume Next
          Set oPrt = refprd.Parent.Part
          Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
          If oArry(ind)(i) <> "" And oArry(ind)(i) <> pArry(ind)(i) Then
                  colls.Item(att_usp_Names(i)).Value = oArry(ind)(i)
          End If
     If Err.Number <> 0 Then
              Err.Clear
          End If
     On Error GoTo 0
End Function
Sub iniPrd(oPrd, oDict)
'Dim product
    If allPN.exists(oPrd.PartNumber) = False Then
        allPN(oPrd.PartNumber) = 1
        Call initMyRefPrd(oPrd)
    End If
    If oPrd.Products.Count > 0 Then
        For Each product In oPrd.Products
            Call iniPrd(product, oDict)
        Next product
    End If
    allPN.RemoveAll
End Sub
Sub initMyRefPrd(oPrd)
    Set refprd = oPrd.ReferenceProduct
    Dim colls
    Dim parameterObj(5), attType, i
    '对于PRD创建参数，"iMaterial" "iMass" "iThickness"
    Set attType = CreateObject("Scripting.Dictionary")
    attType(att_usp_Names(0)) = "String"
    attType(att_usp_Names(1)) = "list"
    attType(att_usp_Names(2)) = "String"
    attType(att_usp_Names(3)) = "Density"
    attType(att_usp_Names(4)) = "Mass"
    attType(att_usp_Names(5)) = "Length"
    Set colls = refprd.UserRefProperties
    For i = 2 To 5
        If hasAtt(colls, att_usp_Names(i)) = False Then
            Select Case i
                Case 2
                   Set parameterObj(i) = colls.CreateString(att_usp_Names(i), "TBD")
                Case 4 To 5
                   Set parameterObj(i) = colls.CreateDimension(att_usp_Names(i), attType(att_usp_Names(i)), 0#)
            End Select
         Else
             Set parameterObj(i) = colls.Item(att_usp_Names(i))
        End If
    Next i
    Set colls = Nothing

    Dim oPrt
    On Error Resume Next
    Set oPrt = refprd.Parent.Part
        If Err.Number = 0 Then
           Call iniPrt(oPrd, att_usp_Names())
        Else
           Err.Clear
           Dim oref, pubid, oPub
            Set colls = refprd.Publications
            i = 4
            If hasAtt(colls, att_usp_Names(i)) = False Then
                Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                Set oPub = colls.Add(att_usp_Names(i)) ' 添加发布
                colls.SetDirect att_usp_Names(i), oref ' 设置发布元素
                Set colls = Nothing
            End If
         End If
    On Error GoTo 0
    oPrd.Update
    '释放对象
    Set colls = Nothing
End Sub
Sub iniPrt(oPrd, att_usp_Names())
    Dim refprd, oPrt
    Set refprd = oPrd.ReferenceProduct
    Set oPrt = refprd.Parent.Part
    Dim MBD: Set MBD = oPrt.MainBody
    '==============创建cm参数集合
    On Error Resume Next
                Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
        If Err.Number <> 0 Then
                Err.Clear
                Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.CreateSet("cm")
                Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
        End If
    On Error GoTo 0
    Dim attType As Object
    Set attType = CreateObject("Scripting.Dictionary")
    attType(att_usp_Names(0)) = "String"
    attType(att_usp_Names(1)) = "list"
    attType(att_usp_Names(2)) = "String"
    attType(att_usp_Names(3)) = "Density"
    attType(att_usp_Names(4)) = "Mass"
    attType(att_usp_Names(5)) = "Length"
'==============创建PRT参数
    Dim parameterObj(5)
    For i = 1 To 5
        If hasAtt(colls, att_usp_Names(i)) = False Then
            Select Case i
                Case 1
                    Set parameterObj(i) = colls.CreateList(att_usp_Names(i)) 'list参数创建  iBodys
                Case 2
                   Set parameterObj(i) = colls.CreateString(att_usp_Names(i), "TBD")  '材料参数创建 iMaterial
                Case 3 To 5
                   Set parameterObj(i) = colls.CreateDimension(att_usp_Names(i), attType(att_usp_Names(i)), 0#) 'dimension参数创建
            End Select
        Else
            Set parameterObj(i) = colls.Item(att_usp_Names(i))
        End If
    Next i
 '=============对于list增加mainbody
        Dim lst, lstobj
        Set lst = parameterObj(1)
        If hasAtt(lst.ValueList, MBD.Name) = False Then
            lst.ValueList.Add MBD
        End If
    Set colls = Nothing
    oPrt.Update:    oPrd.Update
'===============创建发布
    i = 1
    For i = 3 To 5
        Set pubs = refprd.Publications
        If hasAtt(pubs, att_usp_Names(i)) = False Then
            Select Case i
                Case 3, 4, 5
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(att_usp_Names(i)) ' 添加发布
                    pubs.SetDirect att_usp_Names(i), oref ' 设置发布元素
                Case 2
                    Set parameterObj(i) = refprd.UserRefProperties.Item(att_usp_Names(i))
                    Set oref = refprd.CreateReferenceFromName(parameterObj(i).Name)
                    Set oPub = pubs.Add(att_usp_Names(i)) ' 添加发布
                    pubs.SetDirect att_usp_Names(i), oref ' 设置发布元素
            End Select
        End If
    Next
    oPrt.Update: oPrd.Update
' === 创建 EKL 规则 ===
    Dim strEKL(2)
    strEKL(0) = "CalM"
    strEKL(1) = "let lst(list) set lst=cm\iBodys let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1} cm\iMass=Vol*cm\iDensity"
    Set colls = oPrt.Relations
    If hasAtt(colls, strEKL(0)) = False Then
        Set oRule = colls.CreateProgram(strEKL(0), "cal of mass", strEKL(1))
    Else
        Set oRule = colls.Item(strEKL(0))
        If oRule.Value <> strEKL(1) Then
            oRule.Modify strEKL(1)
        End If
    End If
'==============创建关系
    Dim oFormula, RLname, RLnote, RLtarget, RLsource
    Set colls = oPrt.Relations
'----重量计算
    RLname = "CMAS"
    Set RLtarget = refprd.UserRefProperties.Item(att_usp_Names(4))
    RLsource = "cm\iMass"
        If hasAtt(colls, RLname) = False Then
            Set oFormula = colls.CreateFormula(RLname, " ", RLtarget, RLsource)
        Else
            Set oFormula = colls.Item(RLname)
                If oFormula.Value <> RLsource Then
                    oFormula.Modify RLsource
                End If
        End If
        Debug.Print oFormula.Value
    RLname = "CTK"
    Set RLtarget = refprd.UserRefProperties.Item(att_usp_Names(5))
    RLsource = "cm\iThickness"
        If hasAtt(colls, RLname) = False Then
            Set oFormula = colls.CreateFormula(RLname, " ", RLtarget, RLsource)
        Else
            Set oFormula = colls.Item(RLname)
                If oFormula.Value <> RLsource Then
                    oFormula.Modify RLsource
                End If
        End If
End Sub

Function whois2rv()
   '获取要读取或修改的  全局变量 Prd2wt 定义
    Dim infobox: infobox = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
        Select Case infobox
        Case 6  '===选择“是”,进行产品选择====
           On Error Resume Next
           Set whois2rv = selprd()
           If Err.Number <> 0 Then
           Err.Clear
           Exit Function
           End If
        Case 7: Set whois2rv = CATIA.Activedocument.product '===选择“否”====
        Case 2: Exit Function '===选择“取消”====
        End Select
        If Not whois2rv Is Nothing Then
            xlApp.Windows(1).WindowState = xlMaximized
           xlApp.Visible = True
           infobox = "你选择的产品是" & whois2rv.PartNumber & "是否继续"
          MsgBox infobox
        End If
           
End Function
Function selprd()
   Dim oSel, status, iType(0)
            MsgBox "请选择要读取的产品"
            xlApp.Windows(1).WindowState = xlMinimized
            CATIA.ActiveWindow.WindowState = 0
            Set oSel = CATIA.Activedocument.Selection
            oSel.Clear
            iType(0) = "Product"
            If oSel.Count2 = 0 Then
                status = oSel.SelectElement2(iType, "请选择要读取的产品", False)
                'status = oSel.SelectElement3(iType, "请选择要读取的产品", True, 2, False)
            End If
            If status = "Cancel" Then
                xlApp.Windows(1).WindowState = xlMaximized
                Exit Function
            End If
            If status = "Normal" And oSel.Count2 = 1 Then
                    Set selprd = oSel.Item(1).LeafProduct.ReferenceProduct
                    oSel.Clear
            Else
                MsgBox "请只选择一个产品"
                Exit Function
            End If
End Function

Sub noPrd()
Set prd2wt = Nothing
MsgBox "已释放待修改的产品"
End Sub

