  Sub recurPrd(oPrd, xlsht, oRowNb, LV)
        Dim bDict, oArry, i
            Arry2sht bomRowPrd(oPrd, LV), xlsht, oRowNb
            If oPrd.Products.Count > 0 Then '若产品1子级>0
                Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
                For i = 1 To oPrd.Products.Count
                  If bDict.exists(oPrd.Products.Item(i).PartNumber) = False Then
                        bDict(oPrd.Products.Item(i).PartNumber) = 1
                        oRowNb = oRowNb + 1
                        recurPrd oPrd.Products.Item(i), xlsht, oRowNb, LV + 1
                    End If
               Next
            End If
End Sub
Function Assmass(oPrd)
Dim i
    Dim total, children: total = 0
        Set children = oPrd.Products
        If oPrd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass (children.Item(i))
                total = children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value + total
            Next
                oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value 'oprd.Parameters.Item("iMass").Value
        End If
   Assmass = total
End Function
Function bomRowPrd(oPrd, LV)
    Dim oDict, QTy
     QTy = 1
    If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
        Set oDict = Dictbros(oPrd.Parent.Parent)
        QTy = oDict(oPrd.PartNumber)       '获取oprd数量
    End If
    bomRowPrd = Array(LV, attDefault(oPrd), QTy, attUsp(oPrd))
End Function
Function infoPrd(oPrd)
       Dim LV, QTy: LV = 0: QTy = 0
     infoPrd = Array(LV, attDefault(oPrd), QTy, attUsp(oPrd))
End Function

Function attDefault(oPrd)
With oPrd.ReferenceProduct
    Dim att_default(1 To 4)
    att_default(1) = .PartNumber
    att_default(2) = .Nomenclature
    att_default(3) = .definition
    att_default(4) = oPrd.Name
    attDefault = att_default()
End With
End Function
Function attUsp(oPrd)
Dim refprd: Set refprd = oPrd.ReferenceProduct
Dim i, colls
    Dim att_usp(2 To 5)
    For i = 2 To 5
        Select Case i
        Case 2, 4, 5
            Set colls = refprd.UserRefProperties
            att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
            Set colls = Nothing
        Case 3
            On Error Resume Next
                    Dim oPrt
                Set oPrt = refprd.Parent.Part
                Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                    att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
                    Set colls = Nothing
                    If Err.Number <> 0 Then
                         att_usp(i) = "N\A"
                         Err.Clear '非part情况，无cm情况，有cm但无参
                    End If
               On Error GoTo 0
        End Select
    Next
     attUsp = att_usp
End Function
Function Dictbros(oPrd)  '获取兄弟字典
    Dim oDict, i, pn: Set oDict = CreateObject("Scripting.Dictionary")
        For i = 1 To oPrd.Products.Count
            pn = oPrd.Products.Item(i).PartNumber
            If oDict.exists(pn) = True Then
                oDict(pn) = oDict(pn) + 1
            Else
                oDict(pn) = 1
            End If
        Next
    Set Dictbros = oDict
End Function
Function hasAtt(colls, parameterName)
    Dim propertyItem
    On Error Resume Next
        Set propertyItem = colls.Item(parameterName)
        If Err.Number <> 0 Then
            Err.Clear
            hasAtt = False
        Else
            hasAtt = True
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
