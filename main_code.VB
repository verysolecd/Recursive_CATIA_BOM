Option Explicit '初始化对象
Public CATIA As Object, xlApp As Object, prd2wt As Object
Sub myProp(Rw)
     Dim oDoc, rootPrd, oprd
     Dim xlsht, startrow, startcol, currRow, LV, rng,i
     Dim children
     Dim propertyArry()
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    Set oDoc = CATIA.Activedocument
    Set rootPrd = CATIA.Activedocument.product
         If Err.Number <> 0 Then
            MsgBox "请打开CATIA产品，再运行本程序": Err.Clear: Exit Sub
         End If
        CATIA.Visible = True
    On Error GoTo 0
    Set xlApp = GetObject(, "Excel.Application")
    Set xlsht = xlApp.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"
    Set xlsht = xlApp.ActiveSheet: xlsht.Columns(1).NumberFormatLocal = "0.00"
    Set xlsht = xlApp.ActiveSheet: xlsht.Rows(1).NumberFormatLocal = "0"
    xlApp.Visible = True   
  
   startrow = 3: startcol = 3

    Select Case Rw
    Case 0
        Dim bDict: Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
        Call iniPrd(rootPrd, bDict)

    Case 1 '====RW=1是读取产品属性到excel
        Set prd2wt = whois2rv() '弹窗说明读取对象
        Dim Prd2Read: Set Prd2Read = prd2wt
        Set rng = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 14)): rng.ClearContents
            currRow = startrow
            Arry2sht getPropArry(Prd2Read), xlsht, currRow
        Set children = Prd2Read.Products
            For i = 1 To children.Count
             currRow = i + startrow
             Arry2sht getPropArry(children.Item(i)), xlsht, currRow
            Next
        Set Prd2Read = Nothing

    Case 2 '====RW=2是写入产品属性到Catia===
        If TypeOf prd2wt Is product Then
        MsgBox "你选择的产品是" & prd2wt.PartNumber & "是否继续"
        Dim Prd2Rv: Set Prd2Rv = prd2wt
            currRow = startrow
            propertyArry() = sht2Arry(xlsht, currRow)
            changeProp Prd2Rv, propertyArry()

        Set children = Prd2Rv.Products '获取catia活动文档产品集
            For i = 1 To children.Count  '遍历产品,在excel获取产品属性，并修改
                currRow = i + startrow
                propertyArry() = sht2Arry(xlsht, currRow)
                changeProp children.Item(i), propertyArry()
            Next
            Set prd2wt = Nothing

        Else
            MsgBox "请按“获取属性”以确定要读取和修改的产品"
            Exit Sub
        End If
    Case 3 '====RW=3是计算所有分总称和总成重量===
        Assmass rootPrd
    Case 4  '====RW=4是，计算分总成质量并遍历bOM
        Set xlsht = getsht(xlApp)
        'Call Assmass(rootPrd)
        currRow = startrow
        LV = 1
        recurPrd3 rootPrd, xlsht, currRow, LV
        LvMg xlsht
    End Select
    '错误处理
    If Err.Number = 0 Then
            aMsgBox "已执行，0.5 后关闭对话框"
        Else
            aMsgBox " 程序执行异常，请排查错误后重新执行，0.5s后将退出" & Err.Description
            Exit Sub
        Exit Sub
    End If
End Sub

