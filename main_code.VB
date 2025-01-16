Option Explicit '初始化对象
Public CATIA As Object, xlApp As Object, prd2wt As Object
Public att_usp_Names
Public allPN As Dictionary
Sub InitializeArray()
    ' 重新定义数组的大小
    ReDim att_usp_Names(0 To 5)
    att_usp_Names(0) = "cm"
    att_usp_Names(1) = "iBodys"
    att_usp_Names(2) = "iMaterial"
    att_usp_Names(3) = "iDensity"
    att_usp_Names(4) = "iMass"
    att_usp_Names(5) = "iThickness"
End Sub
Sub myProp(Rw)
     Dim oDoc, rootPrd, oPrd
     Dim xlsht, startrow, startcol, currRow, LV, rng
     Dim children
     Dim propertyArry()
     Dim i
    Call InitializeArray
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
        On Error Resume Next
        Set prd2wt = whois2rv() '弹窗说明读取对象
        If Err.Number <> 0 Then
        Err.Clear
         aMsgBox "没有要读取的产品"
        Exit Sub
        End If
        On Error GoTo 0
        Dim Prd2Read: Set Prd2Read = prd2wt
        Set rng = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 14)): rng.ClearContents
            currRow = startrow
            Arry2sht infoPrd(Prd2Read), xlsht, currRow
        Set children = Prd2Read.Products
            For i = 1 To children.Count
             currRow = i + startrow
             Arry2sht infoPrd(children.Item(i)), xlsht, currRow
            Next
        Set Prd2Read = Nothing
    Case 2 '====RW=2是写入产品属性到Catia===
        'If TypeOf prd2wt Is product Then
        If prd2wt Is Nothing Then
            Call myProp(1)
            MsgBox "请填写修改信息并重新按修改按钮"
            Exit Sub
        Else
             MsgBox "你要修改的分总成是" & prd2wt.PartNumber & "是否继续"
                Dim Prd2Rv: Set Prd2Rv = prd2wt
                    currRow = startrow
                    AttModify Prd2Rv, sht2Arry(xlsht, currRow)
                Set children = Prd2Rv.Products
                    For i = 1 To children.Count  '遍历产品,在excel获取产品属性，并修改
                        currRow = i + startrow
                        AttModify children.Item(i), sht2Arry(xlsht, currRow)
                    Next
                    Set Prd2Rv = Nothing
        End If
               
    Case 3 '====RW=3是计算所有分总称和总成重量===
        Assmass rootPrd
    Case 4  '====RW=4是，计算分总成质量并遍历bOM
        Set xlsht = getsht(xlApp)
        'Call Assmass(rootPrd)
        currRow = startrow
        LV = 1
        recurPrd rootPrd, xlsht, currRow, LV
        LvMg xlsht
    End Select
    '错误处理
    If Err.Number = 0 Then
            aMsgBox "已执行，0.5s 后自动关闭对话框"
        Else
            aMsgBox " 程序执行异常，请排查错误后重新执行，0.5s后将退出" & Err.Description
            Exit Sub
        Exit Sub
    End If
End Sub

