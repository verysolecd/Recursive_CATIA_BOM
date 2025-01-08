'Private Sub iniP_Click()
'myProp (0)
'End Sub
'Private Sub readp_Click()
' myProp (1)
'End Sub
'Private Sub cchangeP_Click()
'myProp (2)
'End Sub
'Private Sub rBOM_Click()
' myProp (4)
'End Sub
'Private Sub CommandButton1_Click()
'myProp (3)
'End Sub
Option Explicit 
Public CATIA as object, xlApp as object, prd2wt as object
Sub myProp(Rw)
    Dim oDoc, rootPrd, children, oprd,
    dim xlsht,oRowNb, LV, rng
    Dim i, 
    Dim oArry(), RC(0 To 1)
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
    Set oDoc = CATIA.Activedocument
    Set rootPrd = CATIA.Activedocument.product
         If Err.Number <> 0 Then
            MsgBox "请打开CATIA并打开你的产品，再运行本程序": Err.Clear
            Exit Sub
         End If
     CATIA.Visible = True
    On Error GoTo 0
    Set xlApp = GetObject(, "Excel.Application") '获取excel程序
    xlApp.Visible = True
    Set xlsht = xlApp.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"
    Set xlsht = xlApp.ActiveSheet: xlsht.Columns(1).NumberFormatLocal = "0.00"
    RC(0) = 3: RC(1) = 3
    Select Case Rw
    Case 0
    Dim bDict: Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
    Call iniPrd(rootPrd, bDict)   '遍历写入所有要读取的属性，mass/materia
    Case 1 '====RW=1是读取产品属性到excel
       Call whois2rv '弹窗说明读取对象
       Dim Prd2Read: Set Prd2Read = prd2wt
        Set rng = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 15)): rng.ClearContents
            oArry = getPropArry(Prd2Read)
            oRowNb = RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
        Set children = Prd2Read.Products
            For i = 1 To children.Count
             oArry = getPropArry(children.Item(i))
             oRowNb = i + RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
            Next
    Case 2 '====RW=2是写入产品属性到Catia===
        If TypeOf prd2wt Is product Then
        MsgBox "你选择的产品是" & prd2wt.PartNumber & "是否继续"
        Dim Prd2Rv: Set Prd2Rv = prd2wt
            oRowNb = RC(0): oArry() = sht2Arry(xlsht, oRowNb)
            Call changeProp(Prd2Rv, oArry())
        Set children = Prd2Rv.Products '获取catia活动文档产品集
            For i = 1 To children.Count  '遍历产品,在excel获取产品属性，并修改
                oRowNb = i + RC(0)
                oArry() = sht2Arry(xlsht, oRowNb)
                Call changeProp(children.Item(i), oArry())
            Next
            Set prd2wt = Nothing
        End If
    Case 3 '====RW=3是计算所有分总称和总成重量===
        Call Assmass(rootPrd)
    Case 4  '====RW=4是，计算分总成质量并遍历bOM
        Set xlsht = getsht(xlApp)
        'Call Assmass(rootPrd)
        oRowNb = RC(0): LV = 1: Call recurPrd3(rootPrd, xlsht, oRowNb, LV)
        Call LvMg(xlsht)
    
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
Sub aMsgBox(msgx)
    Dim Msgwd, Title, lg, wd
    Set Msgwd = CreateObject("wscript.shell")
    wd = "0.5s后自动关闭对话框"
    lg = 0.514
    Msgwd.popup msgx, lg, wd
    Set Msgwd = Nothing
End Sub
