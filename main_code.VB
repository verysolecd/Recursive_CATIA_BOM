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
'Private Sub CommandButton2_Click()
'myProp (5)
'End Sub
Option Explicit '初始化对象
Dim CATIA As Object
Dim xlApp As Object
Sub myProp(Rw)
    Dim oprd, xlsht, rootPrd, children, oDoc
    Dim i, oRowNb, LV, RNG
    Dim oArry(), RC(0 To 1)
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
    Set oDoc = CATIA.Activedocument
    Set rootPrd = CATIA.Activedocument.Product
         If Err.Number <> 0 Then
            MsgBox "请打开CATIA并打开你的产品，再运行本程序": Err.Clear
            Exit Sub
         End If
     CATIA.Visible = True
    On Error GoTo 0
    Set xlApp = GetObject(, "Excel.Application") '获取excel程序
    xlApp.Visible = True
    Set xlsht = xlApp.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"
    RC(0) = 3: RC(1) = 3
    Select Case Rw
    Case 0
    Dim bDict: Set bDict = CreateObject("Scripting.Dictionary") '子产品字典2
    Call iniPrd(rootPrd, bDict)   '遍历写入所有要读取的属性，mass/materia
    Case 1 '====RW=1是读取产品属性到excel
        Dim Prd2Read: Set Prd2Read = rootPrd
        Set RNG = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 15)): RNG.ClearContents
            oArry = getPropArry(Prd2Read)
            oRowNb = RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
        Set children = Prd2Read.Products
            For i = 1 To children.Count
             oArry = getPropArry(children.Item(i))
             oRowNb = i + RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
            Next
    Case 2 '====RW=2是写入产品属性到Catia===
        Dim Prd2Rv: Set Prd2Rv = rootPrd
            oRowNb = RC(0): oArry() = sht2Arry(xlsht, oRowNb)
            Call changeProp(Prd2Rv, oArry())
        Set children = Prd2Rv.Products '获取catia活动文档产品集
            For i = 1 To children.Count  '遍历产品,在excel获取产品属性，并修改
                oRowNb = i + RC(0)
                oArry() = sht2Arry(xlsht, oRowNb)
                Call changeProp(children.Item(i), oArry())
            Next
    Case 3 '====RW=3是计算所有分总称和总成重量===
        Call Assmass(rootPrd)
    Case 4  '====RW=4是，计算分总成质量并遍历bOM
        Set xlsht = getsht(xlApp)
        Call Assmass(rootPrd)
        oRowNb = RC(0): LV = 1: Call recurPrd3(rootPrd, xlsht, oRowNb, LV)
        Call LvMg(xlsht)
    Case 5 '====RW=4 读取和修改属性
        Dim oSel, status, iType(0), selPrd
        Set oSel = oDoc.Selection
        oSel.Clear
        iType(0) = "Product"
        If oSel.Count2 = 0 Then
            MsgBox "请在CATIA结构树上选择要修改的产品.": xlApp.Windows(1).WindowState = xlMinimized
            status = oSel.SelectElement3(iType, "选择被连接产品", True, 2, False)
        End If
        If status = "Cancel" Then
            MsgBox "用户取消操作，程序将退出"
            Exit Sub
        End If
        If status = "Normal" And oSel.Count2 = 1 Then
                Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
                oSel.Clear
               xlApp.Windows(1).WindowState = xlMaximized
        Else
            MsgBox "请只选择一个产品"
            Exit Sub
        End If
         '===========读取参数
            'Dim Prd2Read
          Set Prd2Read = selPrd
            Set RNG = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 15)): RNG.ClearContents
            oArry = getPropArry(Prd2Read)
            oRowNb = RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
        Set children = Prd2Read.Products
            For i = 1 To children.Count
             oArry = getPropArry(children.Item(i))
             oRowNb = i + RC(0): Call Arry2sht(oArry, xlsht, oRowNb)
            Next
'        Windows(1).WindowState = xlNormal
'         MsgBox "当前活动工作簿窗口将最大化"
'        Windows(1).WindowState = xlMaximized
'        iType(0) = "Product"
'        status = oSel.SelectElement2(iType, "选择被连接产品", False)
    End Select
    '错误处理
    If Err.Number = 0 Then
            aMsgBox "已执行，0.5 后关闭对话框"
        Else
            aMsgBox " 程序执行异常，请排查错误后重新执行，0.5s后将退出"
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
