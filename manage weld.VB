Option Explicit '初始化对象
Dim CATIA As Object
Dim xlApp As Object
Dim oDoc
Sub intersel()
	Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
	Set xlApp = GetObject(, "Excel.Application") '获取excel程序
	Set oDoc = CATIA.ActiveDocument
	Dim iPrd, rootPrd, oPrd
	Dim oSel As Object
	Dim Status,i, iType(0), oPn, iPn
	on err resume next
	set iPrd=rootPrd.Products.Item("点焊信息")
	if err.number>0 then 
	
	
	
	Set oSel = oDoc.Selection
	oSel.Clear
	iType(0) = "Product"
	Status = oSel.SelectElement3(iType, "选择被连接产品", True, 2, False)
    If Status = "Cancel" Then
        MsgBox "操作取消"
        Exit Sub
    End If	
	
    If Status = "Normal" And oSel.Count2 <= 3 Then
        For i = 1 To oSel.Count
            oPn = oPn & "_" & oSel.Item(i).LeafProduct.PartNumber
        Next
         iPn = "SotWeld_" & oPn
         MsgBox iPn
		 oSel.Clear
    End If
	
	dim wpn
	wpn="12000"
	Set oPrd = iPrd.AddNewComponent("Product", wpn)		
    Set oPrd = iPrd.Products.AddNewComponent("Part", "ipn")

	'oPrd.PartNumber=iPn
	set oprd=nothing

	
	
	
	

MsgBox oName
End Sub

dim wpn
wpn="12000"
Set oPrd = products1.AddNewComponent("Product", wpn)  '插入部件
Set oprd = products1.AddNewProduct("") ‘插入产品









