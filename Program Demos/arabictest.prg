LOCAL loExcel, lnWb, lnSh, lcText
lcText = "تسوية عمولة المذكور حتي 31 12 2014م ؛ سددت نقدا من الصندوق"

loExcel = NEWOBJECT("VFPxWorkbookXLSX", "VFPxWorkbookXLSX.vcx", "", 1256)
lnWb = loExcel.CreateWorkbook("ArabicTest.xlsx")
IF lnWb > 0
 	lnSh = loExcel.AddSheet(lnWb, "Arabic Test")
	loExcel.SetCellValue(lnWb, lnSh, 1, 1, lcText)
	loExcel.SaveWorkbook(lnWb)
ENDIF

