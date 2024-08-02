LOCAL lcFile, lcOutPath, lnWb
PUBLIC goExcel   && to keep it from being destroyed and closing the cursors
lcFile = GETFILE("xlsx", "Workbook", "Load", 0, "Select Workbook to load into Class")
IF !EMPTY(lcFile)
	lcOutPath = ADDBS(JUSTPATH(lcFile))
	goExcel = NEWOBJECT("VFPxWorkbookXLSX", "..\VFPxWorkbookXLSX.vcx")
	lnWb = goExcel.OpenXlsxWorkbook(lcFile)

	IF lnWb > 0




		goExcel.SaveWorkbookAs(lnWb, lcOutPath + JUSTSTEM(lcFile) + "Copy.xlsx")
	ENDIF
ENDIF
