LOCAL loExcel, lcFile, lnWb
lcFile = GETFILE("xlsx", "Workbook", "Load", 0, "Select Workbook to load into Class")
IF !EMPTY(lcFile)
	loExcel = NEWOBJECT("VFPxWorkbookXLSX", "VFPxWorkbookXLSX.vcx")
	loExcel.Debug = .T.

	lnWb = loExcel.OpenXlsxWorkbook(lcFile)

	SET STEP ON
	lcFormula = loExcel.GetCellFormula(lnWB, 1, 2, 3)
	lcFormula = loExcel.GetCellFormula(lnWB, 1, 4, 3)
	lcFormula = loExcel.GetCellFormula(lnWB, 1, 6, 3)
ENDIF
