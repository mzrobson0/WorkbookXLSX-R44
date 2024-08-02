#INCLUDE "..\VFPxWorkbookXLSX.h"
LOCAL lcTable, lcExcel, lnWB, loExcel
lcExcel = SYS(5) + ADDBS(SYS(2003)) + "Northwind Customers.xlsx"
lcTable = ADDBS(SYS(2004)) + "Samples\Northwind\customers.dbf"
lcTable = LOCFILE(lcTable)
IF !ISNULL(lcTable) .AND. FILE(lcTable)
	loExcel = NEWOBJECT("VFPxWorkbookXLSX", "..\VFPxWorkbookXLSX.vcx")
*	loExcel.SaveTabletoWorkbook(lcTable, lcExcel, .T., .T.)
	loExcel.SaveTableToWorkbookEx(lcTable, lcExcel, .NULL., .T., JUSTSTEM(lcTable), 1, 1, TABLE_STYLE_LIGHT13)
	
*	lnWB = loExcel.OpenXLSXWorkbook(lcExcel)
*	lcSheet = 'My New Sheet'
*	loReturn = loExcel.SaveGridToWorkbook(loGrid, lnWB, .T, .T., lcSheet)
ENDIF
